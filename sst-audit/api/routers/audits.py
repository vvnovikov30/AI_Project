"""Роутер сессий аудита: сборка SSOT из загруженных в память файлов.

CLI не вызывается: транспортный слой воспроизводит этапы ingest → matching
на публичных функциях ``parsers``/``matcher``/``db``/``review_service``.
Промежуточные XML/JSON-дампы не создаются — единственный источник истины
SQLite. Парсинг выполняется в рабочем потоке (``run_in_threadpool``), где
открывается собственное соединение: ``sqlite3`` не переносится между потоками.
"""

from __future__ import annotations

from collections.abc import Sequence
from decimal import Decimal
from io import BytesIO
from pathlib import Path
import sqlite3
from typing import Any, Final
import uuid

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool

import db
import review_service
from api.deps import get_db, get_db_path, get_max_upload_bytes, require_pdn_consent
from api.schemas import SessionCreateResponse
from matcher import CascadeMatcher, CompositeMatcher
from models import NormalizedItem
from parsers import LSRParser, PDFLSRParser, get_spec_parser

router = APIRouter(tags=["audits"])

# Шкала уверенности ядра — проценты 0..100 (синхронно с main/review_service).
MIN_MATCH_THRESHOLD: Final[Decimal] = Decimal("40.0")
MATCH_METHOD_CASCADE: Final[str] = "FUZZY_CASCADE"
MATCH_METHOD_NO_MATCH: Final[str] = "NO_MATCH"
MATCH_METHOD_REVERSE: Final[str] = "REVERSE_PASS"


def _safe_filename(filename: str | None, fallback: str) -> str:
    """Оставляет только базовое имя загруженного файла (без каталогов)."""
    name = Path(filename or "").name.strip()
    return name or fallback


def _parse_rd(payload: bytes, display_name: str) -> list[NormalizedItem]:
    """Разбирает РД/ВОР из потока в памяти без временных файлов."""
    parser = get_spec_parser(BytesIO(payload), display_name)
    return list(parser.extract_items())


def _parse_smeta(payload: bytes, display_name: str) -> list[NormalizedItem]:
    """Разбирает смету: PDF-форма 4б либо XLSX-ЛСР (первый лист)."""
    source = BytesIO(payload)
    if Path(display_name).suffix.lower() == ".pdf":
        return list(PDFLSRParser(source, display_name).parse())
    return list(LSRParser(source, display_name).parse())


def _candidates(
    smeta_item: NormalizedItem,
    rd_items: Sequence[NormalizedItem],
) -> list[tuple[NormalizedItem, Decimal]]:
    """Оценивает пары «позиция сметы → позиции РД» в шкале 0..100."""
    scored: list[tuple[NormalizedItem, Decimal]] = []
    for rd_item in rd_items:
        raw_score = CascadeMatcher.calculate_score(smeta_item, rd_item)
        score = Decimal(str(raw_score))
        if score > Decimal("0"):
            scored.append((rd_item, score))
    scored.sort(key=lambda pair: pair[1], reverse=True)
    return scored


def _match_row(
    rd_item: NormalizedItem | None,
    smeta_item: NormalizedItem | None,
    confidence: Decimal,
    method: str,
    status: str,
    rd_id_by_object: dict[int, int],
    smeta_id_by_object: dict[int, int],
) -> dict[str, Any]:
    """Готовит запись ``match_records`` с идентификаторами из SSOT."""
    return {
        "rd_item_id": (
            rd_id_by_object.get(id(rd_item)) if rd_item is not None else None
        ),
        "smeta_item_id": (
            smeta_id_by_object.get(id(smeta_item)) if smeta_item is not None else None
        ),
        "confidence_score": confidence,
        "conversion_ratio": db.DEFAULT_CONVERSION_RATIO,
        "match_method": method,
        "status": status,
    }


def _build_matches(
    rd_items: Sequence[NormalizedItem],
    smeta_items: Sequence[NormalizedItem],
    rd_ids: Sequence[int],
    smeta_ids: Sequence[int],
) -> list[dict[str, Any]]:
    """Сопоставляет смету с РД по контракту ``main.match_items``.

    Классификация: ниже 40% — ``AUTO``/``NO_MATCH``; иначе ``resolve_match``
    решает ``AUTO`` или ``HITL_PENDING`` (порог 70%, отрыв 5%). Обратный
    проход создает карточку только для РД без единого сметного кандидата.
    """
    rd_id_by_object = {id(item): item_id for item, item_id in zip(rd_items, rd_ids)}
    smeta_id_by_object = {
        id(item): item_id for item, item_id in zip(smeta_items, smeta_ids)
    }
    rows: list[dict[str, Any]] = []

    if not rd_items:
        for smeta_item in smeta_items:
            rows.append(
                _match_row(
                    None,
                    smeta_item,
                    Decimal("0"),
                    MATCH_METHOD_NO_MATCH,
                    "AUTO",
                    rd_id_by_object,
                    smeta_id_by_object,
                )
            )
        return rows

    covered_rd: set[int] = set()
    for smeta_item in smeta_items:
        candidates = _candidates(smeta_item, rd_items)
        best_rd: NormalizedItem | None = None
        best_score = Decimal("0")
        if candidates:
            best_rd, best_score = candidates[0]
            # Слабое упоминание уже считается покрытием: обратный проход
            # не должен дублировать шум карточкой «РД отсутствует в смете».
            covered_rd.add(id(best_rd))

        if best_rd is None or best_score < MIN_MATCH_THRESHOLD:
            rows.append(
                _match_row(
                    None,
                    smeta_item,
                    Decimal("0"),
                    MATCH_METHOD_NO_MATCH,
                    "AUTO",
                    rd_id_by_object,
                    smeta_id_by_object,
                )
            )
            continue

        selection = CompositeMatcher.resolve_match(
            smeta_item,
            [
                (item, CompositeMatcher._decimal_score(score))
                for item, score in candidates
            ],
        )
        rows.append(
            _match_row(
                selection.rd_item,
                smeta_item,
                best_score,
                MATCH_METHOD_CASCADE,
                selection.status.value,
                rd_id_by_object,
                smeta_id_by_object,
            )
        )

    for rd_item in rd_items:
        if id(rd_item) in covered_rd:
            continue
        rows.append(
            _match_row(
                rd_item,
                None,
                Decimal("0"),
                MATCH_METHOD_REVERSE,
                "HITL_PENDING",
                rd_id_by_object,
                smeta_id_by_object,
            )
        )
    return rows


def _run_session(
    db_path: str,
    user_id: str,
    rd_payload: bytes,
    rd_name: str,
    smeta_payload: bytes,
    smeta_name: str,
) -> SessionCreateResponse:
    """Выполняет ingest → matching в рабочем потоке на своем соединении."""
    rd_items = _parse_rd(rd_payload, rd_name)
    smeta_items = _parse_smeta(smeta_payload, smeta_name)
    session_id = uuid.uuid4().hex
    project_name = Path(rd_name).stem or "API Upload"

    conn = db.get_connection(db_path)
    try:
        db.init_db(conn)
        with db.transaction(conn):
            db.create_session(
                conn,
                session_id,
                project_name,
                rd_name,
                smeta_name,
                user_id=user_id,
            )
            rd_saved = db.save_rd_items(
                conn,
                session_id,
                [item.to_rd_record() for item in rd_items],
            )
            smeta_saved = db.save_smeta_items(
                conn,
                session_id,
                [item.to_smeta_record() for item in smeta_items],
            )

        rd_ids = list(db.get_rd_item_payloads(conn, session_id))
        if len(rd_ids) != rd_saved:
            raise RuntimeError(
                f"Ожидалось {rd_saved} позиций РД, сохранено {len(rd_ids)}"
            )
        smeta_ids = list(db.get_smeta_items_map(conn, session_id))
        if len(smeta_ids) != smeta_saved:
            raise RuntimeError(
                f"Ожидалось {smeta_saved} позиций сметы, сохранено {len(smeta_ids)}"
            )

        matches = _build_matches(rd_items, smeta_items, rd_ids, smeta_ids)
        db.save_matches(conn, session_id, matches)

        pending = review_service.pending_count(conn, session_id)
        status = "MATCHED" if pending else "REVIEWED"
        db.update_session_status(conn, session_id, status)
    finally:
        conn.close()

    return SessionCreateResponse(
        session_id=session_id,
        status=status,
        pending_count=pending,
        rd_file=rd_name,
        lsr_file=smeta_name,
    )


@router.post("/sessions", response_model=SessionCreateResponse)
async def create_session(
    user_id: str = Form(...),
    rd_file: UploadFile = File(...),
    lsr_file: UploadFile = File(...),
    conn: sqlite3.Connection = Depends(get_db),
) -> SessionCreateResponse:
    """Создает сессию аудита из двух файлов, загруженных в память.

    Порядок проверок: активное согласие 152-ФЗ (403) → лимит размера
    каждого файла (413) → парсинг и сопоставление в рабочем потоке.
    Файлы на диск не сохраняются, SSOT пополняется атомарно.
    """
    require_pdn_consent(conn, user_id)

    limit = get_max_upload_bytes()
    rd_payload = await rd_file.read()
    smeta_payload = await lsr_file.read()
    for payload, original_name in (
        (rd_payload, rd_file.filename),
        (smeta_payload, lsr_file.filename),
    ):
        if len(payload) > limit:
            raise HTTPException(
                status_code=413,
                detail=(
                    f"Файл '{original_name or 'upload'}' превышает лимит "
                    f"{limit // (1024 * 1024)} МБ"
                ),
            )

    rd_name = _safe_filename(rd_file.filename, "rd.pdf")
    smeta_name = _safe_filename(lsr_file.filename, "lsr.xlsx")

    try:
        return await run_in_threadpool(
            _run_session,
            get_db_path(),
            user_id,
            rd_payload,
            rd_name,
            smeta_payload,
            smeta_name,
        )
    except (ValueError, KeyError) as exc:
        raise HTTPException(
            status_code=400,
            detail=f"Не удалось обработать загруженные файлы: {exc}",
        ) from exc

