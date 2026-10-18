"""Оркестратор sst-audit: SQLite как единственный источник истины (SSOT).

Единый конвейер:
1. INGEST   — парсинг РД (PDF) и сметы (XLSX), нормализация и атомарная
              запись позиций в ``rd_items`` / ``smeta_items``.
2. MATCHING — каскадное сопоставление, запись связей в ``match_records``.
              Позиции с максимальным сходством ниже ``MIN_MATCH_THRESHOLD``
              не отвлекают сметчика: они фиксируются как ``AUTO``/``NO_MATCH``.
3. HITL     — при наличии очереди ``HITL_PENDING`` автоматически запускается
              терминальное ревью из ``cli_review.py`` в текущем процессе.
4. RULES    — оценка количественных расхождений в ``rule_evaluations``.
5. EXPORT   — итоговая ведомость формируется из SSOT через ``exporter.py``.

XML/JSON-дампы не формируются: единственный источник истины — SQLite.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
import sqlite3
import sys
from typing import Any, Final, Optional, Sequence
import uuid

from rich.console import Console
from rich.table import Table

import cli_review
import db
import dispatcher
import exporter
from evaluator import DiscrepancyDetector
from matcher import CascadeMatcher
from models import CollisionSeverity, NormalizedItem
from parsers import LSRParser, PDFLSRParser, PDFSpecParser

# Жесткий порог отсечки семантического шума (шкала матчера 0-100).
# Ниже него позиция сметы считается отсутствующей в РД и не попадает в HITL.
MIN_MATCH_THRESHOLD: Final[Decimal] = Decimal("40.0")
# Порог автоматического принятия связи (шкала матчера 0-100).
AUTO_MATCH_THRESHOLD: Final[Decimal] = Decimal("70.0")
# Минимальный отрыв лидера от второго кандидата; иначе связь неоднозначна.
AMBIGUITY_MARGIN: Final[Decimal] = Decimal("5.0")
DEFAULT_CONVERSION_RATIO: Final[Decimal] = db.DEFAULT_CONVERSION_RATIO

MATCH_METHOD_CASCADE: Final[str] = "FUZZY_CASCADE"
MATCH_METHOD_NO_CANDIDATE: Final[str] = "NO_CANDIDATE"
MATCH_METHOD_NO_MATCH: Final[str] = "NO_MATCH"
MATCH_METHOD_REVERSE: Final[str] = "REVERSE_PASS"
BATCH_ERROR_LOG: Final[str] = "batch_errors.log"

RULE_BY_SEVERITY: Final[dict[CollisionSeverity, str]] = {
    CollisionSeverity.CRITICAL: "VOLUME_DELTA_CRITICAL",
    CollisionSeverity.SUSPECT: "VOLUME_DELTA_SUSPECT",
    CollisionSeverity.INCOMPATIBLE: "UNIT_INCOMPATIBLE",
    CollisionSeverity.INFO: "VOLUME_WITHIN_TOLERANCE",
}


@dataclass(frozen=True)
class SessionConfig:
    """Конфигурация сессии аудита."""

    project_name: str
    rd_path: str
    smeta_path: str
    db_path: str = db.DEFAULT_DB_PATH
    session_id: Optional[str] = None
    smeta_sheet: Optional[str] = None

    def resolved_session_id(self) -> str:
        """Возвращает заданный идентификатор либо генерирует новый."""
        return self.session_id or uuid.uuid4().hex


@dataclass(frozen=True)
class MatchDecision:
    """Решение матчера по одной позиции сметы либо РД."""

    smeta_item: Optional[NormalizedItem]
    rd_item: Optional[NormalizedItem]
    confidence: Decimal
    status: str
    method: str
    candidates: tuple[tuple[NormalizedItem, float], ...] = ()


@dataclass(frozen=True)
class PipelineResult:
    """Сводка выполненного прогона."""

    session_id: str
    session_status: str
    rd_saved: int
    smeta_saved: int
    matches_saved: int
    evaluations_saved: int
    auto_matches: int
    hitl_pending: int
    no_match: int = 0


def _candidates(
    smeta_item: NormalizedItem,
    rd_items: Sequence[NormalizedItem],
) -> list[tuple[NormalizedItem, float]]:
    """Возвращает кандидатов РД с ненулевым сходством по убыванию score."""
    if not rd_items:
        # Пустой/нераспознанный РД не должен приводить к IndexError.
        return []
    scored = [
        (rd_item, CascadeMatcher.calculate_score(smeta_item, rd_item))
        for rd_item in rd_items
    ]
    return sorted(
        ((item, score) for item, score in scored if score > 0.0),
        key=lambda pair: pair[1],
        reverse=True,
    )


def match_items(
    rd_items: Sequence[NormalizedItem],
    smeta_items: Sequence[NormalizedItem],
) -> list[MatchDecision]:
    """Сопоставляет позиции сметы с РД и классифицирует связи.

    Классификация:

    * ``score < MIN_MATCH_THRESHOLD`` — позиция сметы отсутствует в РД:
      ``AUTO``/``NO_MATCH``, ``rd_item_id = NULL``, ``confidence = 0``;
    * ``MIN_MATCH_THRESHOLD <= score < AUTO_MATCH_THRESHOLD`` — пограничное
      сходство, требуется решение эксперта (``HITL_PENDING``);
    * ``score >= AUTO_MATCH_THRESHOLD`` с отрывом от второго кандидата не
      меньше ``AMBIGUITY_MARGIN`` — автоматическое принятие связи;
    * неоднозначность (отрыв < ``AMBIGUITY_MARGIN``) отправляется в HITL.

    Обратный проход добавляет ``HITL_PENDING`` только для позиций РД, у
    которых вообще не нашлось ни одного сметного кандидата. Слабые кандидаты
    ниже порога отсечки не создают повторную карточку по той же паре.
    """
    decisions: list[MatchDecision] = []
    covered_rd: set[int] = set()

    if not rd_items:
        # Защита от нераспознанного РД: все сметные позиции фиксируются
        # как AUTO/NO_MATCH, а правило SMETA_NOT_IN_RD формируется ниже.
        return [
            MatchDecision(
                smeta_item=smeta_item,
                rd_item=None,
                confidence=Decimal("0"),
                status="AUTO",
                method=MATCH_METHOD_NO_MATCH,
                candidates=(),
            )
            for smeta_item in smeta_items
        ]

    for smeta_item in smeta_items:
        candidates = _candidates(smeta_item, rd_items)
        best_item: Optional[NormalizedItem] = None
        best_score = Decimal("0")

        if candidates:
            best_item, raw_score = candidates[0]
            best_score = Decimal(str(raw_score))
            # Даже слабое лексическое пересечение считается упоминанием РД
            # в смете: иначе обратный проход продублирует шум карточкой
            # "позиция РД отсутствует в смете".
            covered_rd.add(id(best_item))

        if best_item is None or best_score < MIN_MATCH_THRESHOLD:
            decisions.append(
                MatchDecision(
                    smeta_item=smeta_item,
                    rd_item=None,
                    confidence=Decimal("0"),
                    status="AUTO",
                    method=MATCH_METHOD_NO_MATCH,
                    candidates=(),
                )
            )
            continue

        runner_up = (
            Decimal(str(candidates[1][1])) if len(candidates) > 1 else Decimal("0")
        )
        ambiguous = bool(candidates[1:]) and (
            best_score - runner_up
        ) < AMBIGUITY_MARGIN
        qualified_candidates = tuple(
            (item, score)
            for item, score in candidates
            if Decimal(str(score)) >= MIN_MATCH_THRESHOLD
        )[:3]

        is_auto = best_score >= AUTO_MATCH_THRESHOLD and not ambiguous
        decisions.append(
            MatchDecision(
                smeta_item=smeta_item,
                rd_item=best_item,
                confidence=best_score,
                status="AUTO" if is_auto else "HITL_PENDING",
                method=MATCH_METHOD_CASCADE,
                candidates=qualified_candidates,
            )
        )

    for rd_item in rd_items:
        if id(rd_item) in covered_rd:
            continue
        decisions.append(
            MatchDecision(
                smeta_item=None,
                rd_item=rd_item,
                confidence=Decimal("0"),
                status="HITL_PENDING",
                method=MATCH_METHOD_REVERSE,
            )
        )

    return decisions


def build_rule_evaluations(
    decisions: Sequence[MatchDecision],
    match_ids: Sequence[int],
) -> list[dict[str, Any]]:
    """Готовит записи ``rule_evaluations`` для сохранения в SSOT.

    Для связанных пар применяется ``DiscrepancyDetector``; для позиций без
    пары фиксируются правила ``RD_NOT_IN_SMETA`` / ``SMETA_NOT_IN_RD``.
    """
    if len(decisions) != len(match_ids):
        raise RuntimeError(
            f"Количество решений ({len(decisions)}) не совпадает с количеством "
            f"записей match_records ({len(match_ids)})"
        )

    # Родительские расценки, покрытые сопоставленными дочерними ресурсами:
    # для них правило SMETA_NOT_IN_RD не применяется.
    covered_parents: set[str] = {
        str(decision.smeta_item.parent_id)
        for decision in decisions
        if decision.rd_item is not None
        and decision.smeta_item is not None
        and decision.smeta_item.is_resource
        and decision.smeta_item.parent_id
    }

    evaluations: list[dict[str, Any]] = []
    for decision, match_id in zip(decisions, match_ids):
        if decision.smeta_item is not None and decision.rd_item is not None:
            collision = DiscrepancyDetector.evaluate(
                decision.smeta_item,
                decision.rd_item,
                decision.confidence,
            )
            evaluations.append(
                {
                    "match_record_id": match_id,
                    "rule_name": RULE_BY_SEVERITY[collision.severity],
                    "is_violation": collision.severity is not CollisionSeverity.INFO,
                    "delta_volume": collision.volume_delta,
                    "delta_cost": None,
                    "message": collision.description,
                }
            )
        elif decision.rd_item is not None:
            evaluations.append(
                {
                    "match_record_id": match_id,
                    "rule_name": "RD_NOT_IN_SMETA",
                    "is_violation": True,
                    "delta_volume": decision.rd_item.volume,
                    "delta_cost": None,
                    "message": "Позиция РД не найдена в смете",
                }
            )
        elif decision.smeta_item is not None:
            if str(decision.smeta_item.raw_id) in covered_parents:
                evaluations.append(
                    {
                        "match_record_id": match_id,
                        "rule_name": "SMETA_PARENT_COVERED",
                        "is_violation": False,
                        "delta_volume": Decimal("0"),
                        "delta_cost": None,
                        "message": (
                            "Расценка покрыта сопоставленным дочерним "
                            "ресурсом узла сметы"
                        ),
                    }
                )
            else:
                evaluations.append(
                    {
                        "match_record_id": match_id,
                        "rule_name": "SMETA_NOT_IN_RD",
                        "is_violation": True,
                        "delta_volume": -decision.smeta_item.volume,
                        "delta_cost": None,
                        "message": "Позиция сметы не подтверждена РД",
                    }
                )
    return evaluations


def _load_item_ids(
    conn: sqlite3.Connection,
    table: str,
    session_id: str,
    expected: int,
) -> list[int]:
    """Возвращает id вставленных позиций в порядке вставки (id ASC)."""
    rows = conn.execute(
        f"SELECT id FROM {table} WHERE session_id = ? ORDER BY id ASC",
        (session_id,),
    ).fetchall()
    ids = [int(row["id"]) for row in rows]
    if len(ids) != expected:
        raise RuntimeError(
            f"Ожидалось {expected} записей в {table} для сессии {session_id}, "
            f"найдено {len(ids)}"
        )
    return ids


def _decisions_from_match_rows(
    match_rows: Sequence[sqlite3.Row],
    rd_by_id: dict[int, NormalizedItem],
    smeta_by_id: dict[int, NormalizedItem],
) -> list[MatchDecision]:
    """Пересобирает решения матчера из финального состояния SSOT.

    После терминального ревью выбранные кандидаты уже сохранены в
    ``match_records``. Оценка правил обязана выполняться по этим связям,
    а не по первоначальному предложению матчера.
    """
    decisions: list[MatchDecision] = []
    for row in match_rows:
        rd_id = row["rd_item_id"]
        smeta_id = row["smeta_item_id"]
        decisions.append(
            MatchDecision(
                smeta_item=smeta_by_id.get(int(smeta_id)) if smeta_id else None,
                rd_item=rd_by_id.get(int(rd_id)) if rd_id else None,
                confidence=Decimal(str(row["confidence_score"] or 0)),
                status=str(row["status"]),
                method=str(row["match_method"]),
            )
        )
    return decisions


def _parse_smeta_items(
    smeta_path: str, smeta_sheet: Optional[str]
) -> Sequence[NormalizedItem]:
    """Парсит сметный документ в зависимости от формата файла.

    Поддерживаются два формата: XLSX-ЛСР (``LSRParser``) и PDF-смета
    по форме 4б (``PDFLSRParser``).

    :param smeta_path: путь к файлу сметы.
    :param smeta_sheet: имя листа XLSX; для PDF не используется.
    :return: список нормализованных позиций сметы.
    :raises ValueError: если для PDF-сметы передан ``smeta_sheet``.
    """
    if Path(smeta_path).suffix.lower() == ".pdf":
        if smeta_sheet is not None:
            raise ValueError("Для PDF-сметы параметр smeta_sheet не применяется")
        return PDFLSRParser(smeta_path).parse()
    smeta_parser = LSRParser(smeta_path)
    if smeta_sheet is None:
        return smeta_parser.parse()
    return smeta_parser.parse(sheet_name=smeta_sheet)


def run_pipeline(
    conn: sqlite3.Connection,
    config: SessionConfig,
    rd_items: Optional[Sequence[NormalizedItem]] = None,
    smeta_items: Optional[Sequence[NormalizedItem]] = None,
    *,
    interactive: bool = False,
    reviewer: str = "Оператор ПТО",
) -> PipelineResult:
    """Выполняет полный цикл ingest -> matching -> rules -> status.

    :param conn: соединение SQLite; схема инициализируется идемпотентно.
    :param config: конфигурация сессии (проект, пути, db_path).
    :param rd_items: заранее подготовленные позиции РД (иначе парсятся из PDF).
    :param smeta_items: заранее подготовленные позиции сметы (иначе из XLSX).
    :param interactive: запускать терминальный HITL-опрос по очереди.
    :param reviewer: имя эксперта для записей ``expert_reviews``.
    """
    db.init_db(conn)
    session_id = config.resolved_session_id()

    if rd_items is None:
        rd_items = PDFSpecParser(config.rd_path).extract_items()
    if smeta_items is None:
        smeta_items = _parse_smeta_items(config.smeta_path, config.smeta_sheet)

    rd_items = list(rd_items)
    smeta_items = list(smeta_items)

    # --- ЭТАП INGEST: сессия и позиции сохраняются атомарно -----------------
    with db.transaction(conn):
        db.create_session(
            conn,
            session_id,
            config.project_name,
            config.rd_path,
            config.smeta_path,
            smeta_sheet=config.smeta_sheet,
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

    rd_ids = _load_item_ids(conn, "rd_items", session_id, rd_saved)
    smeta_ids = _load_item_ids(conn, "smeta_items", session_id, smeta_saved)
    rd_id_by_object = {id(item): item_id for item, item_id in zip(rd_items, rd_ids)}
    smeta_id_by_object = {
        id(item): item_id for item, item_id in zip(smeta_items, smeta_ids)
    }
    rd_by_id = {item_id: item for item, item_id in zip(rd_items, rd_ids)}
    smeta_by_id = {item_id: item for item, item_id in zip(smeta_items, smeta_ids)}

    # --- ЭТАП MATCHING ------------------------------------------------------
    decisions = match_items(rd_items, smeta_items)
    db.save_matches(
        conn,
        session_id,
        [
            {
                "rd_item_id": (
                    rd_id_by_object.get(id(decision.rd_item))
                    if decision.rd_item is not None
                    else None
                ),
                "smeta_item_id": (
                    smeta_id_by_object.get(id(decision.smeta_item))
                    if decision.smeta_item is not None
                    else None
                ),
                "confidence_score": decision.confidence,
                "conversion_ratio": DEFAULT_CONVERSION_RATIO,
                "match_method": decision.method,
                "status": decision.status,
            }
            for decision in decisions
        ],
    )

    match_rows = db.get_match_records(conn, session_id)
    if len(match_rows) != len(decisions):
        raise RuntimeError(
            f"Сохранено {len(match_rows)} связей, ожидалось {len(decisions)}"
        )

    # --- ЭТАП HITL: терминальное ревью в текущем процессе -------------------
    if interactive:
        pending_before_review = db.get_pending_reviews(conn, session_id)
        if pending_before_review:
            cli_review.run_review(conn, session_id, reviewer=reviewer)
            match_rows = db.get_match_records(conn, session_id)
            if len(match_rows) != len(decisions):
                raise RuntimeError(
                    f"После ревью сохранено {len(match_rows)} связей, "
                    f"ожидалось {len(decisions)}"
                )
            decisions = _decisions_from_match_rows(match_rows, rd_by_id, smeta_by_id)
        else:
            print("Все позиции сопоставлены автоматически (непоняток нет)")

    # --- ЭТАП RULES ---------------------------------------------------------
    match_ids = [int(row["id"]) for row in match_rows]
    evaluations = build_rule_evaluations(decisions, match_ids)
    evaluations_saved = db.save_rule_evaluations(conn, evaluations)

    # --- ЭТАП STATUS --------------------------------------------------------
    pending = db.get_pending_reviews(conn, session_id)
    session_status = "MATCHED" if pending else "REVIEWED"
    db.update_session_status(conn, session_id, session_status)

    return PipelineResult(
        session_id=session_id,
        session_status=session_status,
        rd_saved=rd_saved,
        smeta_saved=smeta_saved,
        matches_saved=len(match_rows),
        evaluations_saved=evaluations_saved,
        auto_matches=sum(
            1
            for decision in decisions
            if decision.status == "AUTO" and decision.rd_item is not None
        ),
        hitl_pending=len(pending),
        no_match=sum(
            1
            for decision in decisions
            if decision.status == "AUTO"
            and decision.rd_item is None
            and decision.smeta_item is not None
        ),
    )


@dataclass(frozen=True)
class BatchSectionResult:
    """Результат обработки одного раздела в ночном пакетном режиме."""

    pdf_path: Path
    sheet_name: str
    confidence: float
    session_id: str
    result: Optional[PipelineResult] = None
    error: Optional[str] = None

    @property
    def succeeded(self) -> bool:
        """Признак успешной обработки раздела."""
        return self.error is None and self.result is not None


@dataclass(frozen=True)
class BatchSummary:
    """Итоговая статистика пакетного прогона."""

    total: int
    succeeded: int
    failed: int
    hitl_pending: int
    sections: tuple[BatchSectionResult, ...] = ()
    output_path: Optional[Path] = None

    @property
    def successful_session_ids(self) -> list[str]:
        """Идентификаторы успешно завершенных сессий."""
        return [
            section.session_id for section in self.sections if section.succeeded
        ]


def _append_batch_error(
    log_path: Path,
    pdf_path: Path,
    sheet_name: str,
    error: BaseException,
) -> None:
    """Пишет ошибку раздела в stderr и накопительный ``batch_errors.log``."""
    message = f"{type(error).__name__}: {error}"
    print(
        f"[ERROR] {pdf_path.name} -> {sheet_name}: {message}",
        file=sys.stderr,
    )
    try:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8") as stream:
            stream.write(
                f"{datetime.now().isoformat(timespec='seconds')}\t"
                f"{pdf_path}\t{sheet_name}\t{message}\n"
            )
    except OSError as log_error:
        print(
            f"[ERROR] Не удалось записать {log_path}: {log_error}",
            file=sys.stderr,
        )


def _render_mapping_table(
    pairs: Sequence[tuple[Path, str, float]],
    console: Console,
) -> None:
    """Показывает сводную таблицу сопоставления PDF и листов сметы."""
    table = Table(
        title="Сопоставление томов РД и листов сметы",
        show_lines=False,
    )
    table.add_column("Том РД (PDF)", style="cyan", no_wrap=False)
    table.add_column("Лист сметы", style="magenta")
    table.add_column("Уверенность", justify="right", style="green")

    for pdf_path, sheet_name, confidence in pairs:
        table.add_row(pdf_path.name, sheet_name, f"{confidence:.1f}%")
    console.print(table)


def run_batch_pipeline(
    conn: sqlite3.Connection,
    *,
    project_name: str,
    batch_dir: str | Path,
    smeta_path: str | Path,
    db_path: str = db.DEFAULT_DB_PATH,
    headless: bool = True,
    reviewer: str = "Оператор ПТО",
    output_path: Optional[str | Path] = None,
    error_log_path: str | Path = BATCH_ERROR_LOG,
    console: Optional[Console] = None,
    project_dispatcher: Optional[dispatcher.ProjectDispatcher] = None,
) -> BatchSummary:
    """Обрабатывает папку с томами РД в пакетном (ночном) режиме.

    Один раздел = изолированная сессия SQLite. Ошибка парсинга или расчета
    одного тома не прерывает обработку остальных: она пишется в stderr и
    ``batch_errors.log``. В headless-режиме HITL-опрос не запускается,
    непонятные связи остаются в БД со статусом ``HITL_PENDING`` до утра.
    """
    console = console or Console()
    smeta_file = Path(smeta_path)
    resolver = project_dispatcher or dispatcher.ProjectDispatcher()

    pairs = resolver.match_volumes(Path(batch_dir), smeta_file)
    if not pairs:
        console.print(
            "[yellow]Не найдено ни одной пары PDF ↔ лист ЛС. "
            "Проверьте имена томов и заголовки A8/A9 сметы.[/yellow]"
        )
        return BatchSummary(total=0, succeeded=0, failed=0, hitl_pending=0)

    _render_mapping_table(pairs, console)

    results: list[BatchSectionResult] = []
    log_path = Path(error_log_path)

    for pdf_path, sheet_name, confidence in pairs:
        session_id = uuid.uuid4().hex
        config = SessionConfig(
            project_name=project_name,
            rd_path=str(pdf_path),
            smeta_path=str(smeta_file),
            db_path=db_path,
            session_id=session_id,
            smeta_sheet=sheet_name,
        )
        try:
            pipeline_result = run_pipeline(
                conn,
                config,
                interactive=not headless,
                reviewer=reviewer,
            )
        except Exception as error:  # noqa: BLE001 - ночной прогон не должен падать
            _append_batch_error(log_path, pdf_path, sheet_name, error)
            results.append(
                BatchSectionResult(
                    pdf_path=pdf_path,
                    sheet_name=sheet_name,
                    confidence=confidence,
                    session_id=session_id,
                    error=f"{type(error).__name__}: {error}",
                )
            )
            continue

        results.append(
            BatchSectionResult(
                pdf_path=pdf_path,
                sheet_name=sheet_name,
                confidence=confidence,
                session_id=session_id,
                result=pipeline_result,
            )
        )

    succeeded = [section for section in results if section.succeeded]
    failed = [section for section in results if not section.succeeded]
    hitl_pending = sum(
        section.result.hitl_pending
        for section in succeeded
        if section.result is not None
    )

    summary = BatchSummary(
        total=len(results),
        succeeded=len(succeeded),
        failed=len(failed),
        hitl_pending=hitl_pending,
        sections=tuple(results),
    )
    console.print(
        f"Всего разделов: {summary.total} | Успешно: {summary.succeeded} | "
        f"С ошибками: {summary.failed} | "
        f"Требуют верификации утром: {summary.hitl_pending} позиций."
    )

    target: Optional[Path] = None
    if output_path is not None and succeeded:
        target = exporter.export_sessions_to_excel(
            conn,
            summary.successful_session_ids,
            output_path,
        )
        console.print(f"Многолистовая ведомость сформирована: {target}")
    elif output_path is not None:
        console.print(
            "[yellow]Ведомость не сформирована: нет успешно обработанных "
            "разделов.[/yellow]"
        )

    return BatchSummary(
        total=summary.total,
        succeeded=summary.succeeded,
        failed=summary.failed,
        hitl_pending=summary.hitl_pending,
        sections=summary.sections,
        output_path=target,
    )


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Точка входа all-in-one CLI: ingest -> match -> HITL -> export."""
    parser = argparse.ArgumentParser(
        description=(
            "Единый аудит смет (Excel) и РД (PDF): автоматическое HITL-ревью "
            "пограничных связей и формирование итогового XLSX"
        )
    )
    parser.add_argument(
        "--project",
        "--project-name",
        dest="project_name",
        required=True,
        help="Наименование проекта/объекта",
    )
    parser.add_argument(
        "--rd",
        default=None,
        help="Путь к рабочей документации / спецификации (.pdf)",
    )
    parser.add_argument(
        "--batch-dir",
        dest="batch_dir",
        default=None,
        help="Папка с томами РД (.pdf) для ночного пакетного прогона",
    )
    parser.add_argument(
        "--smeta",
        required=True,
        help="Путь к локальной смете (.xlsx или .pdf формы 4б)",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Конкретный лист XLSX-сметы (для PDF не применяется)",
    )
    parser.add_argument(
        "--db",
        dest="db_path",
        default=db.DEFAULT_DB_PATH,
        help="Путь к файлу SQLite (SSOT)",
    )
    parser.add_argument("--session-id", default=None, help="Идентификатор сессии")
    parser.add_argument(
        "--out",
        default=None,
        help=(
            "Итоговая ведомость аудита (.xlsx); в пакетном режиме — "
            "многолистовая книга по разделам"
        ),
    )
    parser.add_argument(
        "--no-review",
        action="store_true",
        help="Не запускать терминальное HITL-ревью (для автоматизации/CI)",
    )
    parser.add_argument(
        "--headless",
        "--no-interactive",
        dest="headless",
        action="store_true",
        help=(
            "Полный запрет опроса оператора: HITL сохраняется в БД "
            "со статусом HITL_PENDING (для ночных прогонов по расписанию)"
        ),
    )
    parser.add_argument(
        "--interactive",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--reviewer",
        default="Оператор ПТО",
        help="Имя эксперта для записей expert_reviews",
    )
    args = parser.parse_args(argv)

    smeta_path = Path(args.smeta)
    if not smeta_path.is_file():
        parser.error(f"Файл не найден: --smeta {smeta_path}")

    if args.batch_dir:
        batch_dir = Path(args.batch_dir)
        if not batch_dir.is_dir():
            parser.error(f"Папка не найдена: --batch-dir {batch_dir}")
        if args.rd:
            parser.error("--rd не используется вместе с --batch-dir")
        if args.sheet:
            parser.error("--sheet не используется с --batch-dir: лист определяется автоматически")
        if args.session_id:
            parser.error("--session-id не используется с --batch-dir: сессии создаются изолированно")

        conn = db.get_connection(args.db_path)
        try:
            headless = bool(args.headless or args.no_review)
            summary = run_batch_pipeline(
                conn,
                project_name=args.project_name,
                batch_dir=batch_dir,
                smeta_path=smeta_path,
                db_path=args.db_path,
                headless=headless,
                reviewer=args.reviewer,
                output_path=args.out,
            )
        except (OSError, ValueError, sqlite3.Error) as error:
            print(f"[ERROR] Пакетный прогон не выполнен: {error}", file=sys.stderr)
            return 1
        finally:
            conn.close()

        return 0 if summary.failed == 0 else 1

    if not args.rd:
        parser.error("--rd обязателен для одиночного запуска (либо укажите --batch-dir)")
    rd_path = Path(args.rd)
    if not rd_path.is_file():
        parser.error(f"Файл не найден: --rd {rd_path}")

    config = SessionConfig(
        project_name=args.project_name,
        rd_path=str(rd_path),
        smeta_path=str(smeta_path),
        db_path=args.db_path,
        session_id=args.session_id,
        smeta_sheet=args.sheet,
    )

    conn = db.get_connection(args.db_path)
    try:
        result = run_pipeline(
            conn,
            config,
            interactive=not (args.no_review or args.headless),
            reviewer=args.reviewer,
        )
        print(
            f"[+] Сессия {result.session_id} завершена "
            f"со статусом {result.session_status}"
        )
        print(
            f"    РД: {result.rd_saved}, смета: {result.smeta_saved}, "
            f"связей: {result.matches_saved}, "
            f"правил: {result.evaluations_saved}"
        )
        print(
            f"    AUTO: {result.auto_matches}, "
            f"NO_MATCH (< {MIN_MATCH_THRESHOLD}%): {result.no_match}, "
            f"HITL_PENDING: {result.hitl_pending}"
        )
        target = exporter.export_audit_to_excel(
            conn,
            result.session_id,
            args.out or exporter.DEFAULT_OUTPUT,
        )
        print(f"Ведомость сформирована: {target}")
    finally:
        conn.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
