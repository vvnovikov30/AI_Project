"""Роутер HITL: карточки очереди, экспертные решения и журнал арбитража.

Слой только валидирует вход, вызывает публичные функции ``review_service``
и сериализует результат. ``bbox``/``page_meta`` не входят в ``ReviewCard``,
поэтому обогащение берется из аудиторского следа ``rd_items.raw_payload``
через ``db.get_rd_item_payloads`` (без SQL в транспортном слое).
"""

from __future__ import annotations

from decimal import Decimal
import json
import sqlite3
from typing import Any

from fastapi import APIRouter, Depends, HTTPException

import db
import review_service
from api.deps import get_db
from api.schemas import (
    ArbitrationEventSchema,
    CandidateSchema,
    DecisionRequest,
    DecisionResponse,
    FinalizeResponse,
    ReviewCardSchema,
    normalize_numeric_strings,
    normalize_page_meta,
)

router = APIRouter(tags=["review"])

CardExtras = tuple[tuple[str, str, str, str] | None, dict[str, Any] | None]


def _resolve_session(conn: sqlite3.Connection, session: str) -> str:
    """Разрешает ``session_id`` (включая ``latest``) или отдает 404."""
    try:
        return review_service.resolve_session_id(conn, session)
    except LookupError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc


def _bbox(value: Any) -> tuple[str, str, str, str] | None:
    """Приводит bbox к кортежу из четырех строк (или ``None``)."""
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        return None
    parts: list[str] = []
    for item in value:
        if isinstance(item, bool):
            return None
        if isinstance(item, (int, float, Decimal)):
            parts.append(str(Decimal(str(item))))
        elif isinstance(item, str) and item.strip():
            parts.append(item)
        else:
            return None
    return (parts[0], parts[1], parts[2], parts[3])


def _card_extras(conn: sqlite3.Connection, session_id: str) -> dict[int, CardExtras]:
    """Собирает ``match_id -> (bbox, page_meta)`` из следа позиций РД."""
    extras: dict[int, CardExtras] = {}
    payloads = db.get_rd_item_payloads(conn, session_id)
    for row in db.get_match_records(conn, session_id):
        rd_item_id = row["rd_item_id"]
        if rd_item_id is None:
            continue
        payload = payloads.get(int(rd_item_id))
        pointer = payload.get("pointer") if isinstance(payload, dict) else None
        if not isinstance(pointer, dict):
            pointer = {}
        extras[int(row["id"])] = (
            _bbox(pointer.get("bbox")),
            normalize_page_meta(pointer.get("page_meta")),
        )
    return extras


def _serialize_card(card: review_service.ReviewCard, extras: CardExtras) -> ReviewCardSchema:
    """Преобразует ``ReviewCard`` в JSON-схему без float."""
    bbox, page_meta = extras
    delta = card.delta_volume()
    return ReviewCardSchema(
        match_id=card.match_id,
        index=card.index,
        total=card.total,
        confidence=str(card.confidence),
        conversion_ratio=str(card.conversion_ratio),
        adjusted_ratio=None,
        rd_code=card.rd_code,
        rd_description=card.rd_description,
        rd_unit=card.rd_unit,
        rd_quantity=None if card.rd_quantity is None else str(card.rd_quantity),
        smeta_code=card.smeta_code,
        smeta_description=card.smeta_description,
        smeta_unit=card.smeta_unit,
        smeta_quantity=None if card.smeta_quantity is None else str(card.smeta_quantity),
        unit_price=None if card.unit_price is None else str(card.unit_price),
        total_cost=None if card.total_cost is None else str(card.total_cost),
        waste_coefficient=str(card.waste_coefficient),
        delta_volume=None if delta is None else str(delta),
        bbox=bbox,
        page_meta=page_meta,
        candidates=[
            CandidateSchema(
                rd_item_id=candidate.rd_item_id,
                score=str(candidate.score),
                description=candidate.description,
                unit=candidate.unit,
                quantity=str(candidate.quantity),
            )
            for candidate in card.candidates
        ],
    )


def _decode_payload(value: Any) -> dict[str, Any] | None:
    """Разбирает JSON-колонку журнала арбитража и нормализует числа."""
    if value is None:
        return None
    if isinstance(value, dict):
        return normalize_page_meta(value)
    try:
        decoded = json.loads(str(value))
    except (TypeError, ValueError):
        return {"raw": str(value)}
    if isinstance(decoded, dict):
        return normalize_page_meta(decoded)
    return {"value": normalize_numeric_strings(decoded)}


def _event_schema(row: sqlite3.Row) -> ArbitrationEventSchema:
    """Преобразует строку ``arbitration_events`` в схему ответа."""
    created_at = row["created_at"]
    return ArbitrationEventSchema(
        id=int(row["id"]),
        session_id=str(row["session_id"]),
        match_record_id=int(row["match_record_id"]),
        actor=str(row["actor"]),
        action=str(row["action"]),
        old_payload=_decode_payload(row["old_payload"]),
        new_payload=_decode_payload(row["new_payload"]),
        comment=None if row["comment"] is None else str(row["comment"]),
        created_at=None if created_at is None else str(created_at),
    )


@router.get("/sessions/{session_id}/cards", response_model=list[ReviewCardSchema])
async def list_cards(
    session_id: str,
    conn: sqlite3.Connection = Depends(get_db),
) -> list[ReviewCardSchema]:
    """Возвращает очередь HITL с геометрией источника (bbox/page_meta)."""
    resolved = _resolve_session(conn, session_id)
    cards = review_service.load_pending_cards(conn, resolved)
    extras = _card_extras(conn, resolved)
    return [
        _serialize_card(card, extras.get(card.match_id, (None, None)))
        for card in cards
    ]


@router.post(
    "/sessions/{session_id}/decisions",
    response_model=DecisionResponse,
)
async def apply_decision(
    session_id: str,
    payload: DecisionRequest,
    conn: sqlite3.Connection = Depends(get_db),
) -> DecisionResponse:
    """Применяет экспертное решение к карточке очереди HITL.

    :raises HTTPException: 404 — сессия/связь/кандидат не найдены;
        409 — связь уже обработана; 400 — некорректное значение.
    """
    resolved = _resolve_session(conn, session_id)
    cards = review_service.load_pending_cards(conn, resolved)
    card = next((item for item in cards if item.match_id == payload.match_id), None)
    if card is None:
        known_ids = {
            int(row["id"]) for row in db.get_match_records(conn, resolved)
        }
        if payload.match_id not in known_ids:
            raise HTTPException(
                status_code=404,
                detail=f"Связь #{payload.match_id} не найдена в сессии",
            )
        raise HTTPException(
            status_code=409,
            detail=f"Связь #{payload.match_id} уже обработана экспертом",
        )

    delta: Decimal | None = None
    try:
        if payload.action == "APPROVE":
            review_service.apply_approval(conn, card, payload.reviewer)
        elif payload.action == "SELECT_CANDIDATE":
            position = next(
                (
                    index
                    for index, candidate in enumerate(card.candidates, start=1)
                    if candidate.rd_item_id == payload.candidate_rd_id
                ),
                None,
            )
            if position is None:
                raise HTTPException(
                    status_code=404,
                    detail=(
                        f"Кандидат РД #{payload.candidate_rd_id} отсутствует "
                        "в карточке"
                    ),
                )
            review_service.apply_candidate_selection(
                conn,
                card,
                card.candidates[position - 1],
                position,
                payload.reviewer,
            )
        elif payload.action == "REJECT":
            review_service.apply_rejection(
                conn,
                card,
                payload.reviewer,
                str(payload.reason_code),
                payload.comment,
            )
        else:
            delta = review_service.apply_adjustment(
                conn,
                card,
                payload.reviewer,
                Decimal(str(payload.adjusted_ratio)),
            )
    except (LookupError, sqlite3.IntegrityError) as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except (ValueError, KeyError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    pending = review_service.pending_count(conn, resolved)
    return DecisionResponse(
        match_id=payload.match_id,
        decision=payload.action,
        status="RESOLVED",
        reviewer=payload.reviewer,
        adjusted_ratio=payload.adjusted_ratio if payload.action == "ADJUST" else None,
        delta_volume=None if delta is None else str(delta),
        pending_count=pending,
    )


@router.post(
    "/sessions/{session_id}/finalize",
    response_model=FinalizeResponse,
)
async def finalize(
    session_id: str,
    conn: sqlite3.Connection = Depends(get_db),
) -> FinalizeResponse:
    """Финализирует сессию, если очередь HITL полностью разобрана."""
    resolved = _resolve_session(conn, session_id)
    pending = review_service.pending_count(conn, resolved)
    if pending:
        raise HTTPException(
            status_code=409,
            detail=f"Осталось необработанных карточек: {pending}",
        )
    try:
        status = review_service.finalize_session(conn, resolved)
    except LookupError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return FinalizeResponse(
        session_id=resolved,
        status=status,
        pending_count=pending,
    )


@router.get(
    "/sessions/{session_id}/events",
    response_model=list[ArbitrationEventSchema],
)
async def list_events(
    session_id: str,
    conn: sqlite3.Connection = Depends(get_db),
) -> list[ArbitrationEventSchema]:
    """Возвращает append-only журнал арбитража в хронологическом порядке."""
    resolved = _resolve_session(conn, session_id)
    return [
        _event_schema(row)
        for row in db.get_arbitration_events(conn, resolved)
    ]
