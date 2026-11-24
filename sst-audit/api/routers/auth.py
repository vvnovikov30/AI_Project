"""Роутер 152-ФЗ: согласие субъекта ПДн и обезличенный профиль."""

from __future__ import annotations

import sqlite3
import uuid

from fastapi import APIRouter, Depends, HTTPException

import db
from api.deps import get_db
from api.schemas import ConsentRequest, ConsentResponse, normalize_page_meta

router = APIRouter(tags=["auth"])


@router.post("/consents", response_model=ConsentResponse)
async def record_consent(
    payload: ConsentRequest,
    conn: sqlite3.Connection = Depends(get_db),
) -> ConsentResponse:
    """Фиксирует субъекта ПДн по необратимому хэшу и записывает согласие.

    Сырой логин не сохраняется: в ``users`` пишется только ``login_hash``,
    в ``consents_152fz`` — тип, версия документа и ``client_meta``.
    Повторный вход тем же логином переиспользует существующего пользователя.
    """
    try:
        login_hash = db.hash_login(payload.login)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    contact_mask = db.mask_login(payload.login)
    existing = db.get_user_by_login_hash(conn, login_hash)
    if existing is None:
        user_id = uuid.uuid4().hex
        db.create_user(conn, user_id, login_hash, contact_mask=contact_mask)
    else:
        user_id = str(existing["id"])
        stored_mask = existing["contact_mask"]
        if stored_mask:
            contact_mask = str(stored_mask)

    try:
        consent_id = db.record_consent_152fz(
            conn,
            payload.consent_type,
            payload.doc_version,
            accepted=payload.accepted,
            user_id=user_id,
            client_meta=normalize_page_meta(payload.client_meta),
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return ConsentResponse(
        user_id=user_id,
        contact_mask=contact_mask or None,
        consent_id=consent_id,
        accepted=payload.accepted,
    )
