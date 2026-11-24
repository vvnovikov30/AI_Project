"""Зависимости FastAPI: динамические настройки, соединение SSOT и 152-ФЗ.

Настройки читаются из окружения при каждом вызове: это позволяет тестам
изолировать БД и лимит загрузки без перезапуска приложения.
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from typing import Final
import os
import sqlite3

from fastapi import HTTPException

import db

DEFAULT_DB_PATH: Final[str] = "sst_audit.db"
DEFAULT_MAX_UPLOAD_MB: Final[int] = 40
DEFAULT_CONSENT_TYPE: Final[str] = "PDN_PROCESSING"


def get_db_path() -> str:
    """Возвращает путь к SSOT из ``SST_DB_PATH`` (читается динамически)."""
    return os.environ.get("SST_DB_PATH", DEFAULT_DB_PATH) or DEFAULT_DB_PATH


def get_max_upload_bytes() -> int:
    """Возвращает лимит размера одного файла из ``SST_MAX_UPLOAD_MB``.

    Некорректное или неположительное значение заменяется безопасным
    значением по умолчанию (40 МБ), чтобы конфигурация не отключала лимит.
    """
    raw = os.environ.get("SST_MAX_UPLOAD_MB", str(DEFAULT_MAX_UPLOAD_MB))
    try:
        megabytes = int(str(raw).strip())
    except (TypeError, ValueError):
        megabytes = DEFAULT_MAX_UPLOAD_MB
    if megabytes <= 0:
        megabytes = DEFAULT_MAX_UPLOAD_MB
    return megabytes * 1024 * 1024


async def get_db() -> AsyncIterator[sqlite3.Connection]:
    """Отдает соединение SQLite на время запроса и закрывает его.

    Асинхронный генератор выполняется в event loop: соединение не мигрирует
    между потоками anyio (``sqlite3`` по умолчанию проверяет поток создания).
    """
    conn = db.get_connection(get_db_path())
    try:
        conn.execute("PRAGMA busy_timeout = 5000")
        conn.execute("PRAGMA foreign_keys = ON")
        yield conn
    finally:
        conn.close()


def require_pdn_consent(
    conn: sqlite3.Connection,
    user_id: str,
    consent_type: str = DEFAULT_CONSENT_TYPE,
) -> None:
    """Проверяет активное согласие 152-ФЗ либо прерывает запрос кодом 403.

    :raises HTTPException: 400 — пустой ``user_id``; 403 — нет согласия.
    """
    if not user_id or not str(user_id).strip():
        raise HTTPException(status_code=400, detail="user_id не может быть пустым")
    if not db.has_active_consent(conn, str(user_id).strip(), consent_type):
        raise HTTPException(
            status_code=403,
            detail=(
                "Нет активного согласия на обработку персональных данных "
                "(152-ФЗ). Сначала зафиксируйте согласие."
            ),
        )
