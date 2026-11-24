"""ASGI-приложение sst-audit: тонкий веб-слой поверх ядра (API 0.1.0).

Эндпоинты сгруппированы под ``/api/v1``:
* ``/consents`` — согласие 152-ФЗ и обезличенный профиль;
* ``/sessions`` — создание сессии из загруженных файлов;
* ``/sessions/{id}/cards|decisions|finalize|events`` — HITL.
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager

from fastapi import FastAPI

import db
from api.deps import get_db_path
from api.routers import audits, auth, review

API_VERSION = "0.1.0"
API_PREFIX = "/api/v1"


@asynccontextmanager
async def lifespan(_: FastAPI) -> AsyncIterator[None]:
    """Инициализирует схему SSOT по динамическому пути из окружения."""
    conn = db.get_connection(get_db_path())
    try:
        db.init_db(conn)
    finally:
        conn.close()
    yield


app = FastAPI(
    title="sst-audit API",
    version=API_VERSION,
    description="Тонкий веб-слой: согласия 152-ФЗ, сессии аудита и HITL.",
    lifespan=lifespan,
)

app.include_router(auth.router, prefix=API_PREFIX)
app.include_router(audits.router, prefix=API_PREFIX)
app.include_router(review.router, prefix=API_PREFIX)


@app.get(f"{API_PREFIX}/health", tags=["health"])
async def health() -> dict[str, str]:
    """Проверка живости сервиса без обращения к данным."""
    return {"status": "ok", "version": API_VERSION}
