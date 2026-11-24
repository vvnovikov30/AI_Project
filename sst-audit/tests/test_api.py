"""Интеграционные тесты веб-слоя FastAPI (TestClient + изолированный SQLite).

Проверяются: health, согласия 152-ФЗ (403/413, отсутствие сырого логина в БД),
очередь HITL, полный цикл решений (APPROVE/ADJUST/REJECT/SELECT_CANDIDATE),
финализация и журнал арбитража. Отдельно контролируется отсутствие float
в JSON-ответах и сохранение bool/ведущих нулей.
"""

from __future__ import annotations

from decimal import Decimal
import json
from pathlib import Path
from typing import Any

from fastapi.testclient import TestClient
import pytest

import db
from api.main import app

API = "/api/v1"


def _assert_no_float(value: Any) -> None:
    """Рекурсивно проверяет, что в JSON нет чисел с плавающей точкой."""
    assert not isinstance(value, float), f"float в ответе API: {value!r}"
    if isinstance(value, dict):
        for item in value.values():
            _assert_no_float(item)
    elif isinstance(value, list):
        for item in value:
            _assert_no_float(item)


def _rd_payload(index: int, title: str, pointer: dict[str, Any]) -> dict[str, Any]:
    """Аудиторский след позиции РД (совместим с review_service)."""
    return {
        "raw_id": f"rd-{index}",
        "raw_title": title,
        "clean_title": title.lower(),
        "raw_unit": "м",
        "canonical_unit": "м",
        "unit_multiplier": "1",
        "dimensions": {},
        "context_section": "Раздел 1",
        "pointer": pointer,
        "is_resource": False,
    }


def _smeta_payload(index: int, title: str) -> dict[str, Any]:
    """Аудиторский след позиции сметы."""
    return {
        "raw_id": f"smeta-{index}",
        "raw_title": title,
        "clean_title": title.lower(),
        "raw_unit": "м",
        "canonical_unit": "м",
        "unit_multiplier": "1",
        "dimensions": {},
        "context_section": "Раздел 1",
        "is_resource": False,
    }


def _seed_pending_session(
    db_path: Path,
    session_id: str = "api-session",
    pending: int = 3,
) -> list[int]:
    """Создает SSOT-сессию с очередью HITL_PENDING и геометрией PDF."""
    pointer: dict[str, Any] = {
        "doc_type": "SPECIFICATION",
        "file_name": "rd.pdf",
        "sheet_name": "Лист 1",
        "row_idx": 1,
        "col_letter": "B",
        "coordinate": "p. 1",
        "bbox": ["12.5", 20.0, "130.25", 40],
        "page_meta": {
            "width": 842.0,
            "height": 595,
            "rotation": 90,
            "scaled": True,
            "label": "00123",
        },
    }
    conn = db.get_connection(str(db_path))
    try:
        db.init_db(conn)
        db.create_session(conn, session_id, "API test", "rd.pdf", "lsr.xlsx")
        db.save_rd_items(
            conn,
            session_id,
            [
                {
                    "row_idx": index,
                    "code": f"ФЕР01-{index:02d}",
                    "description": "Труба стальная DN50",
                    "unit": "м",
                    "quantity": Decimal("10"),
                    "raw_payload": _rd_payload(
                        index, "Труба стальная DN50", pointer
                    ),
                }
                for index in range(1, pending + 1)
            ],
        )
        db.save_smeta_items(
            conn,
            session_id,
            [
                {
                    "row_idx": index,
                    "code": f"01-{index:02d}",
                    "description": "Труба стальная DN50",
                    "unit": "м",
                    "quantity": Decimal("15"),
                    "unit_price": Decimal("100.00"),
                    "total_cost": Decimal("1500.00"),
                    "raw_payload": _smeta_payload(
                        index, "Труба стальная DN50"
                    ),
                }
                for index in range(1, pending + 1)
            ],
        )
        rd_ids = list(db.get_rd_item_payloads(conn, session_id))
        smeta_ids = list(db.get_smeta_items_map(conn, session_id))
        db.save_matches(
            conn,
            session_id,
            [
                {
                    "rd_item_id": rd_id,
                    "smeta_item_id": smeta_id,
                    "confidence_score": Decimal("55.0"),
                    "conversion_ratio": Decimal("1.0"),
                    "match_method": "FUZZY_CASCADE",
                    "status": "HITL_PENDING",
                }
                for rd_id, smeta_id in zip(rd_ids, smeta_ids)
            ],
        )
        return [int(row["id"]) for row in db.get_match_records(conn, session_id)]
    finally:
        conn.close()


@pytest.fixture()
def db_path(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Изолирует SSOT на временный файл через динамическую настройку."""
    path = tmp_path / "api_test.db"
    monkeypatch.setenv("SST_DB_PATH", str(path))
    monkeypatch.delenv("SST_MAX_UPLOAD_MB", raising=False)
    conn = db.get_connection(str(path))
    try:
        db.init_db(conn)
    finally:
        conn.close()
    return path


@pytest.fixture()
def client(db_path: Path) -> TestClient:
    """Поднимает приложение с lifespan-инициализацией схемы."""
    with TestClient(app) as test_client:
        yield test_client


def _consent(client: TestClient, login: str = "engineer@example.ru") -> str:
    """Фиксирует согласие и возвращает user_id."""
    response = client.post(
        f"{API}/consents",
        json={"login": login, "doc_version": "2026-01"},
    )
    assert response.status_code == 200, response.text
    return str(response.json()["user_id"])


def test_health(client: TestClient) -> None:
    """health отвечает версией API и не требует согласия."""
    response = client.get(f"{API}/health")
    assert response.status_code == 200
    assert response.json() == {"status": "ok", "version": "0.1.0"}


def test_consent_is_minimized_and_reused(
    client: TestClient,
    db_path: Path,
) -> None:
    """Сырой логин не попадает в БД, числа нормализуются, user_id стабилен."""
    payload = {
        "login": "Engineer@Example.RU",
        "doc_version": "2026-01",
        "client_meta": {
            "scaled": True,
            "width": 842.0,
            "rotation": "0",
            "label": "00123",
        },
    }
    first = client.post(f"{API}/consents", json=payload)
    assert first.status_code == 200, first.text
    first_body = first.json()
    assert first_body["accepted"] is True
    assert first_body["contact_mask"]
    assert "engineer" not in first_body["contact_mask"].lower()

    second = client.post(f"{API}/consents", json=payload)
    assert second.status_code == 200, second.text
    second_body = second.json()
    assert second_body["user_id"] == first_body["user_id"]
    assert second_body["consent_id"] != first_body["consent_id"]

    raw_bytes = db_path.read_bytes()
    assert b"Engineer@Example.RU" not in raw_bytes
    assert b"engineer@example.ru" not in raw_bytes

    conn = db.get_connection(str(db_path))
    try:
        user = conn.execute(
            "SELECT login_hash, contact_mask FROM users WHERE id = ?",
            (first_body["user_id"],),
        ).fetchone()
        assert user is not None
        assert len(str(user["login_hash"])) == 64
        consent = conn.execute(
            "SELECT accepted, client_meta FROM consents_152fz "
            "ORDER BY id DESC LIMIT 1"
        ).fetchone()
    finally:
        conn.close()

    assert int(consent["accepted"]) == 1
    meta = json.loads(consent["client_meta"])
    assert meta == {
        "scaled": True,
        "width": "842.0",
        "rotation": "0",
        "label": "00123",
    }
    assert isinstance(meta["scaled"], bool)


def test_sessions_require_active_pdn_consent(client: TestClient) -> None:
    """Без согласия или при accepted=False сессия не создается (403)."""
    files = {
        "rd_file": ("rd.pdf", b"%PDF-1.4", "application/pdf"),
        "lsr_file": ("lsr.xlsx", b"PK", "application/octet-stream"),
    }
    missing = client.post(
        f"{API}/sessions",
        data={"user_id": "ghost-user"},
        files=files,
    )
    assert missing.status_code == 403

    declined = client.post(
        f"{API}/consents",
        json={
            "login": "declined@example.ru",
            "doc_version": "2026-01",
            "accepted": False,
        },
    )
    assert declined.status_code == 200
    user_id = declined.json()["user_id"]

    forbidden = client.post(
        f"{API}/sessions",
        data={"user_id": user_id},
        files=files,
    )
    assert forbidden.status_code == 403


def test_upload_size_limit_returns_413(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Превышение SST_MAX_UPLOAD_MB прерывает запрос кодом 413."""
    user_id = _consent(client, "uploader@example.ru")
    monkeypatch.setenv("SST_MAX_UPLOAD_MB", "1")
    oversize = b"0" * (1024 * 1024 + 1)
    response = client.post(
        f"{API}/sessions",
        data={"user_id": user_id},
        files={
            "rd_file": ("rd.pdf", oversize, "application/pdf"),
            "lsr_file": ("lsr.xlsx", b"PK", "application/octet-stream"),
        },
    )
    assert response.status_code == 413


def test_unknown_session_returns_404(client: TestClient) -> None:
    """Неизвестный session_id дает 404 на всех HITL-эндпоинтах."""
    assert client.get(f"{API}/sessions/missing/cards").status_code == 404
    assert client.get(f"{API}/sessions/missing/events").status_code == 404
    assert client.post(f"{API}/sessions/missing/finalize").status_code == 404


def test_hitl_full_cycle(client: TestClient, db_path: Path) -> None:
    """APPROVE/ADJUST/REJECT, повторное решение 409, финализация и журнал."""
    _seed_pending_session(db_path)
    session_id = "api-session"

    cards_response = client.get(f"{API}/sessions/{session_id}/cards")
    assert cards_response.status_code == 200, cards_response.text
    cards = cards_response.json()
    assert len(cards) == 3
    _assert_no_float(cards)

    first_card = cards[0]
    assert first_card["index"] == 1
    assert first_card["total"] == 3
    assert first_card["confidence"] == "55.0"
    assert first_card["conversion_ratio"] == "1.0"
    assert first_card["bbox"] == ["12.5", "20.0", "130.25", "40"]
    assert first_card["page_meta"] == {
        "width": "842.0",
        "height": "595",
        "rotation": "90",
        "scaled": True,
        "label": "00123",
    }
    assert isinstance(first_card["page_meta"]["scaled"], bool)
    assert first_card["candidates"], "кандидаты РД должны присутствовать"
    assert Decimal(first_card["candidates"][0]["score"]) >= Decimal("40")

    early_finalize = client.post(f"{API}/sessions/{session_id}/finalize")
    assert early_finalize.status_code == 409

    approve = client.post(
        f"{API}/sessions/{session_id}/decisions",
        json={
            "match_id": first_card["match_id"],
            "action": "APPROVE",
            "reviewer": "Эксперт ПТО",
        },
    )
    assert approve.status_code == 200, approve.text
    assert approve.json()["pending_count"] == 2
    _assert_no_float(approve.json())

    adjust = client.post(
        f"{API}/sessions/{session_id}/decisions",
        json={
            "match_id": cards[1]["match_id"],
            "action": "ADJUST",
            "reviewer": "Эксперт ПТО",
            "adjusted_ratio": "1.25",
        },
    )
    assert adjust.status_code == 200, adjust.text
    adjust_body = adjust.json()
    assert adjust_body["adjusted_ratio"] == "1.25"
    assert isinstance(adjust_body["delta_volume"], str)
    assert adjust_body["pending_count"] == 1
    _assert_no_float(adjust_body)

    reject = client.post(
        f"{API}/sessions/{session_id}/decisions",
        json={
            "match_id": cards[2]["match_id"],
            "action": "REJECT",
            "reviewer": "Эксперт ПТО",
            "reason_code": "NOT_IN_PROJECT",
            "comment": "Позиция вне проекта",
        },
    )
    assert reject.status_code == 200, reject.text
    assert reject.json()["pending_count"] == 0

    repeated = client.post(
        f"{API}/sessions/{session_id}/decisions",
        json={
            "match_id": first_card["match_id"],
            "action": "APPROVE",
            "reviewer": "Эксперт ПТО",
        },
    )
    assert repeated.status_code == 409

    finalize = client.post(f"{API}/sessions/{session_id}/finalize")
    assert finalize.status_code == 200, finalize.text
    assert finalize.json() == {
        "session_id": session_id,
        "status": "REVIEWED",
        "pending_count": 0,
    }

    assert client.get(f"{API}/sessions/{session_id}/cards").json() == []

    events_response = client.get(f"{API}/sessions/{session_id}/events")
    assert events_response.status_code == 200
    events = events_response.json()
    actions = {event["action"] for event in events}
    assert {"REVIEW_APPROVED", "REVIEW_ADJUSTED", "REVIEW_REJECTED"} <= actions
    _assert_no_float(events)
    for event in events:
        if event["new_payload"] is not None:
            assert isinstance(event["new_payload"], dict)


def test_select_candidate_and_validation(
    client: TestClient,
    db_path: Path,
) -> None:
    """Выбор кандидата обновляет связь; чужой кандидат — 404."""
    _seed_pending_session(db_path, session_id="select-session", pending=1)
    cards = client.get(f"{API}/sessions/select-session/cards").json()
    assert len(cards) == 1
    card = cards[0]
    candidate = card["candidates"][0]

    unknown_candidate = client.post(
        f"{API}/sessions/select-session/decisions",
        json={
            "match_id": card["match_id"],
            "action": "SELECT_CANDIDATE",
            "reviewer": "Эксперт ПТО",
            "candidate_rd_id": 999_999,
        },
    )
    assert unknown_candidate.status_code == 404

    selected = client.post(
        f"{API}/sessions/select-session/decisions",
        json={
            "match_id": card["match_id"],
            "action": "SELECT_CANDIDATE",
            "reviewer": "Эксперт ПТО",
            "candidate_rd_id": candidate["rd_item_id"],
        },
    )
    assert selected.status_code == 200, selected.text
    assert selected.json()["pending_count"] == 0

    invalid = client.post(
        f"{API}/sessions/select-session/decisions",
        json={
            "match_id": card["match_id"],
            "action": "ADJUST",
            "reviewer": "Эксперт ПТО",
        },
    )
    assert invalid.status_code == 422

    finalize = client.post(f"{API}/sessions/select-session/finalize")
    assert finalize.json()["status"] == "REVIEWED"


