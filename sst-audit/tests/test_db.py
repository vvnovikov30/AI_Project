"""Тесты слоя данных SQLite (SSOT): схема, каскады, транзакции и CRUD."""

from collections.abc import Iterator
from decimal import Decimal
import sqlite3

import pytest

from db import (
    create_session,
    get_connection,
    get_pending_reviews,
    init_db,
    record_expert_review,
    save_matches,
    save_rd_items,
    save_smeta_items,
    transaction,
)

SESSION_ID = "session-001"

EXPECTED_TABLES = {
    "sessions",
    "rd_items",
    "smeta_items",
    "match_records",
    "rule_evaluations",
    "expert_reviews",
}


@pytest.fixture()
def conn() -> Iterator[sqlite3.Connection]:
    """Изолированная in-memory БД с инициализированной схемой."""
    connection = get_connection(":memory:")
    init_db(connection)
    try:
        yield connection
    finally:
        connection.close()


def _first_id(connection: sqlite3.Connection, table: str) -> int:
    """Возвращает идентификатор первой записи таблицы."""
    row = connection.execute(
        f"SELECT id FROM {table} ORDER BY id LIMIT 1"
    ).fetchone()
    assert row is not None, f"Таблица {table} пуста"
    return int(row["id"])


def _seed_full_chain(connection: sqlite3.Connection) -> int:
    """Создает цепочку session -> items -> match -> rule_evaluation -> review."""
    create_session(connection, SESSION_ID, "ЖК Север", "rd.pdf", "smeta.xlsx")
    save_rd_items(
        connection,
        SESSION_ID,
        [
            {
                "row_idx": 12,
                "code": "ФССЦ-21.1.06.04-0012",
                "description": "Кабель силовой ВВГнг(A)-LS 3х1.5",
                "unit": "100 м",
                "quantity": Decimal("15.0"),
                "raw_payload": {"coordinate": "D12"},
            }
        ],
    )
    save_smeta_items(
        connection,
        SESSION_ID,
        [
            {
                "row_idx": 40,
                "code": "08-02-001-01",
                "description": "Прокладка кабеля",
                "unit": "100 м",
                "quantity": Decimal("10.0"),
                "unit_price": Decimal("2500.00"),
                "total_cost": Decimal("25000.00"),
                "raw_payload": {"coordinate": "E40"},
            }
        ],
    )
    save_matches(
        connection,
        SESSION_ID,
        [
            {
                "rd_item_id": _first_id(connection, "rd_items"),
                "smeta_item_id": _first_id(connection, "smeta_items"),
                "confidence_score": Decimal("74.5"),
                "conversion_ratio": Decimal("1.0"),
                "match_method": "FUZZY_CASCADE",
                "status": "HITL_PENDING",
            }
        ],
    )

    match_id = _first_id(connection, "match_records")
    connection.execute(
        """
        INSERT INTO rule_evaluations
            (match_record_id, rule_name, is_violation, delta_volume, delta_cost, message)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (match_id, "UNIT_MULTIPLIER", 1, 5.0, 12500.0, "Расхождение кратности"),
    )
    record_expert_review(
        connection,
        match_id,
        "Инженер ПТО",
        "APPROVED",
        reason_code=None,
        comment="Подтверждено",
        adjusted_ratio=None,
    )
    connection.commit()
    return match_id


def test_init_db_creates_tables_and_view(conn: sqlite3.Connection) -> None:
    """Схема создается идемпотентно: таблицы, индексы и представление отчета."""
    init_db(conn)  # повторная инициализация не должна падать

    objects = {
        row["name"]: row["type"]
        for row in conn.execute(
            "SELECT name, type FROM sqlite_master WHERE type IN ('table', 'view')"
        )
    }
    for table in EXPECTED_TABLES:
        assert objects.get(table) == "table", f"Таблица {table} не создана"

    assert objects.get("v_audit_export") == "view"
    assert conn.execute("PRAGMA foreign_keys").fetchone()[0] == 1

    view_columns = {
        row["name"] for row in conn.execute("PRAGMA table_info(v_audit_export)")
    }
    assert {
        "session_id",
        "confidence_score",
        "rd_quantity",
        "smeta_quantity",
        "decision",
        "reviewed_at",
    } <= view_columns


def test_foreign_keys_cascade_deletion(conn: sqlite3.Connection) -> None:
    """Удаление сессии каскадно очищает все дочерние сущности."""
    _seed_full_chain(conn)

    conn.execute("DELETE FROM sessions WHERE id = ?", (SESSION_ID,))
    conn.commit()

    for table in (
        "rd_items",
        "smeta_items",
        "match_records",
        "rule_evaluations",
        "expert_reviews",
    ):
        count = conn.execute(f"SELECT COUNT(*) AS c FROM {table}").fetchone()["c"]
        assert count == 0, f"Каскадное удаление не сработало для таблицы {table}"

    export_count = conn.execute(
        "SELECT COUNT(*) AS c FROM v_audit_export"
    ).fetchone()["c"]
    assert export_count == 0


def test_transaction_rollback_on_error(conn: sqlite3.Connection) -> None:
    """Ошибка внутри транзакции откатывает все изменения."""
    with pytest.raises(sqlite3.IntegrityError):
        with transaction(conn):
            conn.execute(
                """
                INSERT INTO sessions
                    (id, project_name, rd_filename, smeta_filename)
                VALUES (?, ?, ?, ?)
                """,
                ("rollback-session", "Проект", "rd.pdf", "smeta.xlsx"),
            )
            # Нарушение внешнего ключа: сессии "missing-session" не существует
            conn.execute(
                """
                INSERT INTO rd_items (session_id, description, unit, quantity)
                VALUES (?, ?, ?, ?)
                """,
                ("missing-session", "Кабель", "м", 10.0),
            )

    rolled_back = conn.execute(
        "SELECT COUNT(*) AS c FROM sessions WHERE id = 'rollback-session'"
    ).fetchone()["c"]
    assert rolled_back == 0

    # Соединение остается работоспособным после rollback
    create_session(conn, "after-rollback", "Проект", "rd.pdf", "smeta.xlsx")
    assert conn.execute("SELECT COUNT(*) AS c FROM sessions").fetchone()["c"] == 1


def test_save_and_retrieve_audit_data(conn: sqlite3.Connection) -> None:
    """Сохранение сессии, сущностей и чтение сводки через v_audit_export."""
    create_session(conn, SESSION_ID, "ЖК Север, корпус 3", "СОТ.pdf", "СОТ.xlsx")

    rd_saved = save_rd_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 12,
                "code": "ФССЦ-21.1.06.04-0012",
                "description": "Кабель силовой ВВГнг(A)-LS 3х1.5",
                "unit": "100 м",
                "quantity": Decimal("15.0"),
                "raw_payload": {"sheet": "СОТ", "coordinate": "D12"},
            }
        ],
    )
    assert rd_saved == 1

    smeta_saved = save_smeta_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 40,
                "code": "08-02-001-01",
                "description": "Прокладка кабеля",
                "unit": "100 м",
                "quantity": Decimal("10.0"),
                "unit_price": Decimal("2500.00"),
                "total_cost": Decimal("25000.00"),
                "raw_payload": {"sheet": "Смета"},
            }
        ],
    )
    assert smeta_saved == 1

    matches_saved = save_matches(
        conn,
        SESSION_ID,
        [
            {
                "rd_item_id": _first_id(conn, "rd_items"),
                "smeta_item_id": _first_id(conn, "smeta_items"),
                "confidence_score": Decimal("91.5"),
                "conversion_ratio": Decimal("1.0"),
                "match_method": "FUZZY_CASCADE",
                "status": "HITL_PENDING",
            }
        ],
    )
    assert matches_saved == 1

    pending = get_pending_reviews(conn, SESSION_ID)
    assert len(pending) == 1
    assert pending[0]["rd_code"] == "ФССЦ-21.1.06.04-0012"
    assert pending[0]["smeta_code"] == "08-02-001-01"
    assert pending[0]["confidence_score"] == pytest.approx(91.5)

    review_id = record_expert_review(
        conn,
        int(pending[0]["match_id"]),
        "Инженер ПТО",
        "ADJUSTED",
        reason_code="RATIO_MISMATCH",
        comment="Уточнена кратность",
        adjusted_ratio=Decimal("0.1"),
    )
    assert isinstance(review_id, int)
    assert review_id > 0
    assert get_pending_reviews(conn, SESSION_ID) == []

    rows = conn.execute(
        "SELECT * FROM v_audit_export WHERE session_id = ?", (SESSION_ID,)
    ).fetchall()
    assert len(rows) == 1

    row = rows[0]
    assert row["project_name"] == "ЖК Север, корпус 3"
    assert row["rd_code"] == "ФССЦ-21.1.06.04-0012"
    assert row["rd_quantity"] == pytest.approx(15.0)
    assert row["smeta_quantity"] == pytest.approx(10.0)
    assert row["total_cost"] == pytest.approx(25000.0)
    assert row["match_method"] == "FUZZY_CASCADE"
    assert row["match_status"] == "RESOLVED"
    assert row["decision"] == "ADJUSTED"
    assert row["adjusted_ratio"] == pytest.approx(0.1)
    assert row["reject_reason_code"] == "RATIO_MISMATCH"
    assert row["reviewer_name"] == "Инженер ПТО"
