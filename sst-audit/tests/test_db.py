"""Тесты слоя данных SQLite (SSOT): схема, каскады, транзакции и CRUD."""

from collections.abc import Iterator
from decimal import Decimal
import json
import sqlite3

import pytest

from db import (
    SCHEMA_SQL,
    VALID_USER_ROLES,
    backup_database,
    create_session,
    create_user,
    get_arbitration_events,
    get_connection,
    get_pending_reviews,
    get_rule_evaluations,
    get_smeta_items_map,
    get_user_by_login_hash,
    hash_login,
    init_db,
    mask_login,
    normalize_login,
    record_arbitration_event,
    record_consent_152fz,
    record_expert_review,
    save_matches,
    save_rd_items,
    save_rule_evaluations,
    save_smeta_items,
    transaction,
    update_match_link,
    update_session_status,
)
from models import RuleEvaluation
import db

SESSION_ID = "session-001"

EXPECTED_TABLES = {
    "sessions",
    "rd_items",
    "smeta_items",
    "match_records",
    "rule_evaluations",
    "expert_reviews",
    "users",
    "consents_152fz",
    "arbitration_events",
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


def _seed_full_chain(
    connection: sqlite3.Connection,
    *,
    with_rule_evaluation: bool = True,
) -> int:
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
    if with_rule_evaluation:
        save_rule_evaluations(
            connection,
            [
                {
                    "match_record_id": match_id,
                    "rule_name": "UNIT_MULTIPLIER",
                    "is_violation": True,
                    "delta_volume": Decimal("5.0"),
                    "delta_cost": Decimal("12500.0"),
                    "message": "Расхождение кратности",
                }
            ],
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

def test_db_schema_numeric_columns_are_text(conn: sqlite3.Connection) -> None:
    """Числовые колонки объявлены TEXT: SQLite не кастует Decimal во float."""
    expected = {
        "rd_items": {"quantity"},
        "smeta_items": {"quantity", "unit_price", "total_cost"},
        "rule_evaluations": {"score_weight"},
        "match_records": {"confidence_score", "conversion_ratio"},
        "expert_reviews": {"adjusted_ratio"},
    }
    for table, columns in expected.items():
        declared = {
            row["name"]: str(row["type"]).upper()
            for row in conn.execute(f"PRAGMA table_info({table})")
        }
        for column in columns:
            assert declared.get(column) == "TEXT", (
                f"{table}.{column} объявлена как {declared.get(column)!r}, "
                "ожидался TEXT"
            )


def test_db_stores_and_loads_decimal_as_text_without_float_conversion(
    conn: sqlite3.Connection,
) -> None:
    """Точность Decimal сохраняется: запись строкой, чтение через Decimal(str)."""
    create_session(conn, SESSION_ID, "Проект", "rd.pdf", "smeta.xlsx")
    quantity = Decimal("-840.97440001")
    unit_price = Decimal("1234.567890123456789")
    total_cost = Decimal("-1038903.45678901")

    save_rd_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 12,
                "code": "ФССЦ-21.1.06.04-0012",
                "description": "Кабель силовой",
                "unit": "100 м",
                "quantity": quantity,
            }
        ],
    )
    save_smeta_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 40,
                "code": "08-02-001-01",
                "description": "Прокладка кабеля",
                "unit": "100 м",
                "quantity": quantity,
                "unit_price": unit_price,
                "total_cost": total_cost,
            }
        ],
    )

    rd_row = conn.execute("SELECT quantity FROM rd_items").fetchone()
    smeta_row = conn.execute(
        "SELECT quantity, unit_price, total_cost FROM smeta_items"
    ).fetchone()

    assert isinstance(rd_row["quantity"], str)
    assert rd_row["quantity"] == str(quantity)
    assert smeta_row["quantity"] == str(quantity)
    assert smeta_row["unit_price"] == str(unit_price)
    assert smeta_row["total_cost"] == str(total_cost)

    loaded = get_smeta_items_map(conn, SESSION_ID)
    item = next(iter(loaded.values()))
    assert item.quantity == quantity
    assert item.unit_price == unit_price
    assert item.total_cost == total_cost




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
    assert Decimal(str(pending[0]["confidence_score"])) == Decimal("91.5")

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
    assert Decimal(str(row["rd_quantity"])) == Decimal("15.0")
    assert Decimal(str(row["smeta_quantity"])) == Decimal("10.0")
    assert Decimal(str(row["total_cost"])) == Decimal("25000.00")
    assert row["match_method"] == "FUZZY_CASCADE"
    assert row["match_status"] == "RESOLVED"
    assert row["decision"] == "ADJUSTED"
    assert Decimal(str(row["adjusted_ratio"])) == Decimal("0.1")
    assert row["reject_reason_code"] == "RATIO_MISMATCH"
    assert row["reviewer_name"] == "Инженер ПТО"


def test_rule_evaluations_and_session_status_validation(
    conn: sqlite3.Connection,
) -> None:
    """Недопустимые статусы сессии и правила отклоняются до записи."""
    create_session(conn, SESSION_ID, "Проект", "rd.pdf", "smeta.xlsx")

    with pytest.raises(ValueError, match="Недопустимый статус сессии"):
        update_session_status(conn, SESSION_ID, "DONE")

    with pytest.raises(ValueError, match="не найдена"):
        update_session_status(conn, "missing-session", "MATCHED")

    update_session_status(conn, SESSION_ID, "REVIEWED")
    status = conn.execute(
        "SELECT status FROM sessions WHERE id = ?", (SESSION_ID,)
    ).fetchone()["status"]
    assert status == "REVIEWED"

    with pytest.raises(ValueError, match="rule_name"):
        save_rule_evaluations(
            conn,
            [
                {
                    "match_record_id": 1,
                    "rule_name": "   ",
                    "is_violation": True,
                }
            ],
        )

    with pytest.raises(ValueError, match="is_violation"):
        save_rule_evaluations(
            conn,
            [
                {
                    "match_record_id": 1,
                    "rule_name": "VOLUME_DELTA_CRITICAL",
                    "is_violation": "yes",
                }
            ],
        )

    assert conn.execute(
        "SELECT COUNT(*) AS c FROM rule_evaluations"
    ).fetchone()["c"] == 0

    with pytest.raises(ValueError, match="не найдена"):
        update_match_link(conn, 42, None, Decimal("0"))


def test_create_session_persists_smeta_sheet(conn: sqlite3.Connection) -> None:
    """Имя листа раздела сохраняется в сессии для пакетного экспорта."""
    create_session(
        conn,
        "batch-section",
        "ЖК Север",
        "СОТ.pdf",
        "smeta.xlsx",
        smeta_sheet="ЛС 02-01-14",
    )

    row = conn.execute(
        "SELECT smeta_sheet FROM sessions WHERE id = 'batch-section'"
    ).fetchone()
    assert row is not None
    assert row["smeta_sheet"] == "ЛС 02-01-14"


def test_init_db_migrates_legacy_sessions_table(tmp_path) -> None:
    """Старая БД без smeta_sheet обновляется идемпотентно без потери данных."""
    connection = get_connection(tmp_path / "legacy.db")
    try:
        connection.execute(
            """
            CREATE TABLE sessions (
                id TEXT PRIMARY KEY,
                project_name TEXT NOT NULL,
                rd_filename TEXT NOT NULL,
                smeta_filename TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status TEXT DEFAULT 'CREATED'
            )
            """
        )
        connection.execute(
            """
            INSERT INTO sessions (id, project_name, rd_filename, smeta_filename)
            VALUES ('legacy', 'Старый проект', 'rd.pdf', 'smeta.xlsx')
            """
        )
        connection.commit()

        init_db(connection)
        init_db(connection)

        columns = {
            row["name"] for row in connection.execute("PRAGMA table_info(sessions)")
        }
        assert "smeta_sheet" in columns
        row = connection.execute(
            "SELECT project_name, smeta_sheet FROM sessions WHERE id = 'legacy'"
        ).fetchone()
        assert row is not None
        assert row["project_name"] == "Старый проект"
        assert row["smeta_sheet"] is None
    finally:
        connection.close()


def test_save_and_retrieve_rule_evaluations(conn: sqlite3.Connection) -> None:
    """Батч RuleEvaluation пишется и читается с точным Decimal-весом."""
    match_id = _seed_full_chain(conn, with_rule_evaluation=False)

    saved = save_rule_evaluations(
        conn,
        [
            RuleEvaluation(
                session_id=SESSION_ID,
                match_record_id=match_id,
                rule_code="UNIT_MULTIPLIER",
                status="FAILED",
                details="Кратность 10 вместо 15",
                score_weight=Decimal("0.35"),
            ),
            RuleEvaluation(
                session_id=SESSION_ID,
                match_record_id=match_id,
                rule_code="COST_DELTA",
                status="PASSED",
                score_weight=Decimal("0.65000000000000000001"),
            ),
        ],
    )
    assert saved == 2

    rows = get_rule_evaluations(conn, SESSION_ID, match_record_id=match_id)
    assert [row.rule_code for row in rows] == ["UNIT_MULTIPLIER", "COST_DELTA"]
    assert all(isinstance(row, RuleEvaluation) for row in rows)
    assert rows[0].status == "FAILED"
    assert rows[0].details == "Кратность 10 вместо 15"
    # Точность веса не теряется: в колонке TEXT лежит str(Decimal).
    assert rows[1].score_weight == Decimal("0.65000000000000000001")
    assert str(rows[1].score_weight) == "0.65000000000000000001"
    assert rows[1].score_weight is not None
    assert rows[0].created_at

    assert get_rule_evaluations(conn, SESSION_ID, match_record_id=match_id + 1) == []
    assert get_rule_evaluations(conn, "missing-session") == []


def test_save_empty_rule_evaluations_is_noop(conn: sqlite3.Connection) -> None:
    """Пустой батч не открывает транзакцию и не создает строк."""
    create_session(conn, SESSION_ID, "Проект", "rd.pdf", "smeta.xlsx")

    assert save_rule_evaluations(conn, []) == 0
    assert get_rule_evaluations(conn, SESSION_ID) == []


def test_rule_evaluations_batch_rollback_on_error(conn: sqlite3.Connection) -> None:
    """Ошибка FK в середине батча откатывает все уже вставленные предикаты."""
    match_id = _seed_full_chain(conn, with_rule_evaluation=False)

    with pytest.raises(sqlite3.IntegrityError):
        save_rule_evaluations(
            conn,
            [
                RuleEvaluation(
                    session_id=SESSION_ID,
                    match_record_id=match_id,
                    rule_code="VALID_RULE",
                    status="PASSED",
                ),
                RuleEvaluation(
                    session_id=SESSION_ID,
                    match_record_id=999_999,
                    rule_code="BROKEN_FK",
                    status="FAILED",
                ),
            ],
        )

    assert get_rule_evaluations(conn, SESSION_ID) == []


def test_get_smeta_items_map_exposes_parent_hierarchy(
    conn: sqlite3.Connection,
) -> None:
    """Карта сметных строк отдает Decimal-поля и parent_id без N+1."""
    create_session(conn, SESSION_ID, "Проект", "rd.pdf", "smeta.xlsx")
    save_smeta_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 10,
                "code": "01-01-001",
                "description": "Раздел: земляные работы",
                "unit": "шт",
                "quantity": Decimal("1"),
                "unit_price": Decimal("0"),
                "total_cost": Decimal("0"),
            }
        ],
    )
    parent_id = _first_id(conn, "smeta_items")
    save_smeta_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 11,
                "code": "01-01-001-01",
                "description": "Разработка грунта",
                "unit": "100 м3",
                "quantity": Decimal("15.0"),
                "unit_price": Decimal("100.5"),
                "total_cost": Decimal("1507.50"),
                "parent_id": parent_id,
                "raw_payload": {"coordinate": "E11"},
            }
        ],
    )

    items = get_smeta_items_map(conn, SESSION_ID)
    assert len(items) == 2

    parent = items[parent_id]
    assert parent.parent_id is None
    assert parent.quantity == Decimal("1")

    child = next(item for item in items.values() if item.parent_id == parent_id)
    assert child.code == "01-01-001-01"
    assert child.quantity == Decimal("15.0")
    assert child.total_cost == Decimal("1507.50")
    assert child.raw_payload == {"coordinate": "E11"}
    assert get_smeta_items_map(conn, "missing-session") == {}


def test_init_db_migrates_legacy_real_columns_to_text(tmp_path) -> None:
    """Устаревшие REAL-колонки пересобираются в TEXT без потери связей."""
    connection = get_connection(tmp_path / "legacy-real.db")
    try:
        connection.executescript(SCHEMA_SQL)
        connection.executescript(
            """
            DROP TABLE rd_items;
            CREATE TABLE rd_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
                row_idx INTEGER,
                code TEXT,
                description TEXT NOT NULL,
                unit TEXT NOT NULL,
                quantity REAL NOT NULL,
                raw_payload JSON
            );
            DROP TABLE smeta_items;
            CREATE TABLE smeta_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
                row_idx INTEGER,
                code TEXT NOT NULL,
                description TEXT NOT NULL,
                unit TEXT NOT NULL,
                quantity REAL NOT NULL,
                unit_price REAL NOT NULL,
                total_cost REAL NOT NULL,
                parent_id INTEGER,
                raw_payload JSON
            );
            DROP TABLE expert_reviews;
            DROP TABLE match_records;
            CREATE TABLE match_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
                rd_item_id INTEGER REFERENCES rd_items(id) ON DELETE CASCADE,
                smeta_item_id INTEGER REFERENCES smeta_items(id) ON DELETE CASCADE,
                confidence_score REAL NOT NULL,
                conversion_ratio REAL DEFAULT 1.0,
                match_method TEXT NOT NULL,
                status TEXT DEFAULT 'AUTO'
            );
            CREATE TABLE expert_reviews (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                match_record_id INTEGER NOT NULL
                    REFERENCES match_records(id) ON DELETE CASCADE,
                reviewer_name TEXT NOT NULL,
                decision TEXT,
                adjusted_ratio REAL,
                reject_reason_code TEXT,
                comment TEXT,
                reviewed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            INSERT INTO sessions (id, project_name, rd_filename, smeta_filename)
            VALUES ('legacy', 'Старый проект', 'rd.pdf', 'smeta.xlsx');
            INSERT INTO rd_items
                (session_id, code, description, unit, quantity)
            VALUES ('legacy', 'ФССЦ-1', 'Кабель', 'м', 15.0);
            INSERT INTO smeta_items
                (session_id, code, description, unit,
                 quantity, unit_price, total_cost)
            VALUES ('legacy', '01-01', 'Прокладка кабеля', 'м',
                    10.0, 2500.0, 25000.0);
            INSERT INTO match_records
                (session_id, rd_item_id, smeta_item_id,
                 confidence_score, conversion_ratio, match_method)
            VALUES ('legacy', 1, 1, 90.5, 2.0, 'FUZZY_CASCADE');
            INSERT INTO expert_reviews
                (match_record_id, reviewer_name, decision, adjusted_ratio)
            VALUES (1, 'Петров', 'ADJUSTED', 1.25);
            """
        )
        connection.commit()

        init_db(connection)
        init_db(connection)

        for table, columns in {
            "rd_items": ("quantity",),
            "smeta_items": ("quantity", "unit_price", "total_cost"),
            "match_records": ("confidence_score", "conversion_ratio"),
            "expert_reviews": ("adjusted_ratio",),
        }.items():
            declared = {
                row["name"]: str(row["type"]).upper()
                for row in connection.execute(f"PRAGMA table_info({table})")
            }
            assert all(declared[column] == "TEXT" for column in columns)

        match_row = connection.execute(
            "SELECT confidence_score, conversion_ratio FROM match_records"
        ).fetchone()
        assert match_row["confidence_score"] == "90.5"
        assert match_row["conversion_ratio"] == "2.0"
        review_row = connection.execute(
            "SELECT adjusted_ratio FROM expert_reviews"
        ).fetchone()
        assert review_row["adjusted_ratio"] == "1.25"

        # Пересборка таблиц удаляет их индексы: init_db обязан их пересоздать.
        index_names = {
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'index'"
            )
        }
        assert {
            "idx_match_records_session_status",
            "idx_expert_reviews_match",
        } <= index_names

        rd_row = connection.execute(
            "SELECT quantity FROM rd_items WHERE session_id = 'legacy'"
        ).fetchone()
        smeta_row = connection.execute(
            "SELECT quantity, unit_price, total_cost FROM smeta_items "
            "WHERE session_id = 'legacy'"
        ).fetchone()
        assert isinstance(rd_row["quantity"], str)
        assert isinstance(smeta_row["unit_price"], str)
        assert connection.execute(
            "SELECT COUNT(*) FROM match_records WHERE session_id = 'legacy'"
        ).fetchone()[0] == 1
        assert connection.execute("PRAGMA foreign_keys").fetchone()[0] == 1

        loaded = get_smeta_items_map(connection, "legacy")
        item = next(iter(loaded.values()))
        assert item.quantity == Decimal("10.0")
        assert item.unit_price == Decimal("2500.0")
        assert item.total_cost == Decimal("25000.0")

        save_rd_items(
            connection,
            "legacy",
            [
                {
                    "row_idx": 13,
                    "code": "ФССЦ-2",
                    "description": "Кабель силовой",
                    "unit": "100 м",
                    "quantity": Decimal("-840.97440001"),
                }
            ],
        )
        inserted = connection.execute(
            "SELECT quantity FROM rd_items WHERE code = 'ФССЦ-2'"
        ).fetchone()
        assert inserted["quantity"] == "-840.97440001"
    finally:
        connection.close()

def test_init_db_migrates_legacy_rule_evaluations_and_parent_id(
    tmp_path,
) -> None:

    """Устаревшая схема предикатов и сметы обновляется идемпотентно."""
    connection = get_connection(tmp_path / "legacy-rules.db")
    try:
        connection.executescript(SCHEMA_SQL)
        # Откатываем схему к состоянию до появления иерархии и rule_code.
        connection.executescript(
            """
            DROP TABLE smeta_items;
            CREATE TABLE smeta_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
                row_idx INTEGER,
                code TEXT NOT NULL,
                description TEXT NOT NULL,
                unit TEXT NOT NULL,
                quantity REAL NOT NULL,
                unit_price REAL NOT NULL,
                total_cost REAL NOT NULL,
                raw_payload JSON
            );
            CREATE TABLE rule_evaluations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                match_record_id INTEGER REFERENCES match_records(id) ON DELETE CASCADE,
                rule_name TEXT NOT NULL,
                is_violation INTEGER,
                delta_volume REAL,
                delta_cost REAL,
                message TEXT
            );
            INSERT INTO sessions (id, project_name, rd_filename, smeta_filename)
            VALUES ('legacy', 'Старый проект', 'rd.pdf', 'smeta.xlsx');
            INSERT INTO match_records (session_id, confidence_score, match_method)
            VALUES ('legacy', 80, 'FUZZY_CASCADE');
            INSERT INTO rule_evaluations
                (match_record_id, rule_name, is_violation, message)
            VALUES (
                (SELECT id FROM match_records LIMIT 1),
                'UNIT_MULTIPLIER', 1, 'Расхождение кратности'
            );
            """
        )
        connection.commit()

        init_db(connection)
        init_db(connection)

        rule_columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(rule_evaluations)")
        }
        assert {"session_id", "rule_code", "status", "score_weight"} <= rule_columns
        assert "rule_name" not in rule_columns

        migrated = get_rule_evaluations(connection, "legacy")
        assert len(migrated) == 1
        assert migrated[0].rule_code == "UNIT_MULTIPLIER"
        assert migrated[0].status == "FAILED"
        assert migrated[0].details == "Расхождение кратности"
        assert migrated[0].score_weight == Decimal("0.0")

        smeta_columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(smeta_items)")
        }
        assert "parent_id" in smeta_columns

        index_names = {
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'index'"
            )
        }
        assert {
            "idx_rd_items_session",
            "idx_smeta_items_session",
            "idx_smeta_items_parent",
            "idx_rule_evaluations_match",
            "idx_rule_eval_lookup",
        } <= index_names
    finally:
        connection.close()


# ----------------------------------------------------------------------
# Горячее резервирование SSOT (WAL-safe snapshot + ротация)
# ----------------------------------------------------------------------


class TestDatabaseBackup:
    """Проверяет снимок БД через sqlite3 backup API и ротацию файлов."""

    def test_backup_nonexistent_db_returns_none(self, tmp_path) -> None:
        """Первый чистый запуск: файла нет — снимок не создается, исключений нет."""
        missing = tmp_path / "missing.db"

        assert backup_database(db_path=missing) is None
        assert not (tmp_path / "backups").exists()

    def test_backup_creates_valid_copy_with_wal(self, tmp_path) -> None:
        """Копия включает данные, оставшиеся в WAL (без явного checkpoint)."""
        source_path = tmp_path / "wal_source.db"
        source = sqlite3.connect(str(source_path), timeout=30)
        try:
            journal_mode = source.execute("PRAGMA journal_mode=WAL").fetchone()[0]
            assert str(journal_mode).lower() == "wal"
            source.execute(
                "CREATE TABLE items (id INTEGER PRIMARY KEY, code TEXT NOT NULL)"
            )
            source.executemany(
                "INSERT INTO items (code) VALUES (?)",
                [("A-1",), ("B-2",), ("C-3",)],
            )
            source.commit()
            # Соединение не закрыто: часть страниц гарантированно остается в -wal.
            wal_path = source_path.with_name(source_path.name + "-wal")
            assert wal_path.is_file()

            backup_path = backup_database(db_path=source_path)
        finally:
            source.close()

        assert backup_path is not None
        assert backup_path.is_file()
        assert backup_path.parent == source_path.parent / "backups"
        assert backup_path.suffix == ".db"
        assert backup_path.name.startswith("wal_source_")

        restored = sqlite3.connect(str(backup_path))
        try:
            integrity = restored.execute("PRAGMA integrity_check").fetchone()[0]
            rows = restored.execute("SELECT code FROM items ORDER BY id").fetchall()
        finally:
            restored.close()

        assert integrity == "ok"
        assert [row[0] for row in rows] == ["A-1", "B-2", "C-3"]

    def test_backup_rotation_limits_file_count(self, tmp_path) -> None:
        """Ротация по имени оставляет ровно max_backups самых свежих снимков."""
        source_path = tmp_path / "rotate.db"
        source = sqlite3.connect(str(source_path), timeout=30)
        try:
            source.execute("CREATE TABLE items (id INTEGER PRIMARY KEY)")
            source.execute("INSERT INTO items (id) VALUES (1)")
            source.commit()
        finally:
            source.close()

        backup_dir = tmp_path / "backups"
        backup_dir.mkdir()
        for stale_name in (
            "rotate_20200101_000000.db",
            "rotate_20200102_000000.db",
            "rotate_20200103_000000.db",
        ):
            (backup_dir / stale_name).write_bytes(b"stale snapshot")

        fresh = backup_database(
            db_path=source_path,
            backup_dir=backup_dir,
            max_backups=2,
        )

        assert fresh is not None
        assert fresh.is_file()
        snapshots = sorted(backup_dir.glob("rotate_*.db"), reverse=True)
        assert len(snapshots) == 2
        assert snapshots[0].name == fresh.name
        assert snapshots[1].name == "rotate_20200103_000000.db"
        assert not (backup_dir / "rotate_20200101_000000.db").exists()
        assert not (backup_dir / "rotate_20200102_000000.db").exists()


# ----------------------------------------------------------------------
# 152-ФЗ: пользователи по хэшу, согласия и append-only журнал арбитража
# ----------------------------------------------------------------------


def _seed_pending_match(connection: sqlite3.Connection) -> int:
    """Создает цепочку session -> items -> HITL_PENDING match без ревью."""
    create_session(connection, SESSION_ID, "ЖК Север", "rd.pdf", "smeta.xlsx")
    save_rd_items(
        connection,
        SESSION_ID,
        [
            {
                "row_idx": 7,
                "code": "ART-1",
                "description": "Арматура А500С",
                "unit": "т",
                "quantity": Decimal("2.0"),
            },
            {
                "row_idx": 8,
                "code": "CBL-2",
                "description": "Кабель силовой ВВГнг 3х1.5",
                "unit": "м",
                "quantity": Decimal("100.0"),
            },
        ],
    )
    save_smeta_items(
        connection,
        SESSION_ID,
        [
            {
                "row_idx": 3,
                "code": "08-02-001-01",
                "description": "Прокладка кабеля",
                "unit": "м",
                "quantity": Decimal("104.0"),
                "unit_price": Decimal("100.00"),
                "total_cost": Decimal("10400.00"),
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
                "confidence_score": Decimal("62.5"),
                "conversion_ratio": Decimal("1.0"),
                "match_method": "FUZZY_CASCADE",
                "status": "HITL_PENDING",
            }
        ],
    )
    return _first_id(connection, "match_records")


def test_arbitration_events_are_append_only(conn: sqlite3.Connection) -> None:
    """Журнал арбитража доступен только на добавление и чтение (append-only)."""
    for forbidden in (
        "update_arbitration_event",
        "delete_arbitration_event",
        "clear_arbitration_events",
    ):
        assert not hasattr(db, forbidden), forbidden

    match_id = _seed_pending_match(conn)
    first_id = record_arbitration_event(
        conn,
        SESSION_ID,
        match_id,
        "ivanov",
        "LINK_UPDATED",
        old_payload={"rd_item_id": None},
        new_payload={"rd_item_id": 1},
    )
    second_id = record_arbitration_event(
        conn,
        SESSION_ID,
        match_id,
        "petrova",
        "REVIEW_REJECTED",
        comment="Семантика",
    )

    events = get_arbitration_events(conn, SESSION_ID)
    assert [event["id"] for event in events] == [first_id, second_id]
    assert [event["action"] for event in events] == [
        "LINK_UPDATED",
        "REVIEW_REJECTED",
    ]
    # Первое событие не перезаписано вторым (нет UPDATE-семантики).
    assert json.loads(events[0]["new_payload"]) == {"rd_item_id": 1}
    assert get_arbitration_events(conn, SESSION_ID, match_id) == events
    assert get_arbitration_events(conn, "other-session") == []


def test_arbitration_event_payload_keeps_decimal_as_text(
    conn: sqlite3.Connection,
) -> None:
    """Decimal в payload журнала сохраняется строкой (Zero Float Policy)."""
    match_id = _seed_pending_match(conn)
    record_arbitration_event(
        conn,
        SESSION_ID,
        match_id,
        "ivanov",
        "LINK_UPDATED",
        old_payload={"conversion_ratio": Decimal("1.0000")},
        new_payload={"delta_volume": Decimal("-2.750")},
        comment="Пересчет",
    )

    event = get_arbitration_events(conn, SESSION_ID, match_id)[0]
    assert json.loads(event["old_payload"]) == {"conversion_ratio": "1.0000"}
    assert json.loads(event["new_payload"]) == {"delta_volume": "-2.750"}
    assert event["comment"] == "Пересчет"

    with pytest.raises(ValueError, match="actor"):
        record_arbitration_event(conn, SESSION_ID, match_id, "  ", "LINK_UPDATED")
    with pytest.raises(ValueError, match="action"):
        record_arbitration_event(conn, SESSION_ID, match_id, "ivanov", "  ")


def test_update_match_link_writes_arbitration_event(conn: sqlite3.Connection) -> None:
    """Уточнение связи и событие LINK_UPDATED фиксируются одной транзакцией."""
    match_id = _seed_pending_match(conn)

    update_match_link(
        conn,
        match_id,
        None,
        Decimal("0"),
        actor="petrova",
        comment="Отклонено вручную",
    )

    events = get_arbitration_events(conn, SESSION_ID, match_id)
    assert len(events) == 1
    event = events[0]
    assert event["action"] == "LINK_UPDATED"
    assert event["actor"] == "petrova"
    assert event["comment"] == "Отклонено вручную"
    assert json.loads(event["new_payload"]) == {
        "rd_item_id": None,
        "confidence_score": "0",
    }
    assert json.loads(event["old_payload"])["rd_item_id"] is not None

    match = conn.execute(
        "SELECT rd_item_id, confidence_score FROM match_records WHERE id = ?",
        (match_id,),
    ).fetchone()
    assert match["rd_item_id"] is None
    assert match["confidence_score"] == "0"


def test_update_match_link_missing_leaves_no_event(conn: sqlite3.Connection) -> None:
    """Отсутствующая связь не оставляет «висячего» события в журнале."""
    _seed_pending_match(conn)

    with pytest.raises(ValueError, match="не найдена"):
        update_match_link(conn, 4242, None, Decimal("0"))

    assert conn.execute(
        "SELECT COUNT(*) AS c FROM arbitration_events"
    ).fetchone()["c"] == 0


def test_record_expert_review_writes_event_and_rolls_back(
    conn: sqlite3.Connection,
) -> None:
    """Решение эксперта и событие REVIEW_* атомарны: отказ не оставляет следа."""
    match_id = _seed_pending_match(conn)

    review_id = record_expert_review(
        conn,
        match_id,
        "  ivanov  ",
        "APPROVED",
        comment="Подтверждено",
    )

    events = get_arbitration_events(conn, SESSION_ID, match_id)
    assert [event["action"] for event in events] == ["REVIEW_APPROVED"]
    assert events[0]["actor"] == "ivanov"
    payload = json.loads(events[0]["new_payload"])
    assert payload["decision"] == "APPROVED"
    assert payload["status"] == "RESOLVED"
    assert events[0]["comment"] == "Подтверждено"

    with pytest.raises(ValueError, match="Недопустимое решение"):
        record_expert_review(conn, match_id, "ivanov", "MAYBE")

    assert len(get_arbitration_events(conn, SESSION_ID, match_id)) == 1
    assert conn.execute(
        "SELECT COUNT(*) AS c FROM expert_reviews"
    ).fetchone()["c"] == 1
    review = conn.execute(
        "SELECT reviewer_name, decision FROM expert_reviews WHERE id = ?",
        (review_id,),
    ).fetchone()
    assert review["reviewer_name"] == "ivanov"
    assert review["decision"] == "APPROVED"


def test_users_are_stored_only_as_login_hash(conn: sqlite3.Connection) -> None:
    """В ``users`` нет сырого e-mail: только sha256-хэш и маска для UI."""
    assert normalize_login("  Ivanov@Company.RU ") == "ivanov@company.ru"
    digest = hash_login("Ivanov@Company.RU")
    assert digest == hash_login(" ivanov@company.ru ")
    assert len(digest) == 64
    assert "ivanov@company.ru" not in digest
    assert hash_login("ivanov@company.ru", salt="pepper") != digest
    assert mask_login("ivanov@company.ru") == "i***@company.ru"
    assert mask_login("") == ""

    with pytest.raises(ValueError, match="Логин"):
        hash_login("   ")

    create_user(
        conn,
        "user-1",
        digest,
        contact_mask=mask_login("ivanov@company.ru"),
        role="reviewer",
    )
    row = get_user_by_login_hash(conn, digest)
    assert row is not None
    assert row["id"] == "user-1"
    assert row["role"] == "reviewer"
    assert row["contact_mask"] == "i***@company.ru"
    assert get_user_by_login_hash(conn, hash_login("other@company.ru")) is None

    dump = " ".join(
        str(tuple(record)) for record in conn.execute("SELECT * FROM users")
    )
    assert "ivanov@company.ru" not in dump

    assert VALID_USER_ROLES == frozenset({"user", "reviewer", "admin"})
    with pytest.raises(ValueError, match="роль"):
        create_user(conn, "user-2", hash_login("x@y.ru"), role="root")
    with pytest.raises(ValueError, match="login_hash"):
        create_user(conn, "user-3", "   ")
    with pytest.raises(sqlite3.IntegrityError):
        create_user(conn, "user-4", digest)

    create_session(
        conn,
        "session-owned",
        "ЖК Север",
        "rd.pdf",
        "smeta.xlsx",
        user_id="user-1",
    )
    owner = conn.execute(
        "SELECT user_id FROM sessions WHERE id = 'session-owned'"
    ).fetchone()
    assert owner["user_id"] == "user-1"


def test_record_consent_152fz_persists_metadata(conn: sqlite3.Connection) -> None:
    """Согласие фиксируется с версией документа и метаданными клиента."""
    create_user(conn, "user-1", hash_login("ivanov@company.ru"))
    create_session(
        conn,
        "session-owned",
        "ЖК Север",
        "rd.pdf",
        "smeta.xlsx",
        user_id="user-1",
    )

    consent_id = record_consent_152fz(
        conn,
        "PDN_PROCESSING",
        "1.0",
        True,
        user_id="user-1",
        session_id="session-owned",
        client_meta={"attempt": 2, "delta": Decimal("2.5"), "channel": "cli"},
    )
    row = conn.execute(
        "SELECT * FROM consents_152fz WHERE id = ?",
        (consent_id,),
    ).fetchone()
    assert row["accepted"] == 1
    assert row["user_id"] == "user-1"
    assert row["session_id"] == "session-owned"
    assert row["doc_version"] == "1.0"
    meta = json.loads(row["client_meta"])
    assert meta["attempt"] == 2
    assert meta["delta"] == "2.5"
    assert meta["channel"] == "cli"

    declined_id = record_consent_152fz(conn, "OFFER", "2.1", False, user_id="user-1")
    declined = conn.execute(
        "SELECT accepted FROM consents_152fz WHERE id = ?",
        (declined_id,),
    ).fetchone()
    assert declined["accepted"] == 0

    with pytest.raises(ValueError, match="consent_type"):
        record_consent_152fz(conn, "  ", "1.0")
    with pytest.raises(ValueError, match="doc_version"):
        record_consent_152fz(conn, "PDN_PROCESSING", "  ")


def test_init_db_migrates_legacy_sessions_to_152fz(tmp_path) -> None:
    """Легаси-БД получает sessions.user_id и индекс idx_sessions_user."""
    connection = get_connection(tmp_path / "legacy_152fz.db")
    try:
        connection.execute(
            """
            CREATE TABLE sessions (
                id TEXT PRIMARY KEY,
                project_name TEXT NOT NULL,
                rd_filename TEXT NOT NULL,
                smeta_filename TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status TEXT DEFAULT 'CREATED'
            )
            """
        )
        connection.execute(
            """
            INSERT INTO sessions (id, project_name, rd_filename, smeta_filename)
            VALUES ('legacy', 'Старый проект', 'rd.pdf', 'smeta.xlsx')
            """
        )
        connection.commit()

        init_db(connection)
        init_db(connection)

        columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(sessions)")
        }
        assert "user_id" in columns
        assert "smeta_sheet" in columns

        indexes = {
            row["name"]
            for row in connection.execute("PRAGMA index_list(sessions)")
        }
        assert "idx_sessions_user" in indexes

        legacy = connection.execute(
            "SELECT user_id FROM sessions WHERE id = 'legacy'"
        ).fetchone()
        assert legacy["user_id"] is None

        create_user(connection, "user-1", hash_login("ivanov@company.ru"))
        create_session(
            connection,
            "session-owned",
            "Новый проект",
            "rd.pdf",
            "smeta.xlsx",
            user_id="user-1",
        )
        owned = connection.execute(
            "SELECT user_id FROM sessions WHERE id = 'session-owned'"
        ).fetchone()
        assert owned["user_id"] == "user-1"
    finally:
        connection.close()


