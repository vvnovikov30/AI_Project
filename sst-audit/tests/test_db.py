"""Тесты слоя данных SQLite (SSOT): схема, каскады, транзакции и CRUD."""

from collections.abc import Iterator
from decimal import Decimal
import sqlite3

import pytest

from db import (
    SCHEMA_SQL,
    backup_database,
    create_session,
    get_connection,
    get_pending_reviews,
    get_rule_evaluations,
    get_smeta_items_map,
    init_db,
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


