"""Слой данных SQLite как единственный источник истины (SSOT) для sst-audit.

Модуль реализован исключительно на стандартной библиотеке ``sqlite3``
(Public Domain) и не использует облачных API, XML/JSON-дампов или
проприетарных СУБД.

Принципы:
* Схема инициализируется идемпотентно (``CREATE ... IF NOT EXISTS``).
* Все операции записи выполняются в явных транзакциях с rollback при ошибке.
  Вложенные вызовы ``transaction()`` используют SAVEPOINT, поэтому пайплайн
  может атомарно объединить несколько CRUD-операций в одну транзакцию.
* Объемы, стоимости и веса правил хранятся в TEXT-колонках: SQLite не
  выполняет неявный float-кастинг 8-байтного IEEE 754, а точность
  ``Decimal`` сохраняется полностью (``str(Decimal)`` на записи,
  ``Decimal(str)`` на чтении). Расчетное ядро остается целиком на
  ``Decimal``.
"""

from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from contextlib import contextmanager
from datetime import datetime
from decimal import Decimal
from itertools import count
import json
import logging
from pathlib import Path
import re
import sqlite3
from typing import Any, Final, TypeAlias

from models import RULE_EVALUATION_STATUSES, RuleEvaluation, SmetaItem

logger = logging.getLogger(__name__)

# Детерминированный путь к SSOT: всегда каталог проекта, независимо от CWD.
DEFAULT_DB_PATH: Final[Path] = Path(__file__).resolve().parent / "sst_audit.db"
DEFAULT_CONVERSION_RATIO: Final[Decimal] = Decimal("1.0")

# Параметры горячего резервирования SSOT.
DEFAULT_BACKUP_DIRNAME: Final[str] = "backups"
DEFAULT_MAX_BACKUPS: Final[int] = 10
BACKUP_TIMESTAMP_FORMAT: Final[str] = "%Y%m%d_%H%M%S"
BACKUP_CONNECT_TIMEOUT_SECONDS: Final[int] = 30

VALID_SESSION_STATUSES: Final[frozenset[str]] = frozenset(
    {"CREATED", "MATCHED", "REVIEWED", "EXPORTED"}
)
VALID_MATCH_STATUSES: Final[frozenset[str]] = frozenset({"AUTO", "HITL_PENDING", "RESOLVED"})
VALID_REVIEW_DECISIONS: Final[frozenset[str]] = frozenset({"APPROVED", "REJECTED", "ADJUSTED"})
VALID_RULE_STATUSES: Final[frozenset[str]] = RULE_EVALUATION_STATUSES

# Счетчик уникальных имен SAVEPOINT для вложенных транзакций.
_SAVEPOINT_SEQ: Iterator[int] = count(1)

# Тип значения, которое sqlite3 способен принять без дополнительных адаптеров.
SQLNumber: TypeAlias = str | None

# Вход батча предикатов: типизированная модель либо legacy-словарь main.py.
RuleEvaluationInput: TypeAlias = RuleEvaluation | Mapping[str, Any]

SCHEMA_SQL: Final[str] = """
CREATE TABLE IF NOT EXISTS sessions (
    id TEXT PRIMARY KEY,
    project_name TEXT NOT NULL,
    rd_filename TEXT NOT NULL,
    smeta_filename TEXT NOT NULL,
    smeta_sheet TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    status TEXT CHECK (status IN ('CREATED', 'MATCHED', 'REVIEWED', 'EXPORTED')) DEFAULT 'CREATED'
);

CREATE TABLE IF NOT EXISTS rd_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
    row_idx INTEGER,
    code TEXT,
    description TEXT NOT NULL,
    unit TEXT NOT NULL,
    quantity TEXT NOT NULL,
    raw_payload JSON
);

CREATE TABLE IF NOT EXISTS smeta_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
    row_idx INTEGER,
    code TEXT NOT NULL,
    description TEXT NOT NULL,
    unit TEXT NOT NULL,
    quantity TEXT NOT NULL,
    unit_price TEXT NOT NULL,
    total_cost TEXT NOT NULL,
    parent_id INTEGER,
    raw_payload JSON
);

CREATE TABLE IF NOT EXISTS match_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
    rd_item_id INTEGER REFERENCES rd_items(id) ON DELETE CASCADE,
    smeta_item_id INTEGER REFERENCES smeta_items(id) ON DELETE CASCADE,
    confidence_score REAL NOT NULL,
    conversion_ratio REAL DEFAULT 1.0,
    match_method TEXT NOT NULL,
    status TEXT DEFAULT 'AUTO' CHECK (status IN ('AUTO', 'HITL_PENDING', 'RESOLVED'))
);

CREATE TABLE IF NOT EXISTS expert_reviews (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    match_record_id INTEGER NOT NULL REFERENCES match_records(id) ON DELETE CASCADE,
    reviewer_name TEXT NOT NULL,
    decision TEXT CHECK (decision IN ('APPROVED', 'REJECTED', 'ADJUSTED')),
    adjusted_ratio REAL,
    reject_reason_code TEXT,
    comment TEXT,
    reviewed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_rd_items_session ON rd_items(session_id);
CREATE INDEX IF NOT EXISTS idx_smeta_items_session ON smeta_items(session_id);
CREATE INDEX IF NOT EXISTS idx_match_records_session_status ON match_records(session_id, status);
CREATE INDEX IF NOT EXISTS idx_expert_reviews_match ON expert_reviews(match_record_id);

CREATE VIEW IF NOT EXISTS v_audit_export AS
SELECT
    m.session_id                       AS session_id,
    s.project_name                     AS project_name,
    m.id                               AS match_id,
    m.confidence_score                 AS confidence_score,
    m.conversion_ratio                 AS conversion_ratio,
    m.match_method                     AS match_method,
    m.status                           AS match_status,
    rd.id                              AS rd_item_id,
    rd.row_idx                         AS rd_row_idx,
    rd.code                            AS rd_code,
    rd.description                     AS rd_description,
    rd.unit                            AS rd_unit,
    rd.quantity                        AS rd_quantity,
    rd.raw_payload                     AS rd_raw_payload,
    sm.id                              AS smeta_item_id,
    sm.row_idx                         AS smeta_row_idx,
    sm.code                            AS smeta_code,
    sm.description                     AS smeta_description,
    sm.unit                            AS smeta_unit,
    sm.quantity                        AS smeta_quantity,
    sm.unit_price                      AS unit_price,
    sm.total_cost                      AS total_cost,
    sm.raw_payload                     AS smeta_raw_payload,
    er.id                              AS expert_review_id,
    er.reviewer_name                   AS reviewer_name,
    er.decision                        AS decision,
    er.adjusted_ratio                  AS adjusted_ratio,
    er.reject_reason_code              AS reject_reason_code,
    er.comment                         AS review_comment,
    er.reviewed_at                     AS reviewed_at
FROM match_records AS m
JOIN sessions AS s ON s.id = m.session_id
LEFT JOIN rd_items AS rd ON rd.id = m.rd_item_id
LEFT JOIN smeta_items AS sm ON sm.id = m.smeta_item_id
LEFT JOIN expert_reviews AS er ON er.match_record_id = m.id;
"""

PENDING_REVIEWS_SQL: Final[str] = """
SELECT
    m.id               AS match_id,
    m.session_id       AS session_id,
    m.confidence_score AS confidence_score,
    m.conversion_ratio AS conversion_ratio,
    m.match_method     AS match_method,
    m.status           AS status,
    rd.id              AS rd_item_id,
    rd.row_idx         AS rd_row_idx,
    rd.code            AS rd_code,
    rd.description     AS rd_description,
    rd.unit            AS rd_unit,
    rd.quantity        AS rd_quantity,
    sm.id              AS smeta_item_id,
    sm.row_idx         AS smeta_row_idx,
    sm.code            AS smeta_code,
    sm.description     AS smeta_description,
    sm.unit            AS smeta_unit,
    sm.quantity        AS smeta_quantity,
    sm.unit_price      AS unit_price,
    sm.total_cost      AS total_cost
FROM match_records AS m
LEFT JOIN rd_items AS rd ON rd.id = m.rd_item_id
LEFT JOIN smeta_items AS sm ON sm.id = m.smeta_item_id
WHERE m.session_id = ? AND m.status = 'HITL_PENDING'
ORDER BY m.id ASC
"""


# DDL предикатов аудита вынесен из общего SCHEMA_SQL: на устаревших БД
# таблица сначала пересобирается миграцией, и только затем создаются индексы.
RULE_EVALUATIONS_TABLE_SQL: Final[str] = """
CREATE TABLE IF NOT EXISTS rule_evaluations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
    match_record_id INTEGER REFERENCES match_records(id) ON DELETE CASCADE,
    rule_code TEXT NOT NULL,
    status TEXT NOT NULL,
    details TEXT,
    score_weight TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
"""

RULE_EVALUATIONS_SQL: Final[str] = (
    RULE_EVALUATIONS_TABLE_SQL
    + """
CREATE INDEX IF NOT EXISTS idx_rule_evaluations_match ON rule_evaluations(match_record_id);
CREATE INDEX IF NOT EXISTS idx_rule_eval_lookup ON rule_evaluations(session_id, match_record_id);
"""
)

# Индексы, зависящие от колонок, которые добавляются миграциями.
POST_MIGRATION_INDEXES_SQL: Final[str] = """
CREATE INDEX IF NOT EXISTS idx_smeta_items_parent ON smeta_items(session_id, parent_id);
CREATE INDEX IF NOT EXISTS idx_rd_items_session ON rd_items(session_id);
CREATE INDEX IF NOT EXISTS idx_smeta_items_session ON smeta_items(session_id);
"""

# Числовые колонки, которые обязаны иметь TEXT-аффинити: REAL в SQLite
# приводит Decimal к IEEE 754 float уже на записи.
NUMERIC_TEXT_COLUMNS: Final[dict[str, tuple[str, ...]]] = {
    "rd_items": ("quantity",),
    "smeta_items": ("quantity", "unit_price", "total_cost"),
    "rule_evaluations": ("score_weight",),
}


# ----------------------------------------------------------------------
# 1. Соединения, транзакции и инициализация схемы
# ----------------------------------------------------------------------


def get_connection(db_path: str | Path = DEFAULT_DB_PATH) -> sqlite3.Connection:
    """Открывает соединение SQLite с доступом к колонкам по имени.

    :param db_path: путь к файлу БД либо ``":memory:"`` для тестового контура.
    :return: соединение с ``row_factory = sqlite3.Row`` и включенными FK.
    """
    connection = sqlite3.connect(str(db_path), timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


@contextmanager
def transaction(conn: sqlite3.Connection) -> Iterator[sqlite3.Connection]:
    """Контекстный менеджер транзакции: commit при успехе, rollback при ошибке.

    Если транзакция уже открыта (например, пайплайн объединяет несколько
    CRUD-операций), создается SAVEPOINT: внутренний блок может откатиться
    изолированно, а фиксация происходит на внешнем уровне.
    """
    if conn.in_transaction:
        savepoint = f"sst_savepoint_{next(_SAVEPOINT_SEQ)}"
        conn.execute(f"SAVEPOINT {savepoint}")
        try:
            yield conn
        except BaseException:
            conn.execute(f"ROLLBACK TO SAVEPOINT {savepoint}")
            conn.execute(f"RELEASE SAVEPOINT {savepoint}")
            raise
        else:
            conn.execute(f"RELEASE SAVEPOINT {savepoint}")
        return

    # Явный BEGIN обязателен: sqlite3 в legacy-режиме не открывает транзакцию
    # до первого DML, поэтому conn.in_transaction остается False и вложенные
    # вызовы transaction() не увидели бы внешнюю транзакцию.
    conn.execute("BEGIN")
    try:
        yield conn
        conn.commit()
    except BaseException:
        conn.rollback()
        raise


def _table_columns(conn: sqlite3.Connection, table: str) -> set[str]:
    """Возвращает набор имен колонок существующей таблицы (пусто, если нет)."""
    return {str(row["name"]) for row in conn.execute(f"PRAGMA table_info({table})")}


def _migrate_sessions_schema(conn: sqlite3.Connection) -> None:
    """Добавляет новые колонки в уже существующую БД без потери данных."""
    if "smeta_sheet" not in _table_columns(conn, "sessions"):
        conn.execute("ALTER TABLE sessions ADD COLUMN smeta_sheet TEXT")


def _migrate_smeta_items_schema(conn: sqlite3.Connection) -> None:
    """Добавляет колонку иерархии ``parent_id`` в устаревшие сметы."""
    columns = _table_columns(conn, "smeta_items")
    if columns and "parent_id" not in columns:
        conn.execute("ALTER TABLE smeta_items ADD COLUMN parent_id INTEGER")


def _migrate_rule_evaluations_schema(conn: sqlite3.Connection) -> None:
    """Пересобирает legacy-таблицу предикатов под схему ``rule_code``/``status``.

    Старая схема (``rule_name``/``is_violation``/``delta_volume``) не совместима
    с моделью ``RuleEvaluation``. Данные переносятся в новую таблицу с
    привязкой к сессии через ``match_records``; ``score_weight`` остается
    ``NULL`` (вес правил на момент старой схемы не рассчитывался).
    """
    columns = _table_columns(conn, "rule_evaluations")
    if not columns or {"session_id", "rule_code", "status", "score_weight"} <= columns:
        return

    legacy_table = "rule_evaluations_legacy"
    conn.execute(f"DROP TABLE IF EXISTS {legacy_table}")
    conn.execute(f"ALTER TABLE rule_evaluations RENAME TO {legacy_table}")
    conn.executescript(RULE_EVALUATIONS_TABLE_SQL)

    legacy_columns = _table_columns(conn, legacy_table)
    if {"match_record_id", "rule_name", "is_violation"} <= legacy_columns:
        conn.execute(
            """
            INSERT INTO rule_evaluations
                (session_id, match_record_id, rule_code, status, details, score_weight)
            SELECT
                m.session_id,
                legacy.match_record_id,
                legacy.rule_name,
                CASE WHEN legacy.is_violation THEN 'FAILED' ELSE 'PASSED' END,
                legacy.message,
                NULL
            FROM rule_evaluations_legacy AS legacy
            JOIN match_records AS m ON m.id = legacy.match_record_id
            """
        )

    conn.execute(f"DROP TABLE {legacy_table}")


def _needs_text_migration(
    conn: sqlite3.Connection,
    table: str,
    numeric_columns: tuple[str, ...],
) -> bool:
    """Проверяет, объявлена ли хотя бы одна числовая колонка не как TEXT."""
    declared = {
        str(row["name"]): str(row["type"] or "").upper()
        for row in conn.execute(f"PRAGMA table_info({table})")
    }
    return any(
        declared.get(column, "TEXT") != "TEXT" for column in numeric_columns
    )


def _rebuild_numeric_columns_as_text(
    conn: sqlite3.Connection,
    table: str,
    numeric_columns: tuple[str, ...],
) -> None:
    """Пересобирает таблицу с TEXT-аффинити, сохраняя все данные и колонки.

    REAL-значения переводятся в текст через ``CAST(... AS TEXT)``: исходная
    точность уже потеряна при записи во float, но дальнейшие записи идут
    только строкой. Схема берется из ``sqlite_master``, поэтому миграция
    сохраняет legacy-колонки и ограничения.
    """
    source = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = ?",
        (table,),
    ).fetchone()
    if source is None or not source["sql"]:
        return

    new_table = f"{table}__text_migration"
    create_sql = str(source["sql"])
    for column in numeric_columns:
        create_sql = re.sub(
            rf"\b{re.escape(column)}\s+REAL\b",
            f"{column} TEXT",
            create_sql,
            flags=re.IGNORECASE,
        )

    renamed_sql, replacements = re.subn(
        rf"^CREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?[\"']?{re.escape(table)}[\"']?",
        f'CREATE TABLE "{new_table}"',
        create_sql,
        count=1,
        flags=re.IGNORECASE,
    )
    if replacements != 1:
        raise RuntimeError(
            f"Не удалось разобрать DDL таблицы {table}: {create_sql[:120]!r}"
        )

    column_names = [
        str(row["name"]) for row in conn.execute(f"PRAGMA table_info({table})")
    ]
    quoted_columns = ", ".join(f'"{name}"' for name in column_names)
    select_columns = ", ".join(
        f'CAST("{name}" AS TEXT)' if name in numeric_columns else f'"{name}"'
        for name in column_names
    )

    conn.execute(f'DROP TABLE IF EXISTS "{new_table}"')
    conn.execute(renamed_sql)
    conn.execute(
        f'INSERT INTO "{new_table}" ({quoted_columns}) '
        f'SELECT {select_columns} FROM "{table}"'
    )
    conn.execute(f'DROP TABLE "{table}"')
    conn.execute(f'ALTER TABLE "{new_table}" RENAME TO "{table}"')


def _migrate_numeric_columns_to_text(conn: sqlite3.Connection) -> None:
    """Идемпотентно переводит числовые колонки устаревшей БД в TEXT.

    Пересборка выполняется при отключенных FK и в одной транзакции:
    ``match_records`` ссылается на ``rd_items``/``smeta_items``, поэтому
    каскадное удаление при ``DROP TABLE`` недопустимо.
    """
    pending = [
        (table, columns)
        for table, columns in NUMERIC_TEXT_COLUMNS.items()
        if _needs_text_migration(conn, table, columns)
    ]
    if not pending:
        return

    # Представление ссылается на пересобираемые таблицы: на время DDL его
    # нужно снять, иначе SQLite проверяет схему при RENAME и падает.
    view = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'view' AND name = 'v_audit_export'"
    ).fetchone()

    conn.commit()
    conn.execute("PRAGMA foreign_keys = OFF")
    try:
        conn.execute("BEGIN")
        try:
            conn.execute("DROP VIEW IF EXISTS v_audit_export")
            for table, columns in pending:
                _rebuild_numeric_columns_as_text(conn, table, columns)
            if view is not None and view["sql"]:
                conn.execute(str(view["sql"]))
        except BaseException:
            conn.rollback()
            raise
        else:
            conn.commit()
    finally:
        conn.execute("PRAGMA foreign_keys = ON")


def init_db(conn: sqlite3.Connection) -> None:
    """Идемпотентно создает таблицы, индексы и представление v_audit_export."""
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.executescript(SCHEMA_SQL)
    _migrate_sessions_schema(conn)
    _migrate_smeta_items_schema(conn)
    _migrate_rule_evaluations_schema(conn)
    _migrate_numeric_columns_to_text(conn)
    # Индексы по колонкам, которые могли быть добавлены миграциями.
    conn.executescript(RULE_EVALUATIONS_SQL)
    conn.executescript(POST_MIGRATION_INDEXES_SQL)
    conn.commit()


# ----------------------------------------------------------------------
# 2. Внутренние помощники границы персистентности
# ----------------------------------------------------------------------


def _to_sql_number(value: object) -> SQLNumber:
    """Приводит значение ядра к TEXT-представлению для sqlite3.

    SQLite не имеет встроенного адаптера ``Decimal``, поэтому конвертация
    выполняется только здесь — на границе записи. Любое непустое значение
    сериализуется в ``str``: TEXT-аффинити колонки сохраняют точную строку
    без промежуточного Python-float и связанной с ним потери точности.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return str(int(value))
    if isinstance(value, Decimal):
        return str(value)
    return str(value)


def _to_decimal(value: object, default: Decimal = Decimal("0")) -> Decimal:
    """Приводит значение из SQLite (TEXT) обратно к ``Decimal``."""
    if value is None:
        return default
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        return Decimal(int(value))
    return Decimal(str(value).strip())


def _to_optional_int(value: object) -> int | None:
    """Приводит ``parent_id``/идентификатор к ``int`` либо ``None``."""
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError(f"Ожидалось целое число, получено bool: {value!r}")
    if isinstance(value, int):
        return value
    text = str(value).strip()
    if not text:
        return None
    try:
        return int(text)
    except ValueError as exc:
        raise ValueError(f"Ожидалось целое число: {value!r}") from exc


def _deserialize_payload(payload: object) -> dict[str, Any] | None:
    """Разбирает JSON-колонку ``raw_payload`` в словарь аудиторского следа."""
    if payload is None:
        return None
    if isinstance(payload, Mapping):
        return dict(payload)
    text = str(payload)
    try:
        decoded = json.loads(text)
    except (TypeError, ValueError):
        return {"raw": text}
    if isinstance(decoded, Mapping):
        return dict(decoded)
    return {"value": decoded}


def _serialize_payload(payload: object) -> str | None:
    """Сериализует raw_payload в JSON-текст для колонки типа JSON."""
    if payload is None:
        return None
    if isinstance(payload, str):
        return payload
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)


def _require(item: Mapping[str, Any], field_name: str) -> Any:
    """Возвращает обязательное поле записи либо поднимает ValueError."""
    if field_name not in item:
        raise ValueError(
            f"Отсутствует обязательное поле '{field_name}' в записи: {dict(item)!r}"
        )
    return item[field_name]


# ----------------------------------------------------------------------
# 3. CRUD-операции SSOT
# ----------------------------------------------------------------------


def create_session(
    conn: sqlite3.Connection,
    session_id: str,
    project_name: str,
    rd_file: str,
    smeta_file: str,
    smeta_sheet: str | None = None,
) -> str:
    """Регистрирует сессию аудита и возвращает ее идентификатор.

    :param smeta_sheet: имя листа ``ЛС ...``, если сессия относится к одному
        разделу многолистовой сметы (пакетный ночной режим).
    """
    with transaction(conn):
        conn.execute(
            """
            INSERT INTO sessions
                (id, project_name, rd_filename, smeta_filename, smeta_sheet)
            VALUES (?, ?, ?, ?, ?)
            """,
            (session_id, project_name, rd_file, smeta_file, smeta_sheet),
        )
    return session_id


def save_rd_items(
    conn: sqlite3.Connection,
    session_id: str,
    items: Sequence[Mapping[str, Any]],
) -> int:
    """Сохраняет позиции рабочей документации (RD/спецификации)."""
    if not items:
        return 0

    rows = [
        (
            session_id,
            item.get("row_idx"),
            item.get("code"),
            _require(item, "description"),
            _require(item, "unit"),
            _to_sql_number(_require(item, "quantity")),
            _serialize_payload(item.get("raw_payload")),
        )
        for item in items
    ]

    with transaction(conn):
        conn.executemany(
            """
            INSERT INTO rd_items
                (session_id, row_idx, code, description, unit, quantity, raw_payload)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            rows,
        )
    return len(rows)


def save_smeta_items(
    conn: sqlite3.Connection,
    session_id: str,
    items: Sequence[Mapping[str, Any]],
) -> int:
    """Сохраняет позиции сметы. Возвращает количество вставленных строк."""
    if not items:
        return 0

    rows = [
        (
            session_id,
            item.get("row_idx"),
            _require(item, "code"),
            _require(item, "description"),
            _require(item, "unit"),
            _to_sql_number(_require(item, "quantity")),
            _to_sql_number(_require(item, "unit_price")),
            _to_sql_number(_require(item, "total_cost")),
            _to_optional_int(item.get("parent_id")),
            _serialize_payload(item.get("raw_payload")),
        )
        for item in items
    ]

    with transaction(conn):
        conn.executemany(
            """
            INSERT INTO smeta_items
                (session_id, row_idx, code, description, unit,
                 quantity, unit_price, total_cost, parent_id, raw_payload)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            rows,
        )
    return len(rows)


def save_matches(
    conn: sqlite3.Connection,
    session_id: str,
    matches: Sequence[Mapping[str, Any]],
) -> int:
    """Сохраняет результаты сопоставления. Возвращает количество вставленных строк."""
    if not matches:
        return 0

    rows = []
    for match in matches:
        status = str(match.get("status", "AUTO"))
        if status not in VALID_MATCH_STATUSES:
            raise ValueError(
                f"Недопустимый статус match_record '{status}'. "
                f"Разрешены: {sorted(VALID_MATCH_STATUSES)}"
            )
        rows.append(
            (
                session_id,
                match.get("rd_item_id"),
                match.get("smeta_item_id"),
                _to_sql_number(_require(match, "confidence_score")),
                _to_sql_number(match.get("conversion_ratio", DEFAULT_CONVERSION_RATIO)),
                _require(match, "match_method"),
                status,
            )
        )

    with transaction(conn):
        conn.executemany(
            """
            INSERT INTO match_records
                (session_id, rd_item_id, smeta_item_id, confidence_score,
                 conversion_ratio, match_method, status)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            rows,
        )
    return len(rows)


def get_pending_reviews(
    conn: sqlite3.Connection,
    session_id: str,
) -> list[sqlite3.Row]:
    """Возвращает связи сессии, ожидающие решения эксперта (HITL_PENDING)."""
    cursor = conn.execute(PENDING_REVIEWS_SQL, (session_id,))
    return list(cursor.fetchall())


def get_match_records(
    conn: sqlite3.Connection,
    session_id: str,
) -> list[sqlite3.Row]:
    """Возвращает все связи сессии в порядке вставки (id ASC).

    Используется оркестратором для связывания решений матчера с
    идентификаторами ``match_records`` перед оценкой правил.
    """
    cursor = conn.execute(
        "SELECT * FROM match_records WHERE session_id = ? ORDER BY id ASC",
        (session_id,),
    )
    return list(cursor.fetchall())


def _load_session_ids_for_matches(
    conn: sqlite3.Connection,
    match_ids: Sequence[int | None],
) -> dict[int, str]:
    """Разрешает ``match_record_id -> session_id`` одним запросом (без N+1)."""
    unique_ids = sorted({int(match_id) for match_id in match_ids if match_id is not None})
    if not unique_ids:
        return {}
    placeholders = ", ".join("?" for _ in unique_ids)
    rows = conn.execute(
        f"SELECT id, session_id FROM match_records WHERE id IN ({placeholders})",
        unique_ids,
    )
    return {int(row["id"]): str(row["session_id"]) for row in rows}


def _legacy_rule_details(evaluation: Mapping[str, Any]) -> str | None:
    """Формирует ``details`` для legacy-словаря main.py.

    Если переданы числовые дельты (``delta_volume``/``delta_cost``), они
    сохраняются в JSON строковыми значениями, чтобы не потерять точность
    ``Decimal`` в TEXT-колонке.
    """
    details = evaluation.get("details")
    if details is not None:
        return str(details)

    message = evaluation.get("message")
    delta_volume = evaluation.get("delta_volume")
    delta_cost = evaluation.get("delta_cost")
    if delta_volume is None and delta_cost is None:
        return None if message is None else str(message)

    payload: dict[str, str] = {}
    if message is not None:
        payload["message"] = str(message)
    if delta_volume is not None:
        payload["delta_volume"] = str(_to_decimal(delta_volume))
    if delta_cost is not None:
        payload["delta_cost"] = str(_to_decimal(delta_cost))
    return json.dumps(payload, ensure_ascii=False, sort_keys=True)



def _coerce_rule_evaluation(
    evaluation: RuleEvaluationInput,
    session_by_match: Mapping[int, str],
) -> RuleEvaluation:
    """Приводит вход батча (модель либо legacy-словарь) к ``RuleEvaluation``."""
    if isinstance(evaluation, RuleEvaluation):
        if evaluation.session_id.strip():
            return evaluation
        if evaluation.match_record_id is None:
            raise ValueError(
                "Не удалось определить session_id: укажите session_id или match_record_id"
            )
        session_id = session_by_match.get(int(evaluation.match_record_id))
        if session_id is None:
            raise ValueError(
                f"Не найдена сессия для match_record_id={evaluation.match_record_id}"
            )
        return RuleEvaluation(
            session_id=session_id,
            rule_code=evaluation.rule_code,
            status=evaluation.status,
            id=evaluation.id,
            match_record_id=evaluation.match_record_id,
            details=evaluation.details,
            score_weight=evaluation.score_weight,
            created_at=evaluation.created_at,
        )

    match_record_id = _to_optional_int(evaluation.get("match_record_id"))

    rule_code = evaluation.get("rule_code") or evaluation.get("rule_name")
    if not isinstance(rule_code, str) or not rule_code.strip():
        raise ValueError(
            "Поле 'rule_code' (legacy: 'rule_name') должно быть "
            f"непустой строкой: {rule_code!r}"
        )

    raw_status = evaluation.get("status")
    if raw_status is None:
        is_violation = evaluation.get("is_violation", False)
        if not isinstance(is_violation, (bool, int)):
            raise ValueError(
                f"Поле 'is_violation' должно быть bool/int: {is_violation!r}"
            )
        raw_status = "FAILED" if bool(is_violation) else "PASSED"
    status = str(raw_status)
    if status not in VALID_RULE_STATUSES:
        raise ValueError(
            f"Недопустимый статус правила '{status}'. "
            f"Разрешены: {sorted(VALID_RULE_STATUSES)}"
        )

    session_id = evaluation.get("session_id")
    if session_id is None:
        if match_record_id is None:
            raise ValueError(
                "Не удалось определить session_id: укажите session_id или match_record_id"
            )
        session_id = session_by_match.get(match_record_id)
        if session_id is None:
            raise ValueError(
                f"Не найдена сессия для match_record_id={match_record_id}"
            )
    session_id_text = str(session_id).strip()

    raw_weight = evaluation.get("score_weight", Decimal("0"))
    if isinstance(raw_weight, bool) or raw_weight is None:
        raise TypeError("score_weight должен быть Decimal")
    if isinstance(raw_weight, Decimal):
        score_weight = raw_weight
    elif isinstance(raw_weight, (int, str)):
        score_weight = _to_decimal(raw_weight)
    else:
        raise TypeError(
            "score_weight должен быть Decimal: иные числовые типы запрещены"
        )

    created_at = evaluation.get("created_at")
    return RuleEvaluation(
        session_id=session_id_text,
        rule_code=rule_code.strip(),
        status=status,
        match_record_id=match_record_id,
        details=_legacy_rule_details(evaluation),
        score_weight=score_weight,
        created_at=None if created_at is None else str(created_at),
    )


def save_rule_evaluations(
    conn: sqlite3.Connection,
    evaluations: Sequence[RuleEvaluationInput],
) -> int:
    """Батчево сохраняет предикаты аудита в ``rule_evaluations``.

    :param evaluations: модели ``RuleEvaluation`` (основной путь) либо
        legacy-словари main.py с ключами ``rule_name``/``is_violation``/
        ``message``/``delta_volume``/``delta_cost``.
    :return: количество вставленных строк.
    :raises ValueError: некорректный код правила/статус/неразрешимая сессия.
    :raises sqlite3.IntegrityError: нарушение FK (транзакция откатывается).

    ``score_weight`` записывается как ``str(Decimal)``, поэтому точность
    сохраняется полностью (колонка TEXT).
    """
    if not evaluations:
        return 0

    match_ids = [
        item.match_record_id
        if isinstance(item, RuleEvaluation)
        else _to_optional_int(item.get("match_record_id"))
        for item in evaluations
    ]
    session_by_match = _load_session_ids_for_matches(conn, match_ids)

    rows: list[
        tuple[str, int | None, str, str, str | None, str, str | None]
    ] = []
    for evaluation in evaluations:
        normalized = _coerce_rule_evaluation(evaluation, session_by_match)
        rows.append(
            (
                normalized.session_id,
                normalized.match_record_id,
                normalized.rule_code,
                normalized.status,
                normalized.details,
                str(normalized.score_weight),
                normalized.created_at,
            )
        )

    with transaction(conn):
        conn.executemany(
            """
            INSERT INTO rule_evaluations
                (session_id, match_record_id, rule_code, status, details,
                 score_weight, created_at)
            VALUES (?, ?, ?, ?, ?, ?, COALESCE(?, CURRENT_TIMESTAMP))
            """,
            rows,
        )
    return len(rows)


def get_rule_evaluations(
    conn: sqlite3.Connection,
    session_id: str,
    match_record_id: int | None = None,
) -> list[RuleEvaluation]:
    """Читает предикаты сессии (опционально — только по одной связи).

    ``score_weight`` конвертируется обратно в ``Decimal`` без промежуточного
    float: колонка хранит точную строку ``str(Decimal)``.
    """
    sql = (
        "SELECT id, session_id, match_record_id, rule_code, status, details, "
        "score_weight, created_at FROM rule_evaluations WHERE session_id = ?"
    )
    params: list[Any] = [session_id]
    if match_record_id is not None:
        sql += " AND match_record_id = ?"
        params.append(match_record_id)
    sql += " ORDER BY id ASC"

    evaluations: list[RuleEvaluation] = []
    for row in conn.execute(sql, params).fetchall():
        raw_weight = row["score_weight"]
        evaluations.append(
            RuleEvaluation(
                session_id=str(row["session_id"]),
                rule_code=str(row["rule_code"]),
                status=str(row["status"]),
                id=int(row["id"]),
                match_record_id=(
                    None
                    if row["match_record_id"] is None
                    else int(row["match_record_id"])
                ),
                details=None if row["details"] is None else str(row["details"]),
                score_weight=(
                    Decimal("0.0") if raw_weight is None else _to_decimal(raw_weight)
                ),
                created_at=(
                    None if row["created_at"] is None else str(row["created_at"])
                ),
            )
        )
    return evaluations


def get_smeta_items_map(
    conn: sqlite3.Connection,
    session_id: str,
) -> dict[int, SmetaItem]:
    """Загружает все сметные строки сессии в словарь ``{item.id: item}``.

    Словарь позволяет матчеру получать родительскую работу по ``item.parent_id``
    за O(1) без N+1 запросов.
    """
    cursor = conn.execute(
        """
        SELECT id, session_id, row_idx, code, description, unit,
               quantity, unit_price, total_cost, parent_id, raw_payload
        FROM smeta_items
        WHERE session_id = ?
        ORDER BY id ASC
        """,
        (session_id,),
    )

    items: dict[int, SmetaItem] = {}
    for row in cursor.fetchall():
        item_id = int(row["id"])
        items[item_id] = SmetaItem(
            id=item_id,
            session_id=str(row["session_id"]),
            code=str(row["code"]),
            description=str(row["description"]),
            unit=str(row["unit"]),
            quantity=_to_decimal(row["quantity"]),
            unit_price=_to_decimal(row["unit_price"]),
            total_cost=_to_decimal(row["total_cost"]),
            row_idx=None if row["row_idx"] is None else int(row["row_idx"]),
            parent_id=None if row["parent_id"] is None else int(row["parent_id"]),
            raw_payload=_deserialize_payload(row["raw_payload"]),
        )
    return items


def update_session_status(
    conn: sqlite3.Connection,
    session_id: str,
    status: str,
) -> None:
    """Переводит сессию аудита в новый статус.

    :raises ValueError: неизвестный статус либо отсутствующая сессия.
    """
    if status not in VALID_SESSION_STATUSES:
        raise ValueError(
            f"Недопустимый статус сессии '{status}'. "
            f"Разрешены: {sorted(VALID_SESSION_STATUSES)}"
        )

    with transaction(conn):
        cursor = conn.execute(
            "UPDATE sessions SET status = ? WHERE id = ?",
            (status, session_id),
        )

    if cursor.rowcount == 0:
        raise ValueError(f"Сессия '{session_id}' не найдена")


def update_match_link(
    conn: sqlite3.Connection,
    match_id: int,
    rd_item_id: int | None,
    confidence_score: Decimal | int | str,
) -> None:
    """Уточняет связь по решению эксперта (кандидат РД, выбранный в HITL).

    :param rd_item_id: идентификатор выбранной позиции РД либо ``None``,
        если связь отклонена.
    :raises ValueError: связь с указанным id не найдена.
    """
    with transaction(conn):
        cursor = conn.execute(
            """
            UPDATE match_records
            SET rd_item_id = ?, confidence_score = ?
            WHERE id = ?
            """,
            (
                rd_item_id,
                _to_sql_number(confidence_score),
                match_id,
            ),
        )

    if cursor.rowcount == 0:
        raise ValueError(f"Связь '{match_id}' не найдена")


def record_expert_review(
    conn: sqlite3.Connection,
    match_id: int,
    reviewer: str,
    decision: str,
    reason_code: str | None = None,
    comment: str | None = None,
    adjusted_ratio: Decimal | int | str | None = None,
) -> int:
    """Фиксирует экспертизу связи и переводит ее в статус RESOLVED.

    :return: идентификатор созданной записи ``expert_reviews``.
    """
    if decision not in VALID_REVIEW_DECISIONS:
        raise ValueError(
            f"Недопустимое решение '{decision}'. "
            f"Разрешены: {sorted(VALID_REVIEW_DECISIONS)}"
        )
    if not reviewer or not reviewer.strip():
        raise ValueError("reviewer_name не может быть пустым")

    with transaction(conn):
        cursor = conn.execute(
            """
            INSERT INTO expert_reviews
                (match_record_id, reviewer_name, decision,
                 adjusted_ratio, reject_reason_code, comment)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                match_id,
                reviewer.strip(),
                decision,
                _to_sql_number(adjusted_ratio),
                reason_code,
                comment,
            ),
        )
        review_id = int(cursor.lastrowid)
        conn.execute(
            "UPDATE match_records SET status = 'RESOLVED' WHERE id = ?",
            (match_id,),
        )
    return review_id


# ----------------------------------------------------------------------
# 4. Горячее резервирование SSOT (WAL-safe snapshot)
# ----------------------------------------------------------------------


def _rotate_backups(target_dir: Path, stem: str, max_backups: int) -> None:
    """Удаляет устаревшие снимки, оставляя ``max_backups`` самых свежих.

    Сортировка выполняется по имени файла: формат ``YYYYMMDD_HHMMSS``
    лексикографически совпадает с хронологией, поэтому ``st_mtime``
    (float-метка времени) не используется.
    """
    retention = max(max_backups, 1)
    snapshots = sorted(target_dir.glob(f"{stem}_*.db"), reverse=True)
    for stale in snapshots[retention:]:
        stale.unlink(missing_ok=True)


def backup_database(
    db_path: str | Path | None = None,
    backup_dir: str | Path | None = None,
    max_backups: int = DEFAULT_MAX_BACKUPS,
) -> Path | None:
    """Создает целостный горячий снимок рабочей БД SQLite.

    Копирование выполняется через ``sqlite3.Connection.backup``: это
    единственный WAL-safe способ снять снимок «на лету» (учитывает
    незафиксированные в основном файле страницы ``-wal``/``-shm``).

    :param db_path: путь к SSOT; ``None`` — ``DEFAULT_DB_PATH`` проекта.
    :param backup_dir: каталог снимков; по умолчанию ``<db_dir>/backups``.
    :param max_backups: сколько последних снимков хранить (ротация).
    :return: путь к созданному снимку либо ``None``, если база пуста/нет
        доступа (первый чистый запуск либо ошибка резервирования).
    """
    source_path = Path(DEFAULT_DB_PATH if db_path is None else db_path).resolve()
    try:
        if not source_path.is_file() or source_path.stat().st_size == 0:
            return None

        target_dir = (
            Path(backup_dir).resolve()
            if backup_dir is not None
            else source_path.parent / DEFAULT_BACKUP_DIRNAME
        )
        target_dir.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime(BACKUP_TIMESTAMP_FORMAT)
        backup_file = target_dir / f"{source_path.stem}_{timestamp}.db"

        src_conn: sqlite3.Connection | None = None
        dst_conn: sqlite3.Connection | None = None
        try:
            src_conn = sqlite3.connect(str(source_path), timeout=BACKUP_CONNECT_TIMEOUT_SECONDS)
            dst_conn = sqlite3.connect(str(backup_file), timeout=BACKUP_CONNECT_TIMEOUT_SECONDS)
            # pages=0 — скопировать все страницы за один вызов backup API.
            src_conn.backup(dst_conn, pages=0)
        finally:
            if dst_conn is not None:
                dst_conn.close()
            if src_conn is not None:
                src_conn.close()

        _rotate_backups(target_dir, source_path.stem, max_backups)
        return backup_file
    except Exception as exc:  # noqa: BLE001 — резервирование не должно ронять пайплайн
        logger.warning("Не удалось создать резервную копию БД %s: %s", source_path, exc)
        return None
