"""Слой данных SQLite как единственный источник истины (SSOT) для sst-audit.

Модуль реализован исключительно на стандартной библиотеке ``sqlite3``
(Public Domain) и не использует облачных API, XML/JSON-дампов или
проприетарных СУБД.

Принципы:
* Схема инициализируется идемпотентно (``CREATE ... IF NOT EXISTS``).
* Все операции записи выполняются в явных транзакциях с rollback при ошибке.
  Вложенные вызовы ``transaction()`` используют SAVEPOINT, поэтому пайплайн
  может атомарно объединить несколько CRUD-операций в одну транзакцию.
* Объемы и стоимости хранятся в числовых колонках SQLite (REAL) для
  совместимости с отчетными/BI-инструментами. На входе CRUD-методов
  допускается ``Decimal``: конвертация выполняется только на границе
  персистентности, расчетное ядро остается целиком на ``Decimal``.
"""

from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from contextlib import contextmanager
from decimal import Decimal
from itertools import count
import json
from pathlib import Path
import sqlite3
from typing import Any, Final, TypeAlias

DEFAULT_DB_PATH: Final[str] = "sst_audit.db"
DEFAULT_CONVERSION_RATIO: Final[Decimal] = Decimal("1.0")

VALID_SESSION_STATUSES: Final[frozenset[str]] = frozenset(
    {"CREATED", "MATCHED", "REVIEWED", "EXPORTED"}
)
VALID_MATCH_STATUSES: Final[frozenset[str]] = frozenset({"AUTO", "HITL_PENDING", "RESOLVED"})
VALID_REVIEW_DECISIONS: Final[frozenset[str]] = frozenset({"APPROVED", "REJECTED", "ADJUSTED"})

# Счетчик уникальных имен SAVEPOINT для вложенных транзакций.
_SAVEPOINT_SEQ: Iterator[int] = count(1)

# Тип значения, которое sqlite3 способен принять без дополнительных адаптеров.
SQLNumber: TypeAlias = int | float | str | None

SCHEMA_SQL: Final[str] = """
CREATE TABLE IF NOT EXISTS sessions (
    id TEXT PRIMARY KEY,
    project_name TEXT NOT NULL,
    rd_filename TEXT NOT NULL,
    smeta_filename TEXT NOT NULL,
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
    quantity REAL NOT NULL,
    raw_payload JSON
);

CREATE TABLE IF NOT EXISTS smeta_items (
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

CREATE TABLE IF NOT EXISTS rule_evaluations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    match_record_id INTEGER NOT NULL REFERENCES match_records(id) ON DELETE CASCADE,
    rule_name TEXT NOT NULL,
    is_violation BOOLEAN NOT NULL,
    delta_volume REAL,
    delta_cost REAL,
    message TEXT
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
CREATE INDEX IF NOT EXISTS idx_rule_evaluations_match ON rule_evaluations(match_record_id);
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


# ----------------------------------------------------------------------
# 1. Соединения, транзакции и инициализация схемы
# ----------------------------------------------------------------------


def get_connection(db_path: str | Path = DEFAULT_DB_PATH) -> sqlite3.Connection:
    """Открывает соединение SQLite с доступом к колонкам по имени.

    :param db_path: путь к файлу БД либо ``":memory:"`` для тестового контура.
    :return: соединение с ``row_factory = sqlite3.Row`` и включенными FK.
    """
    connection = sqlite3.connect(str(db_path), timeout=10.0)
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


def init_db(conn: sqlite3.Connection) -> None:
    """Идемпотентно создает таблицы, индексы и представление v_audit_export."""
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.executescript(SCHEMA_SQL)
    conn.commit()


# ----------------------------------------------------------------------
# 2. Внутренние помощники границы персистентности
# ----------------------------------------------------------------------


def _to_sql_number(value: Decimal | int | float | str | None) -> SQLNumber:
    """Приводит значение ядра к типу, нативно принимаемому sqlite3.

    SQLite не имеет встроенного адаптера ``Decimal``, поэтому конвертация
    выполняется только здесь — на границе записи в REAL-колонки.
    """
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, bool):
        return int(value)
    return value


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
) -> str:
    """Регистрирует сессию аудита и возвращает ее идентификатор."""
    with transaction(conn):
        conn.execute(
            """
            INSERT INTO sessions (id, project_name, rd_filename, smeta_filename)
            VALUES (?, ?, ?, ?)
            """,
            (session_id, project_name, rd_file, smeta_file),
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
            _serialize_payload(item.get("raw_payload")),
        )
        for item in items
    ]

    with transaction(conn):
        conn.executemany(
            """
            INSERT INTO smeta_items
                (session_id, row_idx, code, description, unit,
                 quantity, unit_price, total_cost, raw_payload)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
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


def save_rule_evaluations(
    conn: sqlite3.Connection,
    evaluations: Sequence[Mapping[str, Any]],
) -> int:
    """Сохраняет результаты проверки правил. Возвращает число строк.

    Обязательные поля записи: ``match_record_id``, ``rule_name``,
    ``is_violation``. Поля ``delta_volume``/``delta_cost`` и ``message``
    опциональны; ``Decimal`` конвертируется на границе персистентности.
    """
    if not evaluations:
        return 0

    rows = []
    for evaluation in evaluations:
        rule_name = _require(evaluation, "rule_name")
        if not isinstance(rule_name, str) or not rule_name.strip():
            raise ValueError(
                f"Поле 'rule_name' должно быть непустой строкой: {rule_name!r}"
            )
        is_violation = _require(evaluation, "is_violation")
        if not isinstance(is_violation, (bool, int)):
            raise ValueError(
                f"Поле 'is_violation' должно быть bool/int: {is_violation!r}"
            )
        rows.append(
            (
                _require(evaluation, "match_record_id"),
                rule_name,
                int(bool(is_violation)),
                _to_sql_number(evaluation.get("delta_volume")),
                _to_sql_number(evaluation.get("delta_cost")),
                evaluation.get("message"),
            )
        )

    with transaction(conn):
        conn.executemany(
            """
            INSERT INTO rule_evaluations
                (match_record_id, rule_name, is_violation,
                 delta_volume, delta_cost, message)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            rows,
        )
    return len(rows)


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
    confidence_score: Decimal | int | float | str,
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
    adjusted_ratio: Decimal | int | float | str | None = None,
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
