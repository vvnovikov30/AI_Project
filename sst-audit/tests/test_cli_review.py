"""Тесты интерактивного терминального ревью сметчика (cli_review.py).

Ввод оператора имитируется через monkeypatch ``builtins.input``; вывод Rich
направляется в StringIO, чтобы тесты не зависели от терминала.
"""

from collections.abc import Iterator
from decimal import Decimal
import io
from pathlib import Path
import sqlite3
from typing import Optional

from rich.console import Console
import pytest

import db
from cli_review import DEFAULT_REVIEWER, resolve_session_id, run_review

SESSION_ID = "cli-session-001"
PROJECT_NAME = "ЖК Север, корпус 3"


@pytest.fixture()
def conn(tmp_path: Path) -> Iterator[sqlite3.Connection]:
    """Изолированная файловая БД SQLite с инициализированной схемой SSOT."""
    connection = db.get_connection(tmp_path / "cli_review.db")
    db.init_db(connection)
    try:
        yield connection
    finally:
        connection.close()


def _console() -> tuple[Console, io.StringIO]:
    """Console, вывод которого можно проверить через возвращенный буфер."""
    stream = io.StringIO()
    console = Console(file=stream, width=120, force_terminal=False)
    return console, stream


def _feed(monkeypatch: pytest.MonkeyPatch, *answers: str) -> None:
    """Подставляет последовательность ответов оператора в input()."""
    queue = iter(answers)
    monkeypatch.setattr("builtins.input", lambda _prompt="": next(queue))


def _first_id(conn: sqlite3.Connection, table: str) -> int:
    """Возвращает id единственной записи таблицы в рамках тестовой сессии."""
    row = conn.execute(
        f"SELECT id FROM {table} WHERE session_id = ? ORDER BY id LIMIT 1",
        (SESSION_ID,),
    ).fetchone()
    assert row is not None
    return int(row["id"])


def _seed_session(
    conn: sqlite3.Connection,
    *,
    pending: bool = True,
    auto: bool = False,
) -> Optional[int]:
    """Создает сессию с парой РД/смета и связями заданных статусов."""
    db.create_session(
        conn,
        SESSION_ID,
        PROJECT_NAME,
        "rd.pdf",
        "smeta.xlsx",
    )
    db.save_rd_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 1,
                "code": "ART-1",
                "description": "Кабель силовой ВВГнг(A)-LS 3х1.5",
                "unit": "м",
                "quantity": Decimal("10.0"),
            }
        ],
    )
    db.save_smeta_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 2,
                "code": "ФССЦ-21.1.06.04-0012",
                "description": "Кабель силовой ВВГнг(A)-LS 3х1.5",
                "unit": "м",
                "quantity": Decimal("10.5"),
                "unit_price": Decimal("100.00"),
                "total_cost": Decimal("1050.00"),
            }
        ],
    )
    rd_id = _first_id(conn, "rd_items")
    smeta_id = _first_id(conn, "smeta_items")

    matches: list[dict[str, object]] = []
    if pending:
        matches.append(
            {
                "rd_item_id": rd_id,
                "smeta_item_id": smeta_id,
                "confidence_score": Decimal("62.5"),
                "conversion_ratio": Decimal("1.0"),
                "match_method": "FUZZY_CASCADE",
                "status": "HITL_PENDING",
            }
        )
    if auto:
        matches.append(
            {
                "rd_item_id": rd_id,
                "smeta_item_id": smeta_id,
                "confidence_score": Decimal("92.0"),
                "conversion_ratio": Decimal("1.0"),
                "match_method": "FUZZY_CASCADE",
                "status": "AUTO",
            }
        )
    db.save_matches(conn, SESSION_ID, matches)

    if not pending:
        return None
    row = conn.execute(
        """
        SELECT id FROM match_records
        WHERE session_id = ? AND status = 'HITL_PENDING'
        ORDER BY id LIMIT 1
        """,
        (SESSION_ID,),
    ).fetchone()
    assert row is not None
    return int(row["id"])


def _review(conn: sqlite3.Connection, match_id: int) -> sqlite3.Row:
    """Возвращает запись expert_reviews по связи."""
    row = conn.execute(
        "SELECT * FROM expert_reviews WHERE match_record_id = ? ORDER BY id DESC",
        (match_id,),
    ).fetchone()
    assert row is not None
    return row


def _match(conn: sqlite3.Connection, match_id: int) -> sqlite3.Row:
    """Возвращает запись match_records по id."""
    row = conn.execute(
        "SELECT * FROM match_records WHERE id = ?", (match_id,)
    ).fetchone()
    assert row is not None
    return row


def _session_status(conn: sqlite3.Connection) -> str:
    """Возвращает текущий статус тестовой сессии."""
    row = conn.execute(
        "SELECT status FROM sessions WHERE id = ?", (SESSION_ID,)
    ).fetchone()
    assert row is not None
    return str(row["status"])


def test_cli_review_approve_flow(
    conn: sqlite3.Connection,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Команда A сохраняет APPROVED и переводит связь в RESOLVED."""
    match_id = _seed_session(conn)
    assert match_id is not None
    _feed(monkeypatch, "A")
    console, _ = _console()

    outcome = run_review(conn, SESSION_ID, reviewer="ivanov", console=console)

    review = _review(conn, match_id)
    assert review["decision"] == "APPROVED"
    assert review["reviewer_name"] == "ivanov"
    assert review["reject_reason_code"] is None
    assert review["adjusted_ratio"] is None

    match = _match(conn, match_id)
    assert match["status"] == "RESOLVED"
    assert match["rd_item_id"] is not None

    assert outcome.approved == 1
    assert outcome.processed == 1
    assert outcome.remaining == 0
    assert outcome.session_status == "REVIEWED"


def test_cli_review_reject_flow_with_reason(
    conn: sqlite3.Connection,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Команда R с кодом причины сохраняет reject_reason_code и снимает связь."""
    match_id = _seed_session(conn)
    assert match_id is not None
    _feed(monkeypatch, "R", "2", "Смета в 100 м, РД в м")
    console, _ = _console()

    outcome = run_review(conn, SESSION_ID, reviewer="petrova", console=console)

    review = _review(conn, match_id)
    assert review["decision"] == "REJECTED"
    assert review["reject_reason_code"] == "UNIT_MULTIPLICITY"
    assert review["comment"] == "Смета в 100 м, РД в м"
    assert review["reviewer_name"] == "petrova"

    match = _match(conn, match_id)
    assert match["status"] == "RESOLVED"
    assert match["rd_item_id"] is None
    assert match["confidence_score"] == 0

    assert outcome.rejected == 1
    assert outcome.remaining == 0


def test_cli_review_adjust_ratio(
    conn: sqlite3.Connection,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Команда C сохраняет adjusted_ratio и пересчитывает дельту объема."""
    match_id = _seed_session(conn)
    assert match_id is not None
    _feed(monkeypatch, "C", "1.25")
    console, _ = _console()

    outcome = run_review(conn, SESSION_ID, reviewer="sidorov", console=console)

    review = _review(conn, match_id)
    assert review["decision"] == "ADJUSTED"
    assert Decimal(str(review["adjusted_ratio"])) == Decimal("1.25")
    # RD 10 * 1.25 = 12.5; с отходом кабеля 6% = 13.25; смета 10.5 => -2.75.
    assert "-2.750" in review["comment"]

    match = _match(conn, match_id)
    assert match["status"] == "RESOLVED"
    assert match["rd_item_id"] is not None

    assert outcome.adjusted == 1
    assert outcome.remaining == 0


def test_cli_review_session_completion_to_reviewed(
    conn: sqlite3.Connection,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """После разметки последней HITL-записи сессия становится REVIEWED."""
    match_id = _seed_session(conn, auto=True)
    assert match_id is not None
    _feed(monkeypatch, "A")
    console, _ = _console()

    outcome = run_review(conn, SESSION_ID, console=console)

    assert _session_status(conn) == "REVIEWED"
    assert outcome.session_status == "REVIEWED"
    assert outcome.total == 1
    assert outcome.remaining == 0
    # AUTO-связь не требует ревью и не блокирует финализацию.
    auto_rows = conn.execute(
        "SELECT COUNT(*) AS c FROM match_records "
        "WHERE session_id = ? AND status = 'AUTO'",
        (SESSION_ID,),
    ).fetchone()
    assert int(auto_rows["c"]) == 1


def test_cli_review_empty_queue_message(
    conn: sqlite3.Connection,
) -> None:
    """Пустая очередь выводит панель и финализирует сессию из AUTO-связей."""
    _seed_session(conn, pending=False, auto=True)
    console, stream = _console()

    outcome = run_review(conn, SESSION_ID, console=console)

    assert "Очередь HITL пуста. Все позиции верифицированы" in stream.getvalue()
    assert outcome.total == 0
    assert outcome.remaining == 0
    assert _session_status(conn) == "REVIEWED"


def test_cli_review_skip_keeps_record_pending(
    conn: sqlite3.Connection,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Команда S не меняет статус связи и не финализирует сессию."""
    match_id = _seed_session(conn)
    assert match_id is not None
    _feed(monkeypatch, "S")
    console, _ = _console()

    outcome = run_review(conn, SESSION_ID, console=console)

    assert _match(conn, match_id)["status"] == "HITL_PENDING"
    assert outcome.skipped == 1
    assert outcome.remaining == 1
    assert outcome.session_status == "CREATED"


def test_resolve_session_id_latest_and_explicit(
    conn: sqlite3.Connection,
) -> None:
    """resolve_session_id поддерживает latest, явный id и ошибку для чужого id."""
    _seed_session(conn)

    assert resolve_session_id(conn, "latest") == SESSION_ID
    assert resolve_session_id(conn, SESSION_ID) == SESSION_ID
    assert DEFAULT_REVIEWER == "estimator_cli"

    with pytest.raises(LookupError):
        resolve_session_id(conn, "missing-session")

