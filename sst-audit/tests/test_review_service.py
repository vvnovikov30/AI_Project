"""Тесты сервисного слоя арбитража (review_service.py).

Проверяются три контракта:

* отсутствие UI-зависимостей (``rich``, ``input()``, ``print()``) — по AST,
  чтобы комментарии и докстринги не давали ложных срабатываний;
* отсутствие N+1 при загрузке очереди HITL (фиксированное число запросов);
* единый источник норматива отходов и запись решений в append-only журнал.
"""

from collections.abc import Iterator
import ast
from decimal import Decimal
import json
from pathlib import Path
import sqlite3

import pytest

import db
import review_service
from evaluator import DiscrepancyDetector
from models import CellPointer, NormalizedItem, UnitClass
from review_service import (
    CandidateMatch,
    ReviewCard,
    apply_adjustment,
    apply_approval,
    apply_candidate_selection,
    apply_rejection,
    finalize_session,
    format_delta,
    load_pending_cards,
    pending_count,
    resolve_session_id,
)

SESSION_ID = "review-session-001"
MODULE_PATH = Path(review_service.__file__).resolve()


@pytest.fixture()
def conn() -> Iterator[sqlite3.Connection]:
    """Изолированная файловая БД SQLite с инициализированной схемой SSOT."""
    connection = db.get_connection(":memory:")
    db.init_db(connection)
    try:
        yield connection
    finally:
        connection.close()


def _item(
    title: str,
    volume: str,
    *,
    unit: UnitClass = UnitClass.MASS,
    raw_unit: str = "т",
) -> NormalizedItem:
    """Нормализованная позиция с контролируемым объемом и наименованием."""
    return NormalizedItem(
        raw_id=f"RAW_{title[:8]}",
        code=None,
        raw_title=title,
        clean_title=title.lower(),
        raw_unit=raw_unit,
        canonical_unit=unit,
        unit_multiplier=Decimal("1.0"),
        volume=Decimal(volume),
        dimensions={},
        context_section="Default_Section",
        pointer=CellPointer("LSR", "test.xlsx", "Sheet1", 1, "A", "A1"),
        is_resource=False,
        parent_id=None,
    )


def _seed_pending_queue(conn: sqlite3.Connection, count: int = 3) -> None:
    """Создает сессию с ``count`` связями HITL_PENDING (РД + смета)."""
    db.create_session(conn, SESSION_ID, "ЖК Север", "rd.pdf", "smeta.xlsx")
    db.save_rd_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": index,
                "code": f"ART-{index}",
                "description": "Арматура А500С",
                "unit": "т",
                "quantity": Decimal("2.0"),
            }
            for index in range(1, count + 1)
        ],
    )
    db.save_smeta_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": index,
                "code": f"08-02-001-{index:02d}",
                "description": "Арматура А500С",
                "unit": "т",
                "quantity": Decimal("2.1"),
                "unit_price": Decimal("100.00"),
                "total_cost": Decimal("210.00"),
            }
            for index in range(1, count + 1)
        ],
    )
    rd_ids = [
        int(row["id"])
        for row in conn.execute(
            "SELECT id FROM rd_items WHERE session_id = ? ORDER BY id", (SESSION_ID,)
        )
    ]
    smeta_ids = [
        int(row["id"])
        for row in conn.execute(
            "SELECT id FROM smeta_items WHERE session_id = ? ORDER BY id",
            (SESSION_ID,),
        )
    ]
    db.save_matches(
        conn,
        SESSION_ID,
        [
            {
                "rd_item_id": rd_id,
                "smeta_item_id": smeta_id,
                "confidence_score": Decimal("62.5"),
                "conversion_ratio": Decimal("1.0"),
                "match_method": "FUZZY_CASCADE",
                "status": "HITL_PENDING",
            }
            for rd_id, smeta_id in zip(rd_ids, smeta_ids)
        ],
    )


def _match_id(conn: sqlite3.Connection, position: int = 0) -> int:
    """Идентификатор связи очереди по порядку id."""
    rows = conn.execute(
        "SELECT id FROM match_records WHERE session_id = ? ORDER BY id", (SESSION_ID,)
    ).fetchall()
    return int(rows[position]["id"])


# ----------------------------------------------------------------------
# 1. Изоляция от UI (проверка по AST, а не по тексту файла)
# ----------------------------------------------------------------------


def test_module_has_no_presentation_dependencies() -> None:
    """В сервисном слое нет rich/argparse и вызовов input()/print()."""
    tree = ast.parse(MODULE_PATH.read_text(encoding="utf-8"))

    imported: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported.update(alias.name.split(".")[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported.add(node.module.split(".")[0])

    assert "rich" not in imported
    assert "argparse" not in imported
    assert "click" not in imported

    called_names = {
        node.func.id
        for node in ast.walk(tree)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
    }
    assert called_names.isdisjoint({"input", "print"})


def test_module_exposes_no_event_mutation_api() -> None:
    """Журнал арбитража append-only: сервис не умеет править/удалять события."""
    for forbidden in (
        "update_arbitration_event",
        "delete_arbitration_event",
        "clear_arbitration_events",
    ):
        assert not hasattr(review_service, forbidden), forbidden
        assert not hasattr(db, forbidden), forbidden


# ----------------------------------------------------------------------
# 2. Загрузка очереди без N+1
# ----------------------------------------------------------------------


def test_load_pending_cards_uses_three_queries(conn: sqlite3.Connection) -> None:
    """Число SQL-запросов не зависит от длины очереди (было N+1)."""
    _seed_pending_queue(conn, count=4)
    statements: list[str] = []

    cards = load_pending_cards(conn, SESSION_ID, trace_callback=statements.append)

    assert len(cards) == 4
    assert len(statements) == 3
    assert "match_records" in statements[0]
    assert "rd_items" in statements[1]
    assert "smeta_items" in statements[2]

    # Кандидаты рассчитаны в памяти: по одному лучшему совпадению на карточку.
    for index, card in enumerate(cards, start=1):
        assert card.index == index
        assert card.total == 4
        assert card.candidates, "кандидаты должны быть рассчитаны"
        assert len(card.candidates) <= review_service.MAX_CANDIDATES
        assert all(
            candidate.score >= review_service.MIN_CANDIDATE_SCORE
            for candidate in card.candidates
        )
        assert isinstance(card.candidates[0], CandidateMatch)


def test_load_pending_cards_query_count_is_constant(conn: sqlite3.Connection) -> None:
    """Очередь из одной карточки и из четырех требует одинаковых 3 запросов."""
    _seed_pending_queue(conn, count=1)
    single: list[str] = []
    load_pending_cards(conn, SESSION_ID, trace_callback=single.append)

    _seed_pending_queue_extra(conn, count=3)
    many: list[str] = []
    load_pending_cards(conn, SESSION_ID, trace_callback=many.append)

    assert len(single) == len(many) == 3


def _seed_pending_queue_extra(conn: sqlite3.Connection, count: int) -> None:
    """Добавляет в существующую сессию еще ``count`` связей HITL_PENDING."""
    db.save_rd_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 100 + index,
                "code": f"BET-{index}",
                "description": "Бетон B25",
                "unit": "м3",
                "quantity": Decimal("5.0"),
            }
            for index in range(1, count + 1)
        ],
    )
    db.save_smeta_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 200 + index,
                "code": f"06-01-001-{index:02d}",
                "description": "Бетон B25",
                "unit": "м3",
                "quantity": Decimal("5.2"),
                "unit_price": Decimal("4500.00"),
                "total_cost": Decimal("23400.00"),
            }
            for index in range(1, count + 1)
        ],
    )
    rd_ids = [
        int(row["id"])
        for row in conn.execute(
            "SELECT id FROM rd_items WHERE session_id = ? AND row_idx >= 100 "
            "ORDER BY id",
            (SESSION_ID,),
        )
    ]
    smeta_ids = [
        int(row["id"])
        for row in conn.execute(
            "SELECT id FROM smeta_items WHERE session_id = ? AND row_idx >= 200 "
            "ORDER BY id",
            (SESSION_ID,),
        )
    ]
    db.save_matches(
        conn,
        SESSION_ID,
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


def test_load_pending_cards_empty_queue_single_query(
    conn: sqlite3.Connection,
) -> None:
    """Пустая очередь читается одним запросом и не требует батча позиций."""
    db.create_session(conn, SESSION_ID, "ЖК Север", "rd.pdf", "smeta.xlsx")
    statements: list[str] = []

    cards = load_pending_cards(conn, SESSION_ID, trace_callback=statements.append)

    assert cards == []
    assert len(statements) == 1
    assert pending_count(conn, SESSION_ID) == 0


# ----------------------------------------------------------------------
# 3. Единый норматив отходов (SSOT с расчетным ядром)
# ----------------------------------------------------------------------


def _card(
    title: str,
    *,
    rd_quantity: str = "10.0",
    smeta_quantity: str = "10.5",
    ratio: str = "1.0",
) -> ReviewCard:
    """Карточка с контролируемыми объемами для проверки расчета дельты."""
    return ReviewCard(
        match_id=1,
        index=1,
        total=1,
        confidence=Decimal("62.5"),
        conversion_ratio=Decimal(ratio),
        rd_code="ART-1",
        rd_description="Арматура А500С",
        rd_unit="т",
        rd_quantity=Decimal(rd_quantity),
        smeta_code="08-02-001-01",
        smeta_description=title,
        smeta_unit="т",
        smeta_quantity=Decimal(smeta_quantity),
        unit_price=Decimal("100.00"),
        total_cost=Decimal("1050.00"),
    )


@pytest.mark.parametrize(
    ("title", "expected"),
    [
        ("Арматура А500С", Decimal("0.05")),
        ("Бетон B25", Decimal("0.02")),
        ("Кабель силовой ВВГнг", Decimal("0.06")),
        ("Кирпич керамический", Decimal("0.00")),
        ("", Decimal("0.00")),
    ],
)
def test_waste_coefficient_is_single_source(
    title: str,
    expected: Decimal,
) -> None:
    """Карточка берет норматив отходов только из ``DiscrepancyDetector``."""
    assert DiscrepancyDetector.waste_coefficient(title) == expected
    assert _card(title).waste_coefficient == expected


def test_card_delta_matches_evaluator_volume_delta() -> None:
    """Дельта карточки совпадает с ``CollisionRecord.volume_delta`` ядра."""
    title = "Кабель силовой ВВГнг"
    card = _card(title, rd_quantity="10.0", smeta_quantity="10.5")

    record = DiscrepancyDetector.evaluate(
        _item(title, "10.5"),
        _item(title, "10.0"),
        Decimal("62.5"),
    )

    assert card.delta_volume() == record.volume_delta == Decimal("-0.1")
    assert card.allowed_delta() == record.allowed_delta
    assert card.adjusted_rd_volume() == Decimal("10.600")
    assert format_delta(card.delta_volume()) == "-0.100"


# ----------------------------------------------------------------------
# 4. Запись решений и журнал арбитража
# ----------------------------------------------------------------------


def test_apply_approval_records_review_event(conn: sqlite3.Connection) -> None:
    """Подтверждение пишет REVIEW_APPROVED и завершает сессию."""
    _seed_pending_queue(conn, count=1)
    match_id = _match_id(conn)
    card = load_pending_cards(conn, SESSION_ID)[0]

    apply_approval(conn, card, "ivanov")

    review = conn.execute(
        "SELECT decision, reviewer_name, comment FROM expert_reviews "
        "WHERE match_record_id = ?",
        (match_id,),
    ).fetchone()
    assert review["decision"] == "APPROVED"
    assert review["reviewer_name"] == "ivanov"
    assert review["comment"] is None

    events = db.get_arbitration_events(conn, SESSION_ID, match_id)
    assert [event["action"] for event in events] == ["REVIEW_APPROVED"]
    assert events[0]["actor"] == "ivanov"
    assert finalize_session(conn, SESSION_ID) == "REVIEWED"


def test_apply_candidate_selection_writes_link_and_review(
    conn: sqlite3.Connection,
) -> None:
    """Выбор кандидата фиксирует LINK_UPDATED с actor=reviewer и решение."""
    _seed_pending_queue(conn, count=2)
    card = load_pending_cards(conn, SESSION_ID)[0]
    assert card.candidates
    candidate = card.candidates[0]

    apply_candidate_selection(conn, card, candidate, 1, "petrova")

    events = db.get_arbitration_events(conn, SESSION_ID, card.match_id)
    assert [event["action"] for event in events] == [
        "LINK_UPDATED",
        "REVIEW_APPROVED",
    ]
    assert {event["actor"] for event in events} == {"petrova"}
    assert "Выбран кандидат №1" in events[0]["comment"]

    match = conn.execute(
        "SELECT rd_item_id, confidence_score, status FROM match_records WHERE id = ?",
        (card.match_id,),
    ).fetchone()
    assert match["rd_item_id"] == candidate.rd_item_id
    assert Decimal(str(match["confidence_score"])) == candidate.score
    assert match["status"] == "RESOLVED"


def test_apply_rejection_keeps_top_candidate(conn: sqlite3.Connection) -> None:
    """Отклонение сохраняет top-1 кандидата и пишет REVIEW_REJECTED."""
    _seed_pending_queue(conn, count=1)
    card = load_pending_cards(conn, SESSION_ID)[0]
    assert card.candidates

    apply_rejection(conn, card, "sidorov", "UNIT_MULTIPLICITY", "Смета в 100 м")

    review = conn.execute(
        "SELECT decision, reject_reason_code, comment, reviewer_name "
        "FROM expert_reviews WHERE match_record_id = ?",
        (card.match_id,),
    ).fetchone()
    assert review["decision"] == "REJECTED"
    assert review["reject_reason_code"] == "UNIT_MULTIPLICITY"
    assert review["comment"] == "Смета в 100 м"
    assert review["reviewer_name"] == "sidorov"

    match = conn.execute(
        "SELECT rd_item_id, status FROM match_records WHERE id = ?",
        (card.match_id,),
    ).fetchone()
    assert match["rd_item_id"] is not None
    assert match["status"] == "RESOLVED"

    events = db.get_arbitration_events(conn, SESSION_ID, card.match_id)
    assert [event["action"] for event in events] == [
        "LINK_UPDATED",
        "REVIEW_REJECTED",
    ]
    assert events[1]["actor"] == "sidorov"


def test_apply_adjustment_records_new_delta(conn: sqlite3.Connection) -> None:
    """Корректировка коэффициента сохраняет adjusted_ratio и новую дельту."""
    _seed_pending_queue(conn, count=1)
    card = load_pending_cards(conn, SESSION_ID)[0]
    assert card.smeta_description == "Арматура А500С"

    delta = apply_adjustment(conn, card, "ivanov", Decimal("1.25"))

    # РД 2.0 * 1.25 * (1 + 0.05) = 2.625; смета 2.1 => -0.525.
    assert delta == Decimal("-0.525")

    review = conn.execute(
        "SELECT decision, adjusted_ratio, comment FROM expert_reviews "
        "WHERE match_record_id = ?",
        (card.match_id,),
    ).fetchone()
    assert review["decision"] == "ADJUSTED"
    assert Decimal(str(review["adjusted_ratio"])) == Decimal("1.25")
    assert "-0.525" in review["comment"]

    events = db.get_arbitration_events(conn, SESSION_ID, card.match_id)
    assert [event["action"] for event in events] == ["REVIEW_ADJUSTED"]
    payload = json.loads(events[0]["new_payload"])
    assert payload["adjusted_ratio"] == "1.25"
    assert payload["conversion_ratio"] == "1.0"


def test_resolve_session_id_and_pending_count(conn: sqlite3.Connection) -> None:
    """Сервис разрешает latest/явный id и считает остаток очереди."""
    _seed_pending_queue(conn, count=2)

    assert resolve_session_id(conn, "latest") == SESSION_ID
    assert resolve_session_id(conn, SESSION_ID) == SESSION_ID
    assert pending_count(conn, SESSION_ID) == 2

    with pytest.raises(LookupError):
        resolve_session_id(conn, "missing-session")
    with pytest.raises(LookupError):
        finalize_session(conn, "missing-session")
