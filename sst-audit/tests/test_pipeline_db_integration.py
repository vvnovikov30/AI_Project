"""Интеграционные тесты сквозного пайплайна main.py с SQLite как SSOT.

Проверяется полный цикл: нормализованные DTO -> run_pipeline() -> записи в
sessions / rd_items / smeta_items / match_records / rule_evaluations,
сшивка отчета через v_audit_export и фильтрация очереди HITL.
"""

from collections.abc import Iterator
from decimal import Decimal
import json
from pathlib import Path
import sqlite3

import pytest

import db
from main import AUTO_MATCH_THRESHOLD, SessionConfig, match_items, run_pipeline
from models import CellPointer, NormalizedItem, UnitClass

PROJECT_NAME = "ЖК Север, корпус 3"
SESSION_ID = "it-session-001"

TITLE_CABLE = "Кабель силовой ВВГнг(A)-LS 3х1.5"
CODE_CABLE = "ФССЦ-21.1.06.04-0012"
TITLE_VALVE = "Задвижка клиновая стальная фланцевая 30с41нж DN50 PN16"
CODE_VALVE = "30с41нж"
TITLE_PUMP = "Насос центробежный консольный"
CODE_PUMP = "PUMP-1"


def _pointer(doc_type: str, row_idx: int) -> CellPointer:
    """Создает аудиторский след ячейки для тестовых DTO."""
    if doc_type == "SPECIFICATION":
        return CellPointer(
            doc_type=doc_type,
            file_name="rd.pdf",
            sheet_name="Page 1",
            row_idx=row_idx,
            col_letter="D",
            coordinate=f"D{row_idx}",
        )
    return CellPointer(
        doc_type=doc_type,
        file_name="smeta.xlsx",
        sheet_name="Смета",
        row_idx=row_idx,
        col_letter="B",
        coordinate=f"B{row_idx}",
    )


def _rd_item(
    raw_id: str,
    title: str,
    code: str | None,
    unit: UnitClass,
    volume: Decimal,
    row_idx: int,
    raw_unit: str | None = None,
) -> NormalizedItem:
    """Позиция рабочей документации (спецификация)."""
    return NormalizedItem(
        raw_id=raw_id,
        code=code,
        raw_title=title,
        clean_title=title.lower(),
        raw_unit=raw_unit or unit.value,
        canonical_unit=unit,
        unit_multiplier=Decimal("1.0"),
        volume=volume,
        dimensions={},
        context_section="Default_Section",
        pointer=_pointer("SPECIFICATION", row_idx),
    )


def _smeta_item(
    raw_id: str,
    title: str,
    code: str | None,
    unit: UnitClass,
    volume: Decimal,
    row_idx: int,
    raw_unit: str | None = None,
    unit_price: Decimal = Decimal("0"),
    total_cost: Decimal = Decimal("0"),
) -> NormalizedItem:
    """Позиция сметы (материальный ресурс)."""
    return NormalizedItem(
        raw_id=raw_id,
        code=code,
        raw_title=title,
        clean_title=title.lower(),
        raw_unit=raw_unit or unit.value,
        canonical_unit=unit,
        unit_multiplier=Decimal("1.0"),
        volume=volume,
        dimensions={},
        context_section="Default_Section",
        pointer=_pointer("LSR", row_idx),
        is_resource=True,
        parent_id="LSR_1",
        unit_price=unit_price,
        total_cost=total_cost,
    )


def _cable_pair() -> tuple[NormalizedItem, NormalizedItem]:
    """Пара «кабель»: полное совпадение шифра и объема."""
    rd_item = _rd_item(
        "SPEC_1",
        TITLE_CABLE,
        CODE_CABLE,
        UnitClass.LENGTH,
        Decimal("150.0"),
        12,
        raw_unit="100 м",
    )
    smeta_item = _smeta_item(
        "LSR_10",
        TITLE_CABLE,
        CODE_CABLE,
        UnitClass.LENGTH,
        Decimal("150.0"),
        40,
        raw_unit="100 м",
        unit_price=Decimal("2500.00"),
        total_cost=Decimal("375000.00"),
    )
    return rd_item, smeta_item


def _three_item_dataset() -> tuple[list[NormalizedItem], list[NormalizedItem]]:
    """2 автоматические пары + 1 позиция РД без пары в смете."""
    rd_cable, smeta_cable = _cable_pair()
    rd_valve = _rd_item(
        "SPEC_2",
        TITLE_VALVE,
        CODE_VALVE,
        UnitClass.COUNT,
        Decimal("8.0"),
        15,
    )
    smeta_valve = _smeta_item(
        "LSR_20",
        TITLE_VALVE,
        CODE_VALVE,
        UnitClass.COUNT,
        Decimal("8.0"),
        45,
        unit_price=Decimal("12000.00"),
        total_cost=Decimal("96000.00"),
    )
    rd_pump = _rd_item(
        "SPEC_3",
        TITLE_PUMP,
        CODE_PUMP,
        UnitClass.COUNT,
        Decimal("2.0"),
        18,
    )
    return [rd_cable, rd_valve, rd_pump], [smeta_cable, smeta_valve]


def _config(session_id: str = SESSION_ID) -> SessionConfig:
    """Конфигурация тестовой сессии (пути не читаются: DTO переданы явно)."""
    return SessionConfig(
        project_name=PROJECT_NAME,
        rd_path="rd.pdf",
        smeta_path="smeta.xlsx",
        session_id=session_id,
    )


@pytest.fixture()
def conn(tmp_path: Path) -> Iterator[sqlite3.Connection]:
    """Реальная файловая БД SQLite (WAL), изолированная на тест."""
    connection = db.get_connection(tmp_path / "sst_audit_it.db")
    try:
        yield connection
    finally:
        connection.close()


def _count(connection: sqlite3.Connection, table: str) -> int:
    return int(
        connection.execute(f"SELECT COUNT(*) AS c FROM {table}").fetchone()["c"]
    )


def test_full_pipeline_persists_all_entities(conn: sqlite3.Connection) -> None:
    """Сквозной прогон создает записи во всех таблицах SSOT."""
    rd_items, smeta_items = _three_item_dataset()

    result = run_pipeline(conn, _config(), rd_items=rd_items, smeta_items=smeta_items)

    assert result.session_id == SESSION_ID
    assert result.rd_saved == 3
    assert result.smeta_saved == 2
    # 2 решения по смете + 1 обратный проход по неосмеченной позиции РД
    assert result.matches_saved == 3
    assert result.evaluations_saved == 3
    assert result.auto_matches == 2
    assert result.hitl_pending == 1
    assert result.session_status == "MATCHED"

    assert _count(conn, "sessions") == 1
    assert _count(conn, "rd_items") == 3
    assert _count(conn, "smeta_items") == 2
    assert _count(conn, "match_records") == 3
    assert _count(conn, "rule_evaluations") == 3

    session = conn.execute(
        "SELECT * FROM sessions WHERE id = ?", (SESSION_ID,)
    ).fetchone()
    assert session is not None
    assert session["project_name"] == PROJECT_NAME
    assert session["status"] == "MATCHED"

    statuses = {
        row["status"]
        for row in conn.execute(
            "SELECT status FROM match_records WHERE session_id = ?", (SESSION_ID,)
        )
    }
    assert statuses == {"AUTO", "HITL_PENDING"}

    rules = {
        row["rule_name"]
        for row in conn.execute("SELECT rule_name FROM rule_evaluations")
    }
    assert "RD_NOT_IN_SMETA" in rules
    assert "VOLUME_WITHIN_TOLERANCE" in rules


def test_audit_export_view_stitches_rows(conn: sqlite3.Connection) -> None:
    """v_audit_export возвращает корректно сшитые строки RD + смета + связь."""
    rd_items, smeta_items = _three_item_dataset()
    run_pipeline(conn, _config(), rd_items=rd_items, smeta_items=smeta_items)

    rows = conn.execute(
        "SELECT * FROM v_audit_export WHERE session_id = ?", (SESSION_ID,)
    ).fetchall()
    assert len(rows) == 3

    by_rd_code = {row["rd_code"]: row for row in rows}

    cable = by_rd_code[CODE_CABLE]
    assert cable["project_name"] == PROJECT_NAME
    assert cable["smeta_code"] == CODE_CABLE
    assert cable["rd_quantity"] == pytest.approx(150.0)
    assert cable["smeta_quantity"] == pytest.approx(150.0)
    assert cable["unit_price"] == pytest.approx(2500.0)
    assert cable["total_cost"] == pytest.approx(375000.0)
    assert cable["match_status"] == "AUTO"
    assert cable["rd_unit"] == "100 м"

    payload = json.loads(cable["rd_raw_payload"])
    assert payload["raw_id"] == "SPEC_1"
    assert payload["volume"] == "150.0"
    assert payload["canonical_unit"] == "м"
    assert payload["pointer"]["coordinate"] == "D12"

    pump = by_rd_code[CODE_PUMP]
    assert pump["smeta_code"] is None
    assert pump["match_status"] == "HITL_PENDING"
    assert pump["match_method"] == "REVERSE_PASS"


def test_pending_reviews_only_hitl_pending(conn: sqlite3.Connection) -> None:
    """get_pending_reviews отдает только связи со статусом HITL_PENDING."""
    rd_items, smeta_items = _three_item_dataset()
    run_pipeline(conn, _config(), rd_items=rd_items, smeta_items=smeta_items)

    pending = db.get_pending_reviews(conn, SESSION_ID)

    assert len(pending) == 1
    assert {row["status"] for row in pending} == {"HITL_PENDING"}
    assert pending[0]["rd_code"] == CODE_PUMP
    assert pending[0]["smeta_item_id"] is None

    db.record_expert_review(
        conn,
        int(pending[0]["match_id"]),
        "Инженер ПТО",
        "REJECTED",
        reason_code="NOT_IN_SCOPE",
        comment="Позиция не требует сверки со сметой",
    )
    assert db.get_pending_reviews(conn, SESSION_ID) == []


def test_session_reviewed_when_no_pending(conn: sqlite3.Connection) -> None:
    """Без очереди HITL сессия получает статус REVIEWED."""
    rd_cable, smeta_cable = _cable_pair()

    result = run_pipeline(
        conn,
        _config("it-session-reviewed"),
        rd_items=[rd_cable],
        smeta_items=[smeta_cable],
    )

    assert result.session_status == "REVIEWED"
    assert result.hitl_pending == 0
    assert result.matches_saved == 1
    assert result.evaluations_saved == 1
    assert db.get_pending_reviews(conn, "it-session-reviewed") == []


def test_ingest_rollback_keeps_database_clean(
    conn: sqlite3.Connection,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Сбой на этапе ingest откатывает сессию и все ранее вставленные позиции."""
    rd_cable, smeta_cable = _cable_pair()

    def _fail_on_smeta(*args: object, **kwargs: object) -> int:
        raise RuntimeError("disk full")

    monkeypatch.setattr(db, "save_smeta_items", _fail_on_smeta)

    with pytest.raises(RuntimeError, match="disk full"):
        run_pipeline(
            conn,
            _config("it-session-rollback"),
            rd_items=[rd_cable],
            smeta_items=[smeta_cable],
        )

    for table in (
        "sessions",
        "rd_items",
        "smeta_items",
        "match_records",
        "rule_evaluations",
    ):
        assert _count(conn, table) == 0, f"Таблица {table} не очищена rollback-ом"


def test_match_items_marks_low_confidence_and_ambiguity() -> None:
    """Матчер классифицирует низкую уверенность и неоднозначность как HITL."""
    weak_smeta = _smeta_item(
        "LSR_W",
        "Кабель силовой",
        None,
        UnitClass.LENGTH,
        Decimal("10.0"),
        50,
    )
    control_rd = _rd_item(
        "SPEC_W",
        "Кабель контрольный КВВГ 3х1.5",
        None,
        UnitClass.LENGTH,
        Decimal("10.0"),
        20,
    )

    weak_decisions = match_items([control_rd], [weak_smeta])
    assert len(weak_decisions) == 1
    assert weak_decisions[0].rd_item is control_rd
    assert weak_decisions[0].confidence < AUTO_MATCH_THRESHOLD
    assert weak_decisions[0].status == "HITL_PENDING"

    ambiguous_smeta = _smeta_item(
        "LSR_A",
        "Кабель силовой ВВГнг 3х1.5",
        None,
        UnitClass.LENGTH,
        Decimal("10.0"),
        51,
    )
    exact_rd = _rd_item(
        "SPEC_A",
        "Кабель силовой ВВГнг 3х1.5",
        None,
        UnitClass.LENGTH,
        Decimal("10.0"),
        21,
    )
    similar_rd = _rd_item(
        "SPEC_B",
        "Кабель силовой ВВГнг 3х2.5",
        None,
        UnitClass.LENGTH,
        Decimal("10.0"),
        22,
    )

    ambiguous_decisions = match_items([exact_rd, similar_rd], [ambiguous_smeta])
    top = ambiguous_decisions[0]
    assert top.rd_item is exact_rd
    assert top.confidence >= AUTO_MATCH_THRESHOLD
    assert len(top.candidates) == 2
    assert top.status == "HITL_PENDING"


def test_interactive_hitl_updates_link_and_rules(
    conn: sqlite3.Connection,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Выбор второго кандидата в HITL обновляет связь и оценку правил."""
    smeta = _smeta_item(
        "LSR_INT",
        "Кабель силовой ВВГнг 3х1.5",
        None,
        UnitClass.LENGTH,
        Decimal("10.0"),
        60,
        unit_price=Decimal("100.00"),
        total_cost=Decimal("1000.00"),
    )
    exact = _rd_item(
        "SPEC_X",
        "Кабель силовой ВВГнг 3х1.5",
        "ART-1.5",
        UnitClass.LENGTH,
        Decimal("10.0"),
        30,
    )
    similar = _rd_item(
        "SPEC_Y",
        "Кабель силовой ВВГнг 3х2.5",
        "ART-2.5",
        UnitClass.LENGTH,
        Decimal("20.0"),
        31,
    )

    monkeypatch.setattr("builtins.input", lambda _prompt: "2")

    result = run_pipeline(
        conn,
        _config("it-session-hitl"),
        rd_items=[exact, similar],
        smeta_items=[smeta],
        interactive=True,
        reviewer="Инженер ПТО",
    )

    # Обратный проход по второму кандидату остается в очереди HITL.
    assert result.session_status == "MATCHED"

    match = conn.execute(
        "SELECT * FROM match_records WHERE session_id = ? ORDER BY id LIMIT 1",
        ("it-session-hitl",),
    ).fetchone()
    assert match is not None
    assert match["status"] == "RESOLVED"

    selected_id = conn.execute(
        "SELECT id FROM rd_items WHERE code = 'ART-2.5'"
    ).fetchone()["id"]
    assert match["rd_item_id"] == selected_id
    assert match["confidence_score"] == pytest.approx(96.153846, rel=1e-4)

    review = conn.execute(
        "SELECT * FROM expert_reviews WHERE match_record_id = ?",
        (match["id"],),
    ).fetchone()
    assert review is not None
    assert review["decision"] == "APPROVED"
    assert review["reviewer_name"] == "Инженер ПТО"

    rule = conn.execute(
        "SELECT * FROM rule_evaluations WHERE match_record_id = ?",
        (match["id"],),
    ).fetchone()
    assert rule is not None
    assert rule["rule_name"] == "VOLUME_DELTA_SUSPECT"
    assert rule["is_violation"] == 1
    # Дельта рассчитана по выбранному кандидату (20 x 1.06), а не по первому.
    assert rule["delta_volume"] == pytest.approx(-11.2)
