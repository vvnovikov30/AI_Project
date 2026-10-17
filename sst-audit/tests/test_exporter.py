"""Тесты генератора ведомости аудита в Excel (exporter.py).

Проверяется структура книги, нативные формулы проверки, скрытая служебная
колонка и приоритет экспертного коэффициента ``adjusted_ratio``.
"""

from collections.abc import Iterator
from decimal import Decimal
from pathlib import Path
import sqlite3

from openpyxl import load_workbook
import pytest

import db
import exporter
from exporter import SHEET_TITLE, export_audit_to_excel

SESSION_ID = "export-session-001"
PROJECT_NAME = "ЖК Север, корпус 3"


@pytest.fixture()
def conn(tmp_path: Path) -> Iterator[sqlite3.Connection]:
    """Изолированная файловая БД SQLite с инициализированной схемой SSOT."""
    connection = db.get_connection(tmp_path / "exporter.db")
    db.init_db(connection)
    try:
        yield connection
    finally:
        connection.close()


def _seed_session(conn: sqlite3.Connection) -> tuple[int, int]:
    """Создает сессию с двумя связями: AUTO и HITL_PENDING.

    :return: идентификаторы связей (AUTO, HITL_PENDING).
    """
    db.create_session(conn, SESSION_ID, PROJECT_NAME, "rd.pdf", "smeta.xlsx")
    db.save_rd_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 1,
                "code": "РД-1",
                "description": "Кабель силовой ВВГнг 3х1.5",
                "unit": "м",
                "quantity": Decimal("100.0"),
            },
            {
                "row_idx": 2,
                "code": "РД-2",
                "description": "Арматура А500С d12",
                "unit": "т",
                "quantity": Decimal("10.0"),
            },
        ],
    )
    db.save_smeta_items(
        conn,
        SESSION_ID,
        [
            {
                "row_idx": 1,
                "code": "ФССЦ-21.1.06.04-0012",
                "description": "Кабель силовой ВВГнг 3х1.5",
                "unit": "м",
                "quantity": Decimal("100.0"),
                "unit_price": Decimal("2500.0"),
                "total_cost": Decimal("250000.0"),
            },
            {
                "row_idx": 2,
                "code": "ФССЦ-22.1.03.01-0001",
                "description": "Арматура А500С d12",
                "unit": "т",
                "quantity": Decimal("12.0"),
                "unit_price": Decimal("45000.0"),
                "total_cost": Decimal("540000.0"),
            },
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
            "SELECT id FROM smeta_items WHERE session_id = ? ORDER BY id", (SESSION_ID,)
        )
    ]

    db.save_matches(
        conn,
        SESSION_ID,
        [
            {
                "rd_item_id": rd_ids[0],
                "smeta_item_id": smeta_ids[0],
                "confidence_score": Decimal("95.0"),
                "conversion_ratio": Decimal("1.0"),
                "match_method": "EXACT",
                "status": "AUTO",
            },
            {
                "rd_item_id": rd_ids[1],
                "smeta_item_id": smeta_ids[1],
                "confidence_score": Decimal("60.0"),
                "conversion_ratio": Decimal("2.0"),
                "match_method": "FUZZY",
                "status": "HITL_PENDING",
            },
        ],
    )

    match_ids = [
        int(row["id"])
        for row in conn.execute(
            "SELECT id FROM match_records WHERE session_id = ? ORDER BY id", (SESSION_ID,)
        )
    ]
    return match_ids[0], match_ids[1]


def _open_sheet(path: Path):
    """Загружает книгу без вычисления формул и возвращает ее активный лист."""
    workbook = load_workbook(path, data_only=False)
    return workbook[SHEET_TITLE]


def test_exporter_creates_valid_xlsx(conn: sqlite3.Connection, tmp_path: Path) -> None:
    """Файл создается, содержит обязательный лист, шапку и строки данных."""
    _seed_session(conn)
    target = export_audit_to_excel(conn, SESSION_ID, tmp_path / "audit.xlsx")

    assert target.is_file()
    assert target.stat().st_size > 0

    workbook = load_workbook(target)
    assert workbook.sheetnames == [SHEET_TITLE]
    worksheet = workbook[SHEET_TITLE]

    assert worksheet["B1"].value == "Проект:"
    assert worksheet["C1"].value == PROJECT_NAME
    assert worksheet["B2"].value == "ID сессии:"
    assert worksheet["C2"].value == SESSION_ID
    assert worksheet["B3"].value == "Дата формирования:"
    assert worksheet["C3"].value
    assert worksheet["A4"].value == "ID"
    assert worksheet["C4"].value == "Позиция РД (наименование)"
    assert worksheet["Q4"].value == "Экспертный комментарий / Причина отклонения"
    assert worksheet.freeze_panes == "C5"
    assert worksheet.max_row == 6
    assert worksheet.max_column == 17

    assert worksheet["B5"].value == 1
    assert worksheet["B6"].value == 2
    assert worksheet["E5"].number_format == exporter.VOLUME_FORMAT
    assert worksheet["J5"].number_format == exporter.MONEY_FORMAT
    assert worksheet["L5"].number_format == exporter.RATIO_FORMAT
    assert worksheet.column_dimensions["C"].width <= 50
    assert worksheet.column_dimensions["Q"].width <= 50

    # Нулевая дельта первой строки — мягкая зеленая заливка.
    assert worksheet["N5"].fill.fill_type == "solid"
    assert worksheet["N5"].fill.fgColor.rgb.endswith("C6EFCE")
    assert worksheet["P5"].value == "AUTO"
    assert worksheet["P6"].value == "HITL_PENDING"


def test_exporter_native_formulas_presence(
    conn: sqlite3.Connection, tmp_path: Path
) -> None:
    """Колонки M, N, O содержат формулы Excel, а не статичные числа."""
    _seed_session(conn)
    target = export_audit_to_excel(conn, SESSION_ID, tmp_path / "formulas.xlsx")
    worksheet = _open_sheet(target)

    for row in (5, 6):
        assert worksheet[f"K{row}"].value == f"=I{row}*J{row}"
        assert worksheet[f"M{row}"].value == f"=E{row}*L{row}"
        assert worksheet[f"N{row}"].value == f"=I{row}-M{row}"
        assert worksheet[f"O{row}"].value == f"=N{row}*J{row}"
        for column in ("M", "N", "O"):
            value = worksheet[f"{column}{row}"].value
            assert isinstance(value, str)
            assert value.startswith("=")
            assert worksheet[f"{column}{row}"].data_type == "f"


def test_exporter_hidden_id_column(conn: sqlite3.Connection, tmp_path: Path) -> None:
    """Служебная колонка A скрыта, но содержит идентификатор связи."""
    _, pending_match_id = _seed_session(conn)
    target = export_audit_to_excel(conn, SESSION_ID, tmp_path / "hidden.xlsx")
    worksheet = _open_sheet(target)

    assert worksheet.column_dimensions["A"].hidden is True
    assert worksheet["A4"].value == "ID"
    assert worksheet["A6"].value == pending_match_id


def test_exporter_adjusted_ratio_applied(
    conn: sqlite3.Connection, tmp_path: Path
) -> None:
    """Экспертная правка коэффициента попадает в колонку L вместо исходной."""
    _, pending_match_id = _seed_session(conn)

    # Две рецензии подряд: в отчет должна попасть последняя (1.25).
    db.record_expert_review(
        conn,
        pending_match_id,
        "ivanov",
        "ADJUSTED",
        comment="Первичное уточнение",
        adjusted_ratio=Decimal("1.5"),
    )
    db.record_expert_review(
        conn,
        pending_match_id,
        "petrov",
        "ADJUSTED",
        comment="Коэффициент уточнен по исполнительной схеме",
        adjusted_ratio=Decimal("1.25"),
    )

    target = export_audit_to_excel(conn, SESSION_ID, tmp_path / "adjusted.xlsx")
    worksheet = _open_sheet(target)

    assert worksheet["L6"].value == pytest.approx(1.25)
    assert worksheet["L6"].value != pytest.approx(2.0)
    assert worksheet["M6"].value == "=E6*L6"
    assert worksheet["P6"].value == "APPROVED"
    assert "исполнительной схеме" in str(worksheet["Q6"].value)

    stored_ratio = conn.execute(
        "SELECT conversion_ratio FROM match_records WHERE id = ?", (pending_match_id,)
    ).fetchone()["conversion_ratio"]
    assert stored_ratio == pytest.approx(2.0)


def test_exporter_rejected_reason_in_comment(
    conn: sqlite3.Connection, tmp_path: Path
) -> None:
    """REJECTED сохраняет наименование РД, причину в Q и не строит дельты."""
    _, pending_match_id = _seed_session(conn)
    db.record_expert_review(
        conn,
        pending_match_id,
        "ivanov",
        "REJECTED",
        reason_code="UNIT_MULTIPLICITY",
        comment="В смете расценка на 100 м",
    )

    target = export_audit_to_excel(conn, SESSION_ID, tmp_path / "rejected.xlsx")
    worksheet = _open_sheet(target)

    assert worksheet["P6"].value == "REJECTED"
    assert worksheet["C6"].value == "Арматура А500С d12"
    assert worksheet["E6"].value == pytest.approx(10.0)
    assert "UNIT_MULTIPLICITY" in str(worksheet["Q6"].value)
    assert "В смете расценка на 100 м" in str(worksheet["Q6"].value)
    # Отклоненная связь: дельтовые формулы не строятся.
    assert worksheet["N6"].value is None
    assert worksheet["O6"].value is None


def test_exporter_resolves_latest_session(conn: sqlite3.Connection, tmp_path: Path) -> None:
    """Сессия ``latest`` разрешается в последнюю созданную запись."""
    _seed_session(conn)
    db.create_session(conn, "export-session-002", "Другой объект", "rd2.pdf", "smeta2.xlsx")

    target = export_audit_to_excel(conn, "latest", tmp_path / "latest.xlsx")
    worksheet = _open_sheet(target)

    assert worksheet["C2"].value == "export-session-002"
    assert worksheet["C1"].value == "Другой объект"
    assert worksheet.max_row == 4

    with pytest.raises(LookupError):
        export_audit_to_excel(conn, "missing-session", tmp_path / "missing.xlsx")

