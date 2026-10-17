"""Генератор ведомости аудита объемов в формате Excel для контролирующих органов.

Отчет строится по данным SQLite (SSOT): строки сшиваются через представление
``v_audit_export``, при наличии нескольких экспертных рецензий выбирается
последняя. Расчетные колонки (``M``, ``N``, ``O``) записываются нативными
формулами Excel, поэтому проверяющий видит не только результат, но и логику
расчета и может пересчитать ведомость средствами самого Excel.

Запуск как самостоятельного скрипта::

    python exporter.py --db sst_audit.db --session latest --out report.xlsx
"""

from __future__ import annotations

import argparse
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import sqlite3
from typing import Any, Final, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

import db
from evaluator import DiscrepancyDetector

SHEET_TITLE: Final[str] = "Ведомость аудита объемов"
LATEST_SESSION: Final[str] = "latest"
DEFAULT_OUTPUT: Final[str] = "audit_report.xlsx"

HEADER_ROW: Final[int] = 4
FIRST_DATA_ROW: Final[int] = 5
HIDDEN_ID_COLUMN: Final[str] = "A"
ONE: Final[Decimal] = Decimal("1.0")

MONEY_FORMAT: Final[str] = "#,##0.00"
VOLUME_FORMAT: Final[str] = "#,##0.000"
RATIO_FORMAT: Final[str] = "#,##0.0000"
INTEGER_FORMAT: Final[str] = "0"

STATUS_AUTO: Final[str] = "AUTO"
STATUS_APPROVED: Final[str] = "APPROVED"
STATUS_REJECTED: Final[str] = "REJECTED"
STATUS_HITL_PENDING: Final[str] = "HITL_PENDING"

REJECT_REASON_TITLES: Final[dict[str, str]] = {
    "SEMANTIC_MISMATCH": "Семантика: описание относится к другому ресурсу",
    "UNIT_MULTIPLICITY": "Единицы измерения или кратность (например, 100 м)",
    "PARSER_ERROR": "Ошибка парсера: данные извлечены неверно",
    "OTHER": "Иное (см. комментарий эксперта)",
}

COLUMN_TITLES: Final[dict[str, str]] = {
    "A": "ID",
    "B": "№ п/п",
    "C": "Позиция РД (наименование)",
    "D": "Ед. изм. РД",
    "E": "Объем по РД",
    "F": "Шифр расценки (Смета)",
    "G": "Наименование расценки (Смета)",
    "H": "Ед. изм. Сметы",
    "I": "Объем по Смете",
    "J": "Цена за ед., руб.",
    "K": "Стоимость по Смете, руб.",
    "L": "Коэфф. приведения",
    "M": "Объем РД приведенный",
    "N": "Дельта объема",
    "O": "Дельта стоимости, руб.",
    "P": "Статус аудита",
    "Q": "Экспертный комментарий / Причина отклонения",
}

# Ограничения автоподбора ширины: длинные текстовые колонки не должны
# вытеснять числовые данные за пределы экрана.
COLUMN_WIDTH_LIMITS: Final[dict[str, int]] = {
    "A": 12,
    "B": 7,
    "C": 50,
    "D": 10,
    "E": 14,
    "F": 24,
    "G": 50,
    "H": 10,
    "I": 14,
    "J": 15,
    "K": 19,
    "L": 14,
    "M": 18,
    "N": 14,
    "O": 19,
    "P": 16,
    "Q": 50,
}

AUDIT_ROWS_SQL: Final[str] = """
SELECT v.*
FROM v_audit_export AS v
WHERE v.session_id = ?
  AND (
      v.expert_review_id IS NULL
      OR v.expert_review_id = (
          SELECT MAX(er.id)
          FROM expert_reviews AS er
          WHERE er.match_record_id = v.match_id
      )
  )
ORDER BY v.match_id
"""


@dataclass(frozen=True)
class AuditRecord:
    """Нормализованная строка ведомости, готовая к записи в Excel."""

    match_id: int
    rd_description: Optional[str]
    rd_unit: Optional[str]
    rd_quantity: Optional[Decimal]
    smeta_code: Optional[str]
    smeta_description: Optional[str]
    smeta_unit: Optional[str]
    smeta_quantity: Optional[Decimal]
    unit_price: Optional[Decimal]
    total_cost: Optional[Decimal]
    ratio: Decimal
    status: str
    comment: Optional[str]

    @property
    def adjusted_rd_volume(self) -> Optional[Decimal]:
        """Приведенный объем РД: E * L (без нормативов отходов)."""
        if self.rd_quantity is None:
            return None
        return self.rd_quantity * self.ratio

    @property
    def delta_volume(self) -> Optional[Decimal]:
        """Дельта объема: смета минус приведенный объем РД."""
        adjusted = self.adjusted_rd_volume
        if adjusted is None or self.smeta_quantity is None:
            return None
        return self.smeta_quantity - adjusted

    @property
    def tolerance(self) -> Optional[Decimal]:
        """Допустимое расхождение по регламенту ``DiscrepancyDetector``."""
        adjusted = self.adjusted_rd_volume
        if adjusted is None:
            return None
        return (
            DiscrepancyDetector.ABS_TOLERANCE
            + DiscrepancyDetector.REL_TOLERANCE * adjusted
        )

    @property
    def is_within_tolerance(self) -> bool:
        """Признак совпадения объемов в пределах допуска."""
        delta = self.delta_volume
        tolerance = self.tolerance
        if delta is None or tolerance is None:
            return False
        return abs(delta) <= tolerance

    @property
    def is_overrun(self) -> bool:
        """Признак превышения сметы относительно приведенного объема РД."""
        delta = self.delta_volume
        tolerance = self.tolerance
        if delta is None or tolerance is None:
            return False
        return delta > tolerance


# ----------------------------------------------------------------------
# 1. Нормализация значений SQLite -> Excel
# ----------------------------------------------------------------------


def _to_decimal(value: Any) -> Optional[Decimal]:
    """Приводит значение SQLite к Decimal без потери десятичного представления."""
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError):
        return None


def _to_excel_number(value: Any) -> Optional[float]:
    """Возвращает число для записи в ячейку openpyxl.

    openpyxl не поддерживает ``Decimal``, поэтому конвертация выполняется
    только на границе формирования отчета; все сравнения и допуски считаются
    в Decimal.
    """
    number = _to_decimal(value)
    return None if number is None else float(number)


def _to_text(value: Any) -> Optional[str]:
    """Возвращает непустую строку либо ``None``."""
    if value is None:
        return None
    text = str(value).strip()
    return text or None


# ----------------------------------------------------------------------
# 2. Чтение SSOT и нормализация строк отчета
# ----------------------------------------------------------------------


def resolve_session_id(conn: sqlite3.Connection, session_id: str) -> str:
    """Разрешает ``latest`` в идентификатор последней созданной сессии.

    :raises LookupError: в базе нет сессий.
    """
    if session_id != LATEST_SESSION:
        return session_id

    row = conn.execute(
        "SELECT id FROM sessions ORDER BY created_at DESC, rowid DESC LIMIT 1"
    ).fetchone()
    if row is None:
        raise LookupError("В базе данных нет ни одной сессии аудита")
    return str(row["id"])


def _load_session(conn: sqlite3.Connection, session_id: str) -> sqlite3.Row:
    """Возвращает метаданные сессии.

    :raises LookupError: сессия с указанным идентификатором не найдена.
    """
    row = conn.execute(
        """
        SELECT id, project_name, rd_filename, smeta_filename, created_at, status
        FROM sessions
        WHERE id = ?
        """,
        (session_id,),
    ).fetchone()
    if row is None:
        raise LookupError(f"Сессия '{session_id}' не найдена")
    return row


def _effective_ratio(row: sqlite3.Row) -> Decimal:
    """Экспертный коэффициент приоритетнее исходного коэффициента матчинга."""
    adjusted = _to_decimal(row["adjusted_ratio"])
    if adjusted is not None and adjusted > 0:
        return adjusted
    ratio = _to_decimal(row["conversion_ratio"])
    if ratio is not None and ratio > 0:
        return ratio
    return ONE


def _audit_status(row: sqlite3.Row) -> str:
    """Сводит статус связи и решение эксперта к статусу ведомости.

    ``ADJUSTED`` показывается как ``APPROVED``: эксперт принял связь после
    уточнения коэффициента. Аномальный ``RESOLVED`` без решения эксперта
    возвращается в очередь, чтобы не маскировать нарушение пайплайна.
    """
    match_status = _to_text(row["match_status"]) or STATUS_HITL_PENDING
    decision = _to_text(row["decision"])

    if decision == "REJECTED":
        return STATUS_REJECTED
    if decision in {"APPROVED", "ADJUSTED"}:
        return STATUS_APPROVED
    if match_status == STATUS_AUTO:
        return STATUS_AUTO
    return STATUS_HITL_PENDING


def _compose_comment(row: sqlite3.Row) -> Optional[str]:
    """Склеивает причину отклонения и экспертный комментарий для колонки Q."""
    parts: list[str] = []
    reason_code = _to_text(row["reject_reason_code"])
    if reason_code:
        reason_title = REJECT_REASON_TITLES.get(reason_code, reason_code)
        parts.append(f"{reason_code}: {reason_title}")

    comment = _to_text(row["review_comment"])
    if comment:
        parts.append(comment)
    return ". ".join(parts) if parts else None


def _record_from_row(row: sqlite3.Row) -> AuditRecord:
    """Преобразует строку ``v_audit_export`` в модель ведомости."""
    return AuditRecord(
        match_id=int(row["match_id"]),
        rd_description=_to_text(row["rd_description"]),
        rd_unit=_to_text(row["rd_unit"]),
        rd_quantity=_to_decimal(row["rd_quantity"]),
        smeta_code=_to_text(row["smeta_code"]),
        smeta_description=_to_text(row["smeta_description"]),
        smeta_unit=_to_text(row["smeta_unit"]),
        smeta_quantity=_to_decimal(row["smeta_quantity"]),
        unit_price=_to_decimal(row["unit_price"]),
        total_cost=_to_decimal(row["total_cost"]),
        ratio=_effective_ratio(row),
        status=_audit_status(row),
        comment=_compose_comment(row),
    )


def load_audit_records(
    conn: sqlite3.Connection,
    session_id: str,
) -> list[AuditRecord]:
    """Возвращает строки ведомости по сессии (по одной на связь)."""
    rows = conn.execute(AUDIT_ROWS_SQL, (session_id,)).fetchall()
    return [_record_from_row(row) for row in rows]


# ----------------------------------------------------------------------
# 3. Стили и оформление листа
# ----------------------------------------------------------------------

_HEADER_FONT: Final[Font] = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
_HEADER_FILL: Final[PatternFill] = PatternFill("solid", fgColor="1F4E78")
_HEADER_ALIGNMENT: Final[Alignment] = Alignment(
    horizontal="center", vertical="center", wrap_text=True
)
_META_LABEL_FONT: Final[Font] = Font(name="Calibri", size=11, bold=True)
_META_VALUE_FONT: Final[Font] = Font(name="Calibri", size=11)
_PROJECT_FONT: Final[Font] = Font(name="Calibri", size=14, bold=True, color="1F4E78")
_BODY_FONT: Final[Font] = Font(name="Calibri", size=10)
_THIN_SIDE: Final[Side] = Side(style="thin", color="B7C3D0")
_CELL_BORDER: Final[Border] = Border(
    left=_THIN_SIDE,
    right=_THIN_SIDE,
    top=_THIN_SIDE,
    bottom=_THIN_SIDE,
)
_GREEN_FILL: Final[PatternFill] = PatternFill("solid", fgColor="C6EFCE")
_RED_FILL: Final[PatternFill] = PatternFill("solid", fgColor="FFC7CE")
_YELLOW_FILL: Final[PatternFill] = PatternFill("solid", fgColor="FFEB9C")
_WHITE_FILL: Final[PatternFill] = PatternFill("solid", fgColor="FFFFFF")

_STATUS_STYLES: Final[dict[str, tuple[PatternFill, Font]]] = {
    STATUS_AUTO: (
        _GREEN_FILL,
        Font(name="Calibri", size=10, bold=True, color="006100"),
    ),
    STATUS_APPROVED: (
        _GREEN_FILL,
        Font(name="Calibri", size=10, bold=True, color="006100"),
    ),
    STATUS_REJECTED: (
        _RED_FILL,
        Font(name="Calibri", size=10, bold=True, color="9C0006"),
    ),
    STATUS_HITL_PENDING: (
        _YELLOW_FILL,
        Font(name="Calibri", size=10, bold=True, color="9C6500"),
    ),
}

def _write_metadata(
    worksheet: Worksheet,
    session: sqlite3.Row,
    generated_at: datetime,
) -> None:
    """Записывает шапку документа: проект, сессию и дату формирования."""
    worksheet["B1"] = "Проект:"
    worksheet["C1"] = str(session["project_name"])
    worksheet.merge_cells("C1:Q1")

    worksheet["B2"] = "ID сессии:"
    worksheet["C2"] = str(session["id"])
    worksheet.merge_cells("C2:Q2")

    worksheet["B3"] = "Дата формирования:"
    worksheet["C3"] = generated_at.strftime("%d.%m.%Y %H:%M")
    worksheet.merge_cells("C3:Q3")

    for row in (1, 2, 3):
        worksheet[f"B{row}"].font = _META_LABEL_FONT
        worksheet[f"B{row}"].alignment = Alignment(horizontal="left", vertical="center")
        worksheet[f"C{row}"].font = _PROJECT_FONT if row == 1 else _META_VALUE_FONT
        worksheet[f"C{row}"].alignment = Alignment(horizontal="left", vertical="center")

    worksheet.row_dimensions[1].height = 24
    worksheet.row_dimensions[2].height = 18
    worksheet.row_dimensions[3].height = 18


def _write_table_header(worksheet: Worksheet) -> None:
    """Формирует строку заголовков таблицы с закреплением под ней."""
    for column, title in COLUMN_TITLES.items():
        cell = worksheet[f"{column}{HEADER_ROW}"]
        cell.value = title
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _CELL_BORDER
    worksheet.row_dimensions[HEADER_ROW].height = 48


def _column_alignment(column: str) -> Alignment:
    """Возвращает выравнивание ячейки в зависимости от типа данных колонки."""
    if column in {"C", "F", "G", "Q"}:
        return Alignment(horizontal="left", vertical="top", wrap_text=True)
    if column in {"A", "B", "D", "H", "P"}:
        return Alignment(horizontal="center", vertical="center", wrap_text=True)
    return Alignment(horizontal="right", vertical="top")


def _apply_number_format(worksheet: Worksheet, row: int) -> None:
    """Назначает числовые форматы колонкам объемов, цен и коэффициентов."""
    for column in ("E", "I", "M", "N"):
        worksheet[f"{column}{row}"].number_format = VOLUME_FORMAT
    for column in ("J", "K", "O"):
        worksheet[f"{column}{row}"].number_format = MONEY_FORMAT
    worksheet[f"L{row}"].number_format = RATIO_FORMAT
    for column in ("A", "B"):
        worksheet[f"{column}{row}"].number_format = INTEGER_FORMAT


def _write_record(
    worksheet: Worksheet,
    row: int,
    index: int,
    record: AuditRecord,
) -> None:
    """Записывает одну строку ведомости и нативные формулы расчета."""
    worksheet[f"A{row}"] = record.match_id
    worksheet[f"B{row}"] = index
    worksheet[f"C{row}"] = record.rd_description
    worksheet[f"D{row}"] = record.rd_unit
    worksheet[f"E{row}"] = _to_excel_number(record.rd_quantity)
    worksheet[f"F{row}"] = record.smeta_code
    worksheet[f"G{row}"] = record.smeta_description
    worksheet[f"H{row}"] = record.smeta_unit
    worksheet[f"I{row}"] = _to_excel_number(record.smeta_quantity)
    worksheet[f"J{row}"] = _to_excel_number(record.unit_price)

    if record.smeta_quantity is not None and record.unit_price is not None:
        worksheet[f"K{row}"] = f"=I{row}*J{row}"
    else:
        worksheet[f"K{row}"] = _to_excel_number(record.total_cost)

    worksheet[f"L{row}"] = _to_excel_number(record.ratio)

    if record.rd_quantity is not None:
        worksheet[f"M{row}"] = f"=E{row}*L{row}"
    if (
        record.status != STATUS_REJECTED
        and record.smeta_quantity is not None
        and record.rd_quantity is not None
    ):
        worksheet[f"N{row}"] = f"=I{row}-M{row}"
        if record.unit_price is not None:
            worksheet[f"O{row}"] = f"=N{row}*J{row}"

    worksheet[f"P{row}"] = record.status
    worksheet[f"Q{row}"] = record.comment

    for column in COLUMN_TITLES:
        cell = worksheet[f"{column}{row}"]
        cell.font = _BODY_FONT
        cell.border = _CELL_BORDER
        cell.alignment = _column_alignment(column)

    _apply_number_format(worksheet, row)

    status_fill, status_font = _STATUS_STYLES.get(
        record.status, (_WHITE_FILL, _BODY_FONT)
    )
    worksheet[f"P{row}"].fill = status_fill
    worksheet[f"P{row}"].font = status_font

    if record.is_within_tolerance or record.is_overrun:
        delta_fill = _GREEN_FILL if record.is_within_tolerance else _RED_FILL
        for column in ("N", "O"):
            worksheet[f"{column}{row}"].fill = delta_fill


def _autosize_columns(worksheet: Worksheet) -> None:
    """Подбирает ширину колонок по контенту с ограничением для описаний."""
    for column, max_width in COLUMN_WIDTH_LIMITS.items():
        longest = len(COLUMN_TITLES[column])
        for row in range(HEADER_ROW, worksheet.max_row + 1):
            value = worksheet[f"{column}{row}"].value
            if value is None:
                continue
            text = str(value)
            if text.startswith("="):
                # Длину формулы Excel отображает результатом, а не текстом.
                continue
            longest = max(longest, *(len(line) for line in text.splitlines()))
        worksheet.column_dimensions[column].width = min(max(longest + 2, 8), max_width)


# ----------------------------------------------------------------------
# 4. Публичный API и точка входа CLI
# ----------------------------------------------------------------------


def export_audit_to_excel(
    conn: sqlite3.Connection,
    session_id: str,
    output_path: str | Path,
) -> Path:
    """Формирует ведомость аудита объемов и сохраняет ее в XLSX.

    :param conn: соединение с SQLite (SSOT).
    :param session_id: идентификатор сессии либо ``latest``.
    :param output_path: путь к создаваемому файлу отчета.
    :return: путь к сохраненному файлу.
    :raises LookupError: сессия не найдена.
    """
    resolved_id = resolve_session_id(conn, session_id)
    session = _load_session(conn, resolved_id)
    records = load_audit_records(conn, resolved_id)
    generated_at = datetime.now()

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = SHEET_TITLE
    worksheet.column_dimensions[HIDDEN_ID_COLUMN].hidden = True

    _write_metadata(worksheet, session, generated_at)
    _write_table_header(worksheet)

    for offset, record in enumerate(records):
        _write_record(worksheet, FIRST_DATA_ROW + offset, offset + 1, record)

    _autosize_columns(worksheet)

    # Служебные строки и колонки остаются на экране при прокрутке.
    worksheet.freeze_panes = f"C{FIRST_DATA_ROW}"
    worksheet.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
    worksheet.sheet_view.showGridLines = False
    # Excel пересчитывает нативные формулы при открытии книги.
    workbook.calculation.fullCalcOnLoad = True

    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(target)
    return target


def build_parser() -> argparse.ArgumentParser:
    """Создает парсер аргументов командной строки экспортера."""
    parser = argparse.ArgumentParser(
        prog="exporter.py",
        description=(
            "Формирование ведомости аудита объемов в Excel "
            "с нативными формулами проверки"
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--db",
        dest="db_path",
        default=db.DEFAULT_DB_PATH,
        help="Путь к файлу SQLite (SSOT)",
    )
    parser.add_argument(
        "--session",
        default=LATEST_SESSION,
        help="Идентификатор сессии либо 'latest' для последней созданной",
    )
    parser.add_argument(
        "--out",
        dest="output_path",
        default=DEFAULT_OUTPUT,
        help="Путь к создаваемому XLSX-отчету",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Запускает экспорт ведомости; возвращает код выхода процесса."""
    args = build_parser().parse_args(argv)
    db_path = Path(args.db_path)

    if not db_path.is_file():
        print(f"Ошибка: файл БД не найден: {db_path}")
        return 2

    conn: Optional[sqlite3.Connection] = None
    try:
        conn = db.get_connection(db_path)
        db.init_db(conn)
        try:
            resolved_id = resolve_session_id(conn, args.session)
            target = export_audit_to_excel(conn, resolved_id, args.output_path)
            row_count = len(load_audit_records(conn, resolved_id))
        except LookupError as exc:
            print(f"Ошибка: {exc}")
            return 2
    except sqlite3.Error as exc:
        print(f"Ошибка SQLite: {exc}")
        return 1
    except OSError as exc:
        print(f"Не удалось сохранить отчет: {exc}")
        return 1
    finally:
        if conn is not None:
            conn.close()

    print(f"Ведомость сформирована: {target} (строк: {row_count})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())




