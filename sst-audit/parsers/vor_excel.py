"""Парсер дефектных ведомостей и ВОР из XLSX и фабрика парсера РД."""

from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
from typing import Any, Callable, Final, List, Optional, Sequence

import openpyxl

from models import CellPointer, NormalizedItem
from normalizer import DimensionLockExtractor, UnitNormalizer

from .common import SourceType, _is_binary_source, _prepare_source
from .spec_pdf import PDFSpecParser

_VOR_HEADER_SCAN_LIMIT: Final[int] = 80
_VOR_NUMBER_PATTERN: Final[re.Pattern[str]] = re.compile(r"[-+]?\d+(?:[.,]\d+)?")
# Сноски в ячейках объема: «12,5*», «7,2¹». Снимаются до разбора числа.
_VOR_FOOTNOTE_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"[*∗†‡]+|(?<=\d)[¹²³⁴⁵⁶⁷⁸⁹⁰]+"
)



def _header_key(value: Any) -> str:
    """Ключ заголовка графы: нижний регистр без пробелов и точек."""
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").casefold().replace("ё", "е")
    return re.sub(r"[\s.]+", "", text)



class ExcelVORParser:
    """Парсер дефектных ведомостей и ВОР из XLSX.

    Шапка ищется по корням заголовков, поэтому поддерживаются как формы
    ГОСТ 21.110, так и таблицы заказчика с графами «Наименование»,
    «Ед. изм.», «Кол.»/«Количество»/«Объем». Объемы читаются строго в
    ``Decimal``; арифметика над ``float`` не выполняется даже для числовых
    ячеек openpyxl (число переводится через его строковое представление).

    Позиции сохраняются в существующую модель ``NormalizedItem`` с
    ``doc_type="VOR_EXCEL"``; новых DTO не вводится.
    """

    TITLE_ROOTS: Final[tuple[str, ...]] = ("наименован",)
    UNIT_ROOTS: Final[tuple[str, ...]] = ("едизм", "измерен")
    QTY_ROOTS: Final[tuple[str, ...]] = ("кол", "колич", "объем")

    def __init__(self, source: SourceType, display_name: Optional[str] = None):
        """Запоминает источник ведомости: путь к ``.xlsx`` либо бинарный поток.

        :param source: путь к книге либо открытый бинарный поток (``BytesIO``).
            Поток принадлежит вызывающему коду и не закрывается парсером.
        :param display_name: безопасное имя файла для ``CellPointer.file_name``;
            для потока по умолчанию используется ``in_memory.xlsx``.
        """
        path, self.file_name, self._buffer = _prepare_source(
            source, display_name, fallback_name="in_memory.xlsx"
        )
        self.display_name = self.file_name
        self.file_path = path if path is not None else self.file_name

    @classmethod
    def _find_header(cls, row: Sequence[Any]) -> Optional[tuple[int, int, int]]:
        """Возвращает индексы граф (наименование, единица, количество)."""
        title_col: Optional[int] = None
        unit_col: Optional[int] = None
        qty_col: Optional[int] = None
        for index, value in enumerate(row):
            key = _header_key(value)
            if not key:
                continue
            if title_col is None and any(root in key for root in cls.TITLE_ROOTS):
                title_col = index
            elif unit_col is None and any(root in key for root in cls.UNIT_ROOTS):
                unit_col = index
            elif qty_col is None and any(root in key for root in cls.QTY_ROOTS):
                qty_col = index
        if title_col is None or unit_col is None or qty_col is None:
            return None
        return title_col, unit_col, qty_col

    @staticmethod
    def _cell_value(row: Sequence[Any], index: int) -> Any:
        """Значение ячейки с защитой от коротких строк read-only итератора."""
        return row[index] if 0 <= index < len(row) else None

    @classmethod
    def _cell_text(cls, row: Sequence[Any], index: int) -> str:
        """Текст ячейки без артефактов Excel и лишних пробелов."""
        value = cls._cell_value(row, index)
        if value is None:
            return ""
        text = (
            str(value)
            .replace("_x000D_", " ")
            .replace("\xa0", " ")
            .replace("\u2009", " ")
            .replace("\u202f", " ")
        )
        return re.sub(r"\s+", " ", text).strip()

    @staticmethod
    def _parse_volume(value: Any) -> Optional[Decimal]:
        """Объем ячейки строго в ``Decimal``: пробелы и сноски удаляются."""
        if value is None or isinstance(value, bool):
            return None
        if isinstance(value, Decimal):
            return value
        if isinstance(value, int):
            return Decimal(value)
        if isinstance(value, float):
            # Единственное допустимое соприкосновение с float: openpyxl
            # отдает числовые ячейки как float. Арифметика не выполняется.
            return Decimal(str(value))

        text = (
            str(value)
            .replace("\xa0", " ")
            .replace("\u2009", " ")
            .replace("\u202f", " ")
            .replace("²", "2")
            .replace("³", "3")
        )
        text = re.sub(r"\s+", "", text)
        text = _VOR_FOOTNOTE_PATTERN.sub("", text)
        match = _VOR_NUMBER_PATTERN.search(text)
        if match is None:
            return None
        try:
            return Decimal(match.group(0).replace(",", "."))
        except InvalidOperation:
            return None

    def extract_items(self) -> List[NormalizedItem]:
        """Извлекает позиции ведомости по всем листам книги.

        Строка шапки определяется по корням заголовков в первых
        ``_VOR_HEADER_SCAN_LIMIT`` строках листа. Строки разделов
        («Раздел 1. Демонтаж») задают ``context_section``, строки без
        числового объема (подзаголовки помещений, примечания) пропускаются.
        """
        workbook = openpyxl.load_workbook(
            self._buffer if self._buffer is not None else str(self.file_path),
            read_only=True,
            data_only=True,
        )
        items: List[NormalizedItem] = []
        try:
            for worksheet in workbook.worksheets:
                header: Optional[tuple[int, int, int]] = None
                section = re.sub(r"\s+", " ", str(worksheet.title)).strip()
                for row_idx, row in enumerate(
                    worksheet.iter_rows(values_only=True), start=1
                ):
                    if header is None:
                        if row_idx > _VOR_HEADER_SCAN_LIMIT:
                            break
                        header = self._find_header(row)
                        continue

                    row_head = " ".join(
                        self._cell_text(row, index)
                        for index in range(min(len(row), 3))
                    )
                    if "раздел" in row_head.casefold():
                        section = row_head.strip()

                    if self._find_header(row) is not None:
                        continue

                    raw_title = self._cell_text(row, header[0])
                    if not raw_title:
                        continue
                    volume = self._parse_volume(
                        self._cell_value(row, header[2])
                    )
                    if volume is None:
                        continue

                    items.append(
                        self._build_item(
                            row,
                            header,
                            raw_title,
                            volume,
                            section,
                            worksheet.title,
                            row_idx,
                            len(items) + 1,
                        )
                    )
        finally:
            workbook.close()
        return items

    def parse(self) -> List[NormalizedItem]:
        """Совместимый с ``PDFSpecParser`` вход для пайплайна."""
        return self.extract_items()

    def _build_item(
        self,
        row: Sequence[Any],
        header: tuple[int, int, int],
        raw_title: str,
        volume: Decimal,
        section: str,
        sheet_name: str,
        row_idx: int,
        sequence: int,
    ) -> NormalizedItem:
        """Создает ``NormalizedItem`` из строки ведомости Excel."""
        title_col, unit_col, qty_col = header
        raw_unit = self._cell_text(row, unit_col)
        canonical_unit, normalized_volume, multiplier = UnitNormalizer.normalize(
            raw_unit, volume
        )
        clean_title = re.sub(r"\s+", " ", raw_title).lower().strip()
        dimensions = DimensionLockExtractor.extract(raw_title)

        title_letter = openpyxl.utils.get_column_letter(title_col + 1)
        qty_letter = openpyxl.utils.get_column_letter(qty_col + 1)
        pointer = CellPointer(
            doc_type="VOR_EXCEL",
            file_name=self.file_name,
            sheet_name=str(sheet_name),
            row_idx=row_idx,
            col_letter=title_letter,
            coordinate=f"{qty_letter}{row_idx}",
        )
        return NormalizedItem(
            raw_id=f"VOR_EXCEL_{sequence}",
            code=None,
            raw_title=raw_title,
            clean_title=clean_title,
            raw_unit=raw_unit,
            canonical_unit=canonical_unit,
            unit_multiplier=multiplier,
            volume=normalized_volume,
            dimensions=dimensions,
            context_section=section,
            pointer=pointer,
        )


def get_spec_parser(
    source: SourceType,
    display_name: Optional[str] = None,
    *,
    pdf_parser_cls: Callable[..., Any] = PDFSpecParser,
) -> "PDFSpecParser | ExcelVORParser":
    """Фабрика парсера РД/ВОР по расширению файла или по имени потока.

    ``.xlsx`` (дефектная ведомость или ВОР) разбирается ``ExcelVORParser``,
    остальные файлы — существующим ``PDFSpecParser``. Класс PDF-парсера
    передается параметром, чтобы вызывающий код мог подменить его в тестах.

    :param source: путь к файлу либо открытый бинарный поток (``BytesIO``).
    :param display_name: безопасное имя файла. Для пути это подпись в отчетах,
        для потока — единственный источник расширения (``upload.xlsx``);
        без имени поток трактуется как PDF.
    :param pdf_parser_cls: фабрика PDF-парсера (по умолчанию ``PDFSpecParser``).
    :return: экземпляр парсера с методом ``extract_items()``.
    """
    if not _is_binary_source(source):
        candidate = Path(source)
        suffix = candidate.suffix.lower()
        if suffix == ".xlsx":
            return ExcelVORParser(candidate, display_name)
        if display_name:
            return pdf_parser_cls(str(candidate), display_name=display_name)
        return pdf_parser_cls(str(candidate))

    stream_suffix = Path(str(display_name or "")).suffix.lower()
    if stream_suffix == ".xlsx":
        return ExcelVORParser(source, display_name)
    if display_name:
        return pdf_parser_cls(source, display_name=display_name)
    return pdf_parser_cls(source)
