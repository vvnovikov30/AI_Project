"""Парсер локальных смет из векторных PDF (форма 4б)."""

from bisect import bisect_right
from dataclasses import dataclass, replace
from decimal import Decimal
import re
from typing import Any, Dict, Final, List, Optional, Tuple

import openpyxl
import pdfplumber

from models import CellPointer, NormalizedItem
from normalizer import DimensionLockExtractor, UnitNormalizer

from .common import (
    SourceType,
    _flush_page_cache,
    _parse_pdf_number,
    _prepare_source,
    _repair_pdf_text,
    _squash_spaces,
)
from .lsr_excel import LSRParser

_PDF_LSR_INT_TOKEN_PATTERN: re.Pattern[str] = re.compile(r"^\d{1,2}$")
_PDF_LSR_POSITION_PATTERN: re.Pattern[str] = re.compile(r"^\d{1,3}(?:[.,]\d{1,2})?$")
_PDF_LSR_LABOR_CODE_PATTERN: re.Pattern[str] = re.compile(r"^\d{1,2}-\d{3}-\d{2,3}$")
_PDF_LSR_MACHINE_CODE_PATTERN: re.Pattern[str] = re.compile(r"^91\.")
# Шифры расценок: сборники ГЭСН/ФЕР/ТЕР/ТСН, ФСНБ и прейскурантные позиции.
_PDF_LSR_RATE_CODE_PATTERN: re.Pattern[str] = re.compile(
    r"^(?:ГЭСН|ФЕР|ТЕР|ТСН|ФСНБ|ФСЭМ|Прайс-лист|[3-6]\.[\d\-]+$)",
    re.IGNORECASE,
)
# Шифры ресурсов: ФСБЦ/ФССЦ, ценники 999-хххх, код ФСБЦ вида 01.7.15.03-0012.
_PDF_LSR_RESOURCE_CODE_PATTERN: re.Pattern[str] = re.compile(
    r"^(?:ФСБЦ|ФССЦ|999-\d{4}|(?:1|2|13)\.[\d\-]+$|"
    r"\d{2}\.\d{1,2}\.\d{1,2}(?:\.\d{1,2})?-\d{3,4}$)",
    re.IGNORECASE,
)
_PDF_LSR_SECTION_PATTERN: re.Pattern[str] = re.compile(
    r"^раздел\b(?!\s*пд)", re.IGNORECASE
)
_PDF_LSR_SERVICE_PATTERN: re.Pattern[str] = re.compile(
    r"^(?:итого|всего|итоги|фот\b|нр\b|нр\s|сп\b|сп\s|сп\(|зтр\b|тр\b|"
    r"в т\.ч\.|в том числе|накладные|сметная прибыль|ндс|прямые затраты|"
    r"затраты труда|оплата труда|эксплуатация машин|материальные ресурсы|"
    r"перевозка\b|справочно\b|оборудование,\s*отсутствующее|"
    r"раздел\s+пд\b|книга\b|пункт\b|(?:от|отм|эм|эмм|м)$)",
    re.IGNORECASE,
)
_PDF_LSR_LINE_TOLERANCE: float = 1.5


@dataclass(frozen=True)
class _LSRColumnLayout:
    """Геометрия таблицы сметы: границы колонок и их назначение.

    :param bounds: возрастающие x-координаты вертикальных линий.
    :param position_col: колонка «№ п.п.».
    :param code_col: колонка «Обоснование» (шифр расценки/ресурса).
    :param title_col: колонка «Наименование».
    :param unit_col: колонка «Единица измерения».
    :param qty_unit_col: колонка «Количество на единицу».
    :param qty_total_col: колонка «Всего с учетом коэффициента».
    :param base_price_col: колонка базисной цены (может отсутствовать).
    :param unit_price_col: колонка цены в текущем уровне цен.
    :param total_cost_col: колонка «Всего затрат».
    """

    bounds: Tuple[float, ...]
    position_col: int
    code_col: int
    title_col: int
    unit_col: int
    qty_unit_col: Optional[int]
    qty_total_col: int
    base_price_col: Optional[int]
    unit_price_col: Optional[int]
    total_cost_col: int

    @property
    def column_count(self) -> int:
        """Возвращает количество колонок таблицы."""
        return len(self.bounds) - 1



class PDFLSRParser:
    """Парсер локальной сметы из векторного PDF (форма 4б, 11-12 граф).

    Парсер работает только с PDF, содержащими текстовый слой: строки
    собираются из отдельных слов ``pdfplumber``, поэтому OCR не требуется.
    Геометрия таблицы определяется динамически для каждой страницы:
    границы колонок берутся из вертикальных линий, при их отсутствии —
    из ячеек ``pdfplumber`` (``vertical_strategy='lines'``) либо из
    кластеризации левых краев слов тела таблицы. Строка-шапка ищется по
    последовательным номерам граф или по ключевым заголовкам.

    Тип строки определяется по шифру обоснования: ГЭСН/ФЕР/ТЕР/ТСН/ФСНБ и
    прейскурант «Прайс-лист» — расценки (родительские позиции), а ФСБЦ/ФССЦ,
    ценники и коды машин/труда — ресурсы. Для строк без распознанного шифра
    действует резервная эвристика по номеру позиции (целое число — расценка,
    вложенный номер — ресурс). Ресурсы получают ``is_resource=True`` и
    ``parent_id`` ближайшей расценки; самостоятельные ресурсные позиции
    с целым номером остаются без родителя.
    """

    # Минимум числовых токенов в строке-шапке «1 2 3 ... 12».
    MIN_HEADER_TOKENS = 6
    # Минимум колонок таблицы для распознавания сметной формы.
    MIN_COLUMNS = 8
    # Максимум колонок: больше — это уже рамки/штампы, а не сетка сметы.
    MAX_COLUMNS = 24
    # Допуск кластеризации вертикальных линий таблицы, пт.
    EDGE_TOLERANCE = 2.0
    # Допуск кластеризации левых краев слов для смет без линий, пт.
    TEXT_ALIGN_TOLERANCE = 8.0
    # Минимальная высота вертикальной линии, чтобы считать ее границей колонки.
    MIN_EDGE_HEIGHT = 3.0
    # Ключевые заголовки шапки для форм без нумерации граф.
    HEADER_TOKENS: Final[Tuple[str, ...]] = (
        "обоснование",
        "наименование",
        "единица измерения",
        "количество",
    )

    def __init__(self, source: SourceType, display_name: Optional[str] = None):
        """Запоминает источник сметы: путь к PDF либо открытый бинарный поток.

        :param source: путь к PDF либо бинарный поток (``BytesIO``). Поток
            принадлежит вызывающему коду: парсер читает его копию и не закрывает
            исходный поток.
        :param display_name: безопасное имя файла для ``CellPointer.file_name``;
            для потока по умолчанию используется ``in_memory.pdf``.
        """
        path, self.file_name, self._buffer = _prepare_source(
            source, display_name, fallback_name="in_memory.pdf"
        )
        self.display_name = self.file_name
        self.file_path = str(path) if path is not None else self.file_name

    def _pdf_stream(self) -> Any:
        """Источник для ``pdfplumber``: путь к файлу либо буфер с позиции 0."""
        if self._buffer is None:
            return self.file_path
        self._buffer.seek(0)
        return self._buffer

    def parse(self, sheet_name: Optional[str] = None) -> List[NormalizedItem]:
        """Разбирает PDF-смету целиком, страница за страницей.

        :param sheet_name: не используется: у PDF нет листов; параметр
            сохранен для единого интерфейса сметных парсеров.
        :return: список нормализованных позиций сметы.
        """
        del sheet_name
        items: List[NormalizedItem] = []
        layout: Optional[_LSRColumnLayout] = None
        section = "Default_Section"
        parent_id: Optional[str] = None
        pending_idx: Optional[int] = None
        last_idx: Optional[int] = None

        with pdfplumber.open(self._pdf_stream()) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                try:
                    lines = self._page_lines(page)
                    page_layout, body_top = self._page_layout(page, lines, layout)
                    if page_layout is None:
                        continue
                    layout = page_layout

                    for line in lines:
                        if line[0]["top"] < body_top:
                            continue
                        cells = self._row_cells(line, layout)
                        position_text = cells.get(layout.position_col, "")
                        code_text = cells.get(layout.code_col, "")
                        title_text = cells.get(layout.title_col, "")

                        if _PDF_LSR_SECTION_PATTERN.match(
                            position_text or title_text
                        ):
                            section = _squash_spaces(
                                " ".join(
                                    part
                                    for part in (
                                        position_text,
                                        code_text,
                                        title_text,
                                    )
                                    if part
                                )
                            )
                            last_idx = None
                            continue

                        if LSRParser.TOTAL_ROW_PATTERN.match(title_text):
                            total = self._cell_number(
                                cells, layout.total_cost_col
                            )
                            if total is None:
                                total = self._cell_number(
                                    cells, layout.unit_price_col
                                )
                            if total is not None and pending_idx is not None:
                                self._apply_position_total(
                                    items[pending_idx], total
                                )
                                pending_idx = None
                            last_idx = None
                            continue

                        if _PDF_LSR_SERVICE_PATTERN.match(title_text):
                            last_idx = None
                            continue

                        position_is_number = bool(
                            _PDF_LSR_POSITION_PATTERN.match(position_text)
                        )
                        if not position_is_number and not code_text:
                            if title_text and last_idx is not None:
                                self._append_title(items[last_idx], title_text)
                            continue

                        unit_text = _squash_spaces(cells.get(layout.unit_col, ""))
                        quantity = self._cell_number(cells, layout.qty_total_col)
                        if quantity is None:
                            quantity = self._cell_number(
                                cells, layout.qty_unit_col
                            )
                        if (
                            code_text
                            and last_idx is not None
                            and not unit_text
                            and quantity is None
                            and not title_text
                        ):
                            # Шифр перенесен на следующую физическую строку
                            # («ГЭСНр 65-02-» + «008-01»).
                            self._append_code(items[last_idx], code_text)
                            continue
                        if not unit_text or quantity is None:
                            last_idx = None
                            continue
                        if _PDF_LSR_LABOR_CODE_PATTERN.match(
                            code_text
                        ) or _PDF_LSR_MACHINE_CODE_PATTERN.match(code_text):
                            last_idx = None
                            continue

                        is_parent = self._is_rate_row(code_text, position_text)
                        is_nested = bool(
                            not is_parent
                            and parent_id is not None
                            and not (
                                position_is_number
                                and not re.search(r"[.,]", position_text)
                            )
                        )
                        sequence = len(items) + 1
                        raw_id = f"PDF_LSR_{sequence}"
                        canonical_unit, volume, multiplier = (
                            UnitNormalizer.normalize(unit_text, quantity)
                        )
                        pointer = CellPointer(
                            doc_type="LSR",
                            file_name=self.file_name,
                            sheet_name=f"Page {page_number}",
                            row_idx=sequence,
                            col_letter=openpyxl.utils.get_column_letter(
                                layout.position_col + 1
                            ),
                            coordinate=f"p.{page_number} [row {sequence}]",
                        )
                        items.append(
                            NormalizedItem(
                                raw_id=raw_id,
                                code=self._normalize_code(code_text) or None,
                                raw_title=title_text,
                                clean_title=LSRParser._clean_estimate_title(
                                    title_text
                                ),
                                raw_unit=unit_text,
                                canonical_unit=canonical_unit,
                                unit_multiplier=multiplier,
                                volume=volume,
                                dimensions=DimensionLockExtractor.extract(
                                    title_text
                                ),
                                context_section=section,
                                pointer=pointer,
                                is_resource=not is_parent,
                                parent_id=parent_id if is_nested else None,
                                unit_price=self._cell_number(
                                    cells, layout.unit_price_col
                                ),
                                total_cost=self._cell_number(
                                    cells, layout.total_cost_col
                                ),
                            )
                        )
                        if is_parent:
                            parent_id = raw_id
                            pending_idx = len(items) - 1
                        elif not is_nested:
                            # Самостоятельная ресурсная позиция без расценки:
                            # блок «Всего по позиции» относится к ней самой.
                            pending_idx = len(items) - 1
                        last_idx = len(items) - 1
                finally:
                    # Кэш объектов страницы pdfplumber держит все линии/символы
                    # листа: освобождаем его сразу после обработки страницы.
                    _flush_page_cache(page)

        return items

    def _page_lines(self, page: Any) -> List[List[Dict[str, Any]]]:
        """Собирает слова страницы в физические строки с восстановлением кодировки.

        :param page: страница ``pdfplumber``.
        :return: список строк, каждая строка — список слов с координатами.
        """
        words: List[Dict[str, Any]] = []
        for raw_word in page.extract_words():
            text = _repair_pdf_text(str(raw_word.get("text", "")))
            if not text:
                continue
            words.append(
                {
                    "text": text,
                    "x0": float(raw_word["x0"]),
                    "x1": float(raw_word["x1"]),
                    "top": float(raw_word["top"]),
                    "bottom": float(raw_word["bottom"]),
                }
            )
        words.sort(key=lambda word: (round(float(word["top"]), 1), float(word["x0"])))

        lines: List[List[Dict[str, Any]]] = []
        current: List[Dict[str, Any]] = []
        line_top: Optional[float] = None
        for word in words:
            top = float(word["top"])
            if line_top is None or abs(top - line_top) > _PDF_LSR_LINE_TOLERANCE:
                if current:
                    lines.append(current)
                current = [word]
                line_top = top
            else:
                current.append(word)
        if current:
            lines.append(current)
        return lines

    def _page_layout(
        self,
        page: Any,
        lines: List[List[Dict[str, Any]]],
        previous: Optional[_LSRColumnLayout],
    ) -> Tuple[Optional[_LSRColumnLayout], float]:
        """Определяет геометрию таблицы и верхнюю границу тела таблицы.

        :param page: страница ``pdfplumber``.
        :param lines: строки страницы из :meth:`_page_lines`.
        :param previous: геометрия предыдущей страницы (для продолжений
            таблицы без повторной шапки).
        :return: пара «геометрия колонок, top начала тела таблицы»;
            ``(None, 0.0)``, если страница не является таблицей сметы.
        """
        header_index = self._header_line_index(lines)
        if header_index is None:
            if previous is not None:
                return previous, 0.0
            return None, 0.0

        header_line = lines[header_index]
        header_top = min(float(word["top"]) for word in header_line)
        body_top = max(float(word["bottom"]) for word in header_line)
        bounds = self._column_bounds(page, lines, header_index, header_top)
        if len(bounds) < self.MIN_COLUMNS:
            if previous is not None:
                return previous, body_top
            return None, 0.0

        band = [
            word
            for line in lines
            for word in line
            if float(word["top"]) <= header_top + _PDF_LSR_LINE_TOLERANCE
        ]
        headers = self._header_cells(band, bounds)
        layout = self._layout_from_header(headers, len(bounds) - 1, bounds)
        if layout is not None:
            return layout, body_top
        if previous is not None and len(previous.bounds) == len(bounds):
            return replace(previous, bounds=bounds), body_top
        return None, 0.0

    @classmethod
    def _header_line_index(cls, lines: List[List[Dict[str, Any]]]) -> Optional[int]:
        """Находит строку-шапку сметы.

        Сначала ищется строка с последовательными номерами граф
        «1 2 3 ... N», затем — строка с ключевыми заголовками
        («Обоснование», «Наименование», «Единица измерения»,
        «Количество») для форм без нумерации граф.

        :param lines: строки страницы.
        :return: индекс строки-шапки либо ``None``.
        """
        for index, line in enumerate(lines):
            values = sorted(
                int(str(word["text"]))
                for word in line
                if _PDF_LSR_INT_TOKEN_PATTERN.match(str(word["text"]))
            )
            if len(values) < cls.MIN_HEADER_TOKENS:
                continue
            if values[0] == 1 and values == list(range(1, len(values) + 1)):
                return index

        for index, line in enumerate(lines):
            text = _squash_spaces(
                " ".join(str(word["text"]) for word in line)
            ).lower()
            hits = sum(1 for token in cls.HEADER_TOKENS if token in text)
            if hits >= 2:
                return index
        return None

    def _column_bounds(
        self,
        page: Any,
        lines: List[List[Dict[str, Any]]],
        header_index: int,
        header_top: float,
    ) -> Tuple[float, ...]:
        """Определяет границы колонок таблицы без привязки к координатам файла.

        Порядок источников геометрии:

        1. вертикальные линии таблицы, начинающиеся выше шапки;
        2. все вертикальные линии страницы (если линии шапки не найдены);
        3. ячейки таблиц ``pdfplumber`` (``vertical_strategy='lines'``);
        4. кластеризация левых краев слов тела таблицы (сметы без линий).

        :param page: страница ``pdfplumber``.
        :param lines: строки страницы из :meth:`_page_lines`.
        :param header_index: индекс строки-шапки.
        :param header_top: верхняя координата шапки.
        :return: возрастающий кортеж x-координат границ колонок.
        """
        bounds = self._line_column_bounds(page, header_top)
        if len(bounds) < self.MIN_COLUMNS:
            bounds = self._line_column_bounds(page, None)
        if not self.MIN_COLUMNS <= len(bounds) <= self.MAX_COLUMNS:
            table_bounds = self._table_column_bounds(page)
            if len(table_bounds) >= self.MIN_COLUMNS:
                bounds = table_bounds
        if len(bounds) < self.MIN_COLUMNS:
            bounds = self._text_column_bounds(lines, header_index)
        return bounds

    def _line_column_bounds(
        self, page: Any, top_limit: Optional[float]
    ) -> Tuple[float, ...]:
        """Кластеризует вертикальные линии страницы в границы колонок.

        :param page: страница ``pdfplumber``.
        :param top_limit: учитывать линии, начинающиеся выше этой
            координаты (шапка таблицы); ``None`` — все линии страницы.
        :return: возрастающий кортеж x-координат границ колонок.
        """
        coordinates: List[float] = []
        for edge in page.edges:
            if str(edge.get("orientation")) != "v":
                continue
            top = float(edge["top"])
            bottom = float(edge["bottom"])
            if bottom - top < self.MIN_EDGE_HEIGHT:
                continue
            if top_limit is not None and top >= top_limit:
                continue
            coordinates.append((float(edge["x0"]) + float(edge["x1"])) / 2)
        return self._cluster_coordinates(coordinates, self.EDGE_TOLERANCE)

    def _table_column_bounds(self, page: Any) -> Tuple[float, ...]:
        """Извлекает границы колонок из таблиц ``pdfplumber`` по линиям.

        :param page: страница ``pdfplumber``.
        :return: возрастающий кортеж x-координат границ колонок.
        """
        find_tables = getattr(page, "find_tables", None)
        if find_tables is None:
            return ()
        try:
            tables = find_tables(
                table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                }
            )
        except (TypeError, ValueError):
            return ()
        coordinates: List[float] = []
        for table in tables:
            for cell in getattr(table, "cells", ()):
                if cell is None:
                    continue
                coordinates.extend((float(cell[0]), float(cell[2])))
        return self._cluster_coordinates(coordinates, self.EDGE_TOLERANCE)

    def _text_column_bounds(
        self, lines: List[List[Dict[str, Any]]], header_index: int
    ) -> Tuple[float, ...]:
        """Строит границы колонок по левым краям слов тела таблицы.

        Резервный способ для смет без линий: слова тела выровнены по левому
        краю граф, поэтому кластеризация их ``x0`` дает начала колонок,
        а границы колонок — середины между соседними кластерами.

        :param lines: строки страницы.
        :param header_index: индекс строки-шапки.
        :return: возрастающий кортеж x-координат границ колонок.
        """
        body = lines[header_index + 1 :]
        if not body:
            return ()
        starts = [float(word["x0"]) for line in body for word in line]
        clusters = self._cluster_coordinates(starts, self.TEXT_ALIGN_TOLERANCE)
        if len(clusters) < self.MIN_COLUMNS:
            return ()
        right = max(float(word["x1"]) for line in body for word in line)
        bounds = [clusters[0]]
        for left, next_left in zip(clusters, clusters[1:]):
            bounds.append((left + next_left) / 2)
        bounds.append(right)
        return tuple(bounds)

    @staticmethod
    def _cluster_coordinates(
        coordinates: List[float], tolerance: float
    ) -> Tuple[float, ...]:
        """Схлопывает близкие координаты в упорядоченный набор границ.

        :param coordinates: исходные x-координаты.
        :param tolerance: максимальное расстояние между границами.
        :return: возрастающий кортеж уникальных координат.
        """
        ordered = sorted(coordinates)
        clustered: List[float] = []
        for value in ordered:
            if clustered and abs(value - clustered[-1]) <= tolerance:
                continue
            clustered.append(value)
        return tuple(clustered)

    @staticmethod
    def _header_cells(
        words: List[Dict[str, Any]], bounds: Tuple[float, ...]
    ) -> Dict[int, str]:
        """Раскладывает слова шапки по колонкам и склеивает текст колонки.

        :param words: слова шапки страницы.
        :param bounds: границы колонок.
        :return: отображение «индекс колонки -> текст заголовка».
        """
        cells: Dict[int, List[str]] = {}
        for word in words:
            center = (float(word["x0"]) + float(word["x1"])) / 2
            index = bisect_right(bounds, center) - 1
            if index < 0 or index >= len(bounds) - 1:
                continue
            cells.setdefault(index, []).append(str(word["text"]))
        return {
            index: _squash_spaces(" ".join(parts))
            for index, parts in cells.items()
        }

    @classmethod
    def _layout_from_header(
        cls,
        headers: Dict[int, str],
        column_count: int,
        bounds: Tuple[float, ...],
    ) -> Optional[_LSRColumnLayout]:
        """Сопоставляет текстовые заголовки граф с индексами колонок.

        :param headers: отображение «индекс колонки -> текст шапки».
        :param column_count: количество колонок таблицы.
        :param bounds: x-координаты границ колонок.
        :return: геометрия колонок либо ``None``, если шапка не распознана.
        """
        if column_count < cls.MIN_COLUMNS:
            return None

        def find(*tokens: str, exclude: Tuple[str, ...] = ()) -> Optional[int]:
            for index in range(column_count):
                text = headers.get(index, "").lower()
                if all(token in text for token in tokens) and not any(
                    token in text for token in exclude
                ):
                    return index
            return None

        graph: Dict[int, int] = {}
        for index in range(column_count):
            match = re.search(r"(\d{1,2})\s*$", headers.get(index, ""))
            if match:
                graph[int(match.group(1))] = index

        position = find("№")
        code = find("обоснование")
        title = find("наименование")
        unit = find("единица измерения")
        if title is None:
            title = graph.get(3)
        if unit is None:
            unit = graph.get(4)
        if position is None:
            position = graph.get(1, 0)
        if code is None:
            code = graph.get(2, 1 if column_count > 1 else 0)
        if title is None or unit is None:
            return None

        qty_unit = find("количество", exclude=("всего",))
        if qty_unit is None:
            qty_unit = graph.get(5)
        qty_total = find("всего", "учет")
        if qty_total is None:
            qty_total = find("всего", exclude=("текущем", "базисном"))
        if qty_total is None:
            qty_total = graph.get(7, column_count - 3)
        base_price = find("базисном", exclude=("всего",))
        if base_price is None:
            base_price = graph.get(8)
        unit_price = find("текущем", exclude=("всего",))
        if unit_price is None:
            unit_price = graph.get(10, column_count - 2)
        total_cost = find("текущем", "всего")
        if total_cost is None:
            total_cost = graph.get(12, column_count - 1)

        required = (position, code, title, unit, qty_total, total_cost)
        if any(index is None or index >= column_count for index in required):
            return None
        for optional in (qty_unit, base_price, unit_price):
            if optional is not None and optional >= column_count:
                return None
        if len(set(required)) < len(required):
            return None

        return _LSRColumnLayout(
            bounds=bounds,
            position_col=int(position),
            code_col=int(code),
            title_col=int(title),
            unit_col=int(unit),
            qty_unit_col=qty_unit,
            qty_total_col=int(qty_total),
            base_price_col=base_price,
            unit_price_col=unit_price,
            total_cost_col=int(total_cost),
        )

    @staticmethod
    def _row_cells(
        line: List[Dict[str, Any]], layout: _LSRColumnLayout
    ) -> Dict[int, str]:
        """Раскладывает слова физической строки по колонкам таблицы.

        :param line: слова строки.
        :param layout: геометрия колонок страницы.
        :return: отображение «индекс колонки -> текст ячейки».
        """
        cells: Dict[int, List[str]] = {}
        for word in line:
            x0 = float(word["x0"])
            x1 = float(word["x1"])
            if x0 < layout.bounds[0] - 3 or x1 > layout.bounds[-1] + 3:
                continue
            index = bisect_right(layout.bounds, (x0 + x1) / 2) - 1
            index = max(0, min(index, layout.column_count - 1))
            cells.setdefault(index, []).append(str(word["text"]))
        return {
            index: _squash_spaces(" ".join(parts))
            for index, parts in cells.items()
        }

    @staticmethod
    def _cell_number(
        cells: Dict[int, str], index: Optional[int]
    ) -> Optional[Decimal]:
        """Извлекает число из указанной колонки строки.

        :param cells: ячейки строки.
        :param index: индекс колонки либо ``None``.
        :return: ``Decimal`` либо ``None``.
        """
        if index is None:
            return None
        return _parse_pdf_number(cells.get(index, ""))

    @staticmethod
    def _apply_position_total(item: NormalizedItem, total: Decimal) -> None:
        """Записывает итог «Всего по позиции» в позицию сметы.

        :param item: позиция, к которой относится итог.
        :param total: стоимость по графе «Всего затрат».
        """
        item.total_cost = total
        if item.unit_price is None and item.volume > 0:
            item.unit_price = (total / item.volume).quantize(Decimal("0.0001"))

    @staticmethod
    def _is_rate_row(code: str, position: str) -> bool:
        """Определяет, является ли строка расценкой (родительской позицией).

        Приоритет отдается шифру обоснования из графы 2: сборники
        ГЭСН/ФЕР/ТЕР/ТСН/ФСНБ, главы 3-6 ТСН-2001 и прейскурант
        «Прайс-лист» — расценки, а ФСБЦ/ФССЦ, ценники 999-хххх,
        главы 1/2/13 ТСН-2001 и коды ФСБЦ — ресурсы. Для строк
        без распознанного шифра применяется резервная эвристика по номеру
        позиции: целое число — расценка, вложенный номер (``2.1``) — ресурс.

        :param code: шифр обоснования из графы 2.
        :param position: номер позиции из графы 1.
        :return: ``True`` для расценки, ``False`` для ресурса.
        """
        if code and _PDF_LSR_RATE_CODE_PATTERN.match(code):
            return True
        if code and _PDF_LSR_RESOURCE_CODE_PATTERN.match(code):
            return False
        if not _PDF_LSR_POSITION_PATTERN.match(position):
            return False
        return not re.search(r"[.,]", position)

    @staticmethod
    def _normalize_code(code: str) -> str:
        """Убирает пробелы вокруг дефисов в шифре, разорванном переносом.

        :param code: шифр обоснования из графы 2.
        :return: нормализованный шифр.
        """
        return re.sub(r"\s*-\s*", "-", _squash_spaces(code))

    @staticmethod
    def _append_code(item: NormalizedItem, extra: str) -> None:
        """Дописывает перенесенный на новую строку шифр обоснования.

        :param item: позиция, чей шифр разорван переносом строки.
        :param extra: продолжение шифра.
        """
        item.code = PDFLSRParser._normalize_code(f"{item.code or ''} {extra}")

    @staticmethod
    def _append_title(item: NormalizedItem, extra: str) -> None:
        """Дописывает продолжение наименования к предыдущей позиции.

        :param item: позиция, чье наименование переносится на новую строку.
        :param extra: текст продолжения.
        """
        item.raw_title = _squash_spaces(f"{item.raw_title} {extra}")
        item.clean_title = LSRParser._clean_estimate_title(item.raw_title)
        item.dimensions = DimensionLockExtractor.extract(item.raw_title)

