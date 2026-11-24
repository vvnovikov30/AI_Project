"""Парсер локальных смет Excel: форма 4б, усеченная 4б и 421/пр."""

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import re
from typing import Any, Callable, Dict, Final, List, Optional

import openpyxl

from models import CellPointer, NormalizedItem
from normalizer import DimensionLockExtractor, UnitNormalizer

from .common import (
    SourceType,
    _clean_estimate_title,
    _close_buffer,
    _prepare_source,
)

_LSR_421_NOISE_PREFIXES: Final[tuple[str, ...]] = (
    "нр ",
    "нр от",
    "сп ",
    "сп от",
    "фот",
    "зп",
    "эм",
    "мр",
    "зтр",
    "зт",
    "в т.ч.",
    "в том числе",
    "итого",
    "всего",
    "накладные",
    "сметная прибыль",
    "ндс",
    "прямые затраты",
    "затраты труда",
    "оплата труда",
    "эксплуатация машин",
    "материальные ресурсы",
    "наименование работ и затрат",
)
_LSR_421_SECTION_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"^раздел\b", re.IGNORECASE
)
_LSR_421_POSITION_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"^\d+(?:[.,]\d+)?$"
)


# ---------------------------------------------------------------------------
# Динамическая детекция шапки формы 4б / усеченной 4б (без графы «Ед. изм.»).
# Корни ключевых слов сравниваются с ячейкой, приведенной к нижнему регистру
# без пробелов и знаков пунктуации: «Ед. изм.» -> "едизм", «№ п/п» -> "№пп".
# ---------------------------------------------------------------------------
_LSR_HEADER_SCAN_ROWS: Final[int] = 35
_LSR_CODE_STEMS: Final[tuple[str, ...]] = ("обоснован", "шифр", "позици")
_LSR_NAME_STEMS: Final[tuple[str, ...]] = ("наименован",)
_LSR_UNIT_STEMS: Final[tuple[str, ...]] = ("единиц", "измерен", "едизм")
_LSR_QTY_STEMS: Final[tuple[str, ...]] = ("колич", "колв", "объем")
_LSR_TOTAL_STEMS: Final[tuple[str, ...]] = ("всего", "стоимост", "сметнаястоим")
_LSR_POSITION_STEMS: Final[tuple[str, ...]] = ("№",)
# Служебные строки сметы: итоги разделов, справки, подписи, трудозатраты.
_LSR_SERVICE_PREFIXES: Final[tuple[str, ...]] = (
    "итого",
    "всего",
    "справочно",
    "составил",
    "проверил",
    "утверждаю",
    "согласовано",
    "подпись",
    "раздел",
    "приложение",
    "начальник",
    "главный инженер",
    "накладные",
    "сметная прибыль",
    "ндс",
    "нр от",
    "сп от",
    "фот",
    "зтр",
    "в т.ч.",
    "в том числе",
    "затраты труда",
    "оплата труда",
    "эксплуатация машин",
    "материальные ресурсы",
    "прямые затраты",
)
# Единица измерения в тексте наименования усеченной формы 4б:
# «Демонтаж светильников, 100 шт», «Кабель ВВГнг 3х1.5 (м)».
UNIT_FALLBACK_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"(?:,\s*|\s*\()((?:100\s+|1000\s+)?(?:шт|м2|м3|м|п\.м|пог\.м|компл|комплект|т|кг|упак|рулон))(?:\.|\)|$)",
    re.IGNORECASE,
)
# Расширенный вариант для наименований вида «(100 м коробок блоков)»: после
# единицы идет поясняющее слово. Применяется только если основной шаблон
# не дал совпадения, поэтому не влияет на штатные случаи.
_UNIT_FALLBACK_TRAILING_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"(?:,\s*|\s*\()((?:100\s+|1000\s+)?(?:шт|м2|м3|м|п\.м|пог\.м|компл|комплект|т|кг|упак|рулон))(?=\s+[А-Яа-яЁёA-Za-z])",
    re.IGNORECASE,
)
# Кратности 1/10 в скобках: «(1 шт.)», «(10 шт.)». Штатный шаблон ловит
# только 100/1000, поэтому эти случаи обрабатываются отдельным проходом.
_UNIT_FALLBACK_SMALL_COUNT_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"(?:,\s*|\s*\()((?:1|10)\s+(?:шт|компл|комплект|упак|рулон))(?:\.|\)|$)",
    re.IGNORECASE,
)



@dataclass(frozen=True)
class _LSR4bLayout:
    """Колонки сметы формы 4б, найденные по корням заголовков.

    :param header_row: номер строки шапки (1-based);
    :param code_col: графа обоснования/шифра (``None`` если не найдена);
    :param name_col: графа наименования работ и затрат;
    :param qty_col: графа количества;
    :param unit_col: графа единицы измерения; ``None`` — усеченная форма 4б;
    :param total_cost_col: графа общей стоимости позиции;
    :param unit_price_col: графа стоимости единицы;
    :param position_col: графа № п/п (``None`` если шапка без нумерации).
    """

    header_row: int
    name_col: int
    qty_col: int
    code_col: Optional[int] = None
    unit_col: Optional[int] = None
    total_cost_col: Optional[int] = None
    unit_price_col: Optional[int] = None
    position_col: Optional[int] = None

    @property
    def is_reduced(self) -> bool:
        """Признак усеченной формы 4б: отдельная графа единиц отсутствует."""
        return self.unit_col is None


@dataclass(frozen=True)
class _LSR421Layout:
    """Фактические графы сметы по Приказу 421/пр в листе Excel.

    Номера граф берутся из шапки, а не из физических колонок: в печатной
    форме 421/пр «Ед. изм.» и «Количество» часто сдвинуты вправо из-за
    объединенных ячеек наименования и блока стоимости.
    """

    header_row: int
    code_col: int
    title_col: int
    unit_col: int
    qty_col: int
    position_col: int = 1
    unit_price_col: Optional[int] = None
    total_col: Optional[int] = None



class LSRParser:
    """Парсер сметных расчетов Excel (форма по Приказу 421/пр и ТСН-2001 г. Москвы)."""

    # Ловит федеральные расценки (ГЭСН, ФЕР, ТЕР, 2-значные главы 421/пр)
    # и расценки на работы ТСН-2001 (главы 3, 4, 5, 6: монтаж, строительные, пусконаладка)
    BASE_RATE_PATTERN = re.compile(r"^(ГЭСН|ФЕР|ТЕР|[3-6]\.[\d\-]+|\d{2}\.\d+\.\d+)")

    # Ресурсные строки ТСН-2001: главы 1 (материалы), 2 (оборудование),
    # 13 (неучтенные материалы) и позиции прейскурантов «Прайс-лист».
    RESOURCE_CODE_PATTERN = re.compile(
        r"^(?:Прайс-лист|(?:1|2|13)\.[\d\-]+)", re.IGNORECASE
    )

    # Итоговая строка блока позиции: «ВСЕГО по позиции:» (графа 11).
    TOTAL_ROW_PATTERN = re.compile(r"^всего\s+по\s+позиции\s*:?", re.IGNORECASE)

    def __init__(self, source: SourceType, display_name: Optional[str] = None):
        """Открывает смету из пути или уже открытого бинарного потока.

        :param source: путь к ``.xlsx`` либо бинарный поток (``BytesIO``).
            Поток принадлежит вызывающему коду и не закрывается парсером.
        :param display_name: безопасное имя файла для ``CellPointer.file_name``;
            для потока по умолчанию используется ``in_memory.xlsx``.
        """
        path, self.file_name, self._buffer = _prepare_source(
            source, display_name, fallback_name="in_memory.xlsx"
        )
        self.display_name = self.file_name
        self.file_path = str(path) if path is not None else self.file_name
        self.wb = openpyxl.load_workbook(
            self._buffer if self._buffer is not None else self.file_path,
            data_only=True,
        )
        self._workbook_closed = False

    def __enter__(self) -> "LSRParser":
        return self

    def __exit__(self, *_args: Any) -> None:
        self.close()

    def close(self) -> None:
        """Закрывает книгу openpyxl и внутренний буфер (идемпотентно, совместимо с ``with``)."""
        if self._workbook_closed:
            return
        self._workbook_closed = True
        workbook = getattr(self, "wb", None)
        if workbook is not None:
            workbook.close()
        _close_buffer(getattr(self, "_buffer", None))

    def _resolve_merged_cells_map(self, ws) -> Dict[tuple, Any]:
        """Создает проекционную карту объединенных ячеек для O(1) поиска значений."""
        merge_map = {}
        for rng in ws.merged_cells.ranges:
            top_left_val = ws.cell(row=rng.min_row, column=rng.min_col).value
            for row in range(rng.min_row, rng.max_row + 1):
                for col in range(rng.min_col, rng.max_col + 1):
                    merge_map[(row, col)] = top_left_val
        return merge_map

    # Очистка наименования живет в ``parsers.common``; атрибут класса сохранен
    # ради обратной совместимости (его использует ``PDFLSRParser``).
    _clean_estimate_title = staticmethod(_clean_estimate_title)

    @staticmethod
    def _to_decimal(value: Any) -> Optional[Decimal]:
        """Приведение ячейки Excel к ``Decimal`` (запятая, пробелы, ``None``)."""
        if value is None or isinstance(value, bool):
            return None
        if isinstance(value, Decimal):
            return value
        if isinstance(value, (int, float)):
            return Decimal(str(value))
        text = (
            str(value)
            .strip()
            .replace("\xa0", "")
            .replace(" ", "")
            .replace(",", ".")
        )
        if not text:
            return None
        try:
            return Decimal(text)
        except (InvalidOperation, ValueError):
            return None

    @classmethod
    def _extract_total_cost(
        cls,
        get_val: Callable[[int], Any],
        preferred_col: Optional[int] = None,
    ) -> Optional[Decimal]:
        """Итог блока «ВСЕГО по позиции:» из графы 11 (fallback: графы 10..6).

        В форме 4б ТСН-2001 итоговая ячейка графы 11 объединена с графой 10
        (``J:K``), поэтому значение читается через карту merged-ячеек.
        Отрицательные значения (исключаемые объемы) сохраняются как есть.
        """
        if preferred_col is not None:
            value = cls._to_decimal(get_val(preferred_col))
            if value is not None and value != 0:
                return value
        for col_idx in (11, 10, 9, 8, 7, 6):
            value = cls._to_decimal(get_val(col_idx))
            if value is not None and value != 0:
                return value
        return None

    def _resolve_worksheet(self, sheet_name: Optional[str]):
        """Лист для разбора: явный ``sheet_name`` → ``ЛС ...`` → активный.

        В комплектах СПб единственный лист называется
        «Раздел ПД №11_ЛСР-...», поэтому fallback на активный (первый) лист
        применяется только при отсутствии листов с префиксом ``ЛС ``.
        """
        if sheet_name:
            return self.wb[sheet_name]
        for name in self.wb.sheetnames:
            if name.startswith("ЛС "):
                return self.wb[name]
        active = self.wb.active
        if active is not None:
            return active
        return self.wb[self.wb.sheetnames[0]]

    @staticmethod
    def _header_cell_text(value: Any) -> str:
        """Текст ячейки шапки: без переносов и в нижнем регистре."""
        return re.sub(r"\s+", " ", str(value or "")).strip().lower()

    @classmethod
    def _find_header_column(
        cls,
        values: List[str],
        predicate: Callable[[str], bool],
        start: int = 0,
    ) -> Optional[int]:
        """Индекс (0-based) первой ячейки шапки, удовлетворяющей предикату."""
        for index in range(start, len(values)):
            if predicate(values[index]):
                return index
        return None

    @staticmethod
    def _compact_header_value(value: Any) -> str:
        """Ячейка шапки без пробелов и пунктуации: «Ед. изм.» -> ``едизм``."""
        text = LSRParser._header_cell_text(value).replace("ё", "е")
        return re.sub(r"[^0-9a-zа-я№]+", "", text)

    @staticmethod
    def _header_starts_with(value: str, stems: tuple[str, ...]) -> bool:
        """Проверяет, начинается ли компактный заголовок с одного из корней."""
        return bool(value) and value.startswith(stems)

    @classmethod
    def _detect_4b_layout(
        cls, ws: Any, merge_map: Dict[tuple, Any]
    ) -> Optional[_LSR4bLayout]:
        """Ищет шапку формы 4б/усеченной 4б по корням ключевых слов.

        Критерий усеченной формы: найдены графа наименования и графа
        количества, но отдельная графа единицы измерения отсутствует
        (``unit_col is None``). Порядок граф в листе произвольный.
        """
        max_row = min(ws.max_row, _LSR_HEADER_SCAN_ROWS)
        max_col = min(ws.max_column, 40)
        for row_idx in range(1, max_row + 1):
            values = [
                cls._compact_header_value(
                    merge_map.get(
                        (row_idx, col), ws.cell(row=row_idx, column=col).value
                    )
                )
                for col in range(1, max_col + 1)
            ]
            name_col = cls._find_header_column(
                values,
                lambda value: any(stem in value for stem in _LSR_NAME_STEMS),
            )
            if name_col is None:
                continue
            qty_col = cls._find_header_column(
                values,
                lambda value: (
                    value.startswith(_LSR_QTY_STEMS)
                    or value in ("кол", "колво")
                ),
            )
            if qty_col is None or qty_col == name_col:
                continue
            position_col = cls._find_header_column(
                values, lambda value: value.startswith(_LSR_POSITION_STEMS)
            )
            if position_col in (name_col, qty_col):
                position_col = None
            code_col = cls._find_header_column(
                values,
                lambda value: cls._header_starts_with(value, _LSR_CODE_STEMS),
            )
            if code_col in (name_col, qty_col, position_col):
                code_col = None
            unit_col = cls._find_header_column(
                values,
                lambda value: cls._header_starts_with(value, _LSR_UNIT_STEMS),
            )
            if unit_col in (name_col, qty_col, code_col):
                unit_col = None
            total_cost_col = cls._find_header_column(
                values,
                lambda value: (
                    value.startswith(_LSR_TOTAL_STEMS)
                    and "единиц" not in value
                )
                or value.startswith("общаястоимост"),
            )
            if total_cost_col in (name_col, qty_col, code_col, unit_col):
                total_cost_col = None
            unit_price_col = cls._find_header_column(
                values,
                lambda value: "стоимост" in value and "единиц" in value,
            )
            if unit_price_col in (name_col, qty_col, code_col, unit_col):
                unit_price_col = None
            return _LSR4bLayout(
                header_row=row_idx,
                name_col=name_col + 1,
                qty_col=qty_col + 1,
                code_col=code_col + 1 if code_col is not None else None,
                unit_col=unit_col + 1 if unit_col is not None else None,
                total_cost_col=(
                    total_cost_col + 1 if total_cost_col is not None else None
                ),
                unit_price_col=(
                    unit_price_col + 1 if unit_price_col is not None else None
                ),
                position_col=(
                    position_col + 1 if position_col is not None else None
                ),
            )
        return None

    @classmethod
    def _extract_unit_from_title(cls, raw_title: str) -> Optional[str]:
        """Извлекает единицу измерения из текста наименования (fallback)."""
        text = re.sub(r"\s+", " ", str(raw_title or "")).strip()
        if not text:
            return None
        match = UNIT_FALLBACK_PATTERN.search(text)
        if match is None:
            match = _UNIT_FALLBACK_TRAILING_PATTERN.search(text)
        if match is None:
            match = _UNIT_FALLBACK_SMALL_COUNT_PATTERN.search(text)
        if match is None:
            return None
        unit = re.sub(r"\s+", " ", match.group(1)).strip()
        return unit or None

    @classmethod
    def _detect_421_layout(
        cls, ws: Any, merge_map: Dict[tuple, Any]
    ) -> Optional[_LSR421Layout]:
        """Распознает шапку сметы по Приказу 421/пр.

        Признак — одновременное присутствие граф «Обоснование»,
        «Наименование», «Единица измерения» и «Количество». Форма 4б
        ТСН-2001 содержит «Шифр расценки и коды ресурсов» и «Кол-во
        единиц», поэтому под эту проверку не попадает.
        """
        max_row = min(ws.max_row, 80)
        max_col = min(ws.max_column, 40)
        for row_idx in range(1, max_row + 1):
            values = [
                cls._header_cell_text(
                    merge_map.get(
                        (row_idx, col), ws.cell(row=row_idx, column=col).value
                    )
                )
                for col in range(1, max_col + 1)
            ]
            code_col = cls._find_header_column(
                values, lambda value: value.startswith("обоснование")
            )
            title_col = cls._find_header_column(
                values, lambda value: value.startswith("наименование")
            )
            unit_col = cls._find_header_column(
                values,
                lambda value: value.startswith(
                    ("единица измерения", "ед. изм", "ед.изм")
                ),
            )
            qty_col = cls._find_header_column(
                values,
                lambda value: value.startswith(("количество", "кол-во", "кол.")),
            )
            if None in (code_col, title_col, unit_col, qty_col):
                continue
            if not (code_col < title_col < unit_col < qty_col):
                continue

            unit_price_col: Optional[int] = None
            total_col: Optional[int] = None
            for sub_row in range(row_idx, min(row_idx + 4, ws.max_row) + 1):
                sub_values = [
                    cls._header_cell_text(
                        merge_map.get(
                            (sub_row, col),
                            ws.cell(row=sub_row, column=col).value,
                        )
                    )
                    for col in range(1, max_col + 1)
                ]
                if unit_price_col is None:
                    unit_price_col = cls._find_header_column(
                        sub_values,
                        lambda value: value == "на единицу",
                        start=qty_col + 1,
                    )
                total_candidates = [
                    index
                    for index in range(qty_col + 1, len(sub_values))
                    if sub_values[index] == "всего"
                ]
                if total_candidates:
                    total_col = total_candidates[-1]

            return _LSR421Layout(
                header_row=row_idx,
                code_col=code_col + 1,
                title_col=title_col + 1,
                unit_col=unit_col + 1,
                qty_col=qty_col + 1,
                unit_price_col=(
                    unit_price_col + 1 if unit_price_col is not None else None
                ),
                total_col=total_col + 1 if total_col is not None else None,
            )
        return None

    @classmethod
    def _resolve_421_section(cls, n_pp: str, title: str) -> str:
        """Заголовок раздела 421/пр: «Раздел 1. ...» в графах 1 или 3."""
        for value in (n_pp, title):
            text = re.sub(r"\s+", " ", value).strip()
            if text and _LSR_421_SECTION_PATTERN.match(text):
                return text
        return ""

    @classmethod
    def _extract_421_total(
        cls, get_val: Callable[[int], Any], layout: _LSR421Layout
    ) -> Optional[Decimal]:
        """Итог блока «Всего по позиции» из графы базисной стоимости."""
        if layout.total_col is None:
            return None
        value = cls._to_decimal(get_val(layout.total_col))
        if value is not None and value != 0:
            return value
        return None

    def _parse_421_items(
        self,
        ws: Any,
        merge_map: Dict[tuple, Any],
        layout: _LSR421Layout,
    ) -> List[NormalizedItem]:
        """Позиции сметы 421/пр: шифр, наименование, ед. изм., количество.

        Читаются строки позиций и ресурсов; ресурсные строки привязываются
        к последней расценке через ``parent_id``. Объемы — только ``Decimal``
        (включая отрицательные значения исключаемых объемов).
        """
        items: List[NormalizedItem] = []
        current_parent_id: Optional[str] = None
        current_section = "Default_Section"
        # Итог «Всего по позиции» относится к последней расценке, а не к
        # последней ресурсной строке внутри нее.
        pending_work_idx: Optional[int] = None

        for row_idx in range(layout.header_row + 1, ws.max_row + 1):
            def get_val(col_idx: int):
                return merge_map.get(
                    (row_idx, col_idx),
                    ws.cell(row=row_idx, column=col_idx).value,
                )

            n_pp = str(get_val(layout.position_col) or "").strip()
            raw_code = str(get_val(layout.code_col) or "").strip()
            raw_title = str(get_val(layout.title_col) or "").strip()
            raw_unit = str(get_val(layout.unit_col) or "").strip()
            raw_qty = get_val(layout.qty_col)
            title_lower = re.sub(r"\s+", " ", raw_title).strip().lower()

            section = self._resolve_421_section(n_pp, raw_title)
            if section:
                current_section = section
                pending_work_idx = None
                continue

            if self.TOTAL_ROW_PATTERN.match(title_lower):
                total_cost = self._extract_421_total(get_val, layout)
                if total_cost is not None and pending_work_idx is not None:
                    pending_item = items[pending_work_idx]
                    pending_item.total_cost = total_cost
                    if pending_item.unit_price is None and pending_item.volume != 0:
                        pending_item.unit_price = (
                            total_cost / pending_item.volume
                        ).quantize(Decimal("0.0001"))
                    pending_work_idx = None
                continue

            # Строка нумерации граф: 1 | 2 | 3 | 4 | 5 ...
            graph_row_values = [
                str(get_val(col) or "").strip()
                for col in (
                    layout.position_col,
                    layout.code_col,
                    layout.title_col,
                    layout.unit_col,
                    layout.qty_col,
                )
            ]
            if all(
                value.isdigit() and len(value) <= 2
                for value in graph_row_values
            ):
                continue
            if not raw_title or not raw_unit:
                continue

            title_lower = re.sub(r"\s+", " ", raw_title).strip().lower()
            if any(title_lower.startswith(p) for p in _LSR_421_NOISE_PREFIXES):
                continue

            qty_decimal = self._to_decimal(raw_qty)
            if qty_decimal is None:
                continue

            first_code_line = raw_code.split("\n")[0].strip()
            is_explicit_resource = bool(
                self.RESOURCE_CODE_PATTERN.match(first_code_line)
            )
            is_numbered_position = bool(
                _LSR_421_POSITION_PATTERN.fullmatch(n_pp)
            )
            is_work_rate = is_numbered_position or (
                bool(self.BASE_RATE_PATTERN.match(first_code_line))
                and not is_explicit_resource
            )
            item_id = f"LSR_{row_idx}"

            if is_work_rate:
                current_parent_id = item_id
                parent_ref = None
            else:
                parent_ref = current_parent_id

            canonical_u, norm_qty, mult = UnitNormalizer.normalize(
                raw_unit, qty_decimal
            )
            clean_title = self._clean_estimate_title(raw_title)
            dimensions = DimensionLockExtractor.extract(raw_title)
            header_price = (
                self._to_decimal(get_val(layout.unit_price_col))
                if layout.unit_price_col is not None
                else None
            )
            col_letter = openpyxl.utils.get_column_letter(layout.qty_col)
            pointer = CellPointer(
                doc_type="LSR",
                file_name=self.file_name,
                sheet_name=str(ws.title),
                row_idx=row_idx,
                col_letter=col_letter,
                coordinate=f"{col_letter}{row_idx}",
            )

            items.append(NormalizedItem(
                raw_id=item_id,
                code=first_code_line if first_code_line else None,
                raw_title=raw_title,
                clean_title=clean_title,
                raw_unit=raw_unit,
                canonical_unit=canonical_u,
                unit_multiplier=mult,
                volume=norm_qty,
                dimensions=dimensions,
                context_section=current_section,
                pointer=pointer,
                is_resource=not is_work_rate,
                parent_id=parent_ref,
                unit_price=(
                    header_price
                    if header_price is not None and header_price > 0
                    else None
                ),
            ))
            if is_work_rate:
                pending_work_idx = len(items) - 1

        return items

    def _parse_4b_items(
        self,
        ws: Any,
        merge_map: Dict[tuple, Any],
        layout: _LSR4bLayout,
    ) -> List[NormalizedItem]:
        """Разбор формы 4б и усеченной 4б (без отдельной графы единиц).

        В усеченной форме единица измерения извлекается из текста
        наименования (``UNIT_FALLBACK_PATTERN``). Если единица не найдена,
        позиция сохраняется с ``raw_unit=""`` и ``canonical_unit=UNKNOWN``:
        она не отбрасывается и далее проходит мягкий Unit Gate.
        """
        items: List[NormalizedItem] = []
        current_parent_id: Optional[str] = None
        current_section = "Default_Section"
        pending_item_idx: Optional[int] = None

        for row_idx in range(layout.header_row + 1, ws.max_row + 1):
            def get_val(col_idx: Optional[int]):
                if col_idx is None:
                    return None
                return merge_map.get(
                    (row_idx, col_idx),
                    ws.cell(row=row_idx, column=col_idx).value,
                )

            n_pp = str(get_val(layout.position_col) or "").strip()
            raw_code = str(get_val(layout.code_col) or "").strip()
            raw_title = str(get_val(layout.name_col) or "").strip()
            raw_qty = get_val(layout.qty_col)
            title_lower = re.sub(r"\s+", " ", raw_title).strip().lower()

            section = self._resolve_421_section(n_pp, raw_title)
            if section:
                current_section = section
                pending_item_idx = None
                continue

            # Итог блока позиции может стоять в графе № п/п либо наименования.
            total_row = next(
                (
                    text
                    for text in (title_lower, n_pp.lower())
                    if self.TOTAL_ROW_PATTERN.match(text)
                ),
                None,
            )
            if total_row is not None:
                total_cost = self._extract_total_cost(
                    get_val, layout.total_cost_col
                )
                if total_cost is not None and pending_item_idx is not None:
                    pending_item = items[pending_item_idx]
                    pending_item.total_cost = total_cost
                    if pending_item.unit_price is None and pending_item.volume != 0:
                        pending_item.unit_price = (
                            total_cost / pending_item.volume
                        ).quantize(Decimal("0.0001"))
                    pending_item_idx = None
                continue

            if any(title_lower.startswith(p) for p in _LSR_SERVICE_PREFIXES):
                continue
            if not raw_title or raw_qty is None:
                continue
            if raw_title.isdigit():
                continue

            # Строка нумерации граф: 1 | 2 | 3 | 4 ...
            graph_values = [
                str(get_val(col) or "").strip()
                for col in (
                    layout.position_col,
                    layout.code_col,
                    layout.name_col,
                    layout.unit_col,
                    layout.qty_col,
                )
                if col is not None
            ]
            if graph_values and all(
                value.isdigit() and len(value) <= 2 for value in graph_values
            ):
                continue

            qty_decimal = self._to_decimal(raw_qty)
            if qty_decimal is None:
                continue

            first_code_line = raw_code.split("\n")[0].strip() if raw_code else ""
            is_explicit_resource = bool(
                self.RESOURCE_CODE_PATTERN.match(first_code_line)
            )
            is_work_rate = (
                bool(self.BASE_RATE_PATTERN.match(first_code_line))
                and not is_explicit_resource
            )
            has_numbered_position = bool(
                _LSR_421_POSITION_PATTERN.fullmatch(n_pp)
            )
            # Ненумерованная строка с шифром ресурса (например, «101-1805»)
            # считается дочерним ресурсом, а не шумом.
            is_unclassified_resource = bool(first_code_line) and not (
                is_explicit_resource or is_work_rate
            )
            if not (
                has_numbered_position
                or is_explicit_resource
                or is_work_rate
                or is_unclassified_resource
            ):
                continue

            raw_unit = ""
            if layout.unit_col is not None:
                raw_unit = str(get_val(layout.unit_col) or "").strip()
            else:
                raw_unit = self._extract_unit_from_title(raw_title) or ""

            item_id = f"LSR_{row_idx}"

            if is_work_rate:
                current_parent_id = item_id
                parent_ref = None
            else:
                parent_ref = current_parent_id

            canonical_u, norm_qty, mult = UnitNormalizer.normalize(
                raw_unit, qty_decimal
            )
            clean_title = self._clean_estimate_title(raw_title)
            dimensions = DimensionLockExtractor.extract(raw_title)

            header_price = (
                self._to_decimal(get_val(layout.unit_price_col))
                if layout.unit_price_col is not None
                else None
            )
            if header_price is not None and header_price <= 0:
                header_price = None
            row_total = (
                self._to_decimal(get_val(layout.total_cost_col))
                if layout.total_cost_col is not None
                else None
            )
            unit_price = header_price
            if unit_price is None and row_total is not None and norm_qty != 0:
                unit_price = (row_total / norm_qty).quantize(
                    Decimal("0.0001")
                )

            col_letter = openpyxl.utils.get_column_letter(layout.qty_col)
            pointer = CellPointer(
                doc_type="LSR",
                file_name=self.file_name,
                sheet_name=str(ws.title),
                row_idx=row_idx,
                col_letter=col_letter,
                coordinate=f"{col_letter}{row_idx}",
            )

            items.append(NormalizedItem(
                raw_id=item_id,
                code=first_code_line if first_code_line else None,
                raw_title=raw_title,
                clean_title=clean_title,
                raw_unit=raw_unit,
                canonical_unit=canonical_u,
                unit_multiplier=mult,
                volume=norm_qty,
                dimensions=dimensions,
                context_section=current_section,
                pointer=pointer,
                is_resource=not is_work_rate,
                parent_id=parent_ref,
                unit_price=unit_price,
                total_cost=row_total,
            ))
            pending_item_idx = len(items) - 1

        return items

    def parse(self, sheet_name: Optional[str] = None) -> List[NormalizedItem]:
        """Парсит лист сметы и закрывает книгу openpyxl после разбора.

        Обертка гарантирует освобождение файла/архива книги даже при
        исключении; интерфейс ``LSRParser(path).parse()`` не изменился.
        """
        try:
            return self._parse_impl(sheet_name)
        finally:
            self.close()

    def _parse_impl(self, sheet_name: Optional[str] = None) -> List[NormalizedItem]:
        """Парсинг листа сметы в нормализованную плоскую коллекцию DTO.

        :param sheet_name: имя листа ``ЛС ...`` для пакетного/точечного
            прогона. Без параметра выбирается первый лист, начинающийся с
            ``ЛС ``; только если таких листов нет — активный/первый лист
            книги (комплект СПб именует лист «Раздел ПД №11_ЛСР-...»).
        :raises ValueError: запрошенный лист отсутствует в книге.

        Методика определяется контекстно по шапке листа: структура 421/пр
        (графы «Обоснование», «Наименование», «Ед. изм.», «Количество»)
        разбирается отдельной веткой, иначе работает алгоритм ТСН-2001
        формы 4б без изменений.
        """
        if sheet_name is not None and sheet_name not in self.wb.sheetnames:
            raise ValueError(
                f"Лист '{sheet_name}' не найден в книге {self.file_path}. "
                f"Доступны: {', '.join(self.wb.sheetnames)}"
            )
        ws = self._resolve_worksheet(sheet_name)
        merge_map = self._resolve_merged_cells_map(ws)

        layout_421 = self._detect_421_layout(ws, merge_map)
        if layout_421 is not None:
            return self._parse_421_items(ws, merge_map, layout_421)

        # Форма 4б / усеченная 4б: графы ищутся по корням заголовков.
        # Отсутствие графы «Ед. изм.» включает fallback-извлечение единицы
        # из текста наименования (Reduced Layout). Классическая раскладка
        # 4б (A/B/C/D/E) остается на legacy-парсере ТСН, чтобы не менять
        # семантику извлечения итогов «ВСЕГО по позиции».
        layout_4b = self._detect_4b_layout(ws, merge_map)
        if layout_4b is not None and not (
            layout_4b.position_col == 1
            and layout_4b.code_col == 2
            and layout_4b.name_col == 3
            and layout_4b.unit_col == 4
            and layout_4b.qty_col == 5
        ):
            return self._parse_4b_items(ws, merge_map, layout_4b)

        items: List[NormalizedItem] = []

        current_parent_id = None
        current_section = "Default_Section"
        # Индекс последней позиции, для которой еще не прочитан итог блока.
        pending_item_idx: Optional[int] = None

        for row_idx in range(1, ws.max_row + 1):
            def get_val(col_idx: int):
                return merge_map.get((row_idx, col_idx), ws.cell(row=row_idx, column=col_idx).value)

            col_n_pp = get_val(1)  # Колонка A (№ п/п)
            raw_code = str(get_val(2) or "").strip()
            raw_title = str(get_val(3) or "").strip()
            raw_unit = str(get_val(4) or "").strip()
            raw_qty = get_val(5)

            # Отслеживание заголовков разделов ("Раздел: Оборудование", "Раздел: Кабельные изделия")
            val_col_a = str(col_n_pp or "")
            if "раздел:" in val_col_a.lower():
                current_section = val_col_a.strip()

            # Итоговая строка блока: «ВСЕГО по позиции:» (графа 11).
            # Закрывает последнюю прочитанную позицию и проставляет стоимость.
            title_cell = str(get_val(3) or "").strip()
            if self.TOTAL_ROW_PATTERN.match(title_cell):
                total_cost = self._extract_total_cost(get_val)
                if total_cost is not None and pending_item_idx is not None:
                    pending_item = items[pending_item_idx]
                    pending_item.total_cost = total_cost
                    if pending_item.unit_price is None and pending_item.volume > 0:
                        pending_item.unit_price = (
                            total_cost / pending_item.volume
                        ).quantize(Decimal("0.0001"))
                    pending_item_idx = None
                continue

            if not raw_title or raw_qty is None:
                continue

            # 1. Отсекаем строки без валидного номера позиции (пропуск шапок, титульников, пустых ячеек)
            # Допустимы форматы: "1", "2", "26,1", "26.2"
            if not col_n_pp or not re.match(r"^\d+(?:[,\.]\d+)?$", str(col_n_pp).strip()):
                continue

            # 2. Отсекаем служебную строку-шапку с номерами граф (1 | 2 | 3 | 4 | 5)
            if raw_title.isdigit():
                continue

            # 3. Фильтрация сметных начислений и трудозатрат
            title_lower = raw_title.lower().strip()
            if any(title_lower.startswith(p) for p in ("нр от фот", "сп от фот", "зтр", "зп", "эм", "мр", "в т.ч.", "всего", "итого")):
                continue

            try:
                qty_decimal = Decimal(str(raw_qty).replace(",", ".").strip())
            except Exception:
                continue

            col_letter = openpyxl.utils.get_column_letter(2)
            pointer = CellPointer(
                doc_type="LSR",
                file_name=self.file_name,
                sheet_name=str(ws.title),
                row_idx=row_idx,
                col_letter=col_letter,
                coordinate=f"{col_letter}{row_idx}"
            )

            # Иерархия «Работа + Ресурс» ТСН-2001: строки глав 3, 4, 6 —
            # родительские расценки; главы 1, 2, 13 и «Прайс-лист» — ресурсы.
            first_code_line = raw_code.split("\n")[0].strip() if raw_code else ""
            is_explicit_resource = bool(self.RESOURCE_CODE_PATTERN.match(first_code_line))
            is_work_rate = (
                bool(self.BASE_RATE_PATTERN.match(first_code_line))
                and not is_explicit_resource
            )
            item_id = f"LSR_{row_idx}"

            if is_work_rate:
                current_parent_id = item_id
                parent_ref = None
            else:
                parent_ref = current_parent_id

            canonical_u, norm_qty, mult = UnitNormalizer.normalize(raw_unit, qty_decimal)
            clean_title = self._clean_estimate_title(raw_title)
            dimensions = DimensionLockExtractor.extract(raw_title)
            # Цена за единицу из графы 6 заголовка (для работ обычно пуста).
            header_price = self._to_decimal(get_val(6))

            items.append(NormalizedItem(
                raw_id=item_id,
                code=first_code_line if first_code_line else None,
                raw_title=raw_title,
                clean_title=clean_title,
                raw_unit=raw_unit,
                canonical_unit=canonical_u,
                unit_multiplier=mult,
                volume=norm_qty,
                dimensions=dimensions,
                context_section=current_section,
                pointer=pointer,
                is_resource=not is_work_rate,
                parent_id=parent_ref,
                unit_price=header_price if header_price is not None and header_price > 0 else None
            ))
            pending_item_idx = len(items) - 1

        return items
