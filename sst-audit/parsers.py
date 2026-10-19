from bisect import bisect_right
from dataclasses import dataclass, replace
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
from typing import Optional, Dict, Any, List, Callable, Sequence, Tuple, Final
import openpyxl
import openpyxl.utils
import pdfplumber
import pypdfium2 as pdfium

from models import NormalizedItem, CellPointer
from normalizer import UnitNormalizer, DimensionLockExtractor


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

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.wb = openpyxl.load_workbook(file_path, data_only=True)

    def _resolve_merged_cells_map(self, ws) -> Dict[tuple, Any]:
        """Создает проекционную карту объединенных ячеек для O(1) поиска значений."""
        merge_map = {}
        for rng in ws.merged_cells.ranges:
            top_left_val = ws.cell(row=rng.min_row, column=rng.min_col).value
            for row in range(rng.min_row, rng.max_row + 1):
                for col in range(rng.min_col, rng.max_col + 1):
                    merge_map[(row, col)] = top_left_val
        return merge_map

    @staticmethod
    def _clean_estimate_title(raw_title: str) -> str:
        """Очистка сметного наименования от формул обоснования цен, артефактов Excel и префиксов."""
        text = raw_title.replace("_x000D_", " ")
        # Отсекаем строки формул цен: "1 236,60 = [10 913 / 1,22 / 7,54] + 3% Трансп..."
        lines = [
            line for line in text.split("\n")
            if not ("=" in line and ("[" in line or "трансп" in line.lower() or "заг.скл" in line.lower()))
        ]
        text = " ".join(lines)
        # Убираем префикс "ОБОРУДОВАНИЕ:"
        text = re.sub(r"^\s*оборудование\s*:\s*", "", text, flags=re.IGNORECASE)
        return re.sub(r"\s+", " ", text).lower().strip()

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
                file_name=self.file_path,
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
                file_name=self.file_path,
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
                file_name=self.file_path,
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


_HEADER_TITLE_TOKENS: tuple[str, ...] = ("наименование",)
_HEADER_UNIT_TOKENS: tuple[str, ...] = (
    "ед.изм",
    "ед. изм",
    "единица измер",
)
_HEADER_QTY_TOKENS: tuple[str, ...] = (
    "количеств",
    "кол-во",
    "объем",
    "объём",
)
_HEADER_SPEC_TOKENS: tuple[str, ...] = (
    "позиция",
    "тип, марка",
    "тип,марка",
    "код оборудования",
    "завод",
)
_HEADER_MASS_TOKENS: tuple[str, ...] = ("масса", "вес")
_COLUMN_NUMBER_PATTERN = re.compile(r"^\d{1,2}$")
# Марка основного комплекта чертежей: «2025-00224-01-0-ИОС1.С».
_SPEC_MARK_PATTERN = re.compile(r"иос\s*\d+(?:[.\-]\d+)*\s*\.?\s*с\b", re.IGNORECASE)
_SPEC_TITLE_PATTERN = re.compile(
    r"спецификаци\w*\s+оборудован\w*", re.IGNORECASE
)
# Быстрые маркеры страницы спецификации. Проверяются по ``extract_text()``
# ДО вызова find_tables()/extract_table(): на чертежах, планах и однолинейных
# схемах десятки тысяч векторных линий, и поиск таблиц там неприемлемо долог.
_SPEC_PAGE_MARKERS: tuple[str, ...] = (
    "спецификация оборудования",
    "иос1.с",
    "наименование и техническая характеристика",
    "код оборудования",
)
# Сильные маркеры шапки таблицы спецификации. В отличие от марки «ИОС1.С»
# и надписи «Спецификация оборудования» в штампе, они встречаются только
# в самой таблице, поэтому позволяют отсечь листы чертежей.
_SPEC_TABLE_PAGE_MARKERS: tuple[str, ...] = (
    "наименование и техническая характеристика",
    "код оборудования",
)
# Маркеры ведомости объемов работ (ВОР). Первый лист ВОР содержит шапку,
# листы-продолжения идут без нее, поэтому факт ВОР запоминается для
# последующих страниц (см. PDFSpecParser._prescreen_pages).
_VOR_PAGE_MARKERS: tuple[str, ...] = (
    "ведомость монтажных работ",
    "ведомость объемов работ",
    "ведомость объёмов работ",
    "добавляемые объемы",
    "добавляемые объёмы",
    "исключаемые объемы",
    "исключаемые объёмы",
    "дефектный акт",
    "наименование вида работ",
    "обоснование объема",
    # Форма ВОР ПК «ГРАНД-Смета»: шапка печатается на каждом листе и служит
    # единственным признаком продолжения таблицы без рамки.
    "наименование работ, ресурсов, затрат",
    "формула расчета объемов",
)
# Строка таблицы в текстовом слое: единица измерения с объемом
# («шт. 1», «м2 18.0», «упак. 2»). Служит дешевым признаком данных:
# листы чертежей со штампом «Спецификация оборудования» его не имеют.
_TABLE_ROW_PATTERN = re.compile(
    r"(?<![0-9a-zа-яё])"
    r"(?:шт\.?|штук[аи]?|компл\.?|комплект[а-я]*|комп\.?|"
    r"кг|мг|г|т|км|мм|см|дм|м2|м3|м²|м³|кв\.?\s?м|куб\.?\s?м|м|"
    r"пог\.?\s?м\.?|п\.?\s?м\.?|пм\.?|л|мл|"
    r"упак[а-я]*|рул[а-я]*|ед\.?|усл\.?\s?ед\.?|кат\.?|"
    r"набор[а-я]*|пар[аы]?|секц[а-я]*|к-т\.?|тыс\.?)"
    r"\s+\*?\d",
    re.IGNORECASE,
)
_VOR_SOURCE_PATTERN = re.compile(r"том\s*\d", re.IGNORECASE)
# Расчет освещенности Dialux: таблицы «Перпендикулярная освещенность»
# (люксы) ошибочно попадают в том со штампом «Спецификация оборудования».
_DIALUX_PAGE_MARKERS: tuple[str, ...] = ("освещенность", "освещённость")
# Номер позиции: «1», «19», «3,1», «101А».
_POSITION_PATTERN = re.compile(r"^\d{1,3}(?:[.,]\d{1,2})?[A-Za-zА-Яа-я]?$")
# Служебный текст штампа и колонтитулов, который не является позицией.
_SERVICE_ROW_TOKENS: tuple[str, ...] = (
    "согласовано",
    "составил",
    "проверил",
    "утверждаю",
    "взам. инв",
    "инв. №",
    "подп. и дата",
    "формат",
    "масштаб",
    "шифр",
    "спецификация оборудования",
)
# Единица измерения в ячейке таблицы: «шт.», «компл.», «пм», «м2», «кг».
_UNIT_VALUE_PATTERN = re.compile(
    r"^(?:\d+\s*)?"
    r"(?:шт\.?|штук[аи]?|компл\.?|комплект[а-я]*|комп\.?|"
    r"кг|мг|г|т|км|мм|см|дм|м2|м3|м²|м³|кв\.?\s?м|куб\.?\s?м|м|"
    r"пог\.?\s?м\.?|п\.?\s?м\.?|м\.?\s?п\.?|пм\.?|"
    r"л|мл|"
    r"упак[а-я]*|рул[а-я]*|ед\.?|усл\.?\s?ед\.?|кат\.?|"
    r"набор[а-я]*|пар[аы]?|секц[а-я]*|лист[а-я]*|к-т\.?|тыс\.?)"
    r"(?:\s*/\s*[а-я0-9]+)?$"
)
# Шапка формы ВОР ПК «ГРАНД-Смета» (7 граф): количество находится в графе 4,
# а графа 7 — комментарий. Без этой проверки 7 колонок ошибочно трактуются
# как спецификация ГОСТ 21.110 с количеством в графе 8.
_GRAND_VOR_HEADER_MARKERS: tuple[str, ...] = (
    "наименование работ, ресурсов, затрат",
    "наименование работ, ресурсов и затрат",
    "формула расчета объемов",
    "ведомость объемов работ",
)
# Строка раздела ведомости: «Раздел: 1. Кровля», «Раздел 2. Фасад».
_VOR_SECTION_PATTERN = re.compile(r"^раздел\b", re.IGNORECASE)


@dataclass(frozen=True)
class _ColumnSchema:
    """Найденная схема таблицы PDF: ГОСТ-спецификация либо ведомость ВОР."""

    kind: str
    title_col: int
    unit_col: int
    qty_col: int
    position_col: Optional[int] = None
    mark_col: Optional[int] = None
    code_col: Optional[int] = None
    mass_col: Optional[int] = None
    column_count: Optional[int] = None


class PDFSpecParser:
    """Извлечение таблиц ГОСТ 21.110 и ведомостей объемов работ (ВОР).

    Схема колонок определяется по строке заголовка, поэтому один парсер
    поддерживает и 4-5-колоночные ВОР, и много-колоночные спецификации
    оборудования. Пустые и короткие строки не индексируются напрямую.
    """

    STAMP_W_PT = 185 * 2.83465
    STAMP_H_PT = 55 * 2.83465

    # Динамические отступы логической области таблицы (в пунктах PDF).
    # Не зависят от формата листа: А4 (595 pt) и А3 (1191 pt) обрабатываются
    # одинаково, поэтому правые графы «Ед. изм.» и «Количество» не срезаются.
    CROP_LEFT_PT = 15
    CROP_TOP_PT = 15
    CROP_RIGHT_PT = 15
    CROP_BOTTOM_STAMP_PT = 65
    # У форм ГРАНД-Сметы нижнего штампа нет: таблица доходит до края листа.
    GRAND_VOR_BOTTOM_PT = 5

    _LINE_SETTINGS: Dict[str, Any] = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 3,
    }
    _TEXT_SETTINGS: Dict[str, Any] = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
    }

    def __init__(self, file_path: str, debug: bool = False):
        self.file_path = file_path
        self.debug = debug
        # Тексты страниц из быстрого слоя PDFium (1-based номер -> текст).
        self._fast_page_texts: Optional[Dict[int, str]] = None

    def _get_crop_box(self, page) -> tuple[float, float, float, float]:
        """Логическая область таблицы для любого формата листа.

        Координаты относительные: правый край считается от фактической
        ширины страницы (А3 ~1190 pt, А4 ~595 pt), поэтому широкие
        ландшафтные спецификации не теряют правые графы.
        """
        return (
            self.CROP_LEFT_PT,                          # отступ слева
            self.CROP_TOP_PT,                           # отступ сверху
            page.width - self.CROP_RIGHT_PT,            # динамическая ширина
            page.height - self.CROP_BOTTOM_STAMP_PT,    # срез нижнего штампа
        )

    @staticmethod
    def _page_text(page) -> str:
        """Текст страницы в нижнем регистре; отсутствие метода не критично."""
        extractor = getattr(page, "extract_text", None)
        if not callable(extractor):
            return ""
        try:
            return (extractor() or "").lower()
        except Exception:
            return ""

    @classmethod
    def _marker_text(cls, text: str) -> str:
        """Нормализует текст для маркерного поиска.

        Mojibake восстанавливается до ``lower()``: понижение регистра ломает
        обратное преобразование cp866 -> cp1251 (Ё/ё). Пробелы и переносы
        схлопываются: в PDF заголовок часто разорван строкой, а nbsp
        приходит вместо обычного пробела.
        """
        if not text:
            return ""
        return re.sub(r"\s+", " ", cls._fix_mojibake(text)).lower()

    @classmethod
    def _text_has_spec_markers(cls, text: str) -> bool:
        """Текст (после ``_marker_text``) содержит маркеры спецификации."""
        if not text:
            return False
        if any(marker in text for marker in _SPEC_PAGE_MARKERS):
            return True
        return bool(
            _SPEC_MARK_PATTERN.search(text) or _SPEC_TITLE_PATTERN.search(text)
        )

    @classmethod
    def _text_has_vor_markers(cls, text: str) -> bool:
        """Текст (после ``_marker_text``) содержит маркеры ВОР."""
        return bool(text) and any(marker in text for marker in _VOR_PAGE_MARKERS)

    @classmethod
    def _text_has_table_rows(cls, text: str) -> bool:
        """В тексте (после ``_marker_text``) есть строки «ед. изм. + объем»."""
        return bool(text) and bool(_TABLE_ROW_PATTERN.search(text))

    @classmethod
    def _text_is_vor_continuation(cls, text: str) -> bool:
        """Страница похожа на продолжение таблицы ВОР без повторной шапки."""
        if not text:
            return False
        return (
            cls._text_has_table_rows(text)
            or bool(_VOR_SOURCE_PATTERN.search(text))
        )

    def _scan_page_texts(self) -> Optional[Dict[int, str]]:
        """Быстрое чтение текстового слоя всех страниц через pypdfium2.

        PDFium существенно быстрее pdfminer на больших томах: полное
        сканирование 50-страничного PDF занимает доли секунды и не строит
        таблицы. ``None`` означает, что PDFium недоступен либо файл не
        читается — тогда вызывающий код использует прежний pdfplumber-фильтр.
        """
        try:
            document = pdfium.PdfDocument(self.file_path)
        except Exception:
            return None

        texts: Dict[int, str] = {}
        try:
            for index in range(len(document)):
                page = document[index]
                try:
                    textpage = page.get_textpage()
                except Exception:
                    texts[index + 1] = ""
                    continue
                try:
                    texts[index + 1] = textpage.get_text_range() or ""
                except Exception:
                    texts[index + 1] = ""
                finally:
                    close_textpage = getattr(textpage, "close", None)
                    if callable(close_textpage):
                        close_textpage()
        finally:
            close_document = getattr(document, "close", None)
            if callable(close_document):
                close_document()
        return texts

    def _pages_with_raster_images(self, pages: set[int]) -> set[int]:
        """Страницы из ``pages`` (1-based) с растровым изображением (альбом).

        Нужны для листов спецификации, где таблица вставлена картинкой:
        текстовый слой содержит только штамп, и по нему лист не отличить от
        расчетов Dialux (у тех текста тоже нет, но и изображения нет).
        """
        result: set[int] = set()
        if not pages:
            return result
        try:
            document = pdfium.PdfDocument(self.file_path)
        except Exception:
            return result
        try:
            total = len(document)
            for page_number in pages:
                if not 1 <= page_number <= total:
                    continue
                try:
                    page = document[page_number - 1]
                    # Листы спецификации — альбомные (A3). Портретные
                    # страницы с растром (титулы, сканы) не разбираем.
                    if page.get_width() <= page.get_height():
                        continue
                    if any(
                        page.get_objects(filter=[pdfium.raw.FPDF_PAGEOBJ_IMAGE])
                    ):
                        result.add(page_number)
                except Exception:
                    continue
        finally:
            close_document = getattr(document, "close", None)
            if callable(close_document):
                close_document()
        return result

    def _prescreen_pages(self) -> Optional[set[int]]:
        """Быстрый предфильтр страниц до детального разбора pdfplumber.

        Возвращает номера страниц (1-based), которые нужно передать в
        pdfplumber, либо ``None``, если быстрый слой недоступен.

        Логика:

        * страница с маркерами спецификации — разбирается (в том числе лист
          с растровой таблицей); исключаются только расчеты освещенности
          Dialux, где в тексте есть «освещенность»;
        * страница с маркерами ВОР открывает режим продолжения: следующие
          страницы таблицы ВОР разбираются даже без повторной шапки, пока
          на них есть строки «ед. изм. + объем» либо ссылки на обоснование;
        * остальные страницы (чертежи, планы, схемы) отсекаются и не
          доходят до дорогих ``find_tables()``/``extract_table()``.
        """
        texts = self._scan_page_texts()
        self._fast_page_texts = texts
        if texts is None:
            return None
        # Тексты PDFium переиспользуются при разборе страниц (проверка
        # кодировки и сильных маркеров) — без повторного extract_text().

        selected: set[int] = set()
        spec_without_table: set[int] = set()
        vor_active = False
        for page_index in sorted(texts):
            text = self._marker_text(texts[page_index])
            if not text:
                # Пустой текстовый слой: поведение как раньше — страница
                # пропускается, таблица без текста все равно не индексируется.
                vor_active = False
                continue
            if self._text_has_spec_markers(text):
                vor_active = False
                # Расчет освещенности Dialux печатается на листах со штампом
                # «Спецификация оборудования», но его таблица — люксы, а не
                # оборудование. Отсекаем по характерному заголовку.
                if any(marker in text for marker in _DIALUX_PAGE_MARKERS):
                    continue
                # Марка «ИОС1.С» и надпись «Спецификация оборудования» есть
                # и в штампе чертежей. Лист разбирается, если в тексте видны
                # строки данных либо шапка самой таблицы.
                if (
                    self._text_has_table_rows(text)
                    or any(marker in text for marker in _SPEC_TABLE_PAGE_MARKERS)
                ):
                    selected.add(page_index)
                else:
                    # Текстового слоя таблицы нет: возможно, она вставлена
                    # растром. Проверяем изображения отдельно, чтобы не
                    # тянуть в pdfplumber пустые листы Dialux.
                    spec_without_table.add(page_index)
                continue
            if self._text_has_vor_markers(text):
                vor_active = True
                selected.add(page_index)
                continue
            if vor_active and self._text_is_vor_continuation(text):
                selected.add(page_index)
                continue
            vor_active = False
        selected |= self._pages_with_raster_images(spec_without_table)
        return selected

    @classmethod
    def _page_has_spec_markers(cls, page) -> Optional[bool]:
        """Есть ли на странице маркеры спецификации (дешевый предфильтр).

        Возвращает ``None``, если текстовый слой недоступен (например, у
        тестового двойника страницы нет ``extract_text``): тогда фильтрация
        не применяется и поведение остается прежним.
        """
        extractor = getattr(page, "extract_text", None)
        if not callable(extractor):
            return None
        try:
            raw_text = extractor() or ""
        except Exception:
            return None
        if not raw_text:
            return False
        return cls._text_has_spec_markers(cls._marker_text(raw_text))

    @classmethod
    def _has_spec_marker(cls, page) -> bool:
        """Марка спецификации «*.С» либо заголовок «Спецификация оборудования»."""
        text = cls._marker_text(cls._page_text(page))
        if not text:
            return False
        return bool(
            _SPEC_MARK_PATTERN.search(text) or _SPEC_TITLE_PATTERN.search(text)
        )

    @staticmethod
    def _row_within_crop(table_row: Any, crop_box: tuple[float, float, float, float]) -> bool:
        """Строка таблицы попадает в логическую область (штамп отсекается)."""
        bbox = getattr(table_row, "bbox", None)
        if bbox is None:
            return True
        top = float(bbox[1])
        return crop_box[1] - 1 <= top <= crop_box[3]

    def _extract_page_table(self, page) -> Optional[List[List[str]]]:
        """Извлекает строки таблицы страницы с учетом динамического CropBox.

        Таблица ищется на полной странице: линии-правила в PDF часто нарисованы
        от края до края листа, и физическое обрезание страницы рвет их, ломая
        распознавание колонок. Поэтому CropBox применяется как логическая
        область — отбрасываются строки, попадающие в нижний штамп.
        """
        crop_box = self._get_crop_box(page)
        rows, tables = self._extract_rows_via_finder(page, crop_box)
        if rows:
            return self._maybe_rebuild_grand_vor(page, rows, crop_box, tables)

        for settings in (self._LINE_SETTINGS, self._TEXT_SETTINGS):
            table = self._extract_raw_table(page, settings)
            if not table:
                continue
            # Текстовая стратегия на чертежах «видит» таблицу в любом наборе
            # подписей, поэтому без распознанной шапки её результат не берем.
            if settings == self._TEXT_SETTINGS and not self._table_has_header(table):
                continue
            return self._maybe_rebuild_grand_vor(page, table, crop_box, [])
        return None

    def _maybe_rebuild_grand_vor(
        self,
        page,
        rows: List[List[str]],
        crop_box: tuple[float, float, float, float],
        tables: List[Any],
    ) -> List[List[str]]:
        """Дополняет шапку ВОР ГРАНД-Сметы строками данных без линий.

        Экспорт ПК «ГРАНД-Смета» рисует рамку только у шапки таблицы: строки
        работ идут без горизонтальных линий, поэтому ``find_tables()`` видит
        две строки шапки и ни одной позиции. Строки данных собираются по
        координатам слов (см. ``_extract_grand_vor_rows``), а шапка и границы
        граф берутся из единственной размеченной таблицы листа.
        """
        # Рамка есть и у данных: штатный разбор уже дал позиции.
        if len(rows) > 4 and self._table_has_item_rows(rows):
            return rows
        layout = self._grand_vor_layout(tables)
        if layout is None:
            return rows
        boundaries, header_bottom, header_rows = layout
        # Нижний штамп в форме ГРАНД-Сметы отсутствует, а обрезка по
        # CROP_BOTTOM_STAMP_PT теряла последнюю позицию листа.
        data_box = (
            crop_box[0],
            crop_box[1],
            crop_box[2],
            float(getattr(page, "height", crop_box[3])) - self.GRAND_VOR_BOTTOM_PT,
        )
        rebuilt = self._extract_grand_vor_rows(
            page, data_box, boundaries, header_bottom
        )
        if not rebuilt:
            return rows
        return header_rows + rebuilt

    @classmethod
    def _is_grand_vor_header(cls, row: List[str]) -> bool:
        """Строка является шапкой формы ВОР ПК «ГРАНД-Смета»."""
        joined = " ".join(cls._cell_text(cell).lower() for cell in row)
        joined = re.sub(r"\s+", " ", joined)
        return any(marker in joined for marker in _GRAND_VOR_HEADER_MARKERS)

    def _extract_grand_vor_rows(
        self,
        page,
        crop_box: tuple[float, float, float, float],
        boundaries: List[float],
        header_bottom: float,
    ) -> List[List[str]]:
        """Собирает строки таблицы ВОР из слов страницы по границам граф.

        Границы граф берутся из рамки шапки (единственной размеченной части
        формы), после чего каждое слово относится к графе по центру своей
        горизонтальной позиции. Слова одной визуальной строки склеиваются,
        а продолжение многострочной ячейки присоединяется к предыдущей позиции.
        """
        word_extractor = getattr(page, "extract_words", None)
        if not callable(word_extractor):
            return []
        try:
            words = word_extractor() or []
        except Exception:
            return []
        if not words:
            return []

        lines: Dict[int, List[Any]] = {}
        for word in words:
            top = float(word.get("top", 0.0))
            left = float(word.get("x0", 0.0))
            right = float(word.get("x1", 0.0))
            if top < header_bottom - 2:
                continue
            if top > crop_box[3] or right < crop_box[0] or left > crop_box[2]:
                continue
            # Одна визуальная строка: допуск 3 pt на разную высоту шрифта.
            lines.setdefault(round(top / 3), []).append(word)

        cell_count = len(boundaries) - 1
        rows: List[List[str]] = []
        current: Optional[List[str]] = None

        for key in sorted(lines):
            cells = [""] * cell_count
            for word in sorted(lines[key], key=lambda item: item["x0"]):
                center = (float(word["x0"]) + float(word["x1"])) / 2
                index = bisect_right(boundaries, center) - 1
                index = min(max(index, 0), cell_count - 1)
                text = str(word.get("text", "")).strip()
                if not text:
                    continue
                cells[index] = (cells[index] + " " + text).strip()

            if not any(cells):
                continue
            line_text = re.sub(
                r"\s+", " ", " ".join(cell for cell in cells if cell)
            )
            if _VOR_SECTION_PATTERN.match(line_text):
                # «Раздел: 1. Кровля» — отдельная строка-контекст, а не
                # продолжение предыдущей позиции.
                if current is not None:
                    rows.append(current)
                    current = None
                rows.append(cells)
                continue

            if _POSITION_PATTERN.fullmatch(cells[0].strip()):
                if current is not None:
                    rows.append(current)
                current = cells
                continue
            if current is None:
                current = cells
                continue
            # Продолжение ячейки: объём «36.58\n31,18+5,4», перенос
            # наименования либо ссылка на обоснование.
            for index, value in enumerate(cells):
                if not value:
                    continue
                current[index] = (
                    f"{current[index]}\n{value}" if current[index] else value
                )
        if current is not None:
            rows.append(current)
        return rows

    def _grand_vor_layout(
        self, tables: List[Any]
    ) -> Optional[tuple[List[float], float, List[List[str]]]]:
        """Границы граф, низ рамки шапки и строки шапки ВОР ГРАНД-Сметы.

        Разметка берется только у таблицы, в шапке которой есть маркеры формы
        ГРАНД-Сметы: иначе координаты граф произвольной таблицы дали бы мусор.
        """
        for table in tables or []:
            try:
                data = table.extract() or []
            except Exception:
                continue
            header_rows = [
                [self._cell_text(cell) for cell in row] for row in data
            ]
            if not any(self._is_grand_vor_header(row) for row in header_rows[:8]):
                continue
            try:
                columns = table.columns
                bottom = float(table.bbox[3])
            except Exception:
                continue
            if not columns or len(columns) < 4:
                continue
            boundaries = [float(column.bbox[0]) for column in columns]
            boundaries.append(float(columns[-1].bbox[2]))
            return boundaries, bottom, header_rows
        return None

    def _table_has_header(self, table: List[List[str]]) -> bool:
        """В таблице есть шапка ГОСТ 21.110 либо ВОР."""
        if self._has_gost_graph_numbers(table):
            return True
        return any(self._detect_schema(row) is not None for row in table[:8])

    def _extract_rows_via_finder(
        self, page, crop_box: tuple[float, float, float, float]
    ) -> Tuple[List[List[str]], List[Any]]:
        """Строки самой крупной таблицы страницы без строк нижнего штампа.

        Возвращает также все найденные таблицы: рамка шапки ВОР ГРАНД-Сметы
        используется как источник границ граф для строк без линий.
        """
        finder = getattr(page, "find_tables", None)
        if not callable(finder):
            return [], []
        try:
            tables = finder(self._LINE_SETTINGS)
        except Exception:
            return [], []

        best: List[List[str]] = []
        for table in tables or []:
            try:
                data = table.extract()
                table_rows = table.rows
            except Exception:
                continue
            if not data:
                continue
            filtered = [
                [self._cell_text(cell) for cell in row]
                for row, table_row in zip(data, table_rows)
                if self._row_within_crop(table_row, crop_box)
            ]
            if self._table_weight(filtered) > self._table_weight(best):
                best = filtered
        return best, list(tables or [])

    @staticmethod
    def _table_weight(table: List[List[str]]) -> int:
        """Вес таблицы: площадь сетки, чтобы штамп (много узких строк) не побеждал."""
        if not table:
            return 0
        return len(table) * max(len(row) for row in table)

    def _extract_raw_table(
        self, page, settings: Dict[str, Any]
    ) -> List[List[str]]:
        """Извлекает таблицу через ``extract_table`` (в т.ч. для простых страниц)."""
        try:
            table = page.extract_table(settings)
        except Exception:
            return []
        if not table:
            return []
        return [
            [self._cell_text(cell) for cell in row]
            for row in table
            if row
        ]

    @staticmethod
    def _cell_text(value: Any) -> str:
        """Очищает ячейку PDF от переносов, мягких дефисов и лишних пробелов."""
        if value is None:
            return ""
        text = str(value).replace("\xa0", " ").replace("\u200b", "")
        # Перенос слова по дефису в PDF не должен превращаться в пробел.
        text = re.sub(r"-\s*\n\s*", "", text)
        text = text.replace("\n", " ")
        return re.sub(r"\s+", " ", text).strip()

    # Символы, допустимые в осмысленном тексте спецификации.
    _PLAUSIBLE_CHARS = frozenset(
        "абвгдеёжзийклмнопрстуфхцчшщъыьэюя"
        "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ"
        "abcdefghijklmnopqrstuvwxyz"
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        "0123456789 .,;:!?()[]{}\"'-–—/\\%*+№<>=_|@#&$\n\t"
    )

    @classmethod
    def _mojibake_score(cls, text: str) -> int:
        """Оценка «осмысленности» строки: псевдографика cp866 штрафуется.

        pdfminer для части шрифтов спецификаций декодирует коды cp1251 как
        cp866: «ВРУ» превращается в «┬╨╙», «состоящее» — в «ёюёЄю∙хх».
        """
        allowed = sum(1 for ch in text if ch in cls._PLAUSIBLE_CHARS)
        exotic = sum(
            1 for ch in text if ch not in cls._PLAUSIBLE_CHARS and not ch.isspace()
        )
        return allowed - 2 * exotic

    @classmethod
    def _fix_mojibake(cls, text: str) -> str:
        """Восстанавливает строку cp1251, ошибочно декодированную как cp866."""
        if not text or not any(ord(ch) > 0x7F for ch in text):
            return text
        try:
            restored = text.encode("cp866").decode("cp1251")
        except (UnicodeEncodeError, UnicodeDecodeError):
            return text
        if cls._mojibake_score(restored) > cls._mojibake_score(text):
            return restored
        return text

    @classmethod
    def _force_fix_mojibake(cls, text: str) -> str:
        """Принудительное восстановление для страницы с подтвержденной кодировкой."""
        if not text or not any(ord(ch) > 0x7F for ch in text):
            return text
        try:
            return text.encode("cp866").decode("cp1251")
        except (UnicodeEncodeError, UnicodeDecodeError):
            return text

    @classmethod
    def _text_is_mojibake(cls, text: str) -> bool:
        """Определяет битую кодировку переданного текста страницы."""
        sample = (text or "")[:6000]
        if not sample:
            return False
        return cls._fix_mojibake(sample) != sample

    @classmethod
    def _page_is_mojibake(cls, page) -> bool:
        """Определяет битую кодировку текстового слоя страницы."""
        return cls._text_is_mojibake(cls._page_text(page))

    def _repair_table_encoding(
        self,
        page,
        table: List[List[str]],
        page_text: Optional[str] = None,
    ) -> List[List[str]]:
        """Восстанавливает кириллицу таблицы, если кодировка страницы битая.

        ``page_text`` — сырой текст из PDFium-предфильтра. Он позволяет не
        вызывать дорогой ``page.extract_text()`` повторно.
        """
        is_mojibake = (
            self._text_is_mojibake(page_text)
            if page_text is not None
            else self._page_is_mojibake(page)
        )
        if not is_mojibake:
            return table
        return [
            [self._force_fix_mojibake(cell) for cell in row]
            for row in table
        ]

    @classmethod
    def _detect_schema(cls, row: List[str]) -> Optional[_ColumnSchema]:
        """Определяет схему таблицы по строке заголовков.

        Поддерживаются оба целевых формата:

        * ГОСТ 21.110 (спецификация оборудования, >= 7 колонок);
        * ВОР: 4-6 колонок (№, наименование, ед. изм., количество) и форма
          ПК «ГРАНД-Смета» из 7 граф (объем в графе 4, комментарий в графе 7).

        Индексы колонок берутся из шапки, а не из фиксированных позиций,
        поэтому дополнительные служебные колонки не ломают разбор.
        """
        cells = [cls._cell_text(cell).lower() for cell in row]
        # Переносы в шапке PDF дают «Количест во», «Ед. изм», «Коли-чество»;
        # сравнение выполняется по склеенному тексту без пробелов и дефисов.
        keys = [re.sub(r"[\s\-]+", "", cell) for cell in cells]
        joined = " ".join(cells)
        # Форма ВОР СПб (5 граф): «№ п/п | Наименование | Ед. изм. | Кол. |
        # Примечание». В такой форме заголовок количества сокращен до «Кол.»,
        # поэтому общий поиск по «количество»/«объем» не срабатывает.
        is_vor_5 = (
            len(cells) == 5
            and any("наименование" in cell for cell in cells)
            and any(cell.startswith("кол") for cell in cells)
        )
        if is_vor_5:
            return _ColumnSchema(
                kind="VOR",
                title_col=1,
                unit_col=2,
                qty_col=3,
                position_col=0,
                column_count=len(cells),
            )
        # Форма ВОР ПК «ГРАНД-Смета»: 7 граф, количество в графе 4.
        is_grand_vor = cls._is_grand_vor_header(row)

        title_col = next(
            (
                index
                for index, cell in enumerate(keys)
                if any(token in cell for token in _HEADER_TITLE_TOKENS)
            ),
            None,
        )
        unit_col = next(
            (
                index
                for index, cell in enumerate(keys)
                if any(token.replace(" ", "") in cell for token in _HEADER_UNIT_TOKENS)
            ),
            None,
        )
        qty_col = next(
            (
                index
                for index, cell in enumerate(keys)
                if any(token in cell for token in _HEADER_QTY_TOKENS)
            ),
            None,
        )
        mass_col = next(
            (
                index
                for index, cell in enumerate(keys)
                if any(token in cell for token in _HEADER_MASS_TOKENS)
            ),
            None,
        )

        is_specification = not is_grand_vor and (
            len(cells) >= 7
            or any(token in joined for token in _HEADER_SPEC_TOKENS)
        )
        if title_col is None:
            return None
        if unit_col is None:
            if is_grand_vor:
                # ГРАНД-Смета: графа 3 (индекс 2) — единица измерения.
                if len(cells) <= 2:
                    return None
                unit_col = 2
            elif not is_specification or len(cells) <= 5:
                # ГОСТ 21.110: графа 6 (индекс 5) — единица измерения.
                return None
            else:
                unit_col = 5
        if qty_col is None:
            if is_grand_vor:
                # ГРАНД-Смета: графа 4 (индекс 3) — объем работ, а не графа 7
                # (там комментарий), и не графа 8 спецификации ГОСТ 21.110.
                if len(cells) <= 3:
                    return None
                qty_col = 3
            elif not is_specification or len(cells) <= 6:
                return None
            else:
                # ГОСТ 21.110: графа 8 (индекс 7); при отсутствии графы массы —
                # индекс 6. Значение всё равно уточняется по содержимому строки.
                qty_col = 7 if len(cells) > 7 else 6

        position_col = next(
            (
                index
                for index, cell in enumerate(keys)
                if cell in {"позиция", "поз", "№", "nп/п", "№п/п"}
            ),
            None,
        )
        mark_col = next(
            (
                index
                for index, cell in enumerate(keys)
                if "тип,марка" in cell
                or "типмарка" in cell
                or cell == "марка"
            ),
            None,
        )
        code_col = next(
            (
                index
                for index, cell in enumerate(keys)
                if cell.startswith("код")
            ),
            None,
        )
        if is_grand_vor:
            # ГРАНД-Смета: графа 1 (индекс 0) — номер позиции.
            position_col = 0
        return _ColumnSchema(
            kind="SPECIFICATION" if is_specification else "VOR",
            title_col=title_col,
            unit_col=unit_col,
            qty_col=qty_col,
            position_col=position_col,
            mark_col=mark_col,
            code_col=code_col,
            mass_col=mass_col,
            column_count=len(cells),
        )

    @classmethod
    def _gost_positional_schema(cls) -> _ColumnSchema:
        """Позиционный маппинг 9 граф ГОСТ 21.110-2013 для «слепой» шапки.

        Используется, когда шрифт шапки не даёт читаемого текста, но структура
        таблицы однозначно определяется строкой нумерации граф 1..9.
        """
        return _ColumnSchema(
            kind="SPECIFICATION",
            title_col=1,        # графа 2 — наименование и техническая характеристика
            unit_col=5,         # графа 6 — единица измерения
            qty_col=7,          # графа 8 — количество
            position_col=0,     # графа 1 — позиция
            mark_col=2,         # графа 3 — тип, марка
            code_col=3,         # графа 4 — код оборудования
            mass_col=6,         # графа 7 — масса единицы, кг
            column_count=9,     # 9 граф по ГОСТ 21.110-2013
        )

    def _detect_table_schema(self, table: List[List[str]]) -> Optional[_ColumnSchema]:
        """Ищет шапку таблицы в первых строках страницы.

        Приоритет отдаётся спецификации ГОСТ 21.110. Если текст шапки
        нечитаем (шрифт без ToUnicode), схема восстанавливается по строке
        нумерации граф — это штатная ситуация для томов в А3-ландшафте.
        """
        fallback: Optional[_ColumnSchema] = None
        for row in table[:8]:
            schema = self._detect_schema(row)
            if schema is None:
                continue
            if schema.kind == "SPECIFICATION":
                return schema
            if fallback is None:
                fallback = schema
        if fallback is not None:
            return fallback

        if self._has_gost_graph_numbers(table):
            return self._gost_positional_schema()
        return None

    @classmethod
    def _has_gost_graph_numbers(cls, table: List[List[str]]) -> bool:
        """Строка нумерации граф 1..9 — структурный признак ГОСТ 21.110."""
        for row in table[:4]:
            values = [cell.strip() for cell in row if cell.strip()]
            if len(values) < 7:
                continue
            if all(_COLUMN_NUMBER_PATTERN.fullmatch(value) for value in values):
                return True
        return False

    @classmethod
    def _parse_vor_volume(cls, raw_value: Any) -> Optional[Decimal]:
        """Объем работ из графы 4 ВОР: первая строка многострочной ячейки.

        ГРАНД-Смета печатает в графе «Количество» итог и формулу его расчета
        («36.58\\n31,18+5,4», «15\\n15»). Объемом является первая строка:
        формула не должна попадать в ``Decimal``. Ячейка, склеенная
        ``_cell_text`` в одну строку, разбирается по первому токену.
        """
        text = "" if raw_value is None else str(raw_value)
        first = next((line for line in text.split("\n") if line.strip()), "")
        if not first:
            return None
        compact = (
            first.replace(" ", "")
            .replace("\u00a0", "")
            .replace("\u200b", "")
            .replace(",", ".")
            .lstrip("*")
            .rstrip("*")
        )
        if compact:
            try:
                return Decimal(compact)
            except InvalidOperation:
                pass
        # «1 020» с пробелом-разделителем уже разобрано выше; здесь — случай
        # склеенной ячейки «36.58 31,18+5,4», где второй токен уже формула.
        token = re.split(r"\s+", first.strip())[0]
        token = token.replace("\u00a0", "").replace(",", ".").strip("*")
        try:
            return Decimal(token)
        except InvalidOperation:
            return None

    @classmethod
    def _clean_vor_unit(cls, value: str) -> str:
        """Локальные исправления OCR-опечаток единиц измерения ВОР.

        ``UnitNormalizer`` общий для всех форматов, поэтому искажения именно
        ВОР («HT» вместо «шт», потерянная «м» в «2», сноска «*») правятся
        здесь, до передачи единицы в нормализатор.
        """
        text = cls._cell_text(value).replace("*", "").strip()
        if not text:
            return ""
        stripped = text.replace(".", "").strip()
        # Латинские «HT»/кириллические «НТ» — OCR-опечатка «шт».
        if re.fullmatch(r"[HН][TТ]", stripped):
            return "шт"
        if stripped == "2":
            return "м2"
        return text

    @classmethod
    def _parse_decimal(cls, raw_value: Any) -> Optional[Decimal]:
        """Извлекает объем из ячеек ``58,3``, ``1 020``, ``123.45`` и ``52.2*``."""
        text = cls._cell_text(raw_value)
        if not text:
            return None
        normalized = text.replace("*", "").replace(" ", "").replace(",", ".")
        match = re.search(r"-?\d+(?:\.\d+)?", normalized)
        if match is None:
            return None
        try:
            return Decimal(match.group(0))
        except InvalidOperation:
            return None

    @staticmethod
    def _is_column_number_row(row: List[str]) -> bool:
        """Отсекает строку нумерации колонок ``1 | 2 | 3 | ...``."""
        values = [cell.strip() for cell in row if cell.strip()]
        return len(values) >= 3 and all(
            _COLUMN_NUMBER_PATTERN.fullmatch(value) for value in values
        )

    @staticmethod
    def _at(row: List[str], index: Optional[int]) -> str:
        """Безопасно возвращает значение колонки либо пустую строку."""
        if index is None or index < 0 or index >= len(row):
            return ""
        return row[index]

    @classmethod
    def _is_section_heading(cls, title: str) -> bool:
        """Отличает заголовок раздела («КАБЕЛИ», «Фасады») от описания изделия."""
        text = re.sub(r"\s+", " ", title).strip()
        lowered = text.lower()
        if not text or len(text) > 80:
            return False
        if "гост" in lowered or ":" in text:
            return False
        if text.endswith((".", ";", ",")):
            return False
        # Марки и позиции содержат цифры («Рубильник 3Ф 20А»), заголовки — нет.
        if re.search(r"\d", text):
            return False
        if text.isupper():
            return True
        return len(text.split()) <= 4

    @classmethod
    def _looks_like_unit(cls, value: str) -> bool:
        """Ячейка содержит только единицу измерения («шт.», «компл.», «пм»)."""
        text = cls._cell_text(value).lower()
        if not text or len(text) > 20:
            return False
        return _UNIT_VALUE_PATTERN.fullmatch(text) is not None

    @classmethod
    def _find_unit_index(
        cls, row: List[str], preferred: Optional[int] = None
    ) -> Optional[int]:
        """Индекс графы «Единица измерения» с учётом сдвига колонок на листе."""
        if (
            preferred is not None
            and 0 <= preferred < len(row)
            and cls._looks_like_unit(row[preferred])
        ):
            return preferred
        for index, cell in enumerate(row):
            if index == preferred:
                continue
            if cls._looks_like_unit(cell):
                return index
        return None

    @classmethod
    def _resolve_unit(
        cls, row: List[str], schema: _ColumnSchema
    ) -> tuple[str, Optional[int]]:
        """Возвращает текст и индекс графы единицы измерения."""
        index = cls._find_unit_index(row, schema.unit_col)
        if index is None:
            return cls._at(row, schema.unit_col), None
        return row[index], index

    @classmethod
    def _effective_schema(
        cls,
        row: List[str],
        schema: _ColumnSchema,
        unit_index: Optional[int],
    ) -> _ColumnSchema:
        """Пересчитывает графы относительно ед. изм. для листов-продолжений.

        Шапка ГОСТ 21.110 обычно только на первом листе и содержит больше
        колонок из-за объединенных ячеек (16 против 11). На таких листах
        графы 3-8 отсчитываются от графы единицы измерения по ГОСТ 21.110:
        марка = -4, код = -3, количество = +1, масса = +2.
        """
        if (
            unit_index is None
            or unit_index < 4
            or schema.column_count is None
            or len(row) >= schema.column_count
        ):
            return schema
        return replace(
            schema,
            qty_col=unit_index + 1,
            mass_col=unit_index + 2,
            mark_col=unit_index - 4,
            code_col=unit_index - 3,
        )

    @classmethod
    def _resolve_quantity(
        cls, row: List[str], schema: _ColumnSchema, unit_index: Optional[int]
    ) -> Optional[Decimal]:
        """Количество из графы объема; при пустой графе — первое число справа.

        Для ВОР (в т.ч. формы ГРАНД-Сметы) объем берется из первой строки
        ячейки: формула расчета («31,18+5,4») не является количеством.
        Графа 7 «Масса единицы, кг» никогда не трактуется как количество,
        даже если графа 8 в конкретной строке не заполнена.
        """
        raw_qty = cls._at(row, schema.qty_col)
        if schema.kind == "VOR":
            qty = cls._parse_vor_volume(raw_qty)
        else:
            qty = cls._parse_decimal(raw_qty)
        if qty is not None:
            return qty
        if unit_index is None:
            return None
        for index in range(unit_index + 1, len(row)):
            if schema.mass_col is not None and index == schema.mass_col:
                continue
            if schema.kind == "VOR":
                value = cls._parse_vor_volume(row[index])
            else:
                value = cls._parse_decimal(row[index])
            if value is not None:
                return value
        return None

    @classmethod
    def _resolve_vor_section(cls, row: List[str], schema: _ColumnSchema) -> str:
        """Заголовок раздела ВОР («Раздел: 1. Кровля»).

        В форме ГРАНД-Сметы слово «Раздел:» печатается в графе 1, а номер и
        название — в графе 2, поэтому графы склеиваются для контекста.
        """
        if schema.kind != "VOR":
            return ""
        first = cls._cell_text(cls._at(row, 0))
        title = cls._cell_text(cls._at(row, schema.title_col))
        text = f"{first} {title}".strip()
        if not _VOR_SECTION_PATTERN.match(text):
            return ""
        return re.sub(r"\s+", " ", text)

    @classmethod
    def _resolve_title(cls, row: List[str], schema: _ColumnSchema) -> str:
        """Наименование позиции; при сдвиге колонок ищется первый текст строки."""
        title = cls._at(row, schema.title_col).strip()
        if title:
            return title
        if schema.title_col < len(row):
            return ""
        for cell in row:
            text = cell.strip()
            if not text or _COLUMN_NUMBER_PATTERN.fullmatch(text):
                continue
            if cls._looks_like_unit(text) or not re.search(r"[A-Za-zА-Яа-я]", text):
                continue
            return text
        return ""

    @classmethod
    def _is_service_value(cls, value: str) -> bool:
        """Текст штампа/колонтитула, а не позиция спецификации."""
        lowered = cls._cell_text(value).lower()
        return any(token in lowered for token in _SERVICE_ROW_TOKENS)

    @classmethod
    def _resolve_mark(cls, row: List[str], schema: _ColumnSchema) -> str:
        """Тип, марка (графа 3) без захвата единицы измерения соседней графы."""
        value = cls._at(row, schema.mark_col).strip()
        if not value or cls._looks_like_unit(value) or cls._is_service_value(value):
            return ""
        return value

    @classmethod
    def _resolve_code(cls, row: List[str], schema: _ColumnSchema) -> str:
        """Код оборудования/материала (графа 4)."""
        value = cls._at(row, schema.code_col).strip()
        if not value or cls._looks_like_unit(value) or cls._is_service_value(value):
            return ""
        return value

    @classmethod
    def _resolve_position(cls, row: List[str], schema: _ColumnSchema) -> str:
        """Позиция (графа 1): «1», «3.1», «101А»."""
        value = cls._at(row, schema.position_col).strip()
        if value and _POSITION_PATTERN.fullmatch(value):
            return value
        first = row[0].strip() if row else ""
        if _POSITION_PATTERN.fullmatch(first):
            return first
        return ""

    def _is_specification_page(
        self,
        page,
        table: List[List[str]],
        schema: Optional[_ColumnSchema],
        has_marker: Optional[bool] = None,
    ) -> bool:
        """Страница относится к спецификации, а не к чертежу или расчету.

        Признаки: марка «*.С» / заголовок «Спецификация оборудования» либо
        таблица с заполненными графами единицы измерения и количества
        (продолжения спецификации идут без повторной шапки).

        ``has_marker`` передается из быстрого предфильтра PDFium, чтобы не
        запускать ``page.extract_text()`` повторно.
        """
        if has_marker is None:
            has_marker = self._has_spec_marker(page)
        if has_marker:
            return True
        if schema is None:
            return False
        return self._table_has_item_rows(table)

    @classmethod
    def _table_has_item_rows(cls, table: List[List[str]]) -> bool:
        """В таблице есть строки «единица измерения + количество»."""
        for row in table:
            unit_index = cls._find_unit_index(row)
            if unit_index is None:
                continue
            if any(
                cls._parse_decimal(cell) is not None
                for cell in row[unit_index + 1:]
            ):
                return True
        return False

    def extract_items(self) -> List[NormalizedItem]:
        """Извлекает позиции ВОР или спецификации с автоопределением колонок.

        Поддерживаются:

        * ВОР с 4-6 колонками (``№ | Наименование | Ед.изм | Объем``);
        * спецификации ГОСТ 21.110 произвольной ширины, включая А3-ландшафт;
        * разреженные таблицы, где часть ячеек пуста, а переносы текста
          оформлены отдельными строками.

        Страницы чертежей, однолинейных схем и светотехнических расчетов
        пропускаются. Пустые страницы и страницы без распознанной таблицы
        не приводят к исключениям: метод возвращает пустой список либо
        пропускает строки.

        Перед обращением к табличным методам выполняется быстрый текстовый
        предфильтр: PDFium читает текстовый слой всех страниц, и в pdfplumber
        передаются только страницы спецификаций/ВОР и продолжения таблиц.
        Поэтому find_tables() не запускается на векторных чертежах.
        """
        items: List[NormalizedItem] = []
        current_schema: Optional[_ColumnSchema] = None
        current_section = "Default_Section"
        # Предварительное сканирование PDFium: список страниц для детального
        # разбора. None — быстрый слой недоступен, работает резервный путь.
        preselected_pages = self._prescreen_pages()
        vor_continuation = False

        with pdfplumber.open(self.file_path) as pdf:
            total_pages = len(pdf.pages)
            if preselected_pages is None:
                page_numbers = list(range(1, total_pages + 1))
            else:
                page_numbers = sorted(
                    index
                    for index in preselected_pages
                    if 1 <= index <= total_pages
                )

            for p_idx in page_numbers:
                page = pdf.pages[p_idx - 1]
                fast_raw_text: Optional[str] = None

                if preselected_pages is None:
                    # Резервный путь (тестовые двойники либо сбой открытия
                    # PDFium): прежний текстовый префильтр pdfplumber.
                    extractor = getattr(page, "extract_text", None)
                    if not callable(extractor):
                        # Текстовый слой недоступен — фильтр не применяется.
                        raw_page_text: Optional[str] = None
                    else:
                        try:
                            raw_page_text = extractor() or ""
                        except Exception:
                            raw_page_text = None
                    if raw_page_text is None:
                        pass  # поведение до оптимизации: страница разбирается
                    else:
                        page_text = self._marker_text(raw_page_text)
                        has_spec = self._text_has_spec_markers(page_text)
                        has_vor = self._text_has_vor_markers(page_text)
                        is_continuation = (
                            vor_continuation
                            and self._text_is_vor_continuation(page_text)
                        )
                        if not has_spec and not has_vor and not is_continuation:
                            # Страницы чертежей, планов и однолинейных схем не
                            # доходят до построения таблиц по тысячам линий.
                            continue
                        vor_continuation = (not has_spec) and (
                            has_vor or is_continuation
                        )
                else:
                    fast_raw_text = (self._fast_page_texts or {}).get(p_idx)
                    fast_text = self._marker_text(fast_raw_text or "")
                    # Лист с растровой таблицей: в текстовом слое только
                    # штамп. pdfplumber не извлечет из картинки ни строк, ни
                    # количеств, поэтому дорогой разбор такого листа
                    # пропускается (для OCR нужен отдельный движок).
                    if (
                        fast_text
                        and self._text_has_spec_markers(fast_text)
                        and not self._text_has_table_rows(fast_text)
                        and not any(
                            marker in fast_text
                            for marker in _SPEC_TABLE_PAGE_MARKERS
                        )
                    ):
                        continue

                # Сшивка строк не переходит через границу листа: мусор
                # с чертежей и схем не перетекает в спецификацию.
                carried_over_row: Optional[List[str]] = None

                table = self._extract_page_table(page)
                if not table:
                    continue
                table = self._repair_table_encoding(page, table, fast_raw_text)

                page_schema = self._detect_table_schema(table)
                if page_schema is not None:
                    current_schema = page_schema

                has_marker: Optional[bool] = None
                if fast_raw_text is not None:
                    marker_text = self._marker_text(fast_raw_text)
                    has_marker = bool(
                        _SPEC_MARK_PATTERN.search(marker_text)
                        or _SPEC_TITLE_PATTERN.search(marker_text)
                    )
                if not self._is_specification_page(
                    page, table, current_schema, has_marker
                ):
                    continue

                for row in table:
                    if not row or not any(row):
                        continue
                    if self._is_column_number_row(row):
                        continue
                    if current_schema is None:
                        continue

                    raw_title = self._resolve_title(row, current_schema)
                    raw_unit, unit_index = self._resolve_unit(row, current_schema)
                    if current_schema.kind == "VOR":
                        # Опечатки единиц ВОР правятся до UnitNormalizer.
                        raw_unit = self._clean_vor_unit(raw_unit)
                    row_schema = self._effective_schema(
                        row, current_schema, unit_index
                    )
                    parsed_qty = self._resolve_quantity(
                        row, row_schema, unit_index
                    )

                    if parsed_qty is None:
                        section = self._resolve_vor_section(row, current_schema)
                        if section:
                            current_section = section
                            carried_over_row = None
                        elif (
                            not raw_unit
                            and raw_title
                            and self._is_section_heading(raw_title)
                        ):
                            # Подзаголовок раздела — только контекст, не позиция
                            # и не продолжение наименования предыдущей строки.
                            current_section = raw_title
                            carried_over_row = None
                        elif raw_title and not self._is_service_value(raw_title):
                            carried_over_row = row
                        continue

                    if not raw_title and carried_over_row is not None:
                        raw_title = self._resolve_title(
                            carried_over_row, current_schema
                        )

                    if not raw_title or self._is_service_value(raw_title):
                        carried_over_row = None
                        continue

                    items.append(
                        self._build_item(
                            row_schema,
                            row,
                            raw_title,
                            raw_unit,
                            parsed_qty,
                            current_section,
                            p_idx,
                            len(items) + 1,
                        )
                    )
                    carried_over_row = None

        return items

    def _build_item(
        self,
        schema: _ColumnSchema,
        row: List[str],
        raw_title: str,
        raw_unit: str,
        parsed_qty: Decimal,
        section: str,
        page_index: int,
        sequence: int,
    ) -> NormalizedItem:
        """Создает нормализованную позицию из строки распознанной таблицы."""
        raw_mark = self._resolve_mark(row, schema)
        raw_code = self._resolve_code(row, schema)
        position = self._resolve_position(row, schema)

        canonical_u, norm_qty, multiplier = UnitNormalizer.normalize(
            raw_unit, parsed_qty
        )
        clean_title = re.sub(r"\s+", " ", raw_title).lower().strip()
        dimensions = DimensionLockExtractor.extract(raw_title)

        if schema.kind == "SPECIFICATION":
            effective_code = raw_code or raw_mark or position or None
        else:
            effective_code = None

        pointer = CellPointer(
            doc_type=schema.kind,
            file_name=self.file_path,
            sheet_name=f"Page {page_index}",
            row_idx=sequence,
            col_letter=openpyxl.utils.get_column_letter(schema.qty_col + 1),
            coordinate=f"p.{page_index} [row {sequence}]",
        )

        return NormalizedItem(
            raw_id=f"{schema.kind}_{sequence}",
            code=effective_code,
            raw_title=raw_title,
            clean_title=clean_title,
            raw_unit=raw_unit,
            canonical_unit=canonical_u,
            unit_multiplier=multiplier,
            volume=norm_qty,
            dimensions=dimensions,
            context_section=section,
            pointer=pointer,
        )


# ----------------------------------------------------------------------
# 4.1. Парсер дефектных ведомостей и ВОР из Excel
# ----------------------------------------------------------------------

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

    def __init__(self, file_path: str | Path):
        self.file_path = Path(file_path)

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
            str(self.file_path), read_only=True, data_only=True
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
            file_name=str(self.file_path),
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
    file_path: str | Path,
    *,
    pdf_parser_cls: Callable[[str], PDFSpecParser] = PDFSpecParser,
) -> "PDFSpecParser | ExcelVORParser":
    """Фабрика парсера РД/ВОР по расширению файла.

    ``.xlsx`` (дефектная ведомость или ВОР) разбирается ``ExcelVORParser``,
    остальные файлы — существующим ``PDFSpecParser``. Класс PDF-парсера
    передается параметром, чтобы вызывающий код мог подменить его в тестах.

    :param file_path: путь к ведомости работ или тому РД.
    :param pdf_parser_cls: фабрика PDF-парсера (по умолчанию ``PDFSpecParser``).
    :return: экземпляр парсера с методом ``extract_items()``.
    """
    candidate = Path(file_path)
    if candidate.suffix.lower() == ".xlsx":
        return ExcelVORParser(candidate)
    return pdf_parser_cls(str(candidate))


# ----------------------------------------------------------------------
# 5. Парсер локальных смет из векторных PDF (форма 4б)
# ----------------------------------------------------------------------

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
_PDF_LSR_NUMBER_PATTERN: re.Pattern[str] = re.compile(r"^-?\d[\d\s.,]*$")
_PDF_LSR_MOJIBAKE_CODECS: Tuple[str, ...] = ("utf-8", "cp1251")
_PDF_LSR_LINE_TOLERANCE: float = 1.5


def _pdf_text_score(value: str) -> int:
    """Оценивает «естественность» текста для выбора кодировки.

    Считает кириллические символы положительным вкладом, а прочие
    не-ASCII символы (латиница с диакритикой, псевдографика cp866) —
    отрицательным.

    :param value: проверяемый фрагмент текста.
    :return: целочисленная оценка: чем больше, тем естественнее текст.
    """
    cyrillic = 0
    exotic = 0
    for char in value:
        if char.isspace() or char in "№«»—–-.,;:()[]/%":
            continue
        lowered = char.lower()
        if lowered == "ё" or "а" <= lowered <= "я":
            cyrillic += 1
        elif not char.isascii():
            exotic += 1
    return cyrillic * 2 - exotic * 2


def _repair_pdf_text(text: str) -> str:
    """Восстанавливает кириллицу после двойного искажения PDF-текста.

    Генераторы PDF смет выдают текст в UTF-8 или cp1251, который
    просмотрщик/извлекатель декодирует как cp866. Функция перебирает
    кандидатов ``encode('cp866').decode(codec)`` и выбирает вариант с
    наибольшей «кириллической» оценкой.

    :param text: исходный фрагмент, полученный от ``pdfplumber``.
    :return: восстановленный текст либо исходный фрагмент, если все
        кандидаты оказались хуже.
    """
    if not text or text.isascii():
        return text

    best = text
    best_score = _pdf_text_score(text)
    for codec in _PDF_LSR_MOJIBAKE_CODECS:
        try:
            candidate = text.encode("cp866").decode(codec)
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        score = _pdf_text_score(candidate)
        if score > best_score:
            best = candidate
            best_score = score
    return best


def _squash_spaces(value: str) -> str:
    """Сворачивает повторные пробелы и неразрывные пробелы в один.

    :param value: исходная строка.
    :return: строка без ведущих/замыкающих пробелов.
    """
    return re.sub(r"\s+", " ", (value or "").replace("\xa0", " ")).strip()


def _parse_pdf_number(value: str) -> Optional[Decimal]:
    """Разбирает число из ячейки PDF-сметы (пробелы — разделители тысяч).

    :param value: текст ячейки, например ``'1 287,00'``.
    :return: ``Decimal`` либо ``None``, если ячейка не является числом.
    """
    text = _squash_spaces(value)
    if not text or not _PDF_LSR_NUMBER_PATTERN.match(text):
        return None
    try:
        return Decimal(text.replace(" ", "").replace(",", "."))
    except InvalidOperation:
        return None


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

    def __init__(self, file_path: str):
        self.file_path = file_path

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

        with pdfplumber.open(self.file_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
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

                    if _PDF_LSR_SECTION_PATTERN.match(position_text or title_text):
                        section = _squash_spaces(
                            " ".join(
                                part
                                for part in (position_text, code_text, title_text)
                                if part
                            )
                        )
                        last_idx = None
                        continue

                    if LSRParser.TOTAL_ROW_PATTERN.match(title_text):
                        total = self._cell_number(cells, layout.total_cost_col)
                        if total is None:
                            total = self._cell_number(cells, layout.unit_price_col)
                        if total is not None and pending_idx is not None:
                            self._apply_position_total(items[pending_idx], total)
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
                        quantity = self._cell_number(cells, layout.qty_unit_col)
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
                    canonical_unit, volume, multiplier = UnitNormalizer.normalize(
                        unit_text, quantity
                    )
                    pointer = CellPointer(
                        doc_type="LSR",
                        file_name=self.file_path,
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
                            clean_title=LSRParser._clean_estimate_title(title_text),
                            raw_unit=unit_text,
                            canonical_unit=canonical_unit,
                            unit_multiplier=multiplier,
                            volume=volume,
                            dimensions=DimensionLockExtractor.extract(title_text),
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

