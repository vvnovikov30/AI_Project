from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import re
from typing import Optional, Dict, Any, List
import openpyxl
import openpyxl.utils
import pdfplumber

from models import NormalizedItem, CellPointer
from normalizer import UnitNormalizer, DimensionLockExtractor


class LSRParser:
    """Парсер сметных расчетов Excel (форма по Приказу 421/пр и ТСН-2001 г. Москвы)."""

    # Ловит федеральные расценки (ГЭСН, ФЕР, ТЕР, 2-значные главы 421/пр)
    # и расценки на работы ТСН-2001 (главы 3, 4, 5, 6: монтаж, строительные, пусконаладка)
    BASE_RATE_PATTERN = re.compile(r"^(ГЭСН|ФЕР|ТЕР|[3-6]\.[\d\-]+|\d{2}\.\d+\.\d+)")

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

    def parse(self, sheet_name: Optional[str] = None) -> List[NormalizedItem]:
        """Парсинг листа сметы в нормализованную плоскую коллекцию DTO.

        :param sheet_name: имя листа ``ЛС ...`` для пакетного/точечного
            прогона. Без параметра используется активный лист книги.
        :raises ValueError: запрошенный лист отсутствует в книге.
        """
        if sheet_name is not None and sheet_name not in self.wb.sheetnames:
            raise ValueError(
                f"Лист '{sheet_name}' не найден в книге {self.file_path}. "
                f"Доступны: {', '.join(self.wb.sheetnames)}"
            )
        ws = self.wb[sheet_name] if sheet_name else self.wb.active
        merge_map = self._resolve_merged_cells_map(ws)
        items: List[NormalizedItem] = []

        current_parent_id = None
        current_section = "Default_Section"

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

            # Проверка, является ли строка базовой расценкой работ
            first_code_line = raw_code.split("\n")[0].strip() if raw_code else ""
            is_base_rate = bool(self.BASE_RATE_PATTERN.match(first_code_line))
            item_id = f"LSR_{row_idx}"

            if is_base_rate:
                current_parent_id = item_id
                parent_ref = None
            else:
                parent_ref = current_parent_id

            canonical_u, norm_qty, mult = UnitNormalizer.normalize(raw_unit, qty_decimal)
            clean_title = self._clean_estimate_title(raw_title)
            dimensions = DimensionLockExtractor.extract(raw_title)

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
                is_resource=not is_base_rate,
                parent_id=parent_ref
            ))

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
_COLUMN_NUMBER_PATTERN = re.compile(r"^\d{1,2}$")


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


class PDFSpecParser:
    """Извлечение таблиц ГОСТ 21.110 и ведомостей объемов работ (ВОР).

    Схема колонок определяется по строке заголовка, поэтому один парсер
    поддерживает и 4-5-колоночные ВОР, и много-колоночные спецификации
    оборудования. Пустые и короткие строки не индексируются напрямую.
    """

    STAMP_W_PT = 185 * 2.83465
    STAMP_H_PT = 55 * 2.83465

    def __init__(self, file_path: str, debug: bool = False):
        self.file_path = file_path
        self.debug = debug

    def _get_crop_box(self, page):
        width = page.width
        height = page.height
        return (0, 0, width, height - self.STAMP_H_PT)

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

    @classmethod
    def _detect_schema(cls, row: List[str]) -> Optional[_ColumnSchema]:
        """Определяет схему таблицы по строке заголовков.

        Поддерживаются оба целевых формата:

        * ГОСТ 21.110 (спецификация оборудования, >= 7 колонок);
        * ВОР (4-6 колонок: №, наименование, ед. изм., количество).

        Индексы колонок берутся из шапки, а не из фиксированных позиций,
        поэтому дополнительные служебные колонки не ломают разбор.
        """
        cells = [cls._cell_text(cell).lower() for cell in row]
        # Переносы в шапке PDF дают «Количест во», «Ед. изм»; сравнение
        # выполняется по склеенному тексту, чтобы не зависеть от пробелов.
        keys = [cell.replace(" ", "") for cell in cells]
        joined = " ".join(cells)

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
        if title_col is None or unit_col is None or qty_col is None:
            return None

        is_specification = (
            len(cells) >= 7
            or any(token in joined for token in _HEADER_SPEC_TOKENS)
        )
        position_col = next(
            (
                index
                for index, cell in enumerate(cells)
                if cell in {"позиция", "№", "n п/п", "№ п/п"}
            ),
            None,
        )
        mark_col = next(
            (
                index
                for index, cell in enumerate(cells)
                if "тип, марка" in cell
                or "тип,марка" in cell
                or cell == "марка"
            ),
            None,
        )
        code_col = next(
            (
                index
                for index, cell in enumerate(cells)
                if cell.startswith("код")
            ),
            None,
        )
        return _ColumnSchema(
            kind="SPECIFICATION" if is_specification else "VOR",
            title_col=title_col,
            unit_col=unit_col,
            qty_col=qty_col,
            position_col=position_col,
            mark_col=mark_col,
            code_col=code_col,
        )

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
        """Отличает короткий заголовок раздела ВОР от описания изделия."""
        lowered = title.lower()
        return (
            len(title) <= 80
            and "гост" not in lowered
            and not title.endswith((".", ";"))
            and ":" not in title
        )

    def extract_items(self) -> List[NormalizedItem]:
        """Извлекает позиции ВОР или спецификации с автоопределением колонок.

        Поддерживаются:

        * ВОР с 4-6 колонками (``№ | Наименование | Ед.изм | Объем``);
        * спецификации ГОСТ 21.110 с произвольным числом колонок;
        * разреженные таблицы, где часть ячеек пуста, а переносы текста
          оформлены отдельными строками.

        Пустые страницы и страницы без распознанной шапки не приводят к
        исключениям: метод возвращает пустой список либо пропускает строки.
        """
        items: List[NormalizedItem] = []
        current_schema: Optional[_ColumnSchema] = None
        current_section = "Default_Section"

        with pdfplumber.open(self.file_path) as pdf:
            for p_idx, page in enumerate(pdf.pages, start=1):
                crop_area = self._get_crop_box(page)
                cropped_page = page.within_bbox(crop_area)

                table = cropped_page.extract_table({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 3,
                })

                if not table:
                    table = cropped_page.extract_table({
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text"
                    })

                if not table:
                    continue

                for raw_row in table:
                    if not raw_row:
                        continue

                    row = [self._cell_text(cell) for cell in raw_row]
                    if not any(row):
                        continue

                    schema = self._detect_schema(row)
                    if schema is not None:
                        current_schema = schema
                        continue

                    if self._is_column_number_row(row):
                        continue
                    if current_schema is None:
                        continue

                    raw_title = self._at(row, current_schema.title_col)
                    raw_unit = self._at(row, current_schema.unit_col)
                    parsed_qty = self._parse_decimal(
                        self._at(row, current_schema.qty_col)
                    )

                    if not raw_title:
                        continue

                    if parsed_qty is None:
                        if not raw_unit and self._is_section_heading(raw_title):
                            current_section = raw_title
                        continue

                    items.append(
                        self._build_item(
                            current_schema,
                            row,
                            raw_title,
                            raw_unit,
                            parsed_qty,
                            current_section,
                            p_idx,
                            len(items) + 1,
                        )
                    )

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
        raw_mark = self._at(row, schema.mark_col)
        raw_code = self._at(row, schema.code_col)
        position = self._at(row, schema.position_col)

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
