import re
from decimal import Decimal
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
        """Парсинг листа сметы в нормализованную плоскую коллекцию DTO с фильтрацией шума."""
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


class PDFSpecParser:
    """Извлечение спецификаций ГОСТ 21.110-2013 с отладкой и фильтрацией штампов."""

    STAMP_W_PT = 185 * 2.83465
    STAMP_H_PT = 55 * 2.83465

    def __init__(self, file_path: str, debug: bool = False):
        self.file_path = file_path
        self.debug = debug

    def _get_crop_box(self, page):
        width = page.width
        height = page.height
        return (0, 0, width, height - self.STAMP_H_PT)

    def extract_items(self) -> List[NormalizedItem]:
        """Извлечение строк спецификации со сшивкой разорванных таблиц и фильтрацией служебных строк."""
        aggregated_rows = []
        carried_over_row = None

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

                for row in table:
                    if not row:
                        continue

                    # 1. Пропуск строки заголовков ("Наименование", "Тип, марка" и т.д.)
                    if any("наименование" in str(c).lower() for c in row if c):
                        continue

                    # 2. Пропуск строки с номерами колонок ГОСТ 21.110: ['1', '2', '3', '4', '5', '6', '7', '8', '9']
                    first_cell = str(row[0] or "").strip()
                    second_cell = str(row[1] or "").strip() if len(row) > 1 else ""
                    if first_cell == "1" and second_cell == "2":
                        continue

                    cleaned_row = [str(c or "").strip() for c in row]

                    # Сшивка переносов текста наименования (если колонка 0 пустая, а в 1 продолжается текст)
                    if not cleaned_row[0] and len(cleaned_row) > 1 and cleaned_row[1] and carried_over_row:
                        carried_over_row[1] += f" {cleaned_row[1]}"
                        continue

                    if carried_over_row:
                        aggregated_rows.append((p_idx, carried_over_row))
                    carried_over_row = cleaned_row

            if carried_over_row:
                aggregated_rows.append((len(pdf.pages), carried_over_row))

        items: List[NormalizedItem] = []
        for idx, (p_idx, r) in enumerate(aggregated_rows, start=1):
            if len(r) < 7:
                continue

            # Структура колонок ГОСТ 21.110:
            # r[0] - № позиции | r[1] - Наименование | r[2] - Марка, тип
            # r[3] - Код | r[4] - Завод | r[5] - Ед. изм. | r[6] - Количество
            raw_title = r[1] if len(r) > 1 else ""
            raw_mark = r[2] if len(r) > 2 else ""
            raw_code = r[3] if len(r) > 3 else ""
            raw_unit = r[5] if len(r) > 5 else "шт"
            raw_qty_str = r[6] if len(r) > 6 else "0"

            try:
                clean_qty_str = raw_qty_str.replace("\n", "").replace(" ", "").replace(",", ".")
                qty_decimal = Decimal(clean_qty_str)
            except Exception:
                continue

            pointer = CellPointer(
                doc_type="SPECIFICATION",
                file_name=self.file_path,
                sheet_name=f"Page {p_idx}",
                row_idx=idx,
                col_letter="D",
                coordinate=f"p.{p_idx} [row {idx}]"
            )

            canonical_u, norm_qty, mult = UnitNormalizer.normalize(raw_unit, qty_decimal)
            clean_title = re.sub(r"\s+", " ", raw_title).lower().strip()
            dimensions = DimensionLockExtractor.extract(raw_title)

            # В качестве кода позиции берем марку оборудования (r[2]) или артикул (r[3]),
            # а при их отсутствии — порядковый номер (r[0])
            effective_code = raw_code or raw_mark or (r[0] if r[0] else None)

            items.append(NormalizedItem(
                raw_id=f"SPEC_{idx}",
                code=effective_code,
                raw_title=raw_title,
                clean_title=clean_title,
                raw_unit=raw_unit,
                canonical_unit=canonical_u,
                unit_multiplier=mult,
                volume=norm_qty,
                dimensions=dimensions,
                context_section="Default_Section",
                pointer=pointer
            ))

        return items