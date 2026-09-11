import re
from decimal import Decimal
from typing import Optional, Dict, Any, List
import openpyxl
import openpyxl.utils
import pdfplumber

from models import NormalizedItem, CellPointer
from normalizer import UnitNormalizer, DimensionLockExtractor

class LSRParser:
    """Парсер сметных расчетов Excel (форма по Приказу 421/пр)."""
    
    BASE_RATE_PATTERN = re.compile(r"^(ГЭСН|ФЕР|ТЕР|\d{2}\.\d+\.\d+)")
    
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

    def parse(self, sheet_name: Optional[str] = None) -> List[NormalizedItem]:
        """Парсинг листа сметы в нормализованную плоскую коллекцию DTO."""
        ws = self.wb[sheet_name] if sheet_name else self.wb.active
        merge_map = self._resolve_merged_cells_map(ws)
        items: List[NormalizedItem] = []
        
        current_parent_id = None
        
        for row_idx in range(1, ws.max_row + 1):
            def get_val(col_idx: int):
                return merge_map.get((row_idx, col_idx), ws.cell(row=row_idx, column=col_idx).value)
            
            raw_code = str(get_val(2) or "").strip()
            raw_title = str(get_val(3) or "").strip()
            raw_unit = str(get_val(4) or "").strip()
            raw_qty = get_val(5)
            
            if not raw_title or raw_qty is None:
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
            
            is_base_rate = bool(self.BASE_RATE_PATTERN.match(raw_code))
            item_id = f"LSR_{row_idx}"
            
            if is_base_rate:
                current_parent_id = item_id
                parent_ref = None
            else:
                parent_ref = current_parent_id

            # Нормализация единиц и извлечение параметров
            canonical_u, norm_qty, mult = UnitNormalizer.normalize(raw_unit, qty_decimal)
            clean_title = re.sub(r"\s+", " ", raw_title).lower().strip()
            dimensions = DimensionLockExtractor.extract(raw_title)

            items.append(NormalizedItem(
                raw_id=item_id,
                code=raw_code if raw_code else None,
                raw_title=raw_title,
                clean_title=clean_title,
                raw_unit=raw_unit,
                canonical_unit=canonical_u,
                unit_multiplier=mult,
                volume=norm_qty,
                dimensions=dimensions,
                context_section="Default_Section",
                pointer=pointer,
                is_resource=not is_base_rate,
                parent_id=parent_ref
            ))
            
        return items


class PDFSpecParser:
    """Извлечение спецификаций ГОСТ 21.110-2013 с фильтрацией штампов."""

    STAMP_W_PT = 185 * 2.83465
    STAMP_H_PT = 55 * 2.83465

    def __init__(self, file_path: str):
        self.file_path = file_path

    def _get_crop_box(self, page):
        width = page.width
        height = page.height
        return (0, 0, width, height - self.STAMP_H_PT)

    def extract_items(self) -> List[NormalizedItem]:
        """Извлечение строк спецификации со сшивкой разорванных таблиц в DTO."""
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
                    if not row or any("наименование" in str(c).lower() for c in row if c):
                        continue
                    
                    cleaned_row = [str(c or "").strip() for c in row]
                    
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
            if len(r) < 4:
                continue
            
            raw_title = r[1] if len(r) > 1 else ""
            raw_unit = r[2] if len(r) > 2 else "шт"
            raw_qty_str = r[3] if len(r) > 3 else "0"
            
            try:
                qty_decimal = Decimal(raw_qty_str.replace(",", ".").strip())
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

            items.append(NormalizedItem(
                raw_id=f"SPEC_{idx}",
                code=r[0] if r[0] else None,
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