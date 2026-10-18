from dataclasses import dataclass, replace
from decimal import Decimal, InvalidOperation
import re
from typing import Optional, Dict, Any, List
import openpyxl
import openpyxl.utils
import pdfplumber
import pypdfium2 as pdfium

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
    "наименование вида работ",
    "обоснование объема",
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
        rows = self._extract_rows_via_finder(page, crop_box)
        if rows:
            return rows

        for settings in (self._LINE_SETTINGS, self._TEXT_SETTINGS):
            table = self._extract_raw_table(page, settings)
            if not table:
                continue
            # Текстовая стратегия на чертежах «видит» таблицу в любом наборе
            # подписей, поэтому без распознанной шапки её результат не берем.
            if settings == self._TEXT_SETTINGS and not self._table_has_header(table):
                continue
            return table
        return None

    def _table_has_header(self, table: List[List[str]]) -> bool:
        """В таблице есть шапка ГОСТ 21.110 либо ВОР."""
        if self._has_gost_graph_numbers(table):
            return True
        return any(self._detect_schema(row) is not None for row in table[:8])

    def _extract_rows_via_finder(
        self, page, crop_box: tuple[float, float, float, float]
    ) -> List[List[str]]:
        """Строки самой крупной таблицы страницы без строк нижнего штампа."""
        finder = getattr(page, "find_tables", None)
        if not callable(finder):
            return []
        try:
            tables = finder(self._LINE_SETTINGS)
        except Exception:
            return []

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
        return best

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
        * ВОР (4-6 колонок: №, наименование, ед. изм., количество).

        Индексы колонок берутся из шапки, а не из фиксированных позиций,
        поэтому дополнительные служебные колонки не ломают разбор.
        """
        cells = [cls._cell_text(cell).lower() for cell in row]
        # Переносы в шапке PDF дают «Количест во», «Ед. изм», «Коли-чество»;
        # сравнение выполняется по склеенному тексту без пробелов и дефисов.
        keys = [re.sub(r"[\s\-]+", "", cell) for cell in cells]
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
        mass_col = next(
            (
                index
                for index, cell in enumerate(keys)
                if any(token in cell for token in _HEADER_MASS_TOKENS)
            ),
            None,
        )

        is_specification = (
            len(cells) >= 7
            or any(token in joined for token in _HEADER_SPEC_TOKENS)
        )
        if title_col is None:
            return None
        if unit_col is None:
            # ГОСТ 21.110: графа 6 (индекс 5) — единица измерения.
            if not is_specification or len(cells) <= 5:
                return None
            unit_col = 5
        if qty_col is None:
            # ГОСТ 21.110: графа 8 (индекс 7); при отсутствии графы массы —
            # индекс 6. Значение всё равно уточняется по содержимому строки.
            if not is_specification or len(cells) <= 6:
                return None
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
        """Количество из графы 8; при пустой графе — первое число справа от ед. изм.

        Графа 7 «Масса единицы, кг» никогда не трактуется как количество,
        даже если графа 8 в конкретной строке не заполнена.
        """
        qty = cls._parse_decimal(cls._at(row, schema.qty_col))
        if qty is not None:
            return qty
        if unit_index is None:
            return None
        for index in range(unit_index + 1, len(row)):
            if schema.mass_col is not None and index == schema.mass_col:
                continue
            value = cls._parse_decimal(row[index])
            if value is not None:
                return value
        return None

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
                    row_schema = self._effective_schema(
                        row, current_schema, unit_index
                    )
                    parsed_qty = self._resolve_quantity(
                        row, row_schema, unit_index
                    )

                    if parsed_qty is None:
                        if (
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
