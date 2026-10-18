import os
import re
from decimal import Decimal
import openpyxl
import pytest

from models import UnitClass
import parsers
from parsers import LSRParser, PDFSpecParser


# ==============================================================================
# 1. МОДУЛЬНЫЕ ТЕСТЫ LSRParser (СИНТЕТИЧЕСКИЕ ДАННЫЕ В ПАМЯТИ)
# ==============================================================================

class TestLSRParserSynthetic:
    """Проверка логики парсера смет без зависимости от внешних файлов."""

    def test_clean_estimate_title(self):
        """Очистка наименований от формул ценообразования, префиксов и артефактов."""
        raw_1 = (
            "ОБОРУДОВАНИЕ:\n"
            "Цилиндрическая IP-видеокамера с объективом 4 Мп\n"
            "1 236,60 = [10 913 / 1,22 / 7,54] + 3% Трансп + 1,2% Заг.скл"
        )
        cleaned_1 = LSRParser._clean_estimate_title(raw_1)
        assert cleaned_1 == "цилиндрическая ip-видеокамера с объективом 4 мп"

        raw_2 = "ОБОРУДОВАНИЕ:_x000D_\nКлавиатура компьютерная USB."
        cleaned_2 = LSRParser._clean_estimate_title(raw_2)
        assert cleaned_2 == "клавиатура компьютерная usb."

        raw_3 = "Монтаж наружной камеры видеонаблюдения"
        cleaned_3 = LSRParser._clean_estimate_title(raw_3)
        assert cleaned_3 == "монтаж наружной камеры видеонаблюдения"

    def test_merged_cells_resolution(self, tmp_path):
        """Корректное извлечение значений из объединенных ячеек Excel."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value="Объединенный текст")
        ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=2)

        file_path = tmp_path / "test_merged.xlsx"
        wb.save(file_path)

        parser = LSRParser(str(file_path))
        merge_map = parser._resolve_merged_cells_map(parser.wb.active)

        assert merge_map[(1, 1)] == "Объединенный текст"
        assert merge_map[(1, 2)] == "Объединенный текст"
        assert merge_map[(2, 1)] == "Объединенный текст"
        assert merge_map[(2, 2)] == "Объединенный текст"

    def test_noise_filtering_and_hierarchy(self, tmp_path):
        """Фильтрация служебных строк (НР, СП, ЗТР, шапки) и построение дерева parent_id."""
        wb = openpyxl.Workbook()
        ws = wb.active

        # Шапка и раздел
        ws.cell(row=1, column=1, value="ЛОКАЛЬНАЯ СМЕТА № 02-01-14")
        ws.cell(row=2, column=1, value="Раздел: Оборудование")

        # 1. Базовая расценка на монтаж (№ 1)
        ws.cell(row=3, column=1, value="1")
        ws.cell(row=3, column=2, value="4.10-174-2\nПоправка: Гл.12")
        ws.cell(row=3, column=3, value="Монтаж наружной камеры видеонаблюдения")
        ws.cell(row=3, column=4, value="10 шт.")
        ws.cell(row=3, column=5, value="0,5")

        # Служебный шум (должен быть проигнорирован)
        ws.cell(row=4, column=1, value=None)
        ws.cell(row=4, column=3, value="НР от ФОТ")
        ws.cell(row=4, column=4, value="%")
        ws.cell(row=4, column=5, value=90)

        ws.cell(row=5, column=1, value=None)
        ws.cell(row=5, column=3, value="СП от ФОТ")
        ws.cell(row=5, column=4, value="%")
        ws.cell(row=5, column=5, value=46)

        ws.cell(row=6, column=1, value=None)
        ws.cell(row=6, column=3, value="ЗТР")
        ws.cell(row=6, column=4, value="чел-ч")
        ws.cell(row=6, column=5, value=25.33)

        # 2. Дочерний ресурс оборудования (№ 2)
        ws.cell(row=7, column=1, value="2")
        ws.cell(row=7, column=2, value="Прайс-лист")
        ws.cell(row=7, column=3, value="ОБОРУДОВАНИЕ:\nКамера уличная 4Мп")
        ws.cell(row=7, column=4, value="ШТ")
        ws.cell(row=7, column=5, value=5)

        # 3. Служебная строка-нумерация граф таблицы: 1 | 2 | 3 | 4 | 5
        ws.cell(row=8, column=1, value="1")
        ws.cell(row=8, column=2, value="2")
        ws.cell(row=8, column=3, value="3")
        ws.cell(row=8, column=4, value="4")
        ws.cell(row=8, column=5, value=5)

        # 4. Новый раздел и расценка
        ws.cell(row=9, column=1, value="Раздел: Кабельные изделия")
        ws.cell(row=10, column=1, value="3")
        ws.cell(row=10, column=2, value="4.8-280-2")
        ws.cell(row=10, column=3, value="Прокладка пластикового кабель-канала")
        ws.cell(row=10, column=4, value="100 м")
        ws.cell(row=10, column=5, value=3)

        # 5. Материал внутри расценки с дробным номером (№ 3,1)
        ws.cell(row=11, column=1, value="3,1")
        ws.cell(row=11, column=2, value="1.1-1-1962")
        ws.cell(row=11, column=3, value="Кабель-канал 10х15 мм")
        ws.cell(row=11, column=4, value="м")
        ws.cell(row=11, column=5, value=300)

        file_path = tmp_path / "test_hierarchy.xlsx"
        wb.save(file_path)

        parser = LSRParser(str(file_path))
        items = parser.parse()

        # Ровно 4 валидные позиции (1, 2, 3, 3,1)
        assert len(items) == 4

        # Проверка базовой расценки № 1
        it_work = items[0]
        assert it_work.code == "4.10-174-2"
        assert not it_work.is_resource
        assert it_work.parent_id is None
        assert it_work.volume == Decimal("5.0")  # 0.5 * 10
        assert it_work.canonical_unit == UnitClass.COUNT
        assert it_work.context_section == "Раздел: Оборудование"

        # Проверка дочернего ресурса № 2
        it_res = items[1]
        assert it_res.code == "Прайс-лист"
        assert it_res.is_resource
        assert it_res.parent_id == "LSR_3"  # Привязан к родительской расценке строки 3
        assert it_res.volume == Decimal("5")
        assert it_res.clean_title == "камера уличная 4мп"

        # Проверка сброса parent_id при новой расценке
        it_work2 = items[2]
        assert it_work2.code == "4.8-280-2"
        assert not it_work2.is_resource
        assert it_work2.parent_id is None
        assert it_work2.volume == Decimal("300.0")  # 3 * 100
        assert it_work2.canonical_unit == UnitClass.LENGTH
        assert it_work2.context_section == "Раздел: Кабельные изделия"

        # Проверка дочернего материала № 3,1
        it_res2 = items[3]
        assert it_res2.code == "1.1-1-1962"
        assert it_res2.is_resource
        assert it_res2.parent_id == "LSR_10"
        assert it_res2.volume == Decimal("300")


# ==============================================================================
# 2. МОДУЛЬНЫЕ ТЕСТЫ PDFSpecParser
# ==============================================================================

class TestPDFSpecParserUnit:
    """Проверка вспомогательных методов извлечения спецификации."""

    def test_crop_box_calculation(self):
        """Проверка динамической области таблицы (отступы 15/15/15 и срез штампа 65 pt)."""
        class DummyPage:
            width = 842.0   # A4 альбомная
            height = 595.0

        parser = PDFSpecParser("dummy.pdf", debug=False)
        box = parser._get_crop_box(DummyPage())

        assert box[0] == PDFSpecParser.CROP_LEFT_PT
        assert box[1] == PDFSpecParser.CROP_TOP_PT
        assert pytest.approx(box[2], 0.01) == 842.0 - PDFSpecParser.CROP_RIGHT_PT
        assert pytest.approx(box[3], 0.01) == 595.0 - PDFSpecParser.CROP_BOTTOM_STAMP_PT

    def test_crop_box_a3_landscape_uses_actual_page_width(self):
        """А3-ландшафт (1191x842): правые графы «Ед. изм.» и «Количество» не срезаются.

        CropBox считается от фактической ширины листа, а не от константы А4,
        поэтому широкие спецификации ЭОМ разбираются полностью.
        """
        class A3LandscapePage:
            width = 1191.0
            height = 842.0

        parser = PDFSpecParser("dummy.pdf", debug=False)
        left, top, right, bottom = parser._get_crop_box(A3LandscapePage())

        assert (left, top) == (15, 15)
        assert pytest.approx(right, 0.01) == 1176.0
        assert pytest.approx(bottom, 0.01) == 777.0
        # Правая граница шире листа А4 — графа количества не отрезана.
        assert right > 595.0

    def test_quantity_mapping_gost_21_110(self):
        """Количество берется из графы 8; масса (графа 7) не становится количеством.

        Дополнительно проверяется сдвиг колонок на страницах-продолжениях:
        при пустой графе 8 количество ищется как первое число справа от ед. изм.
        """
        schema = PDFSpecParser._gost_positional_schema()

        # Графа 8 заполнена — значение берется напрямую.
        row_full = ["8", "Лоток перфорированный", "УЛП-100", "", "", "м", "12", "120", ""]
        assert PDFSpecParser._resolve_quantity(row_full, schema, 5) == Decimal("120")

        # Графа 8 пуста: масса 12 кг не должна попасть в количество.
        row_mass_only = ["7", "Лоток перфорированный", "УЛП-100", "", "", "м", "12", "", ""]
        assert PDFSpecParser._resolve_quantity(row_mass_only, schema, 5) is None

        # Страница-продолжение без шапки: 11 колонок, ед. изм. в графе 7.
        continuation = parsers._ColumnSchema(
            kind="SPECIFICATION",
            title_col=1,
            unit_col=10,
            qty_col=11,
            position_col=0,
            mark_col=2,
            code_col=6,
            mass_col=13,
        )
        row_shifted = ["19", "Щиток учета", "ЯТП-0,25/36", "", "", "", "компл.", "1", "", ""]
        unit_text, unit_index = PDFSpecParser._resolve_unit(row_shifted, continuation)
        assert unit_text == "компл."
        assert unit_index == 6
        assert PDFSpecParser._resolve_quantity(
            row_shifted, continuation, unit_index
        ) == Decimal("1")

    def test_section_headings_are_context_not_titles(self, monkeypatch):
        """Подзаголовки разделов дают context_section и не склеиваются с позициями.

        Сшивка «продолжение наименования» сбрасывается на границе листа, поэтому
        строка с количеством без наименования с новой страницы не создает фантом.
        """
        page_one = [
            [
                "Позиция",
                "Наименование и техническая характеристика",
                "Тип, марка",
                "Код оборудования",
                "",
                "Ед. изм.",
                "Масса",
                "Количество",
            ],
            ["", "КАБЕЛИ", "", "", "", "", "", ""],
            ["1", "Кабель ППГнг(А)-HF 3х2,5", "ППГнг(А)-HF", "", "", "м", "", "120"],
            ["", "Кабель ВВГнг 3х1,5 (продолжение линии 2)", "", "", "", "", "", ""],
        ]
        page_two = [
            ["2", "", "", "", "", "м", "", "15"],
            ["", "ТРУБЫ, ЛОТКИ", "", "", "", "", "", ""],
            ["3", "Труба гофрированная Д=20", "IEK", "", "", "м", "", "1800"],
        ]

        class _FakePage:
            def __init__(self, table):
                self.width = 1191.0
                self.height = 842.0
                self._table = table

            def extract_table(self, _settings):
                return self._table

        class _FakePDF:
            pages = [_FakePage(page_one), _FakePage(page_two)]

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        monkeypatch.setattr(parsers.pdfplumber, "open", lambda _path: _FakePDF())

        items = PDFSpecParser("eom_spec.pdf").extract_items()

        assert len(items) == 2
        assert [item.volume for item in items] == [Decimal("120"), Decimal("1800")]
        assert [item.context_section for item in items] == ["КАБЕЛИ", "ТРУБЫ, ЛОТКИ"]
        # Наименования не содержат приклеенного заголовка раздела.
        assert items[0].raw_title == "Кабель ППГнг(А)-HF 3х2,5"
        assert items[1].raw_title == "Труба гофрированная Д=20"
        # Фантом из строки без наименования на новом листе не появился.
        assert all("продолжение" not in item.clean_title for item in items)

    def test_pdf_parser_extracts_vor_4_columns(self, monkeypatch):
        """ВОР из 4 колонок распознается без жесткого требования 7 колонок."""
        table = [
            ["№", "Наименование вида работ", "Ед.изм", "Объем работ"],
            ["", "Фасады", "", ""],
            ["", "Монтаж водосточного желоба", "пм.", "58,3"],
            ["", "Установка домового знака", "шт", "1 020"],
            ["", "Монтаж люка", "шт.", "52.2*"],
        ]

        class _FakePage:
            width = 842.0
            height = 595.0

            def within_bbox(self, _bbox):
                return self

            def extract_table(self, _settings):
                return table

        class _FakePDF:
            pages = [_FakePage()]

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        monkeypatch.setattr(parsers.pdfplumber, "open", lambda _path: _FakePDF())

        items = PDFSpecParser("vor.pdf").extract_items()

        assert [item.volume for item in items] == [
            Decimal("58.3"),
            Decimal("1020"),
            Decimal("52.2"),
        ]
        assert [item.raw_unit for item in items] == ["пм.", "шт", "шт."]
        assert all(item.context_section == "Фасады" for item in items)

    def test_pdf_parser_empty_pages_returns_empty_list_without_error(self, monkeypatch):
        """Пустая таблица и строка длиной 1 не вызывают IndexError."""
        table = [
            ["№", "Наименование", "Ед.изм", "Количество"],
            [None],
        ]

        class _FakePage:
            width = 842.0
            height = 595.0

            def within_bbox(self, _bbox):
                return self

            def extract_table(self, _settings):
                return table

        class _FakePDF:
            pages = [_FakePage()]

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        monkeypatch.setattr(parsers.pdfplumber, "open", lambda _path: _FakePDF())

        assert PDFSpecParser("empty.pdf").extract_items() == []

    def test_page_without_spec_markers_skips_table_extraction(self, monkeypatch):
        """Страница чертежа без маркеров не доходит до табличных методов.

        Предфильтр по ``extract_text()`` обязан отсечь страницу до вызова
        ``find_tables()``/``extract_table()``: иначе разбор векторных планов
        и однолинейных схем зависает на десятках тысяч линий.
        """
        class _DrawingPage:
            width = 1191.0
            height = 842.0

            def extract_text(self):
                return "План 1-го этажа. Однолинейная схема ВРУ-1"

            def find_tables(self, *_args, **_kwargs):
                raise AssertionError("find_tables не должен вызываться для чертежа")

            def extract_table(self, *_args, **_kwargs):
                raise AssertionError("extract_table не должен вызываться для чертежа")

        class _FakePDF:
            pages = [_DrawingPage()]

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        monkeypatch.setattr(parsers.pdfplumber, "open", lambda _path: _FakePDF())

        assert PDFSpecParser("eom_drawings.pdf").extract_items() == []

    def test_page_with_spec_marker_is_parsed(self, monkeypatch):
        """Страница с маркером спецификации разбирается как обычно."""
        table = [
            [
                "Позиция",
                "Наименование и техническая характеристика",
                "Тип, марка",
                "Код оборудования",
                "",
                "Ед. изм.",
                "Масса",
                "Количество",
            ],
            ["1", "Щиток учета", "ЯТП-0,25/36", "", "", "компл.", "", "1"],
        ]

        class _SpecPage:
            width = 1191.0
            height = 842.0

            def extract_text(self):
                return "2025-00224-01-0-ИОС1.С Спецификация оборудования"

            def extract_table(self, _settings):
                return table

        class _FakePDF:
            pages = [_SpecPage()]

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        monkeypatch.setattr(parsers.pdfplumber, "open", lambda _path: _FakePDF())

        items = PDFSpecParser("eom_spec.pdf").extract_items()

        assert [item.raw_title for item in items] == ["Щиток учета"]
        assert items[0].raw_unit == "компл."
        assert items[0].volume == Decimal("1")

    def test_spec_marker_is_recognized_in_mojibake_text(self):
        """Маркер находится и в битой кодировке текстового слоя PDF."""
        mojibake = "спецификация оборудования".encode("cp1251").decode("cp866")

        class _MojibakePage:
            def extract_text(self):
                return mojibake

        assert PDFSpecParser._page_has_spec_markers(_MojibakePage()) is True

    def test_prescreen_keeps_spec_tables_and_vor_continuation(self, monkeypatch):
        """Предфильтр PDFium отсекает чертежи, но не продолжения ВОР.

        Штамп чертежа содержит «Спецификация оборудования» и марку «ИОС1.С»,
        поэтому лист принимается только при наличии строк данных либо шапки
        таблицы. Продолжения ВОР идут без шапки и удерживаются, пока на
        странице есть строки «ед. изм. + объем».
        """
        texts = {
            1: "Спецификация оборудования 2025-00224-01-0-ИОС1.С лист 2025",
            2: "Наименование и техническая характеристика Код оборудования",
            3: "ИОС1.С Кабель ВВГнг 3х2,5 м 120",
            4: (
                "Ведомость объемов работ Наименование вида работ "
                "Обоснование объема"
            ),
            5: "Установка лотка м 25",
            6: "План на отм. 0,000",
            7: "Прокладка кабеля м 12",
        }
        monkeypatch.setattr(
            PDFSpecParser, "_scan_page_texts", lambda self: texts
        )

        parser = PDFSpecParser("eom.pdf")

        assert parser._prescreen_pages() == {2, 3, 4, 5}
        assert parser._fast_page_texts == texts

    def test_prescreen_excludes_dialux_lighting_report(self, monkeypatch):
        """Расчет освещенности Dialux отсекается по заголовку таблицы."""
        texts = {
            1: (
                "Спецификация оборудования ИОС1.С "
                "Перпендикулярная освещенность 120 лк"
            ),
            2: "Спецификация оборудования ИОС1.С Щит групповой комп. 1",
        }
        monkeypatch.setattr(
            PDFSpecParser, "_scan_page_texts", lambda self: texts
        )

        parser = PDFSpecParser("eom.pdf")

        assert parser._prescreen_pages() == {2}

    def test_prescreen_keeps_raster_spec_sheet_by_image(self, monkeypatch):
        """Лист с растровой таблицей принимается, пустой лист Dialux — нет."""
        texts = {
            5: "Спецификация оборудования формат А3 2025-00224-01-0-ИОС1.С",
            6: "Спецификация оборудования формат А3 2025-00224-01-0-ИОС1.С",
        }
        monkeypatch.setattr(
            PDFSpecParser, "_scan_page_texts", lambda self: texts
        )
        monkeypatch.setattr(
            PDFSpecParser,
            "_pages_with_raster_images",
            lambda self, pages: pages & {5},
        )

        parser = PDFSpecParser("eom.pdf")

        assert parser._prescreen_pages() == {5}

    def test_prescreen_falls_back_when_scan_unavailable(self, monkeypatch):
        """При сбое PDFium возвращается None и тексты не кэшируются."""
        monkeypatch.setattr(
            PDFSpecParser, "_scan_page_texts", lambda self: None
        )

        parser = PDFSpecParser("broken.pdf")

        assert parser._prescreen_pages() is None
        assert parser._fast_page_texts is None

    def test_table_row_pattern_ignores_stamp_year(self):
        """«лист 2025» в штампе не считается строкой данных спецификации."""
        text = PDFSpecParser._marker_text(
            "Спецификация оборудования лист 2025"
        )

        assert PDFSpecParser._text_has_table_rows(text) is False


# ==============================================================================
# 3. ИНТЕГРАЦИОННЫЕ ТЕСТЫ НА РЕАЛЬНЫХ ДАННЫХ ПРОЕКТА
# ==============================================================================

LSR_FILE = "data/сот.xlsx"
SPEC_FILE = "data/сот2.pdf"

@pytest.mark.skipif(not os.path.exists(LSR_FILE), reason="Файл сметы сот.xlsx не найден")
def test_real_lsr_file_clean_extraction():
    """Интеграционный тест: реальная смета ТСН-2001 содержит ровно 44 чистые позиции."""
    parser = LSRParser(LSR_FILE)
    items = parser.parse()

    assert len(items) == 44

    # Проверка ключевых позиций
    camera_work = next(it for it in items if it.code == "4.10-174-2")
    assert camera_work.volume == Decimal("5.0")
    assert not camera_work.is_resource

    cable_item = next(it for it in items if "витая пара" in it.clean_title)
    assert cable_item.volume == Decimal("1326")
    assert cable_item.canonical_unit == UnitClass.LENGTH


@pytest.mark.skipif(not os.path.exists(SPEC_FILE), reason="Файл спецификации сот2.pdf не найден")
def test_real_pdf_spec_file_clean_extraction():
    """Интеграционный тест: реальная спецификация ГОСТ 21.110 содержит 25 позиций.

    Две вложенные позиции комплекта защиты Ethernet (модуль и корпус)
    учитываются отдельно, а не склеиваются с родительской строкой.
    """
    parser = PDFSpecParser(SPEC_FILE, debug=False)
    items = parser.extract_items()

    assert len(items) == 25
    assert any("модуль защиты" in item.clean_title for item in items)

    # Проверка отсутствия фантомных строк с номерами колонок таблицы
    assert not any(it.raw_title.strip() == "2" for it in items)
    assert not any("оборудование:" in it.clean_title and len(it.clean_title) < 18 for it in items)

    # Проверка извлечения ключевых проектных позиций
    pipe_item = next(it for it in items if "гофротруба" in it.clean_title)
    assert pipe_item.volume == Decimal("1000")
    assert pipe_item.canonical_unit == UnitClass.LENGTH


# ==============================================================================
# 4. СПРИНТ 2: РЕАЛЬНЫЕ СТОИМОСТИ И ИЕРАРХИЯ «РАБОТА + РЕСУРС» (ТСН-2001)
# ==============================================================================

REAL_LSR_FILE = "data/all/Сметная_документация.xlsx"
REAL_LSR_SHEET = "ЛС 02-01-14"


@pytest.mark.skipif(
    not os.path.exists(REAL_LSR_FILE),
    reason="Файл Сметная_документация.xlsx не найден",
)
def test_lsr_parser_extracts_real_costs():
    """Блок «ВСЕГО по позиции:» (графа 11) дает total_cost и unit_price > 0."""
    items = LSRParser(REAL_LSR_FILE).parse(REAL_LSR_SHEET)
    assert len(items) == 44

    priced = [it for it in items if it.total_cost is not None and it.total_cost > 0]
    assert len(priced) == 40
    for item in priced:
        assert item.total_cost > 0
        assert item.unit_price is not None
        assert item.unit_price > 0

    # Расценка монтажа: цена в заголовке отсутствует, поэтому
    # unit_price рассчитывается как total_cost / quantity.
    work = next(it for it in items if it.code == "4.10-174-2")
    assert work.total_cost == Decimal("32477.06")
    assert work.unit_price == (Decimal("32477.06") / work.volume).quantize(
        Decimal("0.0001")
    )

    # Ресурс «Прайс-лист»: собственный итог блока и цена из графы 6.
    resource = next(it for it in items if it.raw_id == "LSR_37")
    assert resource.total_cost == Decimal("46619.82")
    assert resource.unit_price == Decimal("1236.6")


@pytest.mark.skipif(
    not os.path.exists(REAL_LSR_FILE),
    reason="Файл Сметная_документация.xlsx не найден",
)
def test_lsr_parser_builds_work_resource_hierarchy():
    """Материалы и оборудование привязаны parent_id к расценке монтажа."""
    items = LSRParser(REAL_LSR_FILE).parse(REAL_LSR_SHEET)

    work_items = [
        it for it in items
        if it.code and it.code.startswith(("3.", "4.", "6."))
    ]
    assert work_items
    assert all(not it.is_resource for it in work_items)
    assert all(it.parent_id is None for it in work_items)

    # Дочерние ресурсы первой расценки монтажа видеокамеры.
    work = next(it for it in items if it.code == "4.10-174-2")
    children = [it for it in items if it.parent_id == work.raw_id]
    assert children
    assert all(it.is_resource for it in children)
    assert any("видеокамера" in it.clean_title for it in children)
    assert all(it.pointer.row_idx > work.pointer.row_idx for it in children)

    # Новая расценка сбрасывает родителя: последующий «Прайс-лист»
    # привязан к ней, а не к предыдущей работе.
    next_work = next(it for it in items if it.code == "4.10-136-1")
    next_children = [it for it in items if it.parent_id == next_work.raw_id]
    assert any(it.code == "Прайс-лист" for it in next_children)
    assert all(it.pointer.row_idx > next_work.pointer.row_idx for it in next_children)

    # Ресурс не может быть родителем другого ресурса.
    resource_ids = {it.raw_id for it in items if it.is_resource}
    assert all(it.parent_id not in resource_ids for it in items)