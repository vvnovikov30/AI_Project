import os
from pathlib import Path
import re
from decimal import Decimal
import openpyxl
import pytest

from models import UnitClass
import parsers
from parsers import (
    ExcelVORParser,
    LSRParser,
    PDFLSRParser,
    PDFSpecParser,
    get_spec_parser,
)


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

    def test_lsr_parser_detects_421pr_format(self, tmp_path):
        """421/пр: графы «Обоснование | Наименование | Ед. изм. | Количество».

        Колонки единицы измерения и количества сдвинуты из-за объединенной
        шапки наименования — парсер обязан найти их по заголовкам, а не по
        фиксированным A..E. Отрицательный объем исключения сохраняется.
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Смета"

        ws.cell(row=1, column=1, value="ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ № ЛСР-02-01-01")
        ws.cell(row=2, column=1, value="№ п/п")
        ws.cell(row=2, column=2, value="Обоснование")
        ws.cell(row=2, column=3, value="Наименование работ и затрат")
        ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=5)
        ws.cell(row=2, column=6, value="Единица измерения")
        ws.cell(row=2, column=7, value="Количество")

        # Строка нумерации граф: 1 | 2 | 3 | 4 | 5
        for column, value in ((1, 1), (2, 2), (3, 3), (6, 4), (7, 5)):
            ws.cell(row=3, column=column, value=value)

        ws.cell(row=4, column=1, value="Раздел 1. Монтаж")

        # Позиция 1: расценка ТЕР, единица «100 м ...», дробный объем.
        ws.cell(row=5, column=1, value="1")
        ws.cell(row=5, column=2, value="ТЕР10-01-060-01")
        ws.cell(row=5, column=3, value="Установка и крепление наличников")
        ws.cell(row=5, column=6, value="100 м коробок блоков")
        ws.cell(row=5, column=7, value="1,314")

        # Ресурс без номера позиции — дочерняя строка расценки.
        ws.cell(row=6, column=2, value="101-1805")
        ws.cell(row=6, column=3, value="Гвозди строительные")
        ws.cell(row=6, column=6, value="т")
        ws.cell(row=6, column=7, value="0,00071")

        # Позиция 2: отрицательный объем исключаемых работ.
        ws.cell(row=7, column=1, value="2")
        ws.cell(row=7, column=2, value="ТЕР15-01-047-15")
        ws.cell(row=7, column=3, value="Демонтаж подвесных потолков")
        ws.cell(row=7, column=6, value="м2")
        ws.cell(row=7, column=7, value="-840,9744")

        file_path = tmp_path / "lsr_421.xlsx"
        wb.save(file_path)

        # Лист не начинается с «ЛС » — должен сработать fallback на активный.
        items = LSRParser(str(file_path)).parse()

        assert len(items) == 3
        work = items[0]
        assert work.code == "ТЕР10-01-060-01"
        assert work.raw_unit == "100 м коробок блоков"
        assert work.volume == Decimal("1.314")
        assert not work.is_resource
        assert work.parent_id is None
        assert work.context_section == "Раздел 1. Монтаж"
        assert work.pointer.sheet_name == "Смета"
        assert work.pointer.col_letter == "G"

        resource = items[1]
        assert resource.code == "101-1805"
        assert resource.is_resource
        assert resource.parent_id == work.raw_id
        assert resource.canonical_unit == UnitClass.MASS
        assert resource.volume == Decimal("0.71")

        excluded = items[2]
        assert excluded.code == "ТЕР15-01-047-15"
        assert excluded.volume == Decimal("-840.9744")
        assert excluded.canonical_unit == UnitClass.AREA

    def test_detect_4b_layout_reduced_without_unit(self, tmp_path):
        """Корни заголовков находят усеченную 4б при произвольном порядке граф."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value="Наименование работ и затрат")
        ws.cell(row=1, column=2, value="Обоснование расценки")
        ws.cell(row=1, column=3, value="Кол-во")
        ws.cell(row=1, column=4, value="Общая стоимость, руб.")
        ws.cell(row=2, column=1, value="Демонтаж светильников, 100 шт")
        ws.cell(row=2, column=3, value=-5)

        file_path = tmp_path / "lsr_reduced_layout.xlsx"
        wb.save(file_path)

        parser = LSRParser(str(file_path))
        merge_map = parser._resolve_merged_cells_map(parser.wb.active)
        layout = parser._detect_4b_layout(parser.wb.active, merge_map)

        assert layout is not None
        assert layout.is_reduced
        assert layout.name_col == 1
        assert layout.code_col == 2
        assert layout.qty_col == 3
        assert layout.total_cost_col == 4
        assert layout.unit_col is None

    def test_reduced_4b_parsing_without_unit_column(self, tmp_path):
        """Усеченная 4б: единица из текста, отрицательные объемы и стоимость."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Смета"
        ws.cell(row=1, column=1, value="№ п/п")
        ws.cell(row=1, column=2, value="Обоснование расценки")
        ws.cell(row=1, column=3, value="Наименование работ и затрат")
        ws.cell(row=1, column=4, value="Кол-во")
        ws.cell(row=1, column=5, value="Общая стоимость, руб.")

        ws.cell(row=2, column=1, value="Раздел 1. Демонтаж")
        ws.cell(row=3, column=1, value="1")
        ws.cell(row=3, column=2, value="ТЕР10-01-060-01")
        ws.cell(row=3, column=3, value="Демонтаж светильников, 100 шт")
        ws.cell(row=3, column=4, value=-5)
        ws.cell(row=3, column=5, value=-1000)

        ws.cell(row=4, column=2, value="101-1805")
        ws.cell(row=4, column=3, value="Гвозди строительные (кг)")
        ws.cell(row=4, column=4, value="0,5")
        ws.cell(row=4, column=5, value=-50)

        ws.cell(row=5, column=1, value="2")
        ws.cell(row=5, column=2, value="ТЕР15-01-047-15")
        ws.cell(row=5, column=3, value="Прокладка кабеля")
        ws.cell(row=5, column=4, value=10)
        ws.cell(row=5, column=5, value=200)

        ws.cell(row=6, column=3, value="НР от ФОТ")
        ws.cell(row=6, column=4, value=90)

        file_path = tmp_path / "lsr_reduced_4b.xlsx"
        wb.save(file_path)

        items = LSRParser(str(file_path)).parse()

        assert len(items) == 3
        work, resource, no_unit = items
        assert work.code == "ТЕР10-01-060-01"
        assert work.raw_unit == "100 шт"
        assert work.canonical_unit == UnitClass.COUNT
        assert work.volume == Decimal("-500")
        assert work.total_cost == Decimal("-1000")
        assert work.unit_price == Decimal("2.0000")
        assert not work.is_resource
        assert work.context_section == "Раздел 1. Демонтаж"

        assert resource.is_resource
        assert resource.parent_id == work.raw_id
        assert resource.raw_unit == "кг"
        assert resource.canonical_unit == UnitClass.MASS
        assert resource.volume == Decimal("0.5")

        assert no_unit.raw_unit == ""
        assert no_unit.canonical_unit == UnitClass.UNKNOWN
        assert no_unit.volume == Decimal("10")
        assert no_unit.total_cost == Decimal("200")
        assert no_unit.unit_price == Decimal("20.0000")

    @pytest.mark.parametrize(
        ("title", "expected"),
        [
            ("Демонтаж светильников, 100 шт", "100 шт"),
            ("Кабель ВВГнг 3х1.5 (м)", "м"),
            ("Приборы ПС (1 шт.)", "1 шт"),
            ("Прокладка кабеля", None),
        ],
    )
    def test_extract_unit_from_title(self, title, expected):
        """Fallback-извлечение единицы измерения из наименования позиции."""
        assert LSRParser._extract_unit_from_title(title) == expected



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

    def test_vor_5_columns_with_negative_volumes(self, monkeypatch):
        """ВОР СПб: 5 граф, отрицательные объемы исключаемых работ.

        Схема определяется по шапке («Наименование» + «Кол-во»), поэтому
        количество читается из графы 4, а не из фиксированной графы 7.
        Служебные строки подписей не должны становиться позициями.
        """
        table = [
            ["№ п/п", "Наименование", "Ед. изм.", "Кол-во", "Примечание"],
            ["1", "2", "3", "4", "5"],
            ["", "Раздел 1. Демонтаж", "", "", ""],
            ["", "Демонтаж подвесных потолков", "м2", "-840,9744", ""],
            ["", "Монтаж подвесных потолков", "м2", "120,5", ""],
            ["Составил:", "Иванов И.И.", "", "", ""],
            ["Проверил:", "Петров П.П.", "", "", ""],
        ]

        class _FakePage:
            width = 842.0
            height = 595.0

            def extract_text(self):
                return (
                    "ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ\n"
                    "Добавляемые объемы работ\n"
                    "Исключаемые объемы работ"
                )

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

        items = PDFSpecParser("vor_5.pdf").extract_items()

        assert len(items) == 2
        assert [item.volume for item in items] == [
            Decimal("-840.9744"),
            Decimal("120.5"),
        ]
        assert [item.raw_unit for item in items] == ["м2", "м2"]
        assert all(item.canonical_unit == UnitClass.AREA for item in items)
        assert all(
            item.context_section == "Раздел 1. Демонтаж" for item in items
        )
        assert all("Иванов" not in item.raw_title for item in items)
        assert all("Петров" not in item.raw_title for item in items)

    def test_grand_smeta_vor_7_columns_mock(self, monkeypatch):
        """Форма ВОР ПК «ГРАНД-Смета»: рамка только у шапки, объем в графе 4.

        Строки данных идут без горизонтальных линий, поэтому find_tables()
        видит лишь шапку. Границы граф берутся из рамки шапки, слова данных
        раскладываются по графам, а перенос ячейки «Обоснование» не должен
        попадать в наименование или количество.
        """
        header = [
            "№\nп.п.",
            "Наименование работ, ресурсов, затрат по проекту",
            "Ед.\nизм.",
            "Объем работ /\nКоличество",
            "Формула расчета объемов работ",
            "Ссылка на чертежи,\nспецификации",
            "Дополнительная\nинформация",
        ]
        graph_numbers = ["1", "2", "3", "4", "5", "6", "7"]

        class _BBoxRow:
            def __init__(self, top, bottom):
                self.bbox = (13.4, top, 829.0, bottom)

        class _Column:
            def __init__(self, x0, x1):
                self.bbox = (x0, 0.2, x1, 41.7)

        class _GrandVorTable:
            """Размеченная часть формы: шапка и строка нумерации граф."""

            bbox = (13.4, 0.2, 829.0, 41.7)

            @property
            def rows(self):
                return [_BBoxRow(0.2, 30.5), _BBoxRow(30.5, 41.7)]

            @property
            def columns(self):
                edges = [13.4, 33.4, 318.1, 341.0, 395.6, 650.6, 745.6, 829.0]
                return [
                    _Column(left, right)
                    for left, right in zip(edges, edges[1:])
                ]

            def extract(self):
                return [list(header), list(graph_numbers)]

        def word(text, x0, x1, top):
            return {"text": text, "x0": x0, "x1": x1, "top": top}

        words = [
            word("Раздел:", 16.5, 50.0, 45.0),
            word("1.", 52.0, 60.0, 45.0),
            word("Кровля", 62.0, 95.0, 45.0),
            word("1", 21.7, 26.0, 55.0),
            word("Разборка", 36.8, 78.0, 55.0),
            word("кровель", 80.0, 120.0, 55.0),
            word("м2", 320.0, 332.0, 55.0),
            word("15", 345.0, 357.0, 55.0),
            word("15", 400.0, 412.0, 55.0),
            word("Файл:", 653.8, 680.0, 55.0),
            word("03-2026.docx", 653.8, 700.0, 62.8),
            word("2", 21.7, 26.0, 74.0),
            word("Смена:", 36.8, 75.0, 74.0),
            word("колен", 77.0, 110.0, 74.0),
            word("шт", 320.0, 330.0, 74.0),
            word("36", 345.0, 357.0, 74.0),
            word("36", 400.0, 412.0, 74.0),
            word("03-2026.docx", 653.8, 700.0, 81.8),
        ]

        class _FakePage:
            width = 842.0
            height = 595.0

            def within_bbox(self, _bbox):
                return self

            def extract_text(self):
                return (
                    "Ведомость объемов работ № ВОР 02-01-01 "
                    "Наименование работ, ресурсов, затрат по проекту"
                )

            def find_tables(self, _settings):
                return [_GrandVorTable()]

            def extract_table(self, _settings):
                return None

            def extract_words(self):
                return words

        class _FakePDF:
            pages = [_FakePage()]

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        monkeypatch.setattr(parsers.pdfplumber, "open", lambda _path: _FakePDF())

        items = PDFSpecParser("vor_grand.pdf").extract_items()

        assert [item.volume for item in items] == [Decimal("15"), Decimal("36")]
        assert [item.raw_unit for item in items] == ["м2", "шт"]
        assert [item.canonical_unit for item in items] == [
            UnitClass.AREA,
            UnitClass.COUNT,
        ]
        assert [item.context_section for item in items] == [
            "Раздел: 1. Кровля",
            "Раздел: 1. Кровля",
        ]
        assert items[0].raw_title == "Разборка кровель"
        assert items[1].raw_title == "Смена: колен"
        # Перенос графы «Ссылка на чертежи» не приклеивается к наименованию.
        assert all("03-2026.docx" not in item.raw_title for item in items)

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


# ==============================================================================
# 4. МОДУЛЬНЫЕ ТЕСТЫ PDFLSRParser (УНИВЕРСАЛЬНАЯ СЕТКА И ШИФРЫ)
# ==============================================================================


def _pdf_word(
    text: str,
    x0: float,
    x1: float,
    top: float = 100.0,
    bottom: float = 110.0,
) -> dict:
    """Формирует слово страницы в формате ``pdfplumber`` для unit-тестов."""
    return {"text": text, "x0": x0, "x1": x1, "top": top, "bottom": bottom}


class _FakePage:
    """Минимальная заглушка страницы ``pdfplumber`` для тестов геометрии."""

    def __init__(self, edges=None, tables=None):
        self.edges = edges or []
        self._tables = tables or []

    def find_tables(self, table_settings=None):
        """Возвращает заранее заданные таблицы (эмуляция pdfplumber)."""
        del table_settings
        return self._tables


class _FakeTable:
    """Заглушка таблицы ``pdfplumber`` с набором ячеек."""

    def __init__(self, cells):
        self.cells = cells


class TestPDFLSRParserUnit:
    """Проверка универсальных правил разбора PDF-смет формы 4б."""

    @pytest.mark.parametrize(
        ("code", "position", "expected"),
        [
            ("ГЭСН 46-04-008-02", "", True),
            ("ГЭСНр 65-02-008-01", "", True),
            ("ФЕР 15-02-001-01", "", True),
            ("ТЕР 12-01-001-01", "", True),
            ("ТСН 3.1-12-1", "", True),
            ("3.1-12-1", "", True),
            ("6.10-174-2", "", True),
            ("1.2-34-5", "", False),
            ("13.1-2-3", "", False),
            ("Прайс-лист", "", True),
            ("ФСБЦ-08.3.09.02-0025", "6", False),
            ("ФССЦ-21.1.06.04-0012", "", False),
            ("999-9900", "", False),
            ("01.7.15.03-0012", "", False),
            ("08.1.02.22", "2.1", False),
            ("", "6", True),
            ("", "2.1", False),
        ],
    )
    def test_rate_row_detected_by_code(self, code, position, expected):
        """Расценка определяется по шифру, номер позиции — резервное правило."""
        assert PDFLSRParser._is_rate_row(code, position) is expected

    def test_code_normalization_glues_wrapped_cipher(self):
        """Перенос шифра на новую строку не оставляет разрывов у дефисов."""
        assert (
            PDFLSRParser._normalize_code("ГЭСНр 65-02- 008-01")
            == "ГЭСНр 65-02-008-01"
        )
        assert (
            PDFLSRParser._normalize_code("ФСБЦ-24.3.05.16 -0131")
            == "ФСБЦ-24.3.05.16-0131"
        )



    def test_header_line_found_by_graph_numbers(self):
        """Шапка распознается по последовательным номерам граф."""
        lines = [
            [_pdf_word("Смета", 20, 80)],
            [_pdf_word(str(i), 20 + i * 30, 30 + i * 30) for i in range(1, 13)],
            [_pdf_word("1", 20, 30)],
        ]
        assert PDFLSRParser._header_line_index(lines) == 1

    def test_header_line_found_by_titles_without_numbering(self):
        """Шапка без нумерации граф распознается по ключевым заголовкам."""
        lines = [
            [_pdf_word("Смета", 20, 80)],
            [
                _pdf_word("Обоснование", 20, 90),
                _pdf_word("Наименование", 100, 200),
                _pdf_word("Единица измерения", 210, 300),
            ],
            [_pdf_word("ГЭСН 46-04-008-02", 20, 90)],
        ]
        assert PDFLSRParser._header_line_index(lines) == 1

    def test_text_column_bounds_from_body_alignment(self):
        """Без линий таблицы границы колонок строятся по левым краям слов."""
        starts = (20.0, 60.0, 140.0, 260.0, 320.0, 360.0, 400.0, 440.0, 480.0)
        lines = [[_pdf_word("шапка", 20, 480)]]
        for _ in range(3):
            lines.append(
                [
                    _pdf_word(f"c{index}", start, start + 20)
                    for index, start in enumerate(starts)
                ]
            )
        bounds = PDFLSRParser("dummy.pdf")._text_column_bounds(lines, 0)

        assert len(bounds) == len(starts) + 1
        assert list(bounds) == sorted(bounds)
        # Слово из колонки с началом 260 попадает в четвертую графу.
        cells = PDFLSRParser._header_cells([_pdf_word("x", 265, 280)], bounds)
        assert 3 in cells

    def test_column_bounds_prefers_vertical_lines(self):
        """Вертикальные линии таблицы имеют приоритет над текстовой сеткой."""
        edges = [
            {
                "orientation": "v",
                "x0": 20.0 + i * 40,
                "x1": 20.0 + i * 40,
                "top": 50.0,
                "bottom": 700.0,
            }
            for i in range(10)
        ]
        page = _FakePage(edges=edges)
        lines = [[_pdf_word("шапка", 20, 100)], [_pdf_word("тело", 20, 100)]]
        parser = PDFLSRParser("dummy.pdf")

        bounds = parser._column_bounds(page, lines, 0, header_top=100.0)
        assert len(bounds) == 10

    def test_column_bounds_falls_back_to_text_without_lines(self):
        """При отсутствии линий и таблиц используется кластеризация слов."""
        starts = (20.0, 70.0, 150.0, 270.0, 330.0, 370.0, 410.0, 450.0, 490.0)
        lines = [[_pdf_word("шапка", 20, 490)]]
        for _ in range(2):
            lines.append(
                [
                    _pdf_word(f"c{index}", start, start + 20)
                    for index, start in enumerate(starts)
                ]
            )
        parser = PDFLSRParser("dummy.pdf")

        bounds = parser._column_bounds(_FakePage(), lines, 0, header_top=100.0)
        assert len(bounds) == len(starts) + 1
        assert list(bounds) == sorted(bounds)

    def test_column_bounds_uses_pdfplumber_table_cells(self):
        """Ячейки pdfplumber (vertical_strategy='lines') дают сетку колонок."""
        cells = [(20.0 + i * 40, 100.0, 60.0 + i * 40, 110.0) for i in range(9)]
        page = _FakePage(tables=[_FakeTable(cells)])
        parser = PDFLSRParser("dummy.pdf")

        bounds = parser._table_column_bounds(page)
        assert len(bounds) == 10
        assert list(bounds) == sorted(bounds)

# ==============================================================================
# 5. ДИСПЕТЧЕРИЗАЦИЯ СМЕТНЫХ ПАРСЕРОВ И ИНТЕГРАЦИЯ PDF-СМЕТЫ
# ==============================================================================


def test_parse_smeta_items_dispatches_pdf(monkeypatch, tmp_path):
    """Файл .pdf разбирается PDFLSRParser без параметра smeta_sheet."""
    import main as main_module

    calls = {}

    class _StubPDFParser:
        def __init__(self, path):
            calls["path"] = path

        def parse(self):
            calls["parsed"] = True
            return ["pdf_item"]

    monkeypatch.setattr(main_module, "PDFLSRParser", _StubPDFParser)
    result = main_module._parse_smeta_items(str(tmp_path / "smeta.pdf"), None)

    assert result == ["pdf_item"]
    assert calls["path"].endswith("smeta.pdf")
    assert calls["parsed"] is True


def test_parse_smeta_items_rejects_sheet_for_pdf(tmp_path):
    """Для PDF-сметы имя листа XLSX запрещено."""
    import main as main_module

    with pytest.raises(ValueError, match="smeta_sheet"):
        main_module._parse_smeta_items(str(tmp_path / "smeta.PDF"), "Лист1")


def test_parse_smeta_items_dispatches_xlsx(monkeypatch, tmp_path):
    """Файл .xlsx разбирается LSRParser с передачей имени листа."""
    import main as main_module

    seen = {}

    class _StubXLSXParser:
        def __init__(self, path):
            seen["path"] = path

        def parse(self, sheet_name=None):
            seen["sheet"] = sheet_name
            return ["xlsx_item"]

    monkeypatch.setattr(main_module, "LSRParser", _StubXLSXParser)
    path = str(tmp_path / "smeta.xlsx")

    assert main_module._parse_smeta_items(path, "Смета") == ["xlsx_item"]
    assert seen["sheet"] == "Смета"
    assert main_module._parse_smeta_items(path, None) == ["xlsx_item"]
    assert seen["sheet"] is None


REAL_PDF_LSR = "data/ДАГ/ЛСР1.pdf"


@pytest.mark.skipif(
    not os.path.exists(REAL_PDF_LSR),
    reason="Файл ЛСР1.pdf не найден",
)
def test_real_pdf_lsr_extracts_hierarchy_without_hardcoded_geometry():
    """Реальная PDF-смета разбирается без привязки к координатам файла."""
    items = PDFLSRParser(REAL_PDF_LSR).parse()

    assert len(items) > 100
    rates = [it for it in items if not it.is_resource]
    resources = [it for it in items if it.is_resource]
    assert len(rates) > 40
    assert resources

    # Расценки распознаны по шифрам сборников, а не по номеру строки.
    coded_rates = [it for it in rates if it.code]
    assert coded_rates
    assert all(it.code.startswith(("ГЭСН", "ФЕР", "ТЕР", "ТСН")) for it in coded_rates)

    # Ссылочная целостность иерархии «расценка -> ресурсы».
    ids = {it.raw_id for it in items}
    assert all(it.parent_id in ids for it in items if it.parent_id)
    rate_ids = {it.raw_id for it in rates}
    assert all(it.parent_id in rate_ids for it in items if it.parent_id)
    # Вложенные ресурсы всегда идут после своего родителя.
    by_id = {it.raw_id: it for it in items}
    assert all(
        it.pointer.row_idx > by_id[it.parent_id].pointer.row_idx
        for it in items
        if it.parent_id
    )

    # Итоги верхнего уровня заполнены, а их сумма совпадает с итогом сметы.
    top_level = [it for it in items if it.parent_id is None]
    assert all(it.total_cost is not None for it in top_level)
    assert sum(it.total_cost for it in top_level) > 0

    # Секции формы 4б не теряются при переходе между страницами.
    sections = {it.context_section for it in rates}
    assert "Раздел: 1 Кровля" in sections


# ==============================================================================
# 6. ПАРСЕР ДЕФЕКТНЫХ ВЕДОМОСТЕЙ И ВОР ИЗ EXCEL
# ==============================================================================


class TestExcelVORParserSynthetic:
    """Разбор дефектной ведомости XLSX на синтетической таблице."""

    @staticmethod
    def _make_vor(path: Path) -> Path:
        """Формирует таблицу с шапкой не в первой строке и сносками."""
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = "Дефектная ведомость"
        worksheet.append(["УТВЕРЖДАЮ"])
        worksheet.append(["ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ № 02-01"])
        worksheet.append(
            ["№ п/п", "Наименование работ", "Ед. изм.", "Кол-во", "Обоснование"]
        )
        worksheet.append(["Раздел 1. Демонтаж"])
        worksheet.append(
            [1, "Демонтаж светильников", "шт", "1 234,5*", "ГЭСНр67-01-004-05"]
        )
        worksheet.append([2, "Разборка покрытий полов", "м2", 2.52, "ГЭСНр57-01-002-01"])
        worksheet.append([3, "Монтаж кабеля", "100 м", "0,75", "ГЭСН16-04-005-02"])
        worksheet.append([4, "Подзаголовок без объема", "", "", ""])
        worksheet.append([5, "Укладка бетона", "м3", 12, "ГЭСН06-01-001-01"])
        workbook.save(path)
        return path

    def test_excel_vor_parser_extracts_items(self, tmp_path: Path) -> None:
        """Позиции извлекаются в NormalizedItem, объемы строго Decimal."""
        path = self._make_vor(tmp_path / "defect.xlsx")

        items = ExcelVORParser(path).extract_items()

        assert len(items) == 4
        assert [item.raw_id for item in items] == [
            "VOR_EXCEL_1",
            "VOR_EXCEL_2",
            "VOR_EXCEL_3",
            "VOR_EXCEL_4",
        ]
        assert [item.volume for item in items] == [
            Decimal("1234.5"),
            Decimal("2.52"),
            Decimal("75"),
            Decimal("12"),
        ]
        assert all(isinstance(item.volume, Decimal) for item in items)
        assert [item.canonical_unit for item in items] == [
            UnitClass.COUNT,
            UnitClass.AREA,
            UnitClass.LENGTH,
            UnitClass.VOLUME,
        ]
        assert items[2].unit_multiplier == Decimal("100.0")

        assert items[0].raw_title == "Демонтаж светильников"
        assert items[0].clean_title == "демонтаж светильников"
        assert items[0].raw_unit == "шт"
        assert items[0].context_section == "Раздел 1. Демонтаж"
        assert items[0].pointer.doc_type == "VOR_EXCEL"
        assert items[0].pointer.sheet_name == "Дефектная ведомость"
        assert items[0].pointer.row_idx == 5
        assert items[0].pointer.col_letter == "B"
        assert items[0].pointer.coordinate == "D5"


def test_get_spec_parser_selects_parser_by_suffix(tmp_path: Path) -> None:
    """Фабрика отдает ExcelVORParser для .xlsx и PDF-парсер для остальных."""
    xlsx = tmp_path / "vor.xlsx"
    xlsx.write_bytes(b"")

    assert isinstance(get_spec_parser(xlsx), ExcelVORParser)

    class _StubPDFParser:
        def __init__(self, path: str) -> None:
            self.path = path

    parser = get_spec_parser(
        tmp_path / "vor.PDF", pdf_parser_cls=_StubPDFParser
    )

    assert isinstance(parser, _StubPDFParser)
    assert parser.path.endswith("vor.PDF")

