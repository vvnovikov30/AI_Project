"""L2 Unit & Gates: Модульное тестирование логических шлюзов и детекторов коллизий.

Проверяет:
- Pydantic v2 строгие схемы и барьер против типа float.
- UnitNormalizer: соответствие классов ОКЕИ и коэффициентов кратности (100 м2, 1000 шт и т.д.).
- DimensionLockExtractor: извлечение инвариантов DN, PN, диаметров и классов бетона.
- DiscrepancyDetector: формулы погрешностей, отходы ГЭСН и ошибки делителей расценок.
- Неизменность пространственных координат CellPointer.
"""

from decimal import Decimal
import pytest
from pydantic import ValidationError

from conftest import StrictCellPointerSchema, StrictNormalizedItemSchema
from evaluator import DiscrepancyDetector
from models import CellPointer, CollisionSeverity, NormalizedItem, UnitClass
from normalizer import DimensionLockExtractor, UnitNormalizer

# ----------------------------------------------------------------------
# 1. Тестирование Pydantic v2 схем и Zero Float Policy
# ----------------------------------------------------------------------


class TestStrictPydanticSchemas:
    """Проверка контрактной чистоты DTO и блокировки неточных типов данных."""

    def test_rejects_float_volume_instantiation(
        self, base_excel_pointer: StrictCellPointerSchema
    ) -> None:
        """Передача типа float в поле volume обязана вызывать ошибку валидации."""
        with pytest.raises((ValidationError, TypeError)) as exc_info:
            StrictNormalizedItemSchema(
                raw_id="ERR_01",
                code=None,
                raw_title="Щебень гранитный фракции 20-40",
                clean_title="щебень гранитный фракции 20 40",
                raw_unit="м3",
                canonical_unit=UnitClass.VOLUME,
                unit_multiplier=Decimal("1.0"),
                volume=125.50,  # Запрещенный float IEEE 754
                dimensions={},
                context_section="Раздел 1",
                pointer=base_excel_pointer,
                is_resource=False,
                parent_id=None,
            )
        assert "Zero Float Policy Violation" in str(exc_info.value)

    def test_rejects_float_multiplier_instantiation(
        self, base_excel_pointer: StrictCellPointerSchema
    ) -> None:
        """Передача типа float в поле unit_multiplier обязана пресекаться."""
        with pytest.raises((ValidationError, TypeError)) as exc_info:
            StrictNormalizedItemSchema(
                raw_id="ERR_02",
                code=None,
                raw_title="Кровля рулонная",
                clean_title="кровля рулонная",
                raw_unit="100 м2",
                canonical_unit=UnitClass.AREA,
                unit_multiplier=100.0,  # Запрещенный float
                volume=Decimal("15.0"),
                dimensions={},
                context_section="Раздел 2",
                pointer=base_excel_pointer,
                is_resource=False,
                parent_id=None,
            )
        assert "Zero Float Policy Violation" in str(exc_info.value)

    def test_schema_immutability_frozen(
        self, sample_cable_item: StrictNormalizedItemSchema
    ) -> None:
        """DTO являются неизменяемыми (frozen=True)."""
        with pytest.raises(ValidationError):
            # Попытка мутации существующего объекта
            sample_cable_item.volume = Decimal("999.0")  # type: ignore[misc]

    def test_schema_rejects_extra_fields(
        self, base_excel_pointer: StrictCellPointerSchema
    ) -> None:
        """Запрещено передавать недекларированные поля (extra='forbid')."""
        with pytest.raises(ValidationError):
            StrictNormalizedItemSchema(
                raw_id="EXTRA_01",
                code=None,
                raw_title="Труба стальная",
                clean_title="труба стальная",
                raw_unit="м",
                canonical_unit=UnitClass.LENGTH,
                unit_multiplier=Decimal("1.0"),
                volume=Decimal("10.0"),
                dimensions={},
                context_section="Раздел 3",
                pointer=base_excel_pointer,
                is_resource=False,
                parent_id=None,
                unexpected_payload="hacked",  # type: ignore[call-arg]
            )


# ----------------------------------------------------------------------
# 2. Тестирование модуля UnitNormalizer
# ----------------------------------------------------------------------


class TestUnitNormalizer:
    """Проверка канонизации строительных единиц и коэффициентов кратности."""

    @pytest.mark.parametrize(
        ("raw_unit", "expected_class", "expected_mult"),
        [
            ("м2", UnitClass.AREA, Decimal("1.0")),
            ("кв. м", UnitClass.AREA, Decimal("1.0")),
            ("м.кв.", UnitClass.AREA, Decimal("1.0")),
            ("100 м2", UnitClass.AREA, Decimal("100.0")),
            ("1000 м2", UnitClass.AREA, Decimal("1000.0")),
            ("м3", UnitClass.VOLUME, Decimal("1.0")),
            ("куб. м", UnitClass.VOLUME, Decimal("1.0")),
            ("м.куб.", UnitClass.VOLUME, Decimal("1.0")),
            ("10 м3", UnitClass.VOLUME, Decimal("10.0")),
            ("100 м3", UnitClass.VOLUME, Decimal("100.0")),
            ("кг", UnitClass.MASS, Decimal("1.0")),
            ("килограмм", UnitClass.MASS, Decimal("1.0")),
            ("т", UnitClass.MASS, Decimal("1000.0")),
            ("тн", UnitClass.MASS, Decimal("1000.0")),
            ("шт", UnitClass.COUNT, Decimal("1.0")),
            ("штук", UnitClass.COUNT, Decimal("1.0")),
            ("100 шт", UnitClass.COUNT, Decimal("100.0")),
            ("1000 шт", UnitClass.COUNT, Decimal("1000.0")),
            ("тыс. шт", UnitClass.COUNT, Decimal("1000.0")),
        ],
    )
    def test_canonical_unit_and_multipliers(
        self,
        raw_unit: str,
        expected_class: UnitClass,
        expected_mult: Decimal,
    ) -> None:
        """Валидирует правильность маппинга единиц измерения и множителей."""
        base_qty = Decimal("5.500")
        u_class, norm_qty, mult = UnitNormalizer.normalize(raw_unit, base_qty)

        assert u_class == expected_class
        assert mult == expected_mult
        assert norm_qty == base_qty * expected_mult
        assert isinstance(norm_qty, Decimal)
        assert isinstance(mult, Decimal)

    def test_unknown_unit_fallback(self) -> None:
        """Неопознанные единицы получают класс UNKNOWN и множитель 1.0."""
        base_qty = Decimal("42.0")
        u_class, norm_qty, mult = UnitNormalizer.normalize("чел.-час", base_qty)

        assert u_class == UnitClass.UNKNOWN
        assert mult == Decimal("1.0")
        assert norm_qty == Decimal("42.0")


# ----------------------------------------------------------------------
# 3. Тестирование модуля DimensionLockExtractor
# ----------------------------------------------------------------------


class TestDimensionLockExtractor:
    """Проверка извлечения ключевых технических параметров через RegEx."""

    def test_extract_valve_dimensions(self) -> None:
        """Извлечение условного прохода (DN) и давления (PN)."""
        text = "Задвижка стальная клиновая фланцевая Ду 100 Ру 16 для воды"
        dims = DimensionLockExtractor.extract(text)

        assert dims.get("DN") == "100"
        assert dims.get("PN") == "16"

    def test_extract_latin_dn_pn(self) -> None:
        """Извлечение латинских обозначений DN/PN."""
        text = "Ball valve DN50 PN25 stainless steel"
        dims = DimensionLockExtractor.extract(text)

        assert dims.get("DN") == "50"
        assert dims.get("PN") == "25"

    def test_extract_concrete_class(self) -> None:
        """Извлечение марки и класса бетона."""
        text = "Бетонная смесь готовая тяжелая В30 (М400) П4 F200 W8"
        dims = DimensionLockExtractor.extract(text)

        assert dims.get("CONCRETE") == "B30"

    def test_extract_rebar_diameter(self) -> None:
        """Извлечение диаметров стержневой арматуры (знаки ø, ф)."""
        text_1 = "Сетка арматурная из стали класса А500С ø12 мм"
        text_2 = "Каркас плоский продольный ф16"

        dims_1 = DimensionLockExtractor.extract(text_1)
        dims_2 = DimensionLockExtractor.extract(text_2)

        assert dims_1.get("REBAR_DIA") == "12"
        assert dims_2.get("REBAR_DIA") == "16"

    def test_empty_dimensions_when_no_match(self) -> None:
        """Если текст не содержит калибров, возвращается пустой словарь."""
        text = "Кирпич керамический полнотелый одинарный рядовой"
        dims = DimensionLockExtractor.extract(text)
        assert dims == {}


# ----------------------------------------------------------------------
# 4. Тестирование модуля DiscrepancyDetector
# ----------------------------------------------------------------------


class TestDiscrepancyDetector:
    """Проверка математики расхождений, отходов ГЭСН и детектора MULTIPLIER_ERROR."""

    @pytest.fixture
    def mock_pointer(self) -> CellPointer:
        return CellPointer(
            doc_type="MOCK",
            file_name="mock.xlsx",
            sheet_name="Лист 1",
            row_idx=10,
            col_letter="B",
            coordinate="B10",
        )

    def _create_item(
        self,
        raw_id: str,
        title: str,
        unit: UnitClass,
        volume: Decimal,
        pointer: CellPointer,
    ) -> NormalizedItem:
        return NormalizedItem(
            raw_id=raw_id,
            code="MOCK_CODE",
            raw_title=title,
            clean_title=title.lower(),
            raw_unit=unit.value,
            canonical_unit=unit,
            unit_multiplier=Decimal("1.0"),
            volume=volume,
            dimensions={},
            context_section="Default",
            pointer=pointer,
        )

    def test_exact_match_within_tolerance(self, mock_pointer: CellPointer) -> None:
        """Точное совпадение объемов классифицируется как INFO."""
        lsr = self._create_item("L1", "Песок строительный", UnitClass.VOLUME, Decimal("100.000"), mock_pointer)
        spec = self._create_item("S1", "Песок строительный", UnitClass.VOLUME, Decimal("100.000"), mock_pointer)

        record = DiscrepancyDetector.evaluate(lsr, spec, score=98.0)

        assert record.severity == CollisionSeverity.INFO
        assert record.volume_delta == Decimal("0.000")
        assert "соответствуют проектным" in record.description

    def test_rebar_waste_coefficient_application(self, mock_pointer: CellPointer) -> None:
        """Для арматуры нормативный объем сметы допускает превышение на 5% (ГЭСН 06)."""
        spec_vol = Decimal("10.000")  # 10 тонн в проекте
        lsr_vol = Decimal("10.500")   # 10.5 тонн в смете (+5% отход)

        lsr = self._create_item("L_REB", "Арматурная сталь гладкая", UnitClass.MASS, lsr_vol, mock_pointer)
        spec = self._create_item("S_REB", "Арматурная сталь гладкая", UnitClass.MASS, spec_vol, mock_pointer)

        record = DiscrepancyDetector.evaluate(lsr, spec, score=95.0)

        # С учетом 5% расхождение нулевое
        assert record.severity == CollisionSeverity.INFO
        assert record.volume_delta == Decimal("0.000")

    def test_critical_volume_overestimation(self, mock_pointer: CellPointer) -> None:
        """Превышение сметного объема сверх норматива и допусков классифицируется как CRITICAL."""
        lsr = self._create_item("L2", "Кирпич силикатный", UnitClass.COUNT, Decimal("15000"), mock_pointer)
        spec = self._create_item("S2", "Кирпич силикатный", UnitClass.COUNT, Decimal("10000"), mock_pointer)

        record = DiscrepancyDetector.evaluate(lsr, spec, score=90.0)

        assert record.severity == CollisionSeverity.CRITICAL
        assert record.volume_delta == Decimal("5000")
        assert "Завышение объема" in record.description

    def test_suspect_volume_underestimation(self, mock_pointer: CellPointer) -> None:
        """Занижение объема в смете классифицируется как SUSPECT (риск недофинансирования)."""
        lsr = self._create_item("L3", "Бетон тяжелый", UnitClass.VOLUME, Decimal("80.000"), mock_pointer)
        spec = self._create_item("S3", "Бетон тяжелый", UnitClass.VOLUME, Decimal("100.000"), mock_pointer)

        record = DiscrepancyDetector.evaluate(lsr, spec, score=92.0)

        assert record.severity == CollisionSeverity.SUSPECT
        assert record.volume_delta < Decimal("0.0")
        assert "Занижение объема" in record.description

    @pytest.mark.parametrize(
        ("spec_qty", "lsr_qty", "expected_multiplier_desc"),
        [
            (Decimal("10.0"), Decimal("100.0"), "10"),
            (Decimal("5.0"), Decimal("500.0"), "100"),
            (Decimal("2.0"), Decimal("2000.0"), "1000"),
        ],
    )
    def test_multiplier_error_detection(
        self,
        mock_pointer: CellPointer,
        spec_qty: Decimal,
        lsr_qty: Decimal,
        expected_multiplier_desc: str,
    ) -> None:
        """Проверяет детектор забытого делителя расценки (ошибки в 10, 100, 1000 раз)."""
        lsr = self._create_item("L_MULT", "Укладка кабеля", UnitClass.LENGTH, lsr_qty, mock_pointer)
        spec = self._create_item("S_MULT", "Укладка кабеля", UnitClass.LENGTH, spec_qty, mock_pointer)

        record = DiscrepancyDetector.evaluate(lsr, spec, score=95.0)

        assert record.severity == CollisionSeverity.CRITICAL
        assert "MULTIPLIER_ERROR" in record.description
        assert expected_multiplier_desc in record.description


# ----------------------------------------------------------------------
# 5. Тестирование неизменности аудиторского следа (CellPointer)
# ----------------------------------------------------------------------


class TestAuditTrailPreservation:
    """Гарантирует, что пространственные координаты ячеек не теряются при сверке."""

    def test_cell_pointer_is_unaltered_after_evaluation(
        self,
        base_excel_pointer: StrictCellPointerSchema,
        base_pdf_pointer: StrictCellPointerSchema,
    ) -> None:
        """Проверяет сохранение имени файла, номеров строк и точных координат ячеек."""
        pointer_lsr = CellPointer(
            doc_type=base_excel_pointer.doc_type,
            file_name=base_excel_pointer.file_name,
            sheet_name=base_excel_pointer.sheet_name,
            row_idx=base_excel_pointer.row_idx,
            col_letter=base_excel_pointer.col_letter,
            coordinate=base_excel_pointer.coordinate,
        )
        pointer_spec = CellPointer(
            doc_type=base_pdf_pointer.doc_type,
            file_name=base_pdf_pointer.file_name,
            sheet_name=base_pdf_pointer.sheet_name,
            row_idx=base_pdf_pointer.row_idx,
            col_letter=base_pdf_pointer.col_letter,
            coordinate=base_pdf_pointer.coordinate,
        )

        lsr = NormalizedItem(
            raw_id="L_TRAIL",
            code="ФЕР-01",
            raw_title="Разработка грунта",
            clean_title="разработка грунта",
            raw_unit="100 м3",
            canonical_unit=UnitClass.VOLUME,
            unit_multiplier=Decimal("100.0"),
            volume=Decimal("500.0"),
            dimensions={},
            context_section="Земляные работы",
            pointer=pointer_lsr,
        )
        spec = NormalizedItem(
            raw_id="S_TRAIL",
            code="ВОР-01",
            raw_title="Разработка грунта",
            clean_title="разработка грунта",
            raw_unit="м3",
            canonical_unit=UnitClass.VOLUME,
            unit_multiplier=Decimal("1.0"),
            volume=Decimal("500.0"),
            dimensions={},
            context_section="Земляные работы",
            pointer=pointer_spec,
        )

        record = DiscrepancyDetector.evaluate(lsr, spec, score=96.0)

        assert record.lsr_item is not None
        assert record.spec_item is not None
        assert record.lsr_item.pointer.coordinate == "C24"
        assert record.lsr_item.pointer.file_name == "LSR_01_Foundation.xlsx"
        assert record.spec_item.pointer.coordinate == "p.12 [row 15]"
        assert record.spec_item.pointer.file_name == "02_2026_KJ_Spec.pdf"