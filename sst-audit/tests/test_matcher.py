import pytest
from decimal import Decimal
from models import NormalizedItem, CellPointer, UnitClass
from matcher import CascadeMatcher
from normalizer import UnitNormalizer


def make_item(raw_id="1", title="Камера", code=None, section="Default_Section", unit=UnitClass.COUNT, dims=None):
    return NormalizedItem(
        raw_id=raw_id,
        code=code,
        raw_title=title,
        clean_title=title.lower(),
        raw_unit="шт",
        canonical_unit=unit,
        unit_multiplier=Decimal("1.0"),
        volume=Decimal("1"),
        dimensions=dims or {},
        context_section=section,
        pointer=CellPointer("LSR", "test.xlsx", "Sheet1", 1, "A", "A1")
    )


class TestCascadeMatcherAdaptive:
    """Полное покрытие всех ветвей CascadeMatcher: коды, контекст, адаптивные веса."""

    def test_clean_code_filters(self):
        """Проверка фильтрации псевдо-кодов ценообразования."""
        assert CascadeMatcher._clean_code(None) is None
        assert CascadeMatcher._clean_code("") is None
        assert CascadeMatcher._clean_code("   ") is None
        assert CascadeMatcher._clean_code("прайс-лист") is None
        assert CascadeMatcher._clean_code("Прайс Лист") is None
        assert CascadeMatcher._clean_code("б/н") is None
        assert CascadeMatcher._clean_code("—") is None
        assert CascadeMatcher._clean_code("-") is None
        assert CascadeMatcher._clean_code("  RVi-1NCT4053  ") == "RVi-1NCT4053"

    def test_unit_gating_and_unknown(self):
        """Блокировка при несовпадении единиц или при статусе UNKNOWN."""
        item_count = make_item(unit=UnitClass.COUNT)
        item_length = make_item(unit=UnitClass.LENGTH)
        item_unknown = make_item(unit=UnitClass.UNKNOWN)

        assert CascadeMatcher.calculate_score(item_count, item_length) == 0.0
        assert CascadeMatcher.calculate_score(item_count, item_unknown) == 0.0

    def test_dimension_lock_rejection(self):
        """Блокировка при конфликте геометрических параметров."""
        item_dn50 = make_item(dims={"DN": "50"})
        item_dn100 = make_item(dims={"DN": "100"})
        item_pn16 = make_item(dims={"PN": "16"})

        assert CascadeMatcher.calculate_score(item_dn50, item_dn100) == 0.0
        assert CascadeMatcher.calculate_score(item_dn100, item_dn50) == 0.0
        # Непересекающиеся параметры совместимы
        assert CascadeMatcher.calculate_score(item_dn50, item_pn16) > 0.0

    def test_code_exact_match(self):
        """Точное совпадение артикулов (s_code = 100)."""
        it1 = make_item(title="Видеокамера IP", code="RVi-1NCT4053")
        it2 = make_item(title="Видеокамера IP", code="rvi-1nct4053")
        score = CascadeMatcher.calculate_score(it1, it2)
        assert score == 100.0

    def test_code_in_spec_title(self):
        """Артикул сметы найден в названии проекта (s_code = 80)."""
        it_lsr = make_item(title="Видеокамера", code="RVi-1NCT4053")
        it_spec = make_item(title="Видеокамера уличная rvi-1nct4053 4Мп", code="Камера-01")
        score = CascadeMatcher.calculate_score(it_lsr, it_spec)
        assert score >= 80.0

    def test_code_in_lsr_title(self):
        """Артикул проекта найден в названии сметы (s_code = 80)."""
        it_lsr = make_item(title="Коммутатор nst ns-lp-1gp/d сетевой", code="Поз. 5")
        it_spec = make_item(title="Коммутатор сетевой", code="NST NS-LP-1GP/D")
        score = CascadeMatcher.calculate_score(it_lsr, it_spec)
        assert score >= 80.0

    def test_code_conflicting_penalty(self):
        """Штраф при прямом конфликте артикулов (s_code = 0)."""
        it1 = make_item(title="Видеокамера", code="DS-2CD2043")
        it2 = make_item(title="Видеокамера", code="RVi-1NCT4053")
        score = CascadeMatcher.calculate_score(it1, it2)
        # s_code=0, вес переходит с дисконтом: 0.45*0 + 0.55*100 = 55.0
        assert score == pytest.approx(55.0, 0.1)

    def test_only_lsr_code_in_spec_title(self):
        """Только у сметы есть код, и он присутствует в названии проекта."""
        it_lsr = make_item(title="Камера", code="RVI-4053")
        it_spec = make_item(title="Камера rvi-4053 уличная", code=None)
        score = CascadeMatcher.calculate_score(it_lsr, it_spec)
        assert score >= 80.0

    def test_only_spec_code_in_lsr_title(self):
        """Только у проекта есть код, и он присутствует в названии сметы."""
        it_lsr = make_item(title="Камера rvi-4053 уличная", code="Прайс-лист")
        it_spec = make_item(title="Камера", code="RVI-4053")
        score = CascadeMatcher.calculate_score(it_lsr, it_spec)
        assert score >= 80.0

    def test_weight_branch_code_and_context(self):
        """Ветка: has_code and has_context (W_CODE*s_code + W_FUZZY*s_fuzzy + W_CONTEXT*s_context)."""
        it1 = make_item(title="Камера", code="SKU-1", section="Раздел: Оборудование")
        it2 = make_item(title="Камера", code="SKU-1", section="Раздел: Оборудование")
        score = CascadeMatcher.calculate_score(it1, it2)
        assert score == 100.0

    def test_weight_branch_code_no_context(self):
        """Ветка: has_code and not has_context (0.45*s_code + 0.55*s_fuzzy)."""
        it1 = make_item(title="Камера", code="SKU-1", section="Default_Section")
        it2 = make_item(title="Камера", code="SKU-1", section="Default_Section")
        score = CascadeMatcher.calculate_score(it1, it2)
        assert score == 100.0

    def test_weight_branch_no_code_with_context(self):
        """Ветка: not has_code and has_context (0.85*s_fuzzy + 0.15*s_context)."""
        it1 = make_item(title="Камера", code=None, section="Раздел: Оборудование")
        it2 = make_item(title="Камера", code="Прайс-лист", section="Раздел: Оборудование")
        score = CascadeMatcher.calculate_score(it1, it2)
        assert score == 100.0

    def test_weight_branch_no_code_no_context(self):
        """Ветка: not has_code and not has_context (100% s_fuzzy)."""
        it1 = make_item(title="Устройство грозозащиты", code=None, section="Default_Section")
        it2 = make_item(title="Устройство грозозащиты", code="Прайс-лист", section="Default_Section")
        score = CascadeMatcher.calculate_score(it1, it2)
        assert score == 100.0


class TestNormalizerEdgeCases:
    """Закрытие граничных веток в normalizer.py до 100% покрытия."""

    def test_empty_raw_unit(self):
        cls, qty, mult = UnitNormalizer.normalize("", Decimal("10"))
        assert cls == UnitClass.UNKNOWN
        assert mult == Decimal("1.0")

    def test_length_kilometer(self):
        cls, qty, mult = UnitNormalizer.normalize("км", Decimal("2"))
        assert cls == UnitClass.LENGTH
        assert qty == Decimal("2000.0")
        assert mult == Decimal("1000.0")

    def test_count_thousand(self):
        cls, qty, mult = UnitNormalizer.normalize("тыс. шт", Decimal("3"))
        assert cls == UnitClass.COUNT
        assert qty == Decimal("3000.0")
        assert mult == Decimal("1000.0")