import pytest
from decimal import Decimal
from models import NormalizedItem, CellPointer, UnitClass
from matcher import CascadeMatcher, CompositeMatcher
from normalizer import UnitNormalizer


def _node(
    raw_id: str,
    title: str,
    unit: UnitClass,
    volume: str,
    raw_unit: str,
    code=None,
    is_resource: bool = False,
    parent_id=None,
) -> NormalizedItem:
    """Сметный узел/позиция РД с полным контролем признаков ресурса."""
    return NormalizedItem(
        raw_id=raw_id,
        code=code,
        raw_title=title,
        clean_title=title.lower(),
        raw_unit=raw_unit,
        canonical_unit=unit,
        unit_multiplier=Decimal("1.0"),
        volume=Decimal(volume),
        dimensions={},
        context_section="Default_Section",
        pointer=CellPointer("LSR", "test.xlsx", "Sheet1", 1, "A", "A1"),
        is_resource=is_resource,
        parent_id=parent_id,
    )


class TestCompositeNodeMatching:
    """Композитные узлы сметы: работа-родитель + дочерний ресурс."""

    @staticmethod
    def _parent_work() -> NormalizedItem:
        return _node(
            "LSR_1",
            "Монтаж строительных конструкций",
            UnitClass.COUNT,
            "4",
            "шт",
        )

    @staticmethod
    def _child_resource() -> NormalizedItem:
        return _node(
            "LSR_2",
            "Камера видеонаблюдения",
            UnitClass.COUNT,
            "4",
            "шт",
            code="RVi-1NCT4053",
            is_resource=True,
            parent_id="LSR_1",
        )

    def test_composite_node_equipment_matches_child_resource(self):
        """Дочерний ресурс наследует родителя и покрывает его расценку."""
        rd_item = _node(
            "SPEC_1",
            "Камера видеонаблюдения",
            UnitClass.COUNT,
            "4",
            "шт",
            code="RVI-1NCT4053",
        )
        records = CompositeMatcher.match_composite(
            [rd_item], [self._parent_work(), self._child_resource()]
        )

        child_records = [
            record
            for record in records
            if record.smeta_item.raw_id == "LSR_2"
        ]
        assert len(child_records) == 1
        child_record = child_records[0]
        assert child_record.parent_work_id == "LSR_1"
        assert child_record.is_resource_covered is True
        assert child_record.status == "AUTO"
        assert child_record.is_auto_confirmed is True
        assert child_record.confidence > Decimal("0.7")
        assert child_record.confidence.as_tuple().exponent == -4
        assert "LSR_1" in CompositeMatcher.covered_parent_ids(records)

    def test_unit_gate_mismatch_blocks_auto_match(self):
        """R05_UNT: несовместимые размерности блокируют автоподтверждение."""
        rd_item = _node(
            "SPEC_1",
            "Камера видеонаблюдения",
            UnitClass.AREA,
            "4",
            "м2",
        )
        smeta_item = _node(
            "LSR_2",
            "Камера видеонаблюдения",
            UnitClass.VOLUME,
            "4",
            "м3",
            is_resource=True,
            parent_id="LSR_1",
        )

        record = CompositeMatcher.match_pair(rd_item, smeta_item)

        assert record.is_unit_blocked is True
        assert record.status == "BLOCKED"
        assert record.is_auto_confirmed is False
        assert record.confidence == Decimal("0.0000")
        unit_gate = record.breakdown.get("R05_UNT")
        assert unit_gate is not None
        assert unit_gate.status == "FAILED"
        assert unit_gate.score == Decimal("0")
        assert "INCOMPATIBLE_DIMENSIONS" in unit_gate.message


    def test_volume_deficit_blocks_composite_match(self):
        """R_CMP_VOLUME_DEFICIT: ресурс больше работы-родителя блокирует связь."""
        parent_work = _node(
            "LSR_1",
            "Монтаж строительных конструкций",
            UnitClass.COUNT,
            "2",
            "шт",
        )
        child_resource = self._child_resource()
        rd_item = _node(
            "SPEC_1",
            "Камера видеонаблюдения",
            UnitClass.COUNT,
            "4",
            "шт",
            code="RVI-1NCT4053",
        )

        record = CompositeMatcher.match_pair(rd_item, child_resource, parent_work)

        volume_gate = next(
            evaluation
            for evaluation in record.evaluations
            if evaluation.rule_code == "R_CMP_VOLUME_DEFICIT"
        )
        assert volume_gate.status == "FAILED"
        assert volume_gate.score == Decimal("0")
        assert record.status == "BLOCKED"
        assert record.is_auto_confirmed is False
        assert record.parent_work_id == "LSR_1"


class TestRuleBreakdownConfidence:
    """Композитная уверенность R01-R05 строго на Decimal."""

    def test_rule_breakdown_confidence_strict_decimal(self):
        """Веса в сумме 1.0000, итог квантуется до 4 знаков без float."""
        from evaluator import PredicateEvaluator
        from models import RuleBreakdown

        rd_item = _node(
            "SPEC_1",
            "Камера видеонаблюдения",
            UnitClass.COUNT,
            "4",
            "шт",
        )
        smeta_item = _node(
            "LSR_2",
            "Камера видеонаблюдения",
            UnitClass.COUNT,
            "4",
            "шт",
            is_resource=True,
            parent_id="LSR_1",
        )

        breakdown = PredicateEvaluator.evaluate_pair(rd_item, smeta_item)

        assert breakdown.weight_sum == Decimal("1.0000")
        assert breakdown.calculate_total_confidence() == Decimal("1.0000")
        assert breakdown.total_confidence == Decimal("1.0000")
        assert breakdown.total_confidence.as_tuple().exponent == -4
        assert breakdown.is_valid is True
        assert breakdown.is_blocked is False
        assert len(breakdown.rule_evaluations) == 5
        for evaluation in breakdown.rule_evaluations:
            assert isinstance(evaluation.weight, Decimal)
            assert isinstance(evaluation.score, Decimal)
            assert isinstance(evaluation.score_weight, Decimal)

        # Ручная проверка взвешенной суммы без float-арифметики.
        expected = (
            RuleBreakdown.R01_CAT * Decimal("1")
            + RuleBreakdown.R02_DIM * Decimal("1")
            + RuleBreakdown.R03_STD * Decimal("1")
            + RuleBreakdown.R04_MAT * Decimal("1")
            + RuleBreakdown.R05_UNT * Decimal("1")
        )
        assert breakdown.total_confidence == expected


class TestParentCoverage:
    """Родительская расценка не должна попадать в SMETA_NOT_IN_RD."""

    def test_parent_work_not_flagged_as_missing_when_child_matched(self):
        """Покрытый ресурсом родитель фиксируется как SMETA_PARENT_COVERED."""
        from main import MatchDecision, build_rule_evaluations

        parent_work = _node(
            "LSR_1",
            "Монтаж строительных конструкций",
            UnitClass.COUNT,
            "4",
            "шт",
        )
        child_resource = _node(
            "LSR_2",
            "Камера видеонаблюдения",
            UnitClass.COUNT,
            "4",
            "шт",
            is_resource=True,
            parent_id="LSR_1",
        )
        rd_item = _node(
            "SPEC_1",
            "Камера видеонаблюдения",
            UnitClass.COUNT,
            "4",
            "шт",
        )

        decisions = [
            MatchDecision(
                smeta_item=child_resource,
                rd_item=rd_item,
                confidence=Decimal("0.95"),
                status="AUTO",
                method="FUZZY_CASCADE",
            ),
            MatchDecision(
                smeta_item=parent_work,
                rd_item=None,
                confidence=Decimal("0"),
                status="AUTO",
                method="NO_MATCH",
            ),
        ]

        evaluations = build_rule_evaluations(decisions, [1, 2])

        assert len(evaluations) == 2
        assert evaluations[0]["rule_name"] != "SMETA_NOT_IN_RD"
        assert evaluations[1]["rule_name"] == "SMETA_PARENT_COVERED"
        assert evaluations[1]["is_violation"] is False


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