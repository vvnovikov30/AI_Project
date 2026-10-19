import pytest
from decimal import Decimal
from types import SimpleNamespace
from models import NormalizedItem, CellPointer, UnitClass
from evaluator import PredicateEvaluator
from matcher import CascadeMatcher, CompositeMatcher, MatchStatus
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


def make_item(raw_id="1", title="Камера", code=None, section="Default_Section", unit=UnitClass.COUNT, dims=None, raw_unit="шт"):
    return NormalizedItem(
        raw_id=raw_id,
        code=code,
        raw_title=title,
        clean_title=title.lower(),
        raw_unit=raw_unit,
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
        """Несовместимые единицы блокируются, UNKNOWN — мягкий пропуск."""
        item_count = make_item(unit=UnitClass.COUNT)
        item_length = make_item(unit=UnitClass.LENGTH)
        item_unknown = make_item(unit=UnitClass.UNKNOWN)

        assert CascadeMatcher.calculate_score(item_count, item_length) == 0.0
        assert CascadeMatcher.calculate_score(item_count, item_unknown) > 0.0

    def test_matching_with_missing_smeta_unit(self):
        """Усеченная 4б: отсутствие единицы в смете не блокирует матчинг."""
        lsr_item = make_item(unit=UnitClass.COUNT, raw_unit="шт")
        smeta_item = make_item(unit=UnitClass.UNKNOWN, raw_unit="")

        assert CascadeMatcher.calculate_score(lsr_item, smeta_item) > 0.0
        breakdown = PredicateEvaluator.evaluate_pair(lsr_item, smeta_item)
        unit_gate = breakdown.get("R05_UNT")
        assert unit_gate is not None
        assert unit_gate.status == "WARNING"
        assert unit_gate.score == Decimal("0.5")
        assert unit_gate.message == "UNIT_MISSING_IN_SMETA"

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


class TestVolumeGateCandidates:
    """Volume Gate и дедупликация Top-N кандидатов (Спринт 5.2)."""

    def test_matcher_discards_candidates_with_heavy_volume_penalty(
        self, monkeypatch
    ):
        """База 46.7% при расхождении 10x дает < 20% и отсекается порогом 0.40."""
        monkeypatch.setattr(
            CascadeMatcher,
            "calculate_score",
            classmethod(lambda cls, lsr, spec: 46.7),
        )
        smeta_item = _node(
            "LSR_1", "Кабель силовой ВВГнг 3х2.5", UnitClass.LENGTH, "25", "м"
        )
        rd_item = _node(
            "SPEC_1", "Кабель силовой ВВГнг 3х2.5", UnitClass.LENGTH, "2.52", "м"
        )

        base_confidence, final_confidence = CompositeMatcher.calculate_pair_confidence(
            rd_item, smeta_item
        )
        assert base_confidence == Decimal("0.4670")
        assert final_confidence == Decimal("0.0934")
        assert final_confidence < Decimal("0.20")
        assert CompositeMatcher.rank_candidates(smeta_item, [rd_item]) == []

        record = CompositeMatcher.match_pair(rd_item, smeta_item)
        assert record.confidence == final_confidence
        assert record.status == "NO_MATCH"

    def test_matcher_deduplicates_identical_candidates_in_top_n(self) -> None:
        """Дубликаты с одинаковыми именем, объемом и единицей не двоятся в Top-3."""
        smeta_item = _node(
            "LSR_1", "Труба стальная 57х3.5", UnitClass.LENGTH, "10", "м"
        )
        first = _node(
            "SPEC_1", "Труба стальная 57х3.5", UnitClass.LENGTH, "10", "м"
        )
        duplicate = _node(
            "SPEC_2", "  труба стальная 57х3.5 ", UnitClass.LENGTH, "10", "м"
        )
        other_a = _node(
            "SPEC_3", "Труба стальная 89х4", UnitClass.LENGTH, "10", "м"
        )
        other_b = _node(
            "SPEC_4", "Труба стальная 108х4", UnitClass.LENGTH, "10", "м"
        )

        ranked = CompositeMatcher.rank_candidates(
            smeta_item, [first, duplicate, other_a, other_b]
        )

        assert len(ranked) == CompositeMatcher.MAX_CANDIDATES
        keys = [CompositeMatcher._candidate_key(item) for item, _ in ranked]
        assert len(keys) == len(set(keys))
        assert ranked[0][0].raw_id == "SPEC_1"
        assert ranked[0][1] == Decimal("1.0000")
        assert all(item.raw_id != "SPEC_2" for item, _ in ranked)

    def test_matcher_keeps_candidates_above_threshold_with_matching_volumes(
        self,
    ) -> None:
        """Равные объемы не штрафуются: кандидат проходит порог 0.40."""
        smeta_item = _node(
            "LSR_1", "Задвижка 30с41нж DN50 PN16", UnitClass.COUNT, "15", "шт"
        )
        rd_item = _node(
            "SPEC_1", "Задвижка 30с41нж DN50 PN16", UnitClass.COUNT, "15", "шт"
        )

        ranked = CompositeMatcher.rank_candidates(smeta_item, [rd_item])

        assert len(ranked) == 1
        assert ranked[0][0] is rd_item
        assert ranked[0][1] == Decimal("1.0000")


class TestExactMatchBypassAndDeduplication:
    """Margin Trap: 100% совпадение и клоны кандидатов (Спринт 5.3)."""

    def test_exact_match_with_zero_delta_gets_auto_status(self) -> None:
        """Top-1 100% + нулевая дельта получает AUTO без проверки отрыва."""
        smeta_item = _node(
            "LSR_1", "Труба стальная 57х3.5", UnitClass.LENGTH, "10", "м"
        )
        exact = _node(
            "SPEC_1", "Труба стальная 57х3.5", UnitClass.LENGTH, "10", "м"
        )
        clone = _node(
            "SPEC_2", "  труба стальная 57х3.5 ", UnitClass.LENGTH, "10", "м"
        )
        other = _node(
            "SPEC_3", "Труба стальная 89х4", UnitClass.LENGTH, "10", "м"
        )

        decision = CompositeMatcher.resolve_match(
            smeta_item,
            [
                (exact, Decimal("1.0")),
                (clone, Decimal("1.0")),
                (other, Decimal("0.96")),
            ],
        )

        assert decision.status is MatchStatus.AUTO
        assert decision.status.value == "AUTO"
        assert decision.reason == "EXACT_MATCH_BYPASS"
        assert decision.rd_item is exact
        # Клон не участвует в отрыве: margin считается от семантически
        # отличного кандидата и равен 0.04, что ниже порога 0.05.
        assert decision.margin == Decimal("0.0400")

    def test_candidate_pool_deduplication_by_clean_title_and_volume(self) -> None:
        """Клоны по clean_title + volume + canonical_unit схлопываются до margin."""
        smeta_item = _node(
            "LSR_1", "Кабель силовой ВВГнг 3х1.5", UnitClass.LENGTH, "10", "м"
        )
        first = _node(
            "SPEC_1", "Кабель силовой ВВГнг 3х1.5", UnitClass.LENGTH, "10", "м"
        )
        duplicate = _node(
            "SPEC_2", "кабель силовой ввгнг 3х1.5 ", UnitClass.LENGTH, "10", "м"
        )
        third = _node(
            "SPEC_3", "Кабель силовой ВВГнг 3х2.5", UnitClass.LENGTH, "10", "м"
        )

        pool = CompositeMatcher.deduplicate_candidates(
            [
                (first, Decimal("0.94")),
                (duplicate, Decimal("0.92")),
                (third, Decimal("0.90")),
            ]
        )

        assert [item.raw_id for item, _ in pool] == ["SPEC_1", "SPEC_3"]
        assert all(score <= Decimal("0.94") for _, score in pool)

        decision = CompositeMatcher.resolve_match(
            smeta_item,
            [
                (first, Decimal("0.94")),
                (duplicate, Decimal("0.92")),
                (third, Decimal("0.90")),
            ],
        )

        assert len(decision.candidates) == 2
        assert decision.rd_item is first
        # Отрыв считается от SPEC_3, а не от клона SPEC_2.
        assert decision.margin == Decimal("0.0400")
        assert decision.status is MatchStatus.HITL_PENDING
        assert decision.reason == "AMBIGUOUS_MARGIN"

    def test_auto_status_records_excluded_from_hitl_pending_queue(self) -> None:
        """AUTO-записи не попадают в очередь ручной проверки HITL_PENDING."""
        from main import match_items

        smeta_exact = _node(
            "LSR_EXACT", "Труба стальная 57х3.5", UnitClass.LENGTH, "10", "м"
        )
        rd_exact = _node(
            "SPEC_EXACT", "Труба стальная 57х3.5", UnitClass.LENGTH, "10", "м"
        )
        smeta_ambiguous = _node(
            "LSR_AMB", "Кабель силовой ВВГнг 3х1.5", UnitClass.LENGTH, "10", "м"
        )
        rd_ambiguous = _node(
            "SPEC_AMB1", "Кабель силовой ВВГнг 3х1.5", UnitClass.LENGTH, "12", "м"
        )
        rd_similar = _node(
            "SPEC_AMB2", "Кабель силовой ВВГнг 3х2.5", UnitClass.LENGTH, "10", "м"
        )

        decisions = match_items(
            [rd_exact, rd_ambiguous, rd_similar],
            [smeta_exact, smeta_ambiguous],
        )

        auto_decisions = [d for d in decisions if d.status == "AUTO"]
        pending = [d for d in decisions if d.status == "HITL_PENDING"]

        exact_decision = next(
            d for d in decisions if d.smeta_item is not None
            and d.smeta_item.raw_id == "LSR_EXACT"
        )
        assert exact_decision.status == "AUTO"
        assert exact_decision.rd_item is rd_exact
        assert exact_decision not in pending

        assert pending, "Неоднозначная пара должна остаться в очереди HITL"
        assert all(d.status == "HITL_PENDING" for d in pending)
        assert all(d not in pending for d in auto_decisions)
        pending_smeta_ids = {
            d.smeta_item.raw_id for d in pending if d.smeta_item is not None
        }
        assert "LSR_EXACT" not in pending_smeta_ids
        assert "LSR_AMB" in pending_smeta_ids

    def test_candidate_payload_unwrap_handles_all_shapes(self) -> None:
        """Хелпер распаковки не падает на ScoredCandidate, кортежах и item."""
        item = _node(
            "SPEC_1", "Труба стальная 57х3.5", UnitClass.LENGTH, "10", "м"
        )
        wrapped = SimpleNamespace(item=item, score=Decimal("0.97"))
        rd_wrapped = SimpleNamespace(rd_item=item, score=0.97)

        assert CompositeMatcher._extract_candidate_payload(wrapped) == (
            item,
            Decimal("0.97"),
        )
        assert CompositeMatcher._unwrap_candidate(rd_wrapped) == (
            item,
            Decimal("0.97"),
        )
        assert CompositeMatcher._extract_candidate_payload(
            (Decimal("0.5"), item)
        ) == (item, Decimal("0.5"))
        assert CompositeMatcher._extract_candidate_payload(
            (item, Decimal("0.8"))
        ) == (item, Decimal("0.8"))
        assert CompositeMatcher._extract_candidate_payload(item) == (
            item,
            Decimal("0"),
        )