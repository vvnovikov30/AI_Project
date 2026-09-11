"""L4 Metrics & Reconciliation: Валидация качества ранжирования и детекции коллизий."""

from decimal import Decimal
import math
from typing import NamedTuple
import pytest

from evaluator import DiscrepancyDetector
from matcher import CascadeMatcher
from models import CellPointer, CollisionSeverity,NormalizedItem, UnitClass


class BenchmarkPair(NamedTuple):
    query: NormalizedItem
    candidates: list[NormalizedItem]
    ground_truth_id: str
    unit_cost: Decimal
    is_true_collision: bool


@pytest.fixture
def benchmark_suite() -> list[BenchmarkPair]:
    """Набор эталонных пар ПСД для оценки ранжирования и детекции коллизий."""
    pointer = CellPointer(
        doc_type="BENCHMARK",
        file_name="benchmark_psd.xlsx",
        sheet_name="Лист 1",
        row_idx=1,
        col_letter="A",
        coordinate="A1",
    )

    def _item(raw_id: str, title: str, unit: UnitClass, vol: Decimal, dims: dict[str, str] | None = None) -> NormalizedItem:
        return NormalizedItem(
            raw_id=raw_id,
            code=None,
            raw_title=title,
            clean_title=title.lower(),
            raw_unit=unit.value,
            canonical_unit=unit,
            unit_multiplier=Decimal("1.0"),
            volume=vol,
            dimensions=dims or {},
            context_section="Общий",
            pointer=pointer,
        )

    return [
        BenchmarkPair(
            query=_item("Q1", "Кабель силовой ВВГнг(A)-LS 3х2.5", UnitClass.LENGTH, Decimal("106.0"), {"CORES": "3", "SECTION": "2.5"}),
            candidates=[
                _item("C1_1", "Кабель силовой ВВГнг(A)-LS 3х2.5 ГОСТ 31996-2012", UnitClass.LENGTH, Decimal("100.0"), {"CORES": "3", "SECTION": "2.5"}),
                _item("C1_2", "Кабель контрольный КВВГнг(A)-LS 4х1.5", UnitClass.LENGTH, Decimal("100.0"), {"CORES": "4", "SECTION": "1.5"}),
                _item("C1_3", "Провод ПВ-1 1х2.5", UnitClass.LENGTH, Decimal("100.0"), {"CORES": "1", "SECTION": "2.5"}),
            ],
            ground_truth_id="C1_1",
            unit_cost=Decimal("120.50"),
            is_true_collision=False,
        ),
        BenchmarkPair(
            query=_item("Q2", "Задвижка 30с41нж DN50 PN16", UnitClass.COUNT, Decimal("15.0"), {"DN": "50", "PN": "16"}),
            candidates=[
                _item("C2_1", "Задвижка чугунная 30ч6бр DN50 PN10", UnitClass.COUNT, Decimal("10.0"), {"DN": "50", "PN": "10"}),
                _item("C2_2", "Задвижка 30с41нж DN50 PN16 стальная фланцевая", UnitClass.COUNT, Decimal("10.0"), {"DN": "50", "PN": "16"}),
                _item("C2_3", "Задвижка стальная 30с41нж DN100 PN16", UnitClass.COUNT, Decimal("10.0"), {"DN": "100", "PN": "16"}),
            ],
            ground_truth_id="C2_2",
            unit_cost=Decimal("8500.00"),
            is_true_collision=True,
        ),
        BenchmarkPair(
            query=_item("Q3", "Бетон тяжелый товарный БСГ В25 F150 W6", UnitClass.VOLUME, Decimal("550.0"), {"CONCRETE": "B25"}),
            candidates=[
                _item("C3_1", "Бетон тяжелый товарный БСГ В25 F150 W6 готовая смесь", UnitClass.VOLUME, Decimal("400.0"), {"CONCRETE": "B25"}),
                _item("C3_2", "Раствор готовый кладочный тяжелый М150", UnitClass.VOLUME, Decimal("100.0"), {}),
                _item("C3_3", "Бетон мелкозернистый В15", UnitClass.VOLUME, Decimal("50.0"), {"CONCRETE": "B15"}),
            ],
            ground_truth_id="C3_1",
            unit_cost=Decimal("5400.00"),
            is_true_collision=True,
        ),
        BenchmarkPair(
            query=_item("Q4", "Арматурная сталь гладкая и периодического профиля d16 А500С", UnitClass.MASS, Decimal("10.5"), {"REBAR_DIA": "16"}),
            candidates=[
                _item("C4_1", "Арматурная сталь гладкая и периодического профиля d16 А500С ГОСТ 34028-2016", UnitClass.MASS, Decimal("10.0"), {"REBAR_DIA": "16"}),
                _item("C4_2", "Арматура d12 А500С", UnitClass.MASS, Decimal("10.0"), {"REBAR_DIA": "12"}),
                _item("C4_3", "Сталь круглая углеродистая d20", UnitClass.MASS, Decimal("10.0"), {}),
            ],
            ground_truth_id="C4_1",
            unit_cost=Decimal("65000.00"),
            is_true_collision=False,
        ),
        BenchmarkPair(
            query=_item("Q5", "Труба профильная стальная 40х40х2", UnitClass.LENGTH, Decimal("1000.0"), {}),
            candidates=[
                _item("C5_1", "Труба профильная стальная 40х40х2 квадратная", UnitClass.LENGTH, Decimal("10.0"), {}),
                _item("C5_2", "Швеллер гнутый равнополочный 40х40х2.5", UnitClass.LENGTH, Decimal("10.0"), {}),
                _item("C5_3", "Уголок равнополочный 40х4", UnitClass.LENGTH, Decimal("10.0"), {}),
            ],
            ground_truth_id="C5_1",
            unit_cost=Decimal("210.00"),
            is_true_collision=True,
        ),
    ]


class TestRankingQualityMetrics:
    """Оценка точности ранжирования кандидатов каскадным скорером."""

    def test_precision_at_k_and_mrr(self, benchmark_suite: list[BenchmarkPair]) -> None:
        hits_at_1 = 0
        hits_at_3 = 0
        reciprocal_ranks: list[Decimal] = []

        for pair in benchmark_suite:
            scored_candidates = [
                (cand, CascadeMatcher.calculate_score(pair.query, cand))
                for cand in pair.candidates
            ]
            scored_candidates.sort(key=lambda x: x[1], reverse=True)
            ranked_ids = [cand.raw_id for cand, _ in scored_candidates]

            rank = ranked_ids.index(pair.ground_truth_id) + 1

            if rank == 1:
                hits_at_1 += 1
            if rank <= 3:
                hits_at_3 += 1

            reciprocal_ranks.append(Decimal("1.0") / Decimal(str(rank)))

        total = Decimal(str(len(benchmark_suite)))
        precision_1 = Decimal(str(hits_at_1)) / total
        precision_3 = Decimal(str(hits_at_3)) / total
        mrr = sum(reciprocal_ranks) / total

        assert precision_1 >= Decimal("0.90"), f"P@1 ниже порога: {precision_1}"
        assert precision_3 >= Decimal("0.98"), f"P@3 ниже порога: {precision_3}"
        assert mrr >= Decimal("0.92"), f"MRR ниже порога: {mrr}"

    def test_ndcg_at_3(self, benchmark_suite: list[BenchmarkPair]) -> None:
        ndcg_scores: list[Decimal] = []

        for pair in benchmark_suite:
            scored_candidates = [
                (cand, CascadeMatcher.calculate_score(pair.query, cand))
                for cand in pair.candidates
            ]
            scored_candidates.sort(key=lambda x: x[1], reverse=True)
            ranked_ids = [cand.raw_id for cand, _ in scored_candidates[:3]]

            idcg = Decimal("1.0")
            actual_dcg = Decimal("0.0")
            if pair.ground_truth_id in ranked_ids:
                rank_idx = ranked_ids.index(pair.ground_truth_id)
                discount = Decimal(str(math.log2(rank_idx + 2)))
                actual_dcg = Decimal("1.0") / discount

            ndcg_scores.append(actual_dcg / idcg)

        mean_ndcg_3 = sum(ndcg_scores) / Decimal(str(len(ndcg_scores)))
        assert mean_ndcg_3 >= Decimal("0.95"), f"NDCG@3 ниже порога: {mean_ndcg_3}"


class TestReconciliationAndFinancialRisk:
    """Проверка полноты выявления дефектов смет (CRR) и суммы предотвращенных переплат."""

    def test_collision_recall_rate_and_budget_risk(
        self, benchmark_suite: list[BenchmarkPair]
    ) -> None:
        detected_true_collisions = 0
        total_true_collisions = sum(1 for p in benchmark_suite if p.is_true_collision)
        total_budget_risk = Decimal("0.00")

        for pair in benchmark_suite:
            target_spec = next(c for c in pair.candidates if c.raw_id == pair.ground_truth_id)
            score = CascadeMatcher.calculate_score(pair.query, target_spec)
            collision_record = DiscrepancyDetector.evaluate(pair.query, target_spec, score)

            is_flagged = collision_record.severity in (
                CollisionSeverity.CRITICAL,
                CollisionSeverity.SUSPECT,
            )

            if pair.is_true_collision and is_flagged:
                detected_true_collisions += 1

            if collision_record.volume_delta > Decimal("0"):
                prevented_leak = collision_record.volume_delta * pair.unit_cost
                total_budget_risk += prevented_leak

        crr = Decimal(str(detected_true_collisions)) / Decimal(str(total_true_collisions))
        assert crr >= Decimal("0.95"), f"CRR ниже целевого порога 95%: {crr * Decimal('100')}%"

        assert total_budget_risk > Decimal("0.00")
        assert isinstance(total_budget_risk, Decimal)

        expected_risk = Decimal("1017200.00")
        assert total_budget_risk == expected_risk, f"Ошибка в расчете риска переплат: {total_budget_risk} != {expected_risk}"