import re
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Dict, Final, Optional, Sequence

from rapidfuzz import fuzz

from evaluator import PredicateEvaluator, VolumeSufficiencyEvaluator
from normalizer import is_dimension_compatible
from models import (
    MatchRecord,
    NormalizedItem,
    RuleBreakdown,
    RuleEvaluation,
    UnitClass,
)


class CascadeMatcher:
    """Каскадный матчер с адаптивным перераспределением весов и шлюзами размерностей."""

    W_CODE = 0.45
    W_FUZZY = 0.40
    W_CONTEXT = 0.15

    IGNORED_CODES = {
        "", "-", "—", "–", "б/н", "б.н.", "б/номер",
        "прайс", "прайс-лист", "прайс лист", "null", "none"
    }

    @classmethod
    def _clean_code(cls, code: Optional[str]) -> Optional[str]:
        """Фильтрация псевдо-шифров ценообразования и пустых значений."""
        if not code:
            return None
        c = str(code).strip()
        if c.lower() in cls.IGNORED_CODES:
            return None
        return c

    @staticmethod
    def _is_dimension_compatible(dim_lsr: Dict[str, Any], dim_spec: Dict[str, Any]) -> bool:
        """Dimensional Lock: блокирует несовпадающие диаметры, давления и классы бетона (симметрично)."""
        common_keys = dim_lsr.keys() & dim_spec.keys()
        return all(dim_lsr[k] == dim_spec[k] for k in common_keys)

    @staticmethod
    def _units_incompatible(lsr: NormalizedItem, spec: NormalizedItem) -> bool:
        """Мягкий Unit Gate: блокирует только доказанную несовместимость.

        Если единица хотя бы одной стороны не определена, сопоставление не
        блокируется (``None``/``UNKNOWN``), а несовместимость фиксируется
        предикатом R05_UNT как WARNING/SKIPPED.
        """
        if (
            lsr.canonical_unit != UnitClass.UNKNOWN
            and spec.canonical_unit != UnitClass.UNKNOWN
        ):
            return not is_dimension_compatible(
                lsr.canonical_unit.value, spec.canonical_unit.value
            )
        raw_lsr = (lsr.raw_unit or "").strip()
        raw_spec = (spec.raw_unit or "").strip()
        if raw_lsr and raw_spec:
            return not is_dimension_compatible(raw_lsr, raw_spec)
        return False

    @classmethod
    def calculate_score(cls, lsr: NormalizedItem, spec: NormalizedItem) -> float:
        # 1. Unit Gating: мягкая блокировка только при доказанной несовместимости
        if cls._units_incompatible(lsr, spec):
            return 0.0

        # 2. Dimensional Lock: блокировка при несовпадении технических параметров (MR4)
        if not cls._is_dimension_compatible(lsr.dimensions, spec.dimensions):
            return 0.0

        # Очистка шифров от сметных пометок
        c1 = cls._clean_code(lsr.code)
        c2 = cls._clean_code(spec.code)

        # 3. Сопоставление шифров / артикулов / марок
        s_code = 0.0
        code_matched = False

        if c1 and c2:
            if c1.lower() == c2.lower():
                s_code = 100.0
                code_matched = True
            elif len(c1) >= 3 and c1.lower() in spec.clean_title:
                s_code = 80.0
                code_matched = True
            elif len(c2) >= 3 and c2.lower() in lsr.clean_title:
                s_code = 80.0
                code_matched = True
            else:
                s_code = 0.0  # Оба шифра заданы, но прямо противоречат друг другу
        elif c1 and len(c1) >= 3 and c1.lower() in spec.clean_title:
            s_code = 80.0
            code_matched = True
        elif c2 and len(c2) >= 3 and c2.lower() in lsr.clean_title:
            s_code = 80.0
            code_matched = True

        # Шифр считается активным признаком, только если заданы оба сопоставимых кода,
        # либо если код одной стороны успешно подтвержден в тексте другой
        has_code = bool(c1 and c2) or code_matched

        # 4. Лексический скоринг (Fuzzy Matching)
        s_sort = float(fuzz.token_sort_ratio(lsr.clean_title, spec.clean_title))
        s_set = float(fuzz.token_set_ratio(lsr.clean_title, spec.clean_title))
        s_fuzzy = max(s_sort, s_set)

        # 5. Контекст разделов проекта/сметы
        has_context = bool(
            lsr.context_section
            and spec.context_section
            and lsr.context_section != "Default_Section"
            and spec.context_section != "Default_Section"
        )
        s_context = 100.0 if (has_context and lsr.context_section == spec.context_section) else 0.0

        # 6. Адаптивное динамическое перераспределение весов (Сумма весов всегда = 1.0)
        if has_code and has_context:
            return (cls.W_CODE * s_code) + (cls.W_FUZZY * s_fuzzy) + (cls.W_CONTEXT * s_context)
        elif has_code and not has_context:
            # Вес неразмеченного раздела (0.15) переходит в лексику
            return (cls.W_CODE * s_code) + ((1.0 - cls.W_CODE) * s_fuzzy)
        elif not has_code and has_context:
            # Вес отсутствующего шифра (0.45) переходит в лексику: 0.85 * fuzzy + 0.15 * context
            return ((1.0 - cls.W_CONTEXT) * s_fuzzy) + (cls.W_CONTEXT * s_context)
        else:
            # Ни шифр, ни раздел не сопоставимы: 100% веса отдается прямому лексическому сходству
            return s_fuzzy


class CompositeMatcher:
    """Композитное сопоставление узлов сметы «работа + дочерние ресурсы».

    Дочерний ресурс (``is_resource=True``) сопоставляется с позицией РД
    самостоятельно, но дополнительно проверяется на принадлежность к
    родительской расценке: найденный родитель фиксируется в
    ``MatchRecord.parent_work_id`` и считается покрытым (не попадает в
    ``SMETA_NOT_IN_RD``). Все оценки уверенности — строго ``Decimal``.
    """

    SCORE_QUANTUM: Final[Decimal] = Decimal("0.0001")
    AUTO_CONFIDENCE: Final[Decimal] = Decimal("0.7000")
    HITL_CONFIDENCE: Final[Decimal] = Decimal("0.4000")
    METHOD: Final[str] = "COMPOSITE_CASCADE"

    @classmethod
    def _decimal_score(cls, raw_score: Any) -> Decimal:
        """Переводит score CascadeMatcher (0..100) в Decimal-долю ``0..1``."""
        return (
            Decimal(str(round(raw_score, 4))) / Decimal("100")
        ).quantize(cls.SCORE_QUANTUM, rounding=ROUND_HALF_UP)

    @staticmethod
    def build_parent_index(
        smeta_items: Sequence[NormalizedItem],
    ) -> Dict[str, NormalizedItem]:
        """Индекс сметных узлов по ``raw_id`` для O(1) поиска родителя."""
        return {str(item.raw_id): item for item in smeta_items}

    @classmethod
    def resolve_parent(
        cls,
        smeta_item: NormalizedItem,
        parent_index: Dict[str, NormalizedItem],
    ) -> Optional[NormalizedItem]:
        """Возвращает расценку-родителя для дочернего ресурса либо ``None``."""
        if not smeta_item.is_resource or not smeta_item.parent_id:
            return None
        return parent_index.get(str(smeta_item.parent_id))

    @classmethod
    def match_pair(
        cls,
        rd_item: NormalizedItem,
        smeta_item: NormalizedItem,
        parent_work: Optional[NormalizedItem] = None,
    ) -> MatchRecord:
        """Строит композитную связь с полным аудитом предикатов R01-R05.

        Дополнительно для дочернего ресурса запускается вето-проверка
        ``R_CMP_VOLUME_DEFICIT``: превышение объема работы-родителя
        блокирует автоподтверждение связи.
        """
        confidence = cls._decimal_score(
            CascadeMatcher.calculate_score(rd_item, smeta_item)
        )
        breakdown: RuleBreakdown = PredicateEvaluator.evaluate_pair(
            rd_item, smeta_item
        )
        evaluations: list[RuleEvaluation] = list(breakdown.rule_evaluations)

        volume_gate: Optional[RuleEvaluation] = None
        if parent_work is not None and smeta_item.is_resource:
            volume_gate = VolumeSufficiencyEvaluator.evaluate(
                parent_work, [smeta_item]
            )
            evaluations.append(volume_gate)

        blocked = breakdown.is_blocked or (
            volume_gate is not None and volume_gate.status == "FAILED"
        )
        if blocked:
            status = "BLOCKED"
        elif confidence == Decimal("0") or confidence < cls.HITL_CONFIDENCE:
            status = "NO_MATCH"
        elif confidence >= cls.AUTO_CONFIDENCE and breakdown.is_valid:
            status = "AUTO"
        else:
            status = "HITL_PENDING"

        return MatchRecord(
            rd_item=rd_item,
            smeta_item=smeta_item,
            confidence=confidence,
            status=status,
            method=cls.METHOD,
            parent_work_id=(
                str(parent_work.raw_id) if parent_work is not None else None
            ),
            evaluations=evaluations,
            breakdown=breakdown,
        )

    @classmethod
    def match_composite(
        cls,
        rd_items: Sequence[NormalizedItem],
        smeta_items: Sequence[NormalizedItem],
    ) -> list[MatchRecord]:
        """Сопоставляет каждый сметный узел с лучшей позицией РД.

        Для дочерних ресурсов родитель поднимается из индекса сметы, поэтому
        связь ресурса автоматически покрывает и его расценку монтажа.
        Позиции без единого кандидата в результат не попадают.
        """
        if not rd_items or not smeta_items:
            return []

        parent_index = cls.build_parent_index(smeta_items)
        records: list[MatchRecord] = []
        for smeta_item in smeta_items:
            best_item: Optional[NormalizedItem] = None
            best_raw = 0.0
            for rd_item in rd_items:
                raw_score = CascadeMatcher.calculate_score(rd_item, smeta_item)
                if raw_score > best_raw:
                    best_item = rd_item
                    best_raw = raw_score
            if best_item is None:
                continue
            parent_work = cls.resolve_parent(smeta_item, parent_index)
            records.append(cls.match_pair(best_item, smeta_item, parent_work))
        return records

    @staticmethod
    def covered_parent_ids(records: Sequence[MatchRecord]) -> set[str]:
        """Идентификаторы расценок, покрытых сопоставленными ресурсами."""
        return {
            str(record.parent_work_id)
            for record in records
            if record.parent_work_id is not None
        }