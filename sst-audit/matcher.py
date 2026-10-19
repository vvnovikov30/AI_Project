import re
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from enum import Enum
from typing import Any, Dict, Final, Optional, Sequence

from rapidfuzz import fuzz

from evaluator import (
    PredicateEvaluator,
    VolumeSufficiencyEvaluator,
    calculate_volume_penalty,
)
from normalizer import is_dimension_compatible
from models import (
    MatchRecord,
    NormalizedItem,
    RuleBreakdown,
    RuleEvaluation,
    UnitClass,
)


# Порог Exact Match Bypass: лексическое сходство, при котором связь
# считается точной без проверки отрыва от второго кандидата.
EXACT_MATCH_THRESHOLD: Final[Decimal] = Decimal("0.95")
# Порог автоматического принятия связи в серой зоне.
AUTO_MATCH_THRESHOLD: Final[Decimal] = Decimal("0.70")
# Минимальный отрыв лидера от второго уникального кандидата.
AMBIGUITY_MARGIN: Final[Decimal] = Decimal("0.05")
# Допуск на погрешность округления объемов при проверке нулевой дельты.
VOLUME_ABS_TOLERANCE: Final[Decimal] = Decimal("0.001")
VOLUME_REL_TOLERANCE: Final[Decimal] = Decimal("0.01")


class MatchStatus(str, Enum):
    """Контракт терминальных статусов связи «смета ↔ РД» (SSOT-строки)."""

    AUTO = "AUTO"
    HITL_PENDING = "HITL_PENDING"


@dataclass(frozen=True)
class CandidateDecision:
    """Решение по дедуплицированному пулу кандидатов одной позиции сметы."""

    smeta_item: NormalizedItem
    rd_item: Optional[NormalizedItem]
    score: Decimal
    status: MatchStatus
    margin: Decimal
    reason: str
    candidates: tuple[tuple[Any, Decimal], ...] = ()


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
    CANDIDATE_MIN_CONFIDENCE: Final[Decimal] = Decimal("0.4000")
    MAX_CANDIDATES: Final[int] = 3
    METHOD: Final[str] = "COMPOSITE_CASCADE"

    @classmethod
    def _decimal_score(cls, raw_score: Any) -> Decimal:
        """Переводит score CascadeMatcher (0..100) в Decimal-долю ``0..1``."""
        return (
            Decimal(str(round(raw_score, 4))) / Decimal("100")
        ).quantize(cls.SCORE_QUANTUM, rounding=ROUND_HALF_UP)

    @classmethod
    def calculate_pair_confidence(
        cls, rd_item: NormalizedItem, smeta_item: NormalizedItem
    ) -> tuple[Decimal, Decimal]:
        """Возвращает базовую и итоговую уверенность пары.

        Итоговая уверенность — это ``base_confidence * K_vol``, где
        ``K_vol`` — мультипликативный :func:`evaluator.calculate_volume_penalty`.
        Результат квантуется до ``0.0001`` строго на ``Decimal``.

        :return: кортеж ``(base_confidence, final_confidence)``.
        """
        base_confidence = cls._decimal_score(
            CascadeMatcher.calculate_score(rd_item, smeta_item)
        )
        vol_penalty = calculate_volume_penalty(rd_item.volume, smeta_item.volume)
        final_confidence = (base_confidence * vol_penalty).quantize(
            cls.SCORE_QUANTUM, rounding=ROUND_HALF_UP
        )
        return base_confidence, final_confidence

    @staticmethod
    def _to_decimal(value: Any) -> Decimal:
        """Приводит score/объем к ``Decimal`` без потери точности ``str``."""
        if isinstance(value, Decimal):
            return value
        if value is None:
            return Decimal("0")
        return Decimal(str(value))

    @classmethod
    def _extract_candidate_payload(cls, cand: Any) -> tuple[Any, Decimal]:
        """Безопасно распаковывает кандидата любой поддерживаемой формы.

        В пуле встречаются ``ScoredCandidate`` (``cand.item`` /
        ``cand.rd_item``), кортежи ``(score, item)`` / ``(item, score)`` и
        «голые» :class:`NormalizedItem`. Хелпер исключает ``AttributeError``
        на любой из этих форм.

        :return: кортеж ``(item, score)``, где ``score`` — строго ``Decimal``.
        """
        if hasattr(cand, "item"):
            return cand.item, cls._to_decimal(getattr(cand, "score", None))
        if hasattr(cand, "rd_item"):
            return cand.rd_item, cls._to_decimal(getattr(cand, "score", None))
        if isinstance(cand, (tuple, list)):
            if len(cand) != 2:
                return cand, cls._to_decimal(getattr(cand, "score", None))
            first, second = cand
            if isinstance(first, (int, float, Decimal)) and not isinstance(first, bool):
                return second, cls._to_decimal(first)
            return first, cls._to_decimal(second)
        return cand, cls._to_decimal(getattr(cand, "score", None))

    # Обратная совместимость с контрактом ``_unwrap_candidate``.
    _unwrap_candidate = _extract_candidate_payload

    @classmethod
    def _candidate_key(cls, item: Any) -> tuple[str, Optional[Decimal], str]:
        """Ключ уникальности кандидата: имя + объем + единица измерения.

        Наименование берется из ``clean_title`` (fallback ``raw_title``),
        единица — из ``canonical_unit`` (fallback ``raw_unit``), поэтому
        косметические различия источников не порождают ложных дубликатов.
        """
        title = getattr(item, "clean_title", None) or getattr(item, "raw_title", "") or ""
        volume = getattr(item, "volume", None)
        if volume is not None and not isinstance(volume, Decimal):
            volume = Decimal(str(volume))
        unit = getattr(item, "canonical_unit", None)
        if unit is None:
            unit = getattr(item, "raw_unit", None)
        unit_value = getattr(unit, "value", unit)
        return (str(title).strip().lower(), volume, str(unit_value or "").strip().lower())

    @classmethod
    def deduplicate_candidates(
        cls, candidates: Sequence[Any]
    ) -> list[tuple[Any, Decimal]]:
        """Схлопывает семантические дубликаты до расчета отрыва (Margin).

        Из группы с идентичным ``(clean_title, volume, canonical_unit)``
        остается только кандидат с максимальным score. Результат отсортирован
        по убыванию score; при равенстве сохраняется исходный порядок пула.
        """
        best: dict[tuple[str, Optional[Decimal], str], tuple[Any, Decimal]] = {}
        order: list[tuple[str, Optional[Decimal], str]] = []
        for cand in candidates:
            item, score = cls._extract_candidate_payload(cand)
            if item is None:
                continue
            key = cls._candidate_key(item)
            current = best.get(key)
            if current is None:
                best[key] = (item, score)
                order.append(key)
            elif score > current[1]:
                best[key] = (item, score)
        return sorted((best[key] for key in order), key=lambda pair: pair[1], reverse=True)

    @classmethod
    def _volume_delta(cls, rd_item: Any, smeta_item: Any) -> Optional[Decimal]:
        """Абсолютная дельта объемов пары либо ``None`` без количественных данных."""
        rd_volume = getattr(rd_item, "volume", None)
        smeta_volume = getattr(smeta_item, "volume", None)
        if rd_volume is None or smeta_volume is None:
            return None
        return abs(cls._to_decimal(rd_volume) - cls._to_decimal(smeta_volume))

    @classmethod
    def _volume_within_tolerance(
        cls,
        rd_item: Any,
        smeta_item: Any,
        tolerance: Optional[Decimal] = None,
    ) -> bool:
        """Zero Float Policy: ``abs(delta_volume) <= tolerance``.

        Без явного допуска применяется регламентированная комбинация
        абсолютного (``0.001``) и относительного (``1%``) порогов.
        """
        delta = cls._volume_delta(rd_item, smeta_item)
        if delta is None:
            return False
        if tolerance is not None:
            return delta <= tolerance
        rd_value = abs(cls._to_decimal(getattr(rd_item, "volume", Decimal("0"))))
        smeta_value = abs(
            cls._to_decimal(getattr(smeta_item, "volume", Decimal("0")))
        )
        allowed = max(
            VOLUME_ABS_TOLERANCE, VOLUME_REL_TOLERANCE * max(rd_value, smeta_value)
        )
        return delta <= allowed

    @staticmethod
    def _units_compatible(smeta_item: Any, rd_item: Any) -> bool:
        """Совместимость размерностей пары (мягкий Unit Gate)."""
        return not CascadeMatcher._units_incompatible(rd_item, smeta_item)

    @classmethod
    def resolve_match(
        cls,
        smeta_item: NormalizedItem,
        candidates: Sequence[Any],
        *,
        exact_threshold: Decimal = EXACT_MATCH_THRESHOLD,
        auto_threshold: Decimal = AUTO_MATCH_THRESHOLD,
        margin_threshold: Decimal = AMBIGUITY_MARGIN,
        tolerance: Optional[Decimal] = None,
    ) -> CandidateDecision:
        """Классифицирует связь по дедуплицированному пулу кандидатов.

        Порядок решений:

        1. Exact Match Bypass: ``score >= exact_threshold``, нулевая дельта
           объема в пределах допуска и совместимые единицы дают ``AUTO``
           без проверки отрыва от Top-2.
        2. Иначе отрыв считается по уникальным кандидатам; отсутствие Top-2
           трактуется как полный отрыв ``Decimal("1.0")``.
        3. ``AUTO`` при ``score >= auto_threshold`` и
           ``margin >= margin_threshold``, иначе ``HITL_PENDING``.

        :param candidates: пул в любой форме, поддерживаемой
            :meth:`_extract_candidate_payload`; score — в долевой шкале ``0..1``.
        """
        pool = cls.deduplicate_candidates(candidates)
        if not pool:
            return CandidateDecision(
                smeta_item=smeta_item,
                rd_item=None,
                score=Decimal("0"),
                status=MatchStatus.AUTO,
                margin=Decimal("1.0"),
                reason="NO_CANDIDATE",
                candidates=(),
            )

        best_item, best_score = pool[0]
        margin = (
            (best_score - pool[1][1]).quantize(
                cls.SCORE_QUANTUM, rounding=ROUND_HALF_UP
            )
            if len(pool) > 1
            else Decimal("1.0")
        )

        if (
            best_score >= exact_threshold
            and cls._volume_within_tolerance(best_item, smeta_item, tolerance)
            and cls._units_compatible(smeta_item, best_item)
        ):
            status = MatchStatus.AUTO
            reason = "EXACT_MATCH_BYPASS"
        elif best_score >= auto_threshold and margin >= margin_threshold:
            status = MatchStatus.AUTO
            reason = "MARGIN_CONFIRMED"
        else:
            status = MatchStatus.HITL_PENDING
            reason = (
                "AMBIGUOUS_MARGIN"
                if best_score >= auto_threshold
                else "BELOW_AUTO_THRESHOLD"
            )

        return CandidateDecision(
            smeta_item=smeta_item,
            rd_item=best_item,
            score=best_score,
            status=status,
            margin=margin,
            reason=reason,
            candidates=tuple(pool),
        )

    @classmethod
    def rank_candidates(
        cls,
        smeta_item: NormalizedItem,
        rd_items: Sequence[NormalizedItem],
        limit: int = MAX_CANDIDATES,
    ) -> list[tuple[NormalizedItem, Decimal]]:
        """Формирует Top-N кандидатов РД с Volume Gate и дедупликацией.

        Кандидат проходит отбор, только если его итоговая уверенность с учетом
        мультипликативного штрафа за объем не ниже
        :attr:`CANDIDATE_MIN_CONFIDENCE` (``0.40``). Дубликаты с одинаковыми
        именем, объемом и единицей измерения схлопываются: в Top-N попадает
        только наиболее уверенный представитель группы.
        """
        if not rd_items or limit <= 0:
            return []

        scored: list[tuple[NormalizedItem, Decimal]] = []
        for rd_item in rd_items:
            _, final_confidence = cls.calculate_pair_confidence(rd_item, smeta_item)
            if final_confidence < cls.CANDIDATE_MIN_CONFIDENCE:
                continue
            scored.append((rd_item, final_confidence))

        scored.sort(key=lambda pair: pair[1], reverse=True)

        selected: list[tuple[NormalizedItem, Decimal]] = []
        seen: set[tuple[str, Decimal, str]] = set()
        for rd_item, final_confidence in scored:
            key = cls._candidate_key(rd_item)
            if key in seen:
                continue
            seen.add(key)
            selected.append((rd_item, final_confidence))
            if len(selected) >= limit:
                break
        return selected

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
        _base_confidence, confidence = cls.calculate_pair_confidence(
            rd_item, smeta_item
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