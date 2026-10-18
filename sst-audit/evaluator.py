"""Расчетное ядро аудита: предикаты R01-R05 и количественные дельты.

Модуль не содержит float-арифметики: результаты RapidFuzz немедленно
переводятся в ``Decimal`` (Zero Float Policy), а все веса и допуски заданы
строками.
"""

from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Final, Optional

from rapidfuzz import fuzz

from models import (
    CollisionRecord,
    CollisionSeverity,
    NormalizedItem,
    RuleBreakdown,
    RuleEvaluation,
    UnitClass,
)
from normalizer import evaluate_unit_gate, is_dimension_compatible

class PredicateEvaluator:
    """Компонентный аудит предикатов R01-R05 для пары «РД ↔ СМР».

    Каждый предикат возвращает ``RuleEvaluation`` с нормированной оценкой
    ``0..1`` и весом из :class:`models.RuleBreakdown`. Итоговая уверенность
    считается только через ``Decimal``.
    """

    IGNORED_CODES: Final[frozenset[str]] = frozenset(
        {
            "",
            "-",
            "—",
            "–",
            "б/н",
            "б.н.",
            "б/номер",
            "прайс",
            "прайс-лист",
            "прайс лист",
            "null",
            "none",
        }
    )

    # Лексические маркеры работ СМР (категория R01_CAT).
    WORK_TITLE_MARKERS: Final[tuple[str, ...]] = (
        "монтаж",
        "установка",
        "прокладка",
        "разработка",
        "устройство",
        "укладка",
        "бетонирование",
        "демонтаж",
        "окраска",
        "штукатур",
    )
    WORK_UNITS: Final[frozenset[UnitClass]] = frozenset(
        {UnitClass.AREA, UnitClass.VOLUME}
    )

    SEMANTIC_PASS: Final[Decimal] = Decimal("0.70")
    SEMANTIC_WARN: Final[Decimal] = Decimal("0.40")
    RATIO_QUANTUM: Final[Decimal] = Decimal("0.0001")

    @staticmethod
    def clean_code(code: Optional[str]) -> Optional[str]:
        """Фильтрация псевдо-шифров ценообразования и пустых значений."""
        if not code:
            return None
        cleaned = str(code).strip()
        if cleaned.lower() in PredicateEvaluator.IGNORED_CODES:
            return None
        return cleaned

    @classmethod
    def _as_unit_ratio(cls, raw_score: Any) -> Decimal:
        """Переводит результат RapidFuzz (0..100) в Decimal-долю ``0..1``."""
        return (
            Decimal(str(round(raw_score, 4))) / Decimal("100")
        ).quantize(cls.RATIO_QUANTUM, rounding=ROUND_HALF_UP)

    @classmethod
    def _is_work_category(cls, item: NormalizedItem) -> bool:
        """Относит позицию к работам СМР либо к материалам/оборудованию."""
        if item.is_resource:
            return False
        title = item.clean_title.lower()
        if any(marker in title for marker in cls.WORK_TITLE_MARKERS):
            return True
        return item.canonical_unit in cls.WORK_UNITS

    @classmethod
    def evaluate_category(
        cls, rd_item: NormalizedItem, smeta_item: NormalizedItem
    ) -> RuleEvaluation:
        """R01_CAT: работа сопоставляется с работой, материал — с материалом."""
        rd_is_work = cls._is_work_category(rd_item)
        smeta_is_work = cls._is_work_category(smeta_item)
        matched = rd_is_work == smeta_is_work
        return RuleEvaluation(
            rule_code="R01_CAT",
            status="PASSED" if matched else "FAILED",
            weight=RuleBreakdown.R01_CAT,
            score=Decimal("1") if matched else Decimal("0"),
            message=(
                f"CATEGORY_MATCH: work={rd_is_work} vs {smeta_is_work}"
                if matched
                else f"CATEGORY_CONFLICT: work={rd_is_work} vs {smeta_is_work}"
            ),
        )

    @classmethod
    def evaluate_dimensions(
        cls, rd_item: NormalizedItem, smeta_item: NormalizedItem
    ) -> RuleEvaluation:
        """R02_DIM: блокировка несовпадающих DN/PN/класса бетона."""
        rd_dims = rd_item.dimensions or {}
        smeta_dims = smeta_item.dimensions or {}
        common_keys = sorted(set(rd_dims) & set(smeta_dims))
        if not common_keys:
            return RuleEvaluation(
                rule_code="R02_DIM",
                status="SKIPPED",
                weight=RuleBreakdown.R02_DIM,
                score=Decimal("1"),
                message="DIMENSION_LOCK_SKIPPED: нет общих параметров",
            )
        conflicts = [
            key
            for key in common_keys
            if str(rd_dims[key]).strip().upper()
            != str(smeta_dims[key]).strip().upper()
        ]
        if conflicts:
            return RuleEvaluation(
                rule_code="R02_DIM",
                status="FAILED",
                weight=RuleBreakdown.R02_DIM,
                score=Decimal("0"),
                message="DIMENSION_CONFLICT: " + ", ".join(conflicts),
            )
        return RuleEvaluation(
            rule_code="R02_DIM",
            status="PASSED",
            weight=RuleBreakdown.R02_DIM,
            score=Decimal("1"),
            message="DIMENSION_LOCK_PASSED: " + ", ".join(common_keys),
        )

    @classmethod
    def evaluate_standard(
        cls, rd_item: NormalizedItem, smeta_item: NormalizedItem
    ) -> RuleEvaluation:
        """R03_STD: совпадение шифра/ГОСТ/ТУ либо подтверждение в наименовании."""
        rd_code = cls.clean_code(rd_item.code)
        smeta_code = cls.clean_code(smeta_item.code)
        if not rd_code and not smeta_code:
            return RuleEvaluation(
                rule_code="R03_STD",
                status="SKIPPED",
                weight=RuleBreakdown.R03_STD,
                score=Decimal("1"),
                message="CODES_ABSENT",
            )
        if rd_code and smeta_code:
            if rd_code.lower() == smeta_code.lower():
                return RuleEvaluation(
                    rule_code="R03_STD",
                    status="PASSED",
                    weight=RuleBreakdown.R03_STD,
                    score=Decimal("1"),
                    message="CODE_EXACT_MATCH",
                )
            rd_in_title = (
                len(rd_code) >= 3 and rd_code.lower() in smeta_item.clean_title.lower()
            )
            smeta_in_title = (
                len(smeta_code) >= 3
                and smeta_code.lower() in rd_item.clean_title.lower()
            )
            if rd_in_title or smeta_in_title:
                return RuleEvaluation(
                    rule_code="R03_STD",
                    status="PASSED",
                    weight=RuleBreakdown.R03_STD,
                    score=Decimal("0.8"),
                    message="CODE_CONFIRMED_IN_TITLE",
                )
            return RuleEvaluation(
                rule_code="R03_STD",
                status="FAILED",
                weight=RuleBreakdown.R03_STD,
                score=Decimal("0"),
                message=f"CODE_CONFLICT: {rd_code} vs {smeta_code}",
            )
        present = rd_code or smeta_code
        other_title = (
            smeta_item.clean_title.lower() if rd_code else rd_item.clean_title.lower()
        )
        if present is not None and len(present) >= 3 and present.lower() in other_title:
            return RuleEvaluation(
                rule_code="R03_STD",
                status="PASSED",
                weight=RuleBreakdown.R03_STD,
                score=Decimal("0.8"),
                message="CODE_CONFIRMED_IN_TITLE",
            )
        return RuleEvaluation(
            rule_code="R03_STD",
            status="SKIPPED",
            weight=RuleBreakdown.R03_STD,
            score=Decimal("1"),
            message="CODE_ONLY_ONE_SIDE",
        )

    @classmethod
    def evaluate_semantic(
        cls, rd_item: NormalizedItem, smeta_item: NormalizedItem
    ) -> RuleEvaluation:
        """R04_MAT: лексическое сходство наименований (RapidFuzz -> Decimal)."""
        sort_ratio = fuzz.token_sort_ratio(
            rd_item.clean_title, smeta_item.clean_title
        )
        set_ratio = fuzz.token_set_ratio(
            rd_item.clean_title, smeta_item.clean_title
        )
        ratio = cls._as_unit_ratio(max(sort_ratio, set_ratio))
        if ratio >= cls.SEMANTIC_PASS:
            status = "PASSED"
        elif ratio >= cls.SEMANTIC_WARN:
            status = "WARNING"
        else:
            status = "FAILED"
        return RuleEvaluation(
            rule_code="R04_MAT",
            status=status,
            weight=RuleBreakdown.R04_MAT,
            score=ratio,
            message=f"SEMANTIC_SIMILARITY: {ratio}",
        )

    @classmethod
    def evaluate_units(
        cls, rd_item: NormalizedItem, smeta_item: NormalizedItem
    ) -> RuleEvaluation:
        """R05_UNT: шлюз размерностей единиц измерения.

        Если единица сметной позиции не определена (усеченная форма 4б без
        графы «Ед. изм.»), жесткая блокировка не применяется: предикат
        возвращает WARNING с весом 0.5 и причиной ``UNIT_MISSING_IN_SMETA``.
        """
        smeta_raw = (smeta_item.raw_unit or "").strip()
        if not smeta_raw and smeta_item.canonical_unit == UnitClass.UNKNOWN:
            return RuleEvaluation(
                rule_code="R05_UNT",
                status="WARNING",
                weight=RuleBreakdown.R05_UNT,
                score=Decimal("0.5"),
                message="UNIT_MISSING_IN_SMETA",
            )
        rd_unit = rd_item.raw_unit or rd_item.canonical_unit.value
        smeta_unit = smeta_raw or smeta_item.canonical_unit.value
        compatible, gate_score, reason = evaluate_unit_gate(rd_unit, smeta_unit)
        if reason.endswith("SKIPPED"):
            status = "SKIPPED"
        else:
            status = "PASSED" if compatible else "FAILED"
        return RuleEvaluation(
            rule_code="R05_UNT",
            status=status,
            weight=RuleBreakdown.R05_UNT,
            score=gate_score,
            message=reason,
        )

    @classmethod
    def evaluate_pair(
        cls, rd_item: NormalizedItem, smeta_item: NormalizedItem
    ) -> RuleBreakdown:
        """Выполняет все пять предикатов и возвращает сводку с весами."""
        breakdown = RuleBreakdown()
        breakdown.add(cls.evaluate_category(rd_item, smeta_item))
        breakdown.add(cls.evaluate_dimensions(rd_item, smeta_item))
        breakdown.add(cls.evaluate_standard(rd_item, smeta_item))
        breakdown.add(cls.evaluate_semantic(rd_item, smeta_item))
        breakdown.add(cls.evaluate_units(rd_item, smeta_item))
        breakdown.details_summary = " | ".join(
            f"{item.rule_code}={item.status}" for item in breakdown.rule_evaluations
        )
        return breakdown


class DiscrepancyDetector:
    """Выявление системных количественных расхождений и критических ошибок кратности."""

    ABS_TOLERANCE = Decimal("0.001")
    REL_TOLERANCE = Decimal("0.01")  # 1%[cite: 1]

    # Нормы технологических отходов по ГЭСН
    WASTE_COEFFICIENTS = {
        "REBAR": Decimal("0.05"),    # Арматура: среднее 5%[cite: 1]
        "CONCRETE": Decimal("0.02"), # Бетон: 2%[cite: 1]
        "CABLE": Decimal("0.06"),    # Кабель: 6%[cite: 1]
        "DEFAULT": Decimal("0.00")
    }

    @classmethod
    def evaluate(cls, lsr: NormalizedItem, spec: NormalizedItem, score: Decimal) -> CollisionRecord:
        v_spec = spec.volume
        v_lsr = lsr.volume

        # Определение коэффициента отхода
        k_waste = cls.WASTE_COEFFICIENTS["DEFAULT"]
        if "арматур" in lsr.clean_title:
            k_waste = cls.WASTE_COEFFICIENTS["REBAR"] #[cite: 1]
        elif "бетон" in lsr.clean_title:
            k_waste = cls.WASTE_COEFFICIENTS["CONCRETE"] #[cite: 1]
        elif "кабел" in lsr.clean_title:
            k_waste = cls.WASTE_COEFFICIENTS["CABLE"] #[cite: 1]

        v_spec_adjusted = v_spec * (Decimal("1.0") + k_waste) #[cite: 1]
        allowed_delta = cls.ABS_TOLERANCE + (cls.REL_TOLERANCE * v_spec_adjusted) #[cite: 1]
        delta = v_lsr - v_spec_adjusted

        # Анализ MULTIPLIER_ERROR (ошибка 10, 100, 1000 раз)
        ratio = v_lsr / v_spec if v_spec != Decimal("0") else Decimal("0")
        is_mult_error = any([
            Decimal("9.8") <= ratio <= Decimal("10.2"),
            Decimal("98.0") <= ratio <= Decimal("102.0"),
            Decimal("980.0") <= ratio <= Decimal("1020.0")
        ]) #[cite: 1]

        if is_mult_error:
            return CollisionRecord(
                lsr, spec, score, delta, allowed_delta,
                CollisionSeverity.CRITICAL,
                f"MULTIPLIER_ERROR: Объем завышен примерно в {round(ratio)} раз (забыт делитель единицы расценки)" #[cite: 1]
            )

        if abs(delta) > allowed_delta:
            severity = CollisionSeverity.CRITICAL if delta > Decimal("0") else CollisionSeverity.SUSPECT
            desc = "Завышение объема в смете" if delta > Decimal("0") else "Занижение объема относительно РД"
            return CollisionRecord(lsr, spec, score, delta, allowed_delta, severity, desc)

        return CollisionRecord(
            lsr, spec, score, delta, allowed_delta,
            CollisionSeverity.INFO,
            "Объемы соответствуют проектным с учетом допустимого расхода и погрешностей"
        )


class VolumeSufficiencyEvaluator:
    """R_CMP_VOLUME_DEFICIT: проверка достаточности объема родительской работы.

    Для композитных узлов (расценка монтажа + дочерние ресурсы) суммарный
    объем сопоставленных ресурсов не должен превышать объем работы-родителя
    с учетом допуска. Предикат является вето-проверкой: в композитную
    уверенность R01-R05 он не входит, но при ``FAILED`` блокирует
    автоподтверждение связи.
    """

    RULE_CODE: Final[str] = "R_CMP_VOLUME_DEFICIT"
    ABS_TOLERANCE: Final[Decimal] = Decimal("0.001")
    REL_TOLERANCE: Final[Decimal] = Decimal("0.01")
    SCORE_QUANTUM: Final[Decimal] = Decimal("0.0001")

    @classmethod
    def evaluate(
        cls,
        parent_work: NormalizedItem,
        resources: list[NormalizedItem],
    ) -> RuleEvaluation:
        """Сравнивает объем работы-родителя с суммой объемов ресурсов.

        :param parent_work: родительская расценка монтажа из сметы;
        :param resources: сопоставленные дочерние ресурсы (материалы).
        :return: ``PASSED``/``FAILED``/``SKIPPED`` с ``Decimal``-оценкой.
        """
        parent_volume = parent_work.volume
        parent_unit = parent_work.raw_unit or parent_work.canonical_unit.value

        comparable = [
            resource
            for resource in resources
            if resource.is_resource
            and resource.volume > Decimal("0")
            and is_dimension_compatible(
                parent_unit, resource.raw_unit or resource.canonical_unit.value
            )
        ]
        if parent_volume <= Decimal("0") or not comparable:
            return RuleEvaluation(
                rule_code=cls.RULE_CODE,
                status="SKIPPED",
                weight=Decimal("0"),
                score=Decimal("1"),
                message="VOLUME_COMPARE_SKIPPED: нет сопоставимых дочерних ресурсов",
            )

        consumed = sum(
            (resource.volume for resource in comparable), Decimal("0")
        )
        allowed = (
            parent_volume * (Decimal("1") + cls.REL_TOLERANCE)
            + cls.ABS_TOLERANCE
        )
        if consumed > allowed:
            return RuleEvaluation(
                rule_code=cls.RULE_CODE,
                status="FAILED",
                weight=Decimal("0"),
                score=Decimal("0"),
                message=(
                    "VOLUME_DEFICIT: ресурсы "
                    f"{consumed} > работа {parent_volume} ({parent_unit})"
                ),
            )

        score = (consumed / parent_volume).quantize(
            cls.SCORE_QUANTUM, rounding=ROUND_HALF_UP
        )
        if score > Decimal("1"):
            score = Decimal("1")
        return RuleEvaluation(
            rule_code=cls.RULE_CODE,
            status="PASSED",
            weight=Decimal("0"),
            score=score,
            message=(
                "VOLUME_SUFFICIENT: ресурсы "
                f"{consumed} <= работа {parent_volume} ({parent_unit})"
            ),
        )