from decimal import Decimal
from models import NormalizedItem, CollisionRecord, CollisionSeverity

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