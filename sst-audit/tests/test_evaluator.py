"""Юнит-тесты мультипликативного Volume Gate (Спринт 5.2).

Проверяется лестница штрафов :func:`evaluator.calculate_volume_penalty` и
гарантия потолка 20% при расхождении объемов на порядок (10x).
"""

from decimal import Decimal

import pytest

from evaluator import calculate_volume_penalty


class TestVolumePenalty:
    """Мультипликатор ``K_vol`` строго на ``Decimal``."""

    def test_volume_penalty_exact_match_has_no_penalty(self) -> None:
        assert calculate_volume_penalty(Decimal("25"), Decimal("25")) == Decimal("1.00")
        assert calculate_volume_penalty(
            Decimal("2.52"), Decimal("2.52")
        ) == Decimal("1.00")
        assert calculate_volume_penalty(
            Decimal("-3.5"), Decimal("3.5")
        ) == Decimal("1.00")

    def test_volume_penalty_within_rounding_tolerance_has_no_penalty(self) -> None:
        # 25 / 20 = 1.25 — верхняя граница зоны отсутствия штрафа.
        assert calculate_volume_penalty(Decimal("25"), Decimal("20")) == Decimal("1.00")

    def test_volume_penalty_graduated_bands(self) -> None:
        assert calculate_volume_penalty(Decimal("3"), Decimal("2")) == Decimal("0.80")
        assert calculate_volume_penalty(Decimal("10"), Decimal("2")) == Decimal("0.50")
        assert calculate_volume_penalty(Decimal("14"), Decimal("2")) == Decimal("0.30")
        assert calculate_volume_penalty(Decimal("24"), Decimal("2")) == Decimal("0.20")
        assert calculate_volume_penalty(Decimal("200"), Decimal("2")) == Decimal("0.05")

    def test_volume_penalty_missing_or_zero_volume(self) -> None:
        assert calculate_volume_penalty(None, Decimal("25")) == Decimal("0.70")
        assert calculate_volume_penalty(Decimal("25"), None) == Decimal("0.70")
        assert calculate_volume_penalty(Decimal("0"), Decimal("25")) == Decimal("0.20")
        assert calculate_volume_penalty(Decimal("25"), Decimal("0")) == Decimal("0.20")

    @pytest.mark.parametrize(
        ("v_rd", "v_smeta"),
        [
            ("25", "2.52"),  # номинальное расхождение 10x с погрешностью 0.8%
            ("25", "0.25"),  # расхождение 100x
            ("25", "250"),   # обратное расхождение 10x
        ],
    )
    def test_volume_penalty_order_of_magnitude_caps_score_at_twenty_percent(
        self, v_rd: str, v_smeta: str
    ) -> None:
        penalty = calculate_volume_penalty(Decimal(v_rd), Decimal(v_smeta))
        assert penalty <= Decimal("0.20")

        base_confidence = Decimal("1.00")
        final_confidence = (base_confidence * penalty).quantize(Decimal("0.0001"))
        assert final_confidence <= Decimal("0.20")

    def test_volume_penalty_is_symmetric(self) -> None:
        assert calculate_volume_penalty(
            Decimal("25"), Decimal("2.52")
        ) == calculate_volume_penalty(Decimal("2.52"), Decimal("25"))
