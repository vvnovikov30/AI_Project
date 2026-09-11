"""L3 Property-Based & Metamorphic Testing.

Реализует:
1. Hypothesis Property-Based тесты:
   - Идемпотентность лексической нормализации: f(f(x)) == f(x).
   - Round-trip сериализация: Pydantic v2 <-> SQLite TEXT без потери разрядности Decimal.
2. Метаморфические отношения (MR-1..MR-5):
   - MR-1: Симметрия сходства S(A, B) == S(B, A).
   - MR-2: Токенизационная инвариантность (перестановка слов в названии).
   - MR-3: Устойчивость ранжирования к нейтральному контексту ("по ГОСТ", "в сборе").
   - MR-4: Разрушение шлюза параметров (Dimensional Lock: сброс скора в 0.0 при несовпадении DN/PN/B).
   - MR-5: Разрушение шлюза единиц (Unit Gating: сброс скора в 0.0 при конфликте классов ОКЕИ).
"""

from decimal import Decimal
import json
import re
import sqlite3
from typing import Any

from hypothesis import given, strategies as st
import pytest

from conftest import StrictCellPointerSchema, StrictNormalizedItemSchema
from matcher import CascadeMatcher
from models import CellPointer, NormalizedItem, UnitClass
from normalizer import DimensionLockExtractor

# ----------------------------------------------------------------------
# 1. Вспомогательные функции и стратегии Hypothesis
# ----------------------------------------------------------------------


def clean_lexical_title(text: str) -> str:
    """Детерминированная очистка наименования от лишних пробелов и регистра."""
    return re.sub(r"\s+", " ", text).lower().strip()


# Генератор строковых представлений чисел с фиксированной точкой
decimal_text_strategy = st.decimals(
    min_value=Decimal("0.001"),
    max_value=Decimal("999999.999"),
    places=3,
).map(str)

# Генератор кириллических строительных наименований
cyrillic_words = st.sampled_from([
    "Кабель", "ВВГнг-LS", "Арматура", "А500С", "Бетон", "В25",
    "Задвижка", "Труба", "Фланец", "Кирпич", "Раствор", "Утеплитель"
])
construction_title_strategy = st.lists(cyrillic_words, min_size=2, max_size=5).map(" ".join)


# ----------------------------------------------------------------------
# 2. Property-Based тесты (Идемпотентность и Round-Trip)
# ----------------------------------------------------------------------


class TestPropertyBasedInvariants:
    """Тестирование фундаментальных математических свойств конвейера."""

    @given(text=st.text(min_size=1, max_size=200))
    def test_normalization_idempotency(self, text: str) -> None:
        """Инвариант идемпотентности: повторная нормализация не изменяет результат f(f(x)) == f(x)."""
        first_pass = clean_lexical_title(text)
        second_pass = clean_lexical_title(first_pass)
        assert first_pass == second_pass

    @given(
        volume_str=decimal_text_strategy,
        mult_str=st.sampled_from(["1.0", "10.0", "100.0", "1000.0"]),
        title=construction_title_strategy,
        row_idx=st.integers(min_value=1, max_value=50000),
    )
    def test_pydantic_sqlite_roundtrip_precision(
        self,
        volume_str: str,
        mult_str: str,
        title: str,
        row_idx: int,
    ) -> None:
        """Инвариант обратимости (Round-trip): Pydantic -> SQLite (TEXT) -> Pydantic.
        
        Гарантирует отсутствие потерь точности Decimal и неизменность CellPointer.
        """
        conn = sqlite3.connect(":memory:")
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE audit_reconciliation_buffer (
                raw_id TEXT PRIMARY KEY,
                code TEXT,
                raw_title TEXT NOT NULL,
                clean_title TEXT NOT NULL,
                raw_unit TEXT NOT NULL,
                canonical_unit TEXT NOT NULL,
                unit_multiplier TEXT NOT NULL,
                volume TEXT NOT NULL,
                dimensions_json TEXT NOT NULL,
                context_section TEXT NOT NULL,
                doc_type TEXT NOT NULL,
                file_name TEXT NOT NULL,
                sheet_name TEXT NOT NULL,
                row_idx INTEGER NOT NULL,
                col_letter TEXT NOT NULL,
                coordinate TEXT NOT NULL,
                is_resource INTEGER NOT NULL,
                parent_id TEXT
            )
        """)

        pointer = StrictCellPointerSchema(
            doc_type="LSR",
            file_name="reconciliation_volume.xlsx",
            sheet_name="Раздел 1",
            row_idx=row_idx,
            col_letter="C",
            coordinate=f"C{row_idx}",
        )
        original_item = StrictNormalizedItemSchema(
            raw_id=f"ITEM_{row_idx}",
            code="ГОСТ 34028-2016",
            raw_title=title,
            clean_title=clean_lexical_title(title),
            raw_unit="т",
            canonical_unit=UnitClass.MASS,
            unit_multiplier=Decimal(mult_str),
            volume=Decimal(volume_str),
            dimensions={"REBAR_DIA": "12"},
            context_section="Конструкции железобетонные",
            pointer=pointer,
            is_resource=False,
            parent_id=None,
        )

        cursor.execute(
            """
            INSERT INTO audit_reconciliation_buffer (
                raw_id, code, raw_title, clean_title, raw_unit, canonical_unit,
                unit_multiplier, volume, dimensions_json, context_section,
                doc_type, file_name, sheet_name, row_idx, col_letter, coordinate,
                is_resource, parent_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                original_item.raw_id,
                original_item.code,
                original_item.raw_title,
                original_item.clean_title,
                original_item.raw_unit,
                original_item.canonical_unit.value,
                str(original_item.unit_multiplier),
                str(original_item.volume),
                json.dumps(original_item.dimensions),
                original_item.context_section,
                original_item.pointer.doc_type,
                original_item.pointer.file_name,
                original_item.pointer.sheet_name,
                original_item.pointer.row_idx,
                original_item.pointer.col_letter,
                original_item.pointer.coordinate,
                1 if original_item.is_resource else 0,
                original_item.parent_id,
            ),
        )
        conn.commit()

        cursor.execute("SELECT * FROM audit_reconciliation_buffer WHERE raw_id = ?", (original_item.raw_id,))
        row = cursor.fetchone()
        assert row is not None

        loaded_pointer = StrictCellPointerSchema(
            doc_type=row["doc_type"],
            file_name=row["file_name"],
            sheet_name=row["sheet_name"],
            row_idx=row["row_idx"],
            col_letter=row["col_letter"],
            coordinate=row["coordinate"],
        )
        loaded_item = StrictNormalizedItemSchema(
            raw_id=row["raw_id"],
            code=row["code"],
            raw_title=row["raw_title"],
            clean_title=row["clean_title"],
            raw_unit=row["raw_unit"],
            canonical_unit=UnitClass(row["canonical_unit"]),
            unit_multiplier=Decimal(row["unit_multiplier"]),
            volume=Decimal(row["volume"]),
            dimensions=json.loads(row["dimensions_json"]),
            context_section=row["context_section"],
            pointer=loaded_pointer,
            is_resource=bool(row["is_resource"]),
            parent_id=row["parent_id"],
        )

        conn.close()

        assert original_item == loaded_item
        assert original_item.volume == loaded_item.volume
        assert str(original_item.volume) == str(loaded_item.volume)
        assert original_item.pointer == loaded_item.pointer


# ----------------------------------------------------------------------
# 3. Метаморфическое тестирование (MR-1..MR-5)
# ----------------------------------------------------------------------


class TestMetamorphicRelations:
    """Верификация метаморфических инвариантов каскадного алгоритма матчинга."""

    @pytest.fixture
    def default_pointer(self) -> CellPointer:
        return CellPointer(
            doc_type="TEST",
            file_name="test.xlsx",
            sheet_name="Лист 1",
            row_idx=1,
            col_letter="A",
            coordinate="A1",
        )

    def _make_item(
        self,
        raw_id: str,
        title: str,
        unit: UnitClass,
        pointer: CellPointer,
        dimensions: dict[str, Any] | None = None,
        code: str | None = None,
        context: str = "Общий раздел",
    ) -> NormalizedItem:
        return NormalizedItem(
            raw_id=raw_id,
            code=code,
            raw_title=title,
            clean_title=clean_lexical_title(title),
            raw_unit=unit.value,
            canonical_unit=unit,
            unit_multiplier=Decimal("1.0"),
            volume=Decimal("100.000"),
            dimensions=dimensions or {},
            context_section=context,
            pointer=pointer,
        )

    def test_mr1_similarity_symmetry(self, default_pointer: CellPointer) -> None:
        """MR-1 (Симметрия сходства): S(A, B) == S(B, A) при совпадении условий кодификации."""
        title_a = "Кабель силовой бронированный ВБбШв 4х10"
        title_b = "ВБбШв 4х10 кабель силовой бронированный"

        item_a = self._make_item("A1", title_a, UnitClass.LENGTH, default_pointer, dimensions={"CORES": "4"})
        item_b = self._make_item("B1", title_b, UnitClass.LENGTH, default_pointer, dimensions={"CORES": "4"})

        score_ab = Decimal(str(round(CascadeMatcher.calculate_score(item_a, item_b), 4)))
        score_ba = Decimal(str(round(CascadeMatcher.calculate_score(item_b, item_a), 4)))

        assert score_ab == score_ba
        assert score_ab >= Decimal("85.0")

    def test_mr2_token_order_invariance(self, default_pointer: CellPointer) -> None:
        """MR-2 (Токенизационная инвариантность): перестановка слов не снижает оценку ниже порога допуска."""
        title_source = "Задвижка чугунная фланцевая 30ч6бр"
        title_permuted = "30ч6бр фланцевая чугунная задвижка"

        item_orig = self._make_item("T1", title_source, UnitClass.COUNT, default_pointer)
        item_perm = self._make_item("T2", title_permuted, UnitClass.COUNT, default_pointer)

        score = Decimal(str(round(CascadeMatcher.calculate_score(item_orig, item_perm), 2)))
        assert score >= Decimal("85.0")

    def test_mr3_neutral_context_noise_resilience(self, default_pointer: CellPointer) -> None:
        """MR-3 (Устойчивость к нейтральному шуму): добавление стоп-слов сохраняет уверенность скоринга."""
        base_title = "Труба стальная электросварная 108х4"
        noisy_title = "Труба стальная электросварная 108х4 по ГОСТ 10704-91 в комплекте с креплениями"

        item_base = self._make_item("P1", base_title, UnitClass.LENGTH, default_pointer)
        item_noisy = self._make_item("P2", noisy_title, UnitClass.LENGTH, default_pointer)

        score = Decimal(str(round(CascadeMatcher.calculate_score(item_base, item_noisy), 2)))
        assert score >= Decimal("70.0")

    def test_mr4_dimensional_lock_annihilation(self, default_pointer: CellPointer) -> None:
        """MR-4 (Разрушение шлюза размерности): конфликт калибра обнуляет скор (I_dim = 0 -> S_total = 0.0)."""
        title_dn50 = "Задвижка стальная фланцевая 30с41нж Ду 50 Ру 16"
        title_dn100 = "Задвижка стальная фланцевая 30с41нж Ду 100 Ру 16"

        dims_dn50 = DimensionLockExtractor.extract(title_dn50)
        dims_dn100 = DimensionLockExtractor.extract(title_dn100)

        assert dims_dn50.get("DN") == "50"
        assert dims_dn100.get("DN") == "100"

        item_dn50 = self._make_item("V1", title_dn50, UnitClass.COUNT, default_pointer, dimensions=dims_dn50)
        item_dn100 = self._make_item("V2", title_dn100, UnitClass.COUNT, default_pointer, dimensions=dims_dn100)

        score = Decimal(str(CascadeMatcher.calculate_score(item_dn50, item_dn100)))
        assert score == Decimal("0.0")

    def test_mr5_unit_gating_annihilation(self, default_pointer: CellPointer) -> None:
        """MR-5 (Разрушение шлюза единиц): несовместимость классов ОКЕИ обнуляет скор (I_unit = 0 -> S_total = 0.0)."""
        title = "Асфальтобетонная смесь плотная мелкозернистая"

        item_area = self._make_item("U1", title, UnitClass.AREA, default_pointer)
        item_vol = self._make_item("U2", title, UnitClass.VOLUME, default_pointer)

        score = Decimal(str(CascadeMatcher.calculate_score(item_area, item_vol)))
        assert score == Decimal("0.0")