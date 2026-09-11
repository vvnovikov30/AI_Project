"""Корневой конфигурационный модуль pytest.

Определяет глобальные профили Hypothesis, worker-изолированные соединения
с in-memory SQLite, Pydantic v2 модели со строгим Zero Float Policy и типовые
фикстуры строительных конструктивных элементов.
"""

from collections.abc import Generator
from decimal import Decimal
import os
import sqlite3
from typing import Any

from hypothesis import Phase, settings
from pydantic import BaseModel, ConfigDict, Field, field_validator
import pytest

from models import UnitClass

# ----------------------------------------------------------------------
# 1. Конфигурация профилей Hypothesis (dev / ci)
# ----------------------------------------------------------------------

settings.register_profile(
    "dev",
    max_examples=25,
    deadline=400,
    phases=[Phase.explicit, Phase.reuse, Phase.generate],
)

settings.register_profile(
    "ci",
    max_examples=300,
    deadline=None,
    phases=[
        Phase.explicit,
        Phase.reuse,
        Phase.generate,
        Phase.target,
        Phase.shrink,
    ],
)

_selected_profile: str = os.getenv("HYPOTHESIS_PROFILE", "dev")
settings.load_profile(_selected_profile)

# ----------------------------------------------------------------------
# 2. PEP 695 Псевдонимы типов и Pydantic v2 DTO для тестов
# ----------------------------------------------------------------------

type Quantity = Decimal
type Multiplier = Decimal
type Price = Decimal
type DimensionDict = dict[str, str]


class StrictCellPointerSchema(BaseModel):
    """Строгая схема валидации аудиторского следа ячейки."""

    model_config = ConfigDict(
        frozen=True,
        strict=True,
        extra="forbid",
    )

    doc_type: str = Field(min_length=3, max_length=32)
    file_name: str = Field(min_length=1)
    sheet_name: str = Field(min_length=1)
    row_idx: int = Field(ge=1)
    col_letter: str = Field(min_length=1, max_length=4)
    coordinate: str = Field(min_length=2)


class StrictNormalizedItemSchema(BaseModel):
    """Строгая схема валидированного строительного ресурса с Zero Float Policy."""

    model_config = ConfigDict(
        frozen=True,
        strict=True,
        extra="forbid",
        arbitrary_types_allowed=True,
    )

    raw_id: str = Field(min_length=1)
    code: str | None = None
    raw_title: str = Field(min_length=1)
    clean_title: str = Field(min_length=1)
    raw_unit: str = Field(min_length=1)
    canonical_unit: UnitClass
    unit_multiplier: Multiplier
    volume: Quantity
    dimensions: DimensionDict
    context_section: str
    pointer: StrictCellPointerSchema
    is_resource: bool = False
    parent_id: str | None = None

    @field_validator("volume", "unit_multiplier", mode="before")
    @classmethod
    def reject_float_types(cls, value: Any) -> Decimal:
        """Перехватывает и блокирует передачу аппаратного float."""
        if isinstance(value, float):
            raise TypeError(
                f"Zero Float Policy Violation: передано значение типа float ({value!r}). "
                "Используйте строго str или Decimal."
            )
        if isinstance(value, Decimal):
            return value
        if isinstance(value, str | int):
            return Decimal(str(value))
        raise TypeError(f"Недопустимый тип для числового сметного поля: {type(value)}")


# ----------------------------------------------------------------------
# 3. Фикстуры SQLite с изоляцией по pytest-xdist worker_id
# ----------------------------------------------------------------------


@pytest.fixture(scope="session")
def sqlite_worker_uri(request: pytest.FixtureRequest) -> str:
    """Генерирует уникальный URI разделяемой памяти SQLite для параллельного воркера."""
    # Безопасное получение worker_id даже без pytest-xdist
    worker_id = getattr(request.config, "workerinput", {}).get("workerid", "master")
    clean_worker_id = worker_id.replace("-", "_")
    return f"file:memdb_{clean_worker_id}?mode=memory&cache=shared"


@pytest.fixture
def sqlite_conn(sqlite_worker_uri: str) -> Generator[sqlite3.Connection, None, None]:
    """Изолированное соединение с SQLite БД аудита.
    
    Все числовые поля объемов и коэффициентов хранятся строго в типе TEXT,
    чтобы исключить неявное приведение к IEEE 754 float на уровне СУБД.
    """
    conn = sqlite3.connect(sqlite_worker_uri, uri=True)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_reconciliation_buffer (
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
    conn.commit()

    try:
        yield conn
    finally:
        cursor.execute("DELETE FROM audit_reconciliation_buffer")
        conn.commit()
        conn.close()


# ----------------------------------------------------------------------
# 4. Базовые фикстуры аудиторских указателей и строительных элементов
# ----------------------------------------------------------------------


@pytest.fixture
def base_excel_pointer() -> StrictCellPointerSchema:
    """Указатель на ячейку в сметном расчете Excel."""
    return StrictCellPointerSchema(
        doc_type="LSR",
        file_name="LSR_01_Foundation.xlsx",
        sheet_name="Раздел 1",
        row_idx=24,
        col_letter="C",
        coordinate="C24",
    )


@pytest.fixture
def base_pdf_pointer() -> StrictCellPointerSchema:
    """Указатель на строку в проектной спецификации PDF."""
    return StrictCellPointerSchema(
        doc_type="SPECIFICATION",
        file_name="02_2026_KJ_Spec.pdf",
        sheet_name="Page 12",
        row_idx=15,
        col_letter="D",
        coordinate="p.12 [row 15]",
    )


@pytest.fixture
def sample_cable_item(base_excel_pointer: StrictCellPointerSchema) -> StrictNormalizedItemSchema:
    """Кабельная продукция с нормативным коэффициентом отхода 6% (ГЭСНм 08)."""
    return StrictNormalizedItemSchema(
        raw_id="LSR_101",
        code="ФССЦ-21.1.06.04-0012",
        raw_title="Кабель силовой ВВГнг(A)-LS 3х1.5",
        clean_title="кабель силовой ввгнг a ls 3х1.5",
        raw_unit="100 м",
        canonical_unit=UnitClass.LENGTH,
        unit_multiplier=Decimal("100.0"),
        volume=Decimal("150.0"),
        dimensions={"CORES": "3", "SECTION": "1.5"},
        context_section="Раздел ЭОМ",
        pointer=base_excel_pointer,
        is_resource=True,
        parent_id="LSR_100",
    )


@pytest.fixture
def sample_rebar_item(base_pdf_pointer: StrictCellPointerSchema) -> StrictNormalizedItemSchema:
    """Стержневая арматура класса А500С с отходом 5% (ГЭСН 06)."""
    return StrictNormalizedItemSchema(
        raw_id="SPEC_201",
        code="ГОСТ 34028-2016",
        raw_title="Арматура горячекатаная гладкая и периодического профиля d12 А500С",
        clean_title="арматура горячекатаная гладкая периодического профиля d12 а500с",
        raw_unit="т",
        canonical_unit=UnitClass.MASS,
        unit_multiplier=Decimal("1.0"),
        volume=Decimal("14.500"),
        dimensions={"REBAR_DIA": "12", "CLASS": "А500С"},
        context_section="Раздел КЖ",
        pointer=base_pdf_pointer,
        is_resource=False,
        parent_id=None,
    )


@pytest.fixture
def sample_concrete_item(base_excel_pointer: StrictCellPointerSchema) -> StrictNormalizedItemSchema:
    """Бетонная смесь класса В25 с отходом 2% (ГЭСН 06)."""
    return StrictNormalizedItemSchema(
        raw_id="LSR_301",
        code="ФССЦ-04.1.02.05-0008",
        raw_title="Бетон тяжелый товарный БСГ В25 (М350) F150 W6",
        clean_title="бетон тяжелый товарный бсг в25 м350 f150 w6",
        raw_unit="м3",
        canonical_unit=UnitClass.VOLUME,
        unit_multiplier=Decimal("1.0"),
        volume=Decimal("245.500"),
        dimensions={"CONCRETE": "B25", "F": "150", "W": "6"},
        context_section="Раздел КЖ",
        pointer=base_excel_pointer,
        is_resource=False,
        parent_id=None,
    )


@pytest.fixture
def sample_pipe_valve_item(base_pdf_pointer: StrictCellPointerSchema) -> StrictNormalizedItemSchema:
    """Трубопроводная арматура с жестким шлюзом размерности DN50 PN16."""
    return StrictNormalizedItemSchema(
        raw_id="SPEC_401",
        code=None,
        raw_title="Задвижка клиновая стальная фланцевая 30с41нж DN50 PN16",
        clean_title="задвижка клиновая стальная фланцевая 30с41нж dn50 pn16",
        raw_unit="шт",
        canonical_unit=UnitClass.COUNT,
        unit_multiplier=Decimal("1.0"),
        volume=Decimal("8.0"),
        dimensions={"DN": "50", "PN": "16"},
        context_section="Раздел ОВ",
        pointer=base_pdf_pointer,
        is_resource=False,
        parent_id=None,
    )


@pytest.fixture
def standard_tolerances() -> dict[str, Decimal]:
    """Регламентированные числовые допуски сравнения объемов."""
    return {
        "abs_tolerance": Decimal("0.001"),
        "rel_tolerance": Decimal("0.01"),
        "waste_rebar": Decimal("0.05"),
        "waste_concrete": Decimal("0.02"),
        "waste_cable": Decimal("0.06"),
    }