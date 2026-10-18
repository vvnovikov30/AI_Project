import re
from decimal import Decimal
from typing import Any, Dict, Final, Optional, Tuple
from models import DimensionFamily, UnitClass, UnitDimension


def _canonical_unit_key(raw_unit: Optional[str]) -> Optional[str]:
    """Канонический ключ единицы для словаря размерностей.

    Приводит написание к нижнему регистру, убирает все пробелы (включая
    неразрывные) и точки: ``'пог. м' -> 'погм'``, ``'кв. м' -> 'квм'``,
    ``'100 м2' -> '100м2'``. Символы ``²``/``³`` нормализуются в ``2``/``3``.
    """
    if raw_unit is None:
        return None
    text = str(raw_unit).strip().lower()
    if not text:
        return None
    text = text.replace("\xa0", " ").replace("\u2009", " ").replace("\u202f", " ")
    text = text.replace("²", "2").replace("³", "3").replace("ё", "е")
    text = re.sub(r"\s+", "", text)
    text = text.replace(".", "")
    return text or None


def _dimension(
    family: DimensionFamily, base_unit: str, multiplier: str = "1"
) -> UnitDimension:
    """Фабрика описания единицы: кратность передается строкой в ``Decimal``."""
    return UnitDimension(
        family=family,
        base_unit=base_unit,
        multiplier=Decimal(multiplier),
    )


# Словарь физических размерностей (SSOT Unit Gate).
# Ключи — результат ``_canonical_unit_key``, кратность — относительно base_unit.
_UNIT_DIMENSIONS: Final[Dict[str, UnitDimension]] = {
    # --- Длина (базовая единица — метр) ---
    "м": _dimension(DimensionFamily.LENGTH, "м"),
    "метр": _dimension(DimensionFamily.LENGTH, "м"),
    "метров": _dimension(DimensionFamily.LENGTH, "м"),
    "погм": _dimension(DimensionFamily.LENGTH, "м"),
    "пм": _dimension(DimensionFamily.LENGTH, "м"),
    "мп": _dimension(DimensionFamily.LENGTH, "м"),
    "м/п": _dimension(DimensionFamily.LENGTH, "м"),
    "100м": _dimension(DimensionFamily.LENGTH, "м", "100"),
    "1000м": _dimension(DimensionFamily.LENGTH, "м", "1000"),
    "км": _dimension(DimensionFamily.LENGTH, "м", "1000"),
    "мм": _dimension(DimensionFamily.LENGTH, "м", "0.001"),
    # --- Площадь ---
    "м2": _dimension(DimensionFamily.AREA, "м2"),
    "квм": _dimension(DimensionFamily.AREA, "м2"),
    "мкв": _dimension(DimensionFamily.AREA, "м2"),
    "100м2": _dimension(DimensionFamily.AREA, "м2", "100"),
    "100квм": _dimension(DimensionFamily.AREA, "м2", "100"),
    "1000м2": _dimension(DimensionFamily.AREA, "м2", "1000"),
    "1000квм": _dimension(DimensionFamily.AREA, "м2", "1000"),
    # --- Объем ---
    "м3": _dimension(DimensionFamily.VOLUME, "м3"),
    "кубм": _dimension(DimensionFamily.VOLUME, "м3"),
    "мкуб": _dimension(DimensionFamily.VOLUME, "м3"),
    "10м3": _dimension(DimensionFamily.VOLUME, "м3", "10"),
    "10кубм": _dimension(DimensionFamily.VOLUME, "м3", "10"),
    "100м3": _dimension(DimensionFamily.VOLUME, "м3", "100"),
    "100кубм": _dimension(DimensionFamily.VOLUME, "м3", "100"),
    "л": _dimension(DimensionFamily.VOLUME, "м3", "0.001"),
    "литр": _dimension(DimensionFamily.VOLUME, "м3", "0.001"),
    "литров": _dimension(DimensionFamily.VOLUME, "м3", "0.001"),
    # --- Масса (базовая единица — кг) ---
    "кг": _dimension(DimensionFamily.MASS, "кг"),
    "килограмм": _dimension(DimensionFamily.MASS, "кг"),
    "килограммов": _dimension(DimensionFamily.MASS, "кг"),
    "г": _dimension(DimensionFamily.MASS, "кг", "0.001"),
    "грамм": _dimension(DimensionFamily.MASS, "кг", "0.001"),
    "граммов": _dimension(DimensionFamily.MASS, "кг", "0.001"),
    "т": _dimension(DimensionFamily.MASS, "кг", "1000"),
    "тн": _dimension(DimensionFamily.MASS, "кг", "1000"),
    "тонн": _dimension(DimensionFamily.MASS, "кг", "1000"),
    "тонна": _dimension(DimensionFamily.MASS, "кг", "1000"),
    "тонны": _dimension(DimensionFamily.MASS, "кг", "1000"),
    # --- Штучные изделия (COUNT) ---
    "шт": _dimension(DimensionFamily.COUNT, "шт"),
    "штук": _dimension(DimensionFamily.COUNT, "шт"),
    "штука": _dimension(DimensionFamily.COUNT, "шт"),
    "штуки": _dimension(DimensionFamily.COUNT, "шт"),
    "1шт": _dimension(DimensionFamily.COUNT, "шт"),
    "10шт": _dimension(DimensionFamily.COUNT, "шт", "10"),
    "100шт": _dimension(DimensionFamily.COUNT, "шт", "100"),
    "100перемычек": _dimension(DimensionFamily.COUNT, "шт", "100"),
    "1000шт": _dimension(DimensionFamily.COUNT, "шт", "1000"),
    "тысшт": _dimension(DimensionFamily.COUNT, "шт", "1000"),
    "тысштук": _dimension(DimensionFamily.COUNT, "шт", "1000"),
    "компл": _dimension(DimensionFamily.COUNT, "компл"),
    "комплект": _dimension(DimensionFamily.COUNT, "компл"),
    "комплектов": _dimension(DimensionFamily.COUNT, "компл"),
    "комп": _dimension(DimensionFamily.COUNT, "компл"),
    "упак": _dimension(DimensionFamily.COUNT, "упак"),
    "упаковка": _dimension(DimensionFamily.COUNT, "упак"),
    "упаковок": _dimension(DimensionFamily.COUNT, "упак"),
    "уп": _dimension(DimensionFamily.COUNT, "упак"),
    "рулон": _dimension(DimensionFamily.COUNT, "рулон"),
    "рулонов": _dimension(DimensionFamily.COUNT, "рулон"),
    "рул": _dimension(DimensionFamily.COUNT, "рулон"),
    "бобина": _dimension(DimensionFamily.COUNT, "бобина"),
    "бобин": _dimension(DimensionFamily.COUNT, "бобина"),
    "боб": _dimension(DimensionFamily.COUNT, "бобина"),
    "пара": _dimension(DimensionFamily.COUNT, "пара"),
    "пары": _dimension(DimensionFamily.COUNT, "пара"),
    "пар": _dimension(DimensionFamily.COUNT, "пара"),
    # --- Электромонтаж / работы (ELECTRICAL) ---
    "точка": _dimension(DimensionFamily.ELECTRICAL, "точка"),
    "точки": _dimension(DimensionFamily.ELECTRICAL, "точка"),
    "точек": _dimension(DimensionFamily.ELECTRICAL, "точка"),
    "ввод": _dimension(DimensionFamily.ELECTRICAL, "ввод"),
    "вводов": _dimension(DimensionFamily.ELECTRICAL, "ввод"),
    "панель": _dimension(DimensionFamily.ELECTRICAL, "панель"),
    "панелей": _dimension(DimensionFamily.ELECTRICAL, "панель"),
    "шкаф": _dimension(DimensionFamily.ELECTRICAL, "шкаф"),
    "шкафов": _dimension(DimensionFamily.ELECTRICAL, "шкаф"),
    "система": _dimension(DimensionFamily.ELECTRICAL, "система"),
    "системы": _dimension(DimensionFamily.ELECTRICAL, "система"),
    "систем": _dimension(DimensionFamily.ELECTRICAL, "система"),
}

# Допустимые межсемейственные связки (комплекты/системы против штук).
_CROSS_FAMILY_COMPATIBLE: Final[frozenset[frozenset[DimensionFamily]]] = frozenset(
    {
        frozenset({DimensionFamily.COUNT, DimensionFamily.ELECTRICAL}),
    }
)

# Веса предиката R05_UNT (Decimal — Zero Float Policy).
_UNIT_GATE_EXACT_WEIGHT: Final[Decimal] = Decimal("1.0")
_UNIT_GATE_CROSS_WEIGHT: Final[Decimal] = Decimal("0.7")
_UNIT_GATE_SKIP_WEIGHT: Final[Decimal] = Decimal("0.5")
_UNIT_GATE_BLOCK_WEIGHT: Final[Decimal] = Decimal("0.0")


class UnitNormalizer:
    """Канонизация единиц измерения ОКЕИ и расчет коэффициентов кратности."""

    @staticmethod
    def normalize(raw_unit: str, base_qty: Decimal) -> Tuple[UnitClass, Decimal, Decimal]:
        if not raw_unit:
            return UnitClass.UNKNOWN, base_qty, Decimal("1.0")

        # 1. Первичная нормализация: нижний регистр, схлопывание пробелов и неразрывных пробелов
        u_raw = re.sub(r"\s+", " ", str(raw_unit).replace("\xa0", " ")).lower().strip()

        # 2. Вспомогательная форма без точек для гибкого сопоставления ('10 шт.' -> '10 шт')
        u_no_dots = re.sub(r"\s+", " ", u_raw.replace(".", "")).strip()

        # --- Площадь (AREA) ---
        if u_raw in ("м2", "кв. м", "кв.м", "м.кв.") or u_no_dots in ("кв м", "м кв"):
            return UnitClass.AREA, base_qty, Decimal("1.0")
        if u_raw in ("100 м2", "100м2", "100 кв. м", "100 кв.м") or u_no_dots in ("100 кв м", "100кв м"):
            return UnitClass.AREA, base_qty * Decimal("100.0"), Decimal("100.0")
        if u_raw in ("1000 м2", "1000м2", "1000 кв. м") or u_no_dots == "1000 кв м":
            return UnitClass.AREA, base_qty * Decimal("1000.0"), Decimal("1000.0")

        # --- Объем (VOLUME) ---
        if u_raw in ("м3", "куб. м", "куб.м", "м.куб.") or u_no_dots in ("куб м", "м куб"):
            return UnitClass.VOLUME, base_qty, Decimal("1.0")
        if u_raw in ("10 м3", "10м3", "10 куб. м") or u_no_dots == "10 куб м":
            return UnitClass.VOLUME, base_qty * Decimal("10.0"), Decimal("10.0")
        if u_raw in ("100 м3", "100м3", "100 куб. м") or u_no_dots == "100 куб м":
            return UnitClass.VOLUME, base_qty * Decimal("100.0"), Decimal("100.0")

        # --- Масса (MASS) ---
        if u_raw in ("кг", "килограмм") or u_no_dots == "кг":
            return UnitClass.MASS, base_qty, Decimal("1.0")
        if u_raw in ("т", "тн", "тонн", "тонна") or u_no_dots in ("т", "тн"):
            return UnitClass.MASS, base_qty * Decimal("1000.0"), Decimal("1000.0")

        # --- Длина (LENGTH) ---
        if u_raw in ("м", "м.", "м.п.", "пог. м", "пог.м") or u_no_dots in ("м", "мп", "пог м"):
            return UnitClass.LENGTH, base_qty, Decimal("1.0")
        if u_raw in ("100 м", "100м", "100 м.", "100 м.п.", "100 пог. м") or u_no_dots in ("100 м", "100м", "100 мп", "100 пог м"):
            return UnitClass.LENGTH, base_qty * Decimal("100.0"), Decimal("100.0")
        if u_raw in ("км", "1000 м", "1000м") or u_no_dots in ("км", "1000 м", "1000м"):
            return UnitClass.LENGTH, base_qty * Decimal("1000.0"), Decimal("1000.0")

        # --- Количество / Штучные изделия (COUNT) ---
        # Базовая единица (1 шт / компл / устройство / wm)
        if (
            u_raw in ("шт", "шт.", "штук", "штука", "1 шт", "1 шт.", "устройство", "1 устройство", "компл", "компл.", "комплект", "комп", "комп.", "wm")
            or u_no_dots in ("шт", "1 шт", "устройство", "1 устройство", "компл", "комплект", "комп")
        ):
            return UnitClass.COUNT, base_qty, Decimal("1.0")

        # 10 шт. (монтажные расценки ТСН/ГЭСН)
        if u_raw in ("10 шт", "10 шт.", "10шт") or u_no_dots in ("10 шт", "10шт"):
            return UnitClass.COUNT, base_qty * Decimal("10.0"), Decimal("10.0")

        # 100 шт. / 100 перемычек
        if (
            u_raw in ("100 шт", "100 шт.", "100шт", "100 перемычек")
            or u_no_dots in ("100 шт", "100шт", "100 перемычек")
        ):
            return UnitClass.COUNT, base_qty * Decimal("100.0"), Decimal("100.0")

        # 1000 шт. / тыс. шт
        if (
            u_raw in ("1000 шт", "1000 шт.", "1000шт", "тыс. шт", "тыс.шт", "тыс шт")
            or u_no_dots in ("1000 шт", "1000шт", "тыс шт")
        ):
            return UnitClass.COUNT, base_qty * Decimal("1000.0"), Decimal("1000.0")

        return UnitClass.UNKNOWN, base_qty, Decimal("1.0")

    # ------------------------------------------------------------------
    # R05_UNT: Unit Gate (физические размерности)
    # ------------------------------------------------------------------

    @staticmethod
    def classify_unit(raw_unit: Optional[str]) -> Optional[UnitDimension]:
        """Классифицирует единицу в каноническое описание размерности.

        :return: ``UnitDimension`` либо ``None``, если единица не определена
            (None/пустая строка) или отсутствует в словаре физических единиц.
        """
        key = _canonical_unit_key(raw_unit)
        if key is None:
            return None
        return _UNIT_DIMENSIONS.get(key)

    @staticmethod
    def dimension_family(raw_unit: Optional[str]) -> Optional[DimensionFamily]:
        """Возвращает семейство размерности единицы (или ``None``)."""
        dimension = UnitNormalizer.classify_unit(raw_unit)
        return dimension.family if dimension is not None else None

    @staticmethod
    def is_dimension_compatible(
        unit_rd: Optional[str], unit_smeta: Optional[str]
    ) -> bool:
        """Проверяет совместимость физических размерностей пары единиц.

        Правила:
        * одна/обе единицы не определены либо не распознаны -> ``True``
          (мягкий пропуск в HITL, жесткая блокировка не выполняется);
        * единицы одного семейства -> ``True`` (включая кратности);
        * межсемейственная связка COUNT <-> ELECTRICAL -> ``True``
          (комплект/система/точка/шкаф считаются штучными позициями);
        * остальные комбинации семейств -> ``False``
          (AREA vs VOLUME, LENGTH vs COUNT, MASS vs COUNT и т.д.).
        """
        dim_rd = UnitNormalizer.classify_unit(unit_rd)
        dim_smeta = UnitNormalizer.classify_unit(unit_smeta)
        if dim_rd is None or dim_smeta is None:
            return True
        if dim_rd.family is dim_smeta.family:
            return True
        return frozenset({dim_rd.family, dim_smeta.family}) in _CROSS_FAMILY_COMPATIBLE

    @staticmethod
    def evaluate_unit_gate(
        unit_rd: Optional[str], unit_smeta: Optional[str]
    ) -> Tuple[bool, Decimal, str]:
        """Предикат R05_UNT: вердикт, вес и причина совместимости единиц.

        :return: кортеж ``(is_compatible, weight_score, reason)``.

        * несовместимые семейства -> ``(False, 0.0,
          "INCOMPATIBLE_DIMENSIONS: {FAM1} vs {FAM2}")``;
        * одна/обе единицы не определены -> ``(True, 0.5, "..._SKIPPED")``;
        * совпадение с учетом кратности -> ``(True, 1.0,
          "EXACT_OR_SCALED_MATCH")``;
        * допустимая конверсия (шт <-> компл/система, т <-> кг) ->
          ``(True, 0.7, "COMPATIBLE_CROSS_UNIT")``.

        ``weight_score`` возвращается ``Decimal`` в соответствии с Zero Float
        Policy проекта; сравнение с ``1.0``/``0.0`` работает штатно.
        """
        dim_rd = UnitNormalizer.classify_unit(unit_rd)
        dim_smeta = UnitNormalizer.classify_unit(unit_smeta)

        if dim_rd is None and dim_smeta is None:
            return True, _UNIT_GATE_SKIP_WEIGHT, "UNDEFINED_UNITS_SKIPPED"
        if dim_rd is None or dim_smeta is None:
            return True, _UNIT_GATE_SKIP_WEIGHT, "UNKNOWN_UNIT_SKIPPED"

        if dim_rd.family is not dim_smeta.family:
            pair = frozenset({dim_rd.family, dim_smeta.family})
            if pair in _CROSS_FAMILY_COMPATIBLE:
                return True, _UNIT_GATE_CROSS_WEIGHT, "COMPATIBLE_CROSS_UNIT"
            return (
                False,
                _UNIT_GATE_BLOCK_WEIGHT,
                f"INCOMPATIBLE_DIMENSIONS: {dim_rd.family.name} vs {dim_smeta.family.name}",
            )

        if dim_rd.base_unit == dim_smeta.base_unit:
            return True, _UNIT_GATE_EXACT_WEIGHT, "EXACT_OR_SCALED_MATCH"
        return True, _UNIT_GATE_CROSS_WEIGHT, "COMPATIBLE_CROSS_UNIT"


def is_dimension_compatible(unit_rd: Optional[str], unit_smeta: Optional[str]) -> bool:
    """Модульная обертка над :meth:`UnitNormalizer.is_dimension_compatible`."""
    return UnitNormalizer.is_dimension_compatible(unit_rd, unit_smeta)


def evaluate_unit_gate(
    unit_rd: Optional[str], unit_smeta: Optional[str]
) -> Tuple[bool, Decimal, str]:
    """Модульная обертка над :meth:`UnitNormalizer.evaluate_unit_gate`."""
    return UnitNormalizer.evaluate_unit_gate(unit_rd, unit_smeta)


class DimensionLockExtractor:
    """Извлечение технических параметров (DN, PN, класс бетона, диаметр арматуры/труб) через RegEx."""

    @staticmethod
    def extract(text: str) -> Dict[str, str]:
        dims: Dict[str, str] = {}
        upper_text = text.upper()

        dn_match = re.search(r"(?:DN|ДУ)\s*[-=:]?\s*(\d+)", upper_text)
        if dn_match:
            dims["DN"] = dn_match.group(1)

        pn_match = re.search(r"(?:PN|РУ)\s*[-=:]?\s*(\d+)", upper_text)
        if pn_match:
            dims["PN"] = pn_match.group(1)

        # Строгий приоритет класса бетона (B / В) над маркой (М)
        concrete_match = re.search(r"\b[BВ]\s*(\d+(?:\.\d+)?)\b", upper_text)
        if concrete_match:
            num = concrete_match.group(1)
            dims["CONCRETE"] = f"B{num}"
        else:
            m_match = re.search(r"\bМ\s*(\d+)\b", upper_text)
            if m_match:
                dims["CONCRETE"] = f"М{m_match.group(1)}"

        # Диаметры: арматура (ø, ф), трубы и гофротрубы (d-20, диам 20, диаметр 20)
        rebar_match = re.search(r"(?:[ØøФф]|(?:\b(?:D|ДИАМЕТР|ДИАМ)\b))\s*[-=:]?\s*(\d+(?:\.\d+)?)", upper_text)
        if rebar_match:
            dims["REBAR_DIA"] = rebar_match.group(1)

        return dims