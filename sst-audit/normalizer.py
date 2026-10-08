import re
from decimal import Decimal
from typing import Tuple, Dict, Any
from models import UnitClass


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
            u_raw in ("шт", "шт.", "штук", "штука", "1 шт", "1 шт.", "устройство", "1 устройство", "компл", "компл.", "комплект", "wm")
            or u_no_dots in ("шт", "1 шт", "устройство", "1 устройство", "компл", "комплект")
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