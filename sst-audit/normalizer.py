import re
from decimal import Decimal
from typing import Tuple, Dict, Any
from models import UnitClass


class UnitNormalizer:
    """Канонизация единиц измерения ОКЕИ и расчет коэффициентов кратности."""

    @staticmethod
    def normalize(raw_unit: str, base_qty: Decimal) -> Tuple[UnitClass, Decimal, Decimal]:
        u_lower = raw_unit.lower().strip()
        
        # Area
        if u_lower in ("м2", "кв. м", "м.кв."):
            return UnitClass.AREA, base_qty, Decimal("1.0")
        if u_lower == "100 м2":
            return UnitClass.AREA, base_qty * Decimal("100.0"), Decimal("100.0")
        if u_lower == "1000 м2":
            return UnitClass.AREA, base_qty * Decimal("1000.0"), Decimal("1000.0")
            
        # Volume
        if u_lower in ("м3", "куб. м", "м.куб."):
            return UnitClass.VOLUME, base_qty, Decimal("1.0")
        if u_lower == "10 м3":
            return UnitClass.VOLUME, base_qty * Decimal("10.0"), Decimal("10.0")
        if u_lower == "100 м3":
            return UnitClass.VOLUME, base_qty * Decimal("100.0"), Decimal("100.0")
            
        # Mass
        if u_lower in ("кг", "килограмм"):
            return UnitClass.MASS, base_qty, Decimal("1.0")
        if u_lower in ("т", "тн"):
            return UnitClass.MASS, base_qty * Decimal("1000.0"), Decimal("1000.0")
            
        # Count
        if u_lower in ("шт", "штук"):
            return UnitClass.COUNT, base_qty, Decimal("1.0")
        if u_lower == "100 шт":
            return UnitClass.COUNT, base_qty * Decimal("100.0"), Decimal("100.0")
        if u_lower == "1000 шт" or u_lower == "тыс. шт":
            return UnitClass.COUNT, base_qty * Decimal("1000.0"), Decimal("1000.0")
            
        return UnitClass.UNKNOWN, base_qty, Decimal("1.0")


class DimensionLockExtractor:
    """Извлечение технических параметров (DN, PN, класс бетона, диаметр арматуры) через RegEx."""

    @staticmethod
    def extract(text: str) -> Dict[str, str]:
        """Извлекает ключевые технические параметры (DN, PN, класс бетона, диаметр арматуры) через RegEx."""
        dims: Dict[str, str] = {}
        upper_text = text.upper()

        dn_match = re.search(r"(?:DN|ДУ)\s*(\d+)", upper_text)
        if dn_match:
            dims["DN"] = dn_match.group(1)

        pn_match = re.search(r"(?:PN|РУ)\s*(\d+)", upper_text)
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

        # Учет верхнего регистра для знаков ø (Ø) и ф (Ф) после вызова .upper()
        rebar_match = re.search(r"(?:[øØфФdD]|диаметр)\s*(\d+(?:\.\d+)?)", upper_text)
        if rebar_match:
            dims["REBAR_DIA"] = rebar_match.group(1)

        return dims