from typing import Dict, Any
from rapidfuzz import fuzz
from models import NormalizedItem, UnitClass

class CascadeMatcher:
    """Каскадный матчер с аппаратом фильтрации размерностей и параметров."""

    W_CODE = 0.45
    W_FUZZY = 0.40
    W_CONTEXT = 0.15

    @staticmethod
    def _is_dimension_compatible(dim_lsr: Dict[str, Any], dim_spec: Dict[str, Any]) -> bool:
        """Dimensional Lock: блокирует несовпадающие диаметры, давления и классы."""
        for key in dim_lsr:
            if key in dim_spec and dim_lsr[key] != dim_spec[key]:
                return False
        return True

    @classmethod
    def calculate_score(cls, lsr: NormalizedItem, spec: NormalizedItem) -> float:
        # 1. Unit Gating: сброс при несовместимости физических групп
        if lsr.canonical_unit != spec.canonical_unit or lsr.canonical_unit == UnitClass.UNKNOWN:
            return 0.0

        # 2. Dimensional Lock
        if not cls._is_dimension_compatible(lsr.dimensions, spec.dimensions):
            return 0.0

        # 3. Code match (ГОСТ, шифры материалов, артикулы)
        has_code = bool(lsr.code or spec.code)
        s_code = 0.0
        if lsr.code and spec.code and (lsr.code == spec.code):
            s_code = 100.0
        elif lsr.code and lsr.code in spec.clean_title:
            s_code = 80.0
        elif spec.code and spec.code in lsr.clean_title:
            s_code = 80.0

        # 4. Fuzzy lexical match (устойчив и к перестановкам, и к добавке стоп-слов)
        s_sort = float(fuzz.token_sort_ratio(lsr.clean_title, spec.clean_title))
        s_set = float(fuzz.token_set_ratio(lsr.clean_title, spec.clean_title))
        s_fuzzy = max(s_sort, s_set)

        # 5. Context match (раздел / шифр тома)
        s_context = 100.0 if lsr.context_section == spec.context_section else 0.0

        # Динамическое взвешивание: если шифр не задан у обеих сторон, вес отдается лексике
        if not has_code:
            return (0.85 * s_fuzzy) + (0.15 * s_context)

        return (cls.W_CODE * s_code) + (cls.W_FUZZY * s_fuzzy) + (cls.W_CONTEXT * s_context)