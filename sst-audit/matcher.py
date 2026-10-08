import re
from typing import Dict, Any, Optional
from rapidfuzz import fuzz
from models import NormalizedItem, UnitClass


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

    @classmethod
    def calculate_score(cls, lsr: NormalizedItem, spec: NormalizedItem) -> float:
        # 1. Unit Gating: мгновенная блокировка при несовместимости физических групп единиц
        if lsr.canonical_unit != spec.canonical_unit or lsr.canonical_unit == UnitClass.UNKNOWN:
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