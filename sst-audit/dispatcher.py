"""Автоматическое сопоставление томов РД (PDF) с листами локальной сметы.

Диспетчер решает одну задачу: определить, какой PDF-том рабочей документации
относится к какому листу ``ЛС ...`` многолистовой сметы. Сопоставление
выполняется по названию файла (аббревиатура раздела, шифр, дата) и заголовку
сметы в ячейках ``A8``/``A9``.

Модуль использует только открытые библиотеки: ``openpyxl`` (MIT) и
``RapidFuzz`` (MIT). Промежуточные JSON/XML-дампы не создаются: результат
возвращается вызывающему коду, а SSOT остается SQLite.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Final

import openpyxl
from rapidfuzz import fuzz

# Аббревиатуры и синонимы разделов проектной документации.
# Значения с шаблоном ``NN-NN-NN`` являются шифрами разделов сметы.
MAPPING_RULES: Final[dict[str, list[str]]] = {
    "АПС": ["пожарная сигнализация", "автоматическая пожарная", "02-01-08"],
    "СОУЭ": ["оповещени", "эвакуаци", "02-01-09"],
    "РТС": ["радиотрансляцион", "радиофикац", "02-01-11"],
    "СКТ": ["кабельного телевидения", "телевиден", "02-01-12"],
    "СКС": ["структурированная кабельная", "02-01-13"],
    "СОТ": ["охранного теленаблюдения", "видеонаблюдени", "02-01-14"],
    "СКУД": ["контроля и управления доступом", "02-01-15"],
    "СДС": ["двусторонней связи", "02-01-16"],
    "ЭОМ": ["электроосвещени", "электроснабжени", "силовое электро", "02-01-03"],
    "ОВ": ["отоплени", "вентиляци", "02-01-06"],
    "ВК": ["водоснабжени", "водоотведени", "канализаци", "02-01-04", "02-01-05"],
    "АР": ["архитектурные решения", "02-01-01"],
}

# Минимальная уверенность, при которой пара PDF/лист считается найденной.
MIN_CONFIDENCE: Final[float] = 50.0
SECTION_SHEET_PREFIX: Final[str] = "ЛС "
TITLE_CELLS: Final[tuple[str, ...]] = ("A8", "A9")
_CODE_PATTERN: Final[re.Pattern[str]] = re.compile(r"^\d{2}-\d{2}-\d{2}$")


@dataclass(frozen=True)
class SheetCandidate:
    """Лист сметы, участвующий в сопоставлении."""

    name: str
    title: str

    @property
    def searchable_text(self) -> str:
        """Название листа и заголовок сметы, приведенные к нижнему регистру."""
        return _normalize(f"{self.name} {self.title}")


@dataclass(frozen=True)
class _ScoredPair:
    """Пара PDF/лист с оценкой сходства."""

    score: float
    pdf_index: int
    sheet_index: int


def _normalize(text: str) -> str:
    """Приводит текст к сравнимому виду, сохраняя шифры разделов и дефисы."""
    lowered = text.lower().replace("ё", "е")
    # Дефисы внутри шифров (02-01-14) значимы, остальная пунктуация — шум.
    lowered = re.sub(r"\s*-\s*", "-", lowered)
    lowered = re.sub(r"[^\w\-]+", " ", lowered, flags=re.UNICODE)
    return re.sub(r"\s+", " ", lowered).strip()


def _has_token(text: str, token: str) -> bool:
    """Проверяет вхождение аббревиатуры как отдельного токена."""
    pattern = rf"(?<![\w]){re.escape(_normalize(token))}(?![\w])"
    return re.search(pattern, text) is not None


def _rule_evidence(pdf_text: str, sheet_text: str) -> float:
    """Вычисляет смысловую поддержку пары по словарю разделов.

    Оценка 0-100: аббревиатура и/или шифр в имени PDF должны подтверждаться
    аббревиатурой, шифром либо профильным синонимом в заголовке листа.
    """
    best = 0.0
    for abbreviation, patterns in MAPPING_RULES.items():
        abbr = _normalize(abbreviation)
        codes = [pattern for pattern in patterns if _CODE_PATTERN.fullmatch(pattern)]
        keywords = [
            _normalize(pattern)
            for pattern in patterns
            if not _CODE_PATTERN.fullmatch(pattern)
        ]

        pdf_has_abbr = _has_token(pdf_text, abbr)
        sheet_has_abbr = _has_token(sheet_text, abbr)
        pdf_has_code = any(code in pdf_text for code in codes)
        sheet_has_code = any(code in sheet_text for code in codes)
        pdf_has_keyword = any(keyword in pdf_text for keyword in keywords)
        sheet_has_keyword = any(keyword in sheet_text for keyword in keywords)

        if pdf_has_abbr:
            if sheet_has_abbr:
                best = max(best, 55.0)
            elif sheet_has_code:
                best = max(best, 50.0)
            elif sheet_has_keyword:
                best = max(best, 45.0)

        if pdf_has_code:
            if sheet_has_code:
                best = max(best, 55.0)
            elif sheet_has_abbr:
                best = max(best, 45.0)
            elif sheet_has_keyword:
                best = max(best, 40.0)

        if pdf_has_keyword:
            if sheet_has_keyword:
                best = max(best, 35.0)
            elif sheet_has_abbr:
                best = max(best, 30.0)

    return best


def score_pair(pdf_stem: str, sheet_name: str, sheet_title: str = "") -> float:
    """Оценивает сходство имени PDF и листа сметы по шкале 0-100.

    Лексическое сходство RapidFuzz является базой, а словарь разделов дает
    решающую смысловую поддержку. Без подтверждения аббревиатуры/шифра/синонима
    пара не может преодолеть ``MIN_CONFIDENCE``.
    """
    pdf_text = _normalize(pdf_stem)
    sheet_text = _normalize(f"{sheet_name} {sheet_title}")
    lexical = float(fuzz.token_set_ratio(pdf_text, sheet_text))
    evidence = _rule_evidence(pdf_text, sheet_text)
    return round(min(100.0, lexical * 0.45 + evidence), 2)


def _read_sheet_title(worksheet: openpyxl.worksheet.worksheet.Worksheet) -> str:
    """Возвращает первый непустой заголовок из A8/A9 (либо их объединение)."""
    titles: list[str] = []
    for coordinate in TITLE_CELLS:
        value = worksheet[coordinate].value
        if value is None:
            continue
        text = re.sub(r"\s+", " ", str(value).replace("_x000D_", " ")).strip()
        if text and text not in titles:
            titles.append(text)
    return " ".join(titles)


class ProjectDispatcher:
    """Сопоставляет PDF-тома РД с листами многолистовой сметы."""

    MAPPING_RULES: Final[dict[str, list[str]]] = MAPPING_RULES
    MIN_CONFIDENCE: Final[float] = MIN_CONFIDENCE

    def read_sections(self, smeta_path: Path) -> list[SheetCandidate]:
        """Читает только листы ``ЛС ...`` и их заголовки A8/A9.

        Листы ССРСС, ОСР, возвратных сумм и прочие служебные листы
        игнорируются, поскольку они не описывают отдельный раздел проекта.
        """
        workbook = openpyxl.load_workbook(
            str(smeta_path), read_only=True, data_only=True
        )
        try:
            sections: list[SheetCandidate] = []
            for sheet_name in workbook.sheetnames:
                if not sheet_name.upper().startswith(SECTION_SHEET_PREFIX):
                    continue
                worksheet = workbook[sheet_name]
                sections.append(
                    SheetCandidate(
                        name=sheet_name,
                        title=_read_sheet_title(worksheet),
                    )
                )
            return sections
        finally:
            workbook.close()

    def match_volumes(
        self, rd_dir: Path, smeta_path: Path
    ) -> list[tuple[Path, str, float]]:
        """Возвращает пары ``(pdf_path, sheet_name, confidence)``.

        Пары назначаются глобально по убыванию уверенности, поэтому один лист
        сметы не может быть привязан к двум разным томам. PDF без надежного
        соответствия пропускаются и могут быть разобраны вручную.
        """
        rd_path = Path(rd_dir)
        smeta_file = Path(smeta_path)
        if not rd_path.is_dir():
            raise FileNotFoundError(f"Папка с томами РД не найдена: {rd_path}")
        if not smeta_file.is_file():
            raise FileNotFoundError(f"Файл сметы не найден: {smeta_file}")

        pdf_files = sorted(
            (
                path
                for path in rd_path.iterdir()
                if path.is_file() and path.suffix.lower() == ".pdf"
            ),
            key=lambda path: path.name.casefold(),
        )
        sections = self.read_sections(smeta_file)

        ranked: list[_ScoredPair] = []
        for pdf_index, pdf_path in enumerate(pdf_files):
            for sheet_index, section in enumerate(sections):
                score = score_pair(pdf_path.stem, section.name, section.title)
                if score >= self.MIN_CONFIDENCE:
                    ranked.append(_ScoredPair(score, pdf_index, sheet_index))

        ranked.sort(key=lambda pair: (-pair.score, pair.pdf_index, pair.sheet_index))
        used_pdfs: set[int] = set()
        used_sheets: set[int] = set()
        matched: list[tuple[Path, str, float]] = []

        for pair in ranked:
            if pair.pdf_index in used_pdfs or pair.sheet_index in used_sheets:
                continue
            used_pdfs.add(pair.pdf_index)
            used_sheets.add(pair.sheet_index)
            matched.append(
                (
                    pdf_files[pair.pdf_index],
                    sections[pair.sheet_index].name,
                    pair.score,
                )
            )

        return sorted(matched, key=lambda item: item[0].name.casefold())


__all__ = [
    "MAPPING_RULES",
    "MIN_CONFIDENCE",
    "ProjectDispatcher",
    "SheetCandidate",
    "score_pair",
]
