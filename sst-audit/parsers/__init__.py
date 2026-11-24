"""Парсеры смет и ведомостей: единая точка входа пакета ``parsers``.

Пакет разбит по форматам источников, но сохраняет прежний модульный API
(``from parsers import LSRParser, PDFSpecParser, ...``): все классы, хелперы и
константы реэкспортируются здесь. Модули библиотек также попадают в
пространство имен пакета намеренно: тесты подменяют
``parsers.pdfplumber.open`` и ``parsers._ColumnSchema``.

Структура:

* :mod:`parsers.common` — io-источники (путь или ``BytesIO``), кодировка PDF,
  фильтры табличных артефактов;
* :mod:`parsers.lsr_excel` — локальные сметы Excel (4б, 421/пр);
* :mod:`parsers.spec_pdf` — спецификации ГОСТ 21.110 и ВОР из PDF;
* :mod:`parsers.vor_excel` — дефектные ведомости и ВОР из XLSX + фабрика;
* :mod:`parsers.lsr_pdf` — локальные сметы из векторных PDF.
"""

import openpyxl
import pdfplumber
import pypdfium2 as pdfium

from .common import (
    SourceType,
    _PDF_LSR_MOJIBAKE_CODECS,
    _PDF_LSR_NUMBER_PATTERN,
    _TABLE_HEADER_LITERALS,
    _TABLE_HEADER_STEMS,
    _TABLE_HEADER_WORDS,
    _artifact_text,
    _clean_estimate_title,
    _close_buffer,
    _coordinate_text,
    _flush_page_cache,
    _has_measured_title,
    _is_binary_source,
    _is_stamp_or_title_artifact,
    _is_table_header_artifact,
    _page_meta_value,
    _parse_pdf_number,
    _pdf_bbox,
    _pdf_text_score,
    _prepare_source,
    _read_source_bytes,
    _repair_pdf_text,
    _squash_spaces,
)
from .lsr_excel import (
    UNIT_FALLBACK_PATTERN,
    LSRParser,
    _LSR421Layout,
    _LSR4bLayout,
    _LSR_421_NOISE_PREFIXES,
    _LSR_421_POSITION_PATTERN,
    _LSR_421_SECTION_PATTERN,
    _LSR_CODE_STEMS,
    _LSR_HEADER_SCAN_ROWS,
    _LSR_NAME_STEMS,
    _LSR_POSITION_STEMS,
    _LSR_QTY_STEMS,
    _LSR_SERVICE_PREFIXES,
    _LSR_TOTAL_STEMS,
    _LSR_UNIT_STEMS,
    _UNIT_FALLBACK_SMALL_COUNT_PATTERN,
    _UNIT_FALLBACK_TRAILING_PATTERN,
)
from .spec_pdf import (
    PDFSpecParser,
    _ARITHMETIC_FORBIDDEN_PATTERN,
    _ARITHMETIC_OPERATOR_PATTERN,
    _ARITHMETIC_TOKEN_PATTERN,
    _COLUMN_NUMBER_PATTERN,
    _ColumnSchema,
    _DATE_OR_CODE_PATTERN,
    _DIALUX_PAGE_MARKERS,
    _GRAND_VOR_HEADER_MARKERS,
    _HEADER_MASS_TOKENS,
    _HEADER_QTY_TOKENS,
    _HEADER_SPEC_TOKENS,
    _HEADER_TITLE_TOKENS,
    _HEADER_UNIT_TOKENS,
    _POSITION_PATTERN,
    _SERVICE_ROW_TOKENS,
    _SPEC_MARK_PATTERN,
    _SPEC_PAGE_MARKERS,
    _SPEC_TABLE_PAGE_MARKERS,
    _SPEC_TITLE_PATTERN,
    _TABLE_ROW_PATTERN,
    _UNIT_VALUE_PATTERN,
    _VOR_PAGE_MARKERS,
    _VOR_SECTION_PATTERN,
    _VOR_SOURCE_PATTERN,
)
from .vor_excel import (
    ExcelVORParser,
    _header_key,
    _VOR_FOOTNOTE_PATTERN,
    _VOR_HEADER_SCAN_LIMIT,
    _VOR_NUMBER_PATTERN,
    get_spec_parser,
)
from .lsr_pdf import (
    PDFLSRParser,
    _LSRColumnLayout,
    _PDF_LSR_INT_TOKEN_PATTERN,
    _PDF_LSR_LABOR_CODE_PATTERN,
    _PDF_LSR_LINE_TOLERANCE,
    _PDF_LSR_MACHINE_CODE_PATTERN,
    _PDF_LSR_POSITION_PATTERN,
    _PDF_LSR_RATE_CODE_PATTERN,
    _PDF_LSR_RESOURCE_CODE_PATTERN,
    _PDF_LSR_SECTION_PATTERN,
    _PDF_LSR_SERVICE_PATTERN,
)

__all__ = [
    "SourceType",
    "LSRParser",
    "ExcelVORParser",
    "PDFSpecParser",
    "PDFLSRParser",
    "get_spec_parser",
    "UNIT_FALLBACK_PATTERN",
    "_ColumnSchema",
    "_LSR4bLayout",
    "_LSR421Layout",
    "_LSRColumnLayout",
    "_artifact_text",
    "_clean_estimate_title",
    "_close_buffer",
    "_coordinate_text",
    "_flush_page_cache",
    "_has_measured_title",
    "_header_key",
    "_is_binary_source",
    "_is_stamp_or_title_artifact",
    "_is_table_header_artifact",
    "_page_meta_value",
    "_parse_pdf_number",
    "_pdf_bbox",
    "_pdf_text_score",
    "_prepare_source",
    "_read_source_bytes",
    "_repair_pdf_text",
    "_squash_spaces",
]
