"""Общие хелперы парсеров: источники в памяти, кодировка PDF, артефакты.

Листовой модуль пакета: не импортирует другие модули ``parsers``, поэтому
циклические импорты между ``lsr_excel``/``spec_pdf``/``lsr_pdf`` невозможны.

Zero Float Policy: расчетные величины приводятся к ``Decimal``; координаты
``bbox`` и параметры листа сериализуются строками (``_coordinate_text``,
``_pdf_bbox``, ``_page_meta_value``), чтобы в бизнес-моделях не появлялись
``float``.
"""

from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
import re
from typing import Any, BinaryIO, Dict, Final, Optional, Tuple, TypeAlias, Union

# Внешний источник парсера: путь к файлу либо уже открытый бинарный поток
# (например, ``tempfile.SpooledTemporaryFile`` из веб-слоя). Поток принадлежит
# вызывающему коду: пакет его не закрывает.
SourceType: TypeAlias = Union[str, Path, BytesIO, BinaryIO]

# Кодировки, которыми восстанавливается битый текстовый слой PDF.
_PDF_LSR_MOJIBAKE_CODECS: Tuple[str, ...] = ("utf-8", "cp1251")
# Числовая ячейка сметы: «-1 234,5», «12.5»; буквы и сноски снимаются заранее.
_PDF_LSR_NUMBER_PATTERN: re.Pattern[str] = re.compile(r"^-?\d[\d\s.,]*$")


def _is_binary_source(source: Any) -> bool:
    """Проверяет, что источник — открытый бинарный поток, а не путь к файлу."""
    return callable(getattr(source, "read", None)) and not isinstance(
        source, (str, bytes, bytearray, Path)
    )


def _read_source_bytes(source: Any) -> bytes:
    """Читает байты потока, не закрывая его.

    Позиция потока сбрасывается в начало, если поток позиционируемый: parsers
    читают PDF дважды (pypdfium2 + pdfplumber), поэтому данные должны
    начинаться с начала документа.
    """
    seek = getattr(source, "seek", None)
    seekable = getattr(source, "seekable", None)
    if callable(seek) and callable(seekable) and seekable():
        try:
            seek(0)
        except (OSError, ValueError):  # pragma: no cover - защитный сценарий
            pass
    data = source.read()
    if isinstance(data, str):
        raise TypeError(
            "Источник открыт в текстовом режиме; требуется бинарный поток"
        )
    return bytes(data)


def _prepare_source(
    source: SourceType,
    display_name: Optional[str] = None,
    fallback_name: str = "in_memory.pdf",
) -> Tuple[Optional[Path], str, Optional[BytesIO]]:
    """Готовит источник парсера: путь либо внутренний буфер и безопасное имя.

    Для потока данные копируются в собственный :class:`~io.BytesIO`: parsers
    открывают документ дважды и не должны закрывать поток вызывающего кода.
    ``display_name`` попадает в ``CellPointer.file_name`` вместо реального
    пути/временного имени (``/dev/shm/tmp_xyz.pdf``), чтобы отчеты не содержали
    служебных путей.

    :param source: путь к файлу или бинарный поток с документом.
    :param display_name: безопасное имя файла для протокола разбора.
    :param fallback_name: имя по умолчанию для потока без ``display_name``.
    :return: кортеж ``(путь или None, display_name, внутренний буфер или None)``.
    """
    if _is_binary_source(source):
        buffer = BytesIO(_read_source_bytes(source))
        return None, str(display_name or fallback_name), buffer
    path = Path(source)
    return path, str(display_name or source), None


def _close_buffer(buffer: Optional[BytesIO]) -> None:
    """Закрывает внутренний буфер парсера (идемпотентно, ``None`` допустим)."""
    if buffer is not None and not buffer.closed:
        buffer.close()


def _coordinate_text(value: Any) -> Optional[str]:
    """Сериализует координату PDF/Excel строкой без ``float`` в модели.

    :param value: число (``int``/``Decimal``/``float`` из pdfplumber) или строка.
    :return: десятичная строка либо ``None``, если значение не разбирается.
    """
    if value is None:
        return None
    try:
        return str(Decimal(str(value)))
    except (InvalidOperation, ValueError):
        return None


def _pdf_bbox(bbox: Any) -> Optional[Tuple[str, str, str, str]]:
    """Преобразует ``bbox`` pdfplumber (``x0, top, x1, bottom``) в строки.

    :param bbox: последовательность из четырех координат либо ``None``.
    :return: кортеж координат-строк либо ``None``, если геометрия недоступна.
    """
    try:
        parts = tuple(bbox) if bbox is not None else ()
    except TypeError:
        return None
    if len(parts) != 4:
        return None
    texts = tuple(_coordinate_text(part) for part in parts)
    if any(text is None for text in texts):
        return None
    return (texts[0], texts[1], texts[2], texts[3])


def _page_meta_value(page: Any) -> Optional[Dict[str, Any]]:
    """Возвращает метаданные листа PDF: размеры строками и поворот целым.

    :param page: страница pdfplumber (``width``/``height``/``rotation``).
    :return: словарь метаданных либо ``None`` для двойников без геометрии.
    """
    width = _coordinate_text(getattr(page, "width", None))
    height = _coordinate_text(getattr(page, "height", None))
    rotation = getattr(page, "rotation", None)
    if width is None and height is None and rotation is None:
        return None
    meta: Dict[str, Any] = {}
    if width is not None:
        meta["width"] = width
    if height is not None:
        meta["height"] = height
    if rotation is not None:
        try:
            meta["rotation"] = int(rotation)
        except (TypeError, ValueError):  # pragma: no cover - защитный сценарий
            meta["rotation"] = str(rotation)
    return meta


def _flush_page_cache(page: Any) -> None:
    """Освобождает кэш объектов pdfplumber-страницы после ее обработки.

    Тестовые двойники страниц не имеют ``flush_cache``, поэтому вызов
    выполняется только для реального ``PDFPage``.
    """
    flush = getattr(page, "flush_cache", None)
    if callable(flush):
        flush()


def _clean_estimate_title(raw_title: str) -> str:
    """Очистка сметного наименования от формул обоснования цен, артефактов Excel и префиксов."""
    text = raw_title.replace("_x000D_", " ")
    # Отсекаем строки формул цен: "1 236,60 = [10 913 / 1,22 / 7,54] + 3% Трансп..."
    lines = [
        line for line in text.split("\n")
        if not ("=" in line and ("[" in line or "трансп" in line.lower() or "заг.скл" in line.lower()))
    ]
    text = " ".join(lines)
    # Убираем префикс "ОБОРУДОВАНИЕ:"
    text = re.sub(r"^\s*оборудование\s*:\s*", "", text, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).lower().strip()


# ----------------------------------------------------------------------
# Метрологические фильтры табличного мусора: шапки таблиц, штампы томов
# и строки без осмысленного наименования не являются позициями ведомости
# ----------------------------------------------------------------------
# Служебные слова шапок ВОР/спецификаций. Строка, целиком состоящая из
# них, является артефактом распознавания таблицы, а не позицией.
_TABLE_HEADER_WORDS: Final[frozenset[str]] = frozenset(
    {
        "наименование",
        "наименования",
        "наименований",
        "обоснование",
        "обоснования",
        "обоснований",
        "единица",
        "единицы",
        "единиц",
        "измерение",
        "измерения",
        "измерений",
        "количество",
        "количества",
        "количеств",
        "примечание",
        "примечания",
        "примечаний",
        "масса",
        "массы",
        "код",
        "кода",
        "колво",
        "кол",
        "ед",
        "изм",
        "номер",
        "номера",
        "пп",
        "работ",
        "работы",
        "затрат",
        "затраты",
    }
)
# Основания служебных слов: «наименование вида работ», «количество».
_TABLE_HEADER_STEMS: Final[tuple[str, ...]] = (
    "наименован",
    "обоснован",
    "количеств",
    "измерен",
    "примечан",
)
# Явные надписи шапок, которые встречаются без знаков препинания.
_TABLE_HEADER_LITERALS: Final[frozenset[str]] = frozenset(
    {
        "наименование",
        "наименование работ",
        "наименование вида работ",
        "наименование работ, ресурсов, затрат",
        "кол-во",
        "ед. изм.",
        "ед.изм.",
        "единица измерения",
    }
)
# Системные надписи томов РД: колонтитулы, штампы и титулы разделов.
_STAMP_OR_TITLE_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"^(?:(?:раздел|том)(?:\s*\d+(?:[.\-]\d+)*)?"
    r"|проектная\s+документация"
    r"|архитектурные\s+решения"
    r"|ведомость\s+объемов"
    r"|лист\s*\d+)$"
)



def _artifact_text(value: str) -> str:
    """Нормализует надпись для сравнения с шапками и штампами таблиц.

    Восстанавливает битую кодировку PDF (cp866/cp1251), приводит строку к
    нижнему регистру, заменяет ``ё`` на ``е`` и схлопывает пробелы.
    """
    text = _repair_pdf_text(str(value or "")).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text.casefold().replace("ё", "е")).strip()


def _is_table_header_artifact(name: str) -> bool:
    """Является ли наименование шапкой таблицы, а не позицией ВОР.

    Строка отсекается, когда все ее токены (после удаления пунктуации)
    входят в набор служебных слов шапки либо начинаются с их оснований.
    """
    text = _artifact_text(name)
    if not text:
        return False
    if text in _TABLE_HEADER_LITERALS:
        return True
    tokens = [
        token for token in re.sub(r"[^\w\s]", " ", text).split() if token
    ]
    if not tokens:
        return False
    return all(
        token in _TABLE_HEADER_WORDS
        or any(token.startswith(stem) for stem in _TABLE_HEADER_STEMS)
        for token in tokens
    )


def _is_stamp_or_title_artifact(name: str, code: Optional[str] = None) -> bool:
    """Отсекает колонтитулы и штампы томов РД («Лист 1», «Том 1»).

    ``code`` принимается для совместимости с единым барьером валидности:
    номер графы не подтверждает физический объем, поэтому надпись
    отсекается независимо от заполненного кода позиции.
    """
    text = _artifact_text(name)
    if not text:
        return False
    return _STAMP_OR_TITLE_PATTERN.fullmatch(text) is not None


def _has_measured_title(name: str) -> bool:
    """Проверяет, что в наименовании есть осмысленная буквенная часть."""
    letters = re.sub(r"[^A-Za-zА-Яа-яЁё]", "", name or "")
    return len(letters) >= 3


def _pdf_text_score(value: str) -> int:
    """Оценивает «естественность» текста для выбора кодировки.

    Считает кириллические символы положительным вкладом, а прочие
    не-ASCII символы (латиница с диакритикой, псевдографика cp866) —
    отрицательным.

    :param value: проверяемый фрагмент текста.
    :return: целочисленная оценка: чем больше, тем естественнее текст.
    """
    cyrillic = 0
    exotic = 0
    for char in value:
        if char.isspace() or char in "№«»—–-.,;:()[]/%":
            continue
        lowered = char.lower()
        if lowered == "ё" or "а" <= lowered <= "я":
            cyrillic += 1
        elif not char.isascii():
            exotic += 1
    return cyrillic * 2 - exotic * 2


def _repair_pdf_text(text: str) -> str:
    """Восстанавливает кириллицу после двойного искажения PDF-текста.

    Генераторы PDF смет выдают текст в UTF-8 или cp1251, который
    просмотрщик/извлекатель декодирует как cp866. Функция перебирает
    кандидатов ``encode('cp866').decode(codec)`` и выбирает вариант с
    наибольшей «кириллической» оценкой.

    :param text: исходный фрагмент, полученный от ``pdfplumber``.
    :return: восстановленный текст либо исходный фрагмент, если все
        кандидаты оказались хуже.
    """
    if not text or text.isascii():
        return text

    best = text
    best_score = _pdf_text_score(text)
    for codec in _PDF_LSR_MOJIBAKE_CODECS:
        try:
            candidate = text.encode("cp866").decode(codec)
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        score = _pdf_text_score(candidate)
        if score > best_score:
            best = candidate
            best_score = score
    return best


def _squash_spaces(value: str) -> str:
    """Сворачивает повторные пробелы и неразрывные пробелы в один.

    :param value: исходная строка.
    :return: строка без ведущих/замыкающих пробелов.
    """
    return re.sub(r"\s+", " ", (value or "").replace("\xa0", " ")).strip()


def _parse_pdf_number(value: str) -> Optional[Decimal]:
    """Разбирает число из ячейки PDF-сметы (пробелы — разделители тысяч).

    :param value: текст ячейки, например ``'1 287,00'``.
    :return: ``Decimal`` либо ``None``, если ячейка не является числом.
    """
    text = _squash_spaces(value)
    if not text or not _PDF_LSR_NUMBER_PATTERN.match(text):
        return None
    try:
        return Decimal(text.replace(" ", "").replace(",", "."))
    except InvalidOperation:
        return None

