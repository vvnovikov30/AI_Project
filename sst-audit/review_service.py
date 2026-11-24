"""Сервисный слой арбитража HITL (домен без UI-зависимостей).

Модуль отделяет бизнес-логику ревью от терминального интерфейса
(``cli_review.py``): здесь нет ``rich``, ``input()`` и ``print()``, поэтому
логику можно использовать из ночного пайплайна, API или тестов.

Принципы:

* **SSOT**: данные читаются и пишутся только через ``db`` (SQLite);
  промежуточных JSON/XML-дампов нет.
* **Zero Float**: все количества и коэффициенты — ``Decimal``.
* **Без N+1**: ``load_pending_cards`` выполняет фиксированные 3 SQL-запроса
  (очередь + позиции РД + позиции сметы) независимо от числа карточек.

Запуск через терминал остается в ``cli_review.py``::

    python cli_review.py --db sst_audit.db --session latest --reviewer ivanov
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import json
import sqlite3
from typing import Final, Optional

import db
from evaluator import DiscrepancyDetector
from matcher import CascadeMatcher
from models import CellPointer, NormalizedItem, UnitClass
from normalizer import DimensionLockExtractor, UnitNormalizer

DEFAULT_REVIEWER: Final[str] = "estimator_cli"
LATEST_SESSION: Final[str] = "latest"
CONFIDENCE_WARN_THRESHOLD: Final[Decimal] = Decimal("50.0")
# Синхронизировано с main.MIN_MATCH_THRESHOLD: шум ниже 40% не показываем.
MIN_CANDIDATE_SCORE: Final[Decimal] = Decimal("40.0")
MAX_CANDIDATES: Final[int] = 3
ONE: Final[Decimal] = Decimal("1.0")
PERCENT: Final[Decimal] = Decimal("100")

# Коды причин отклонения: клавиша -> (код в БД, пояснение для сметчика).
REJECT_REASONS: Final[dict[str, tuple[str, str]]] = {
    "1": ("SEMANTIC_MISMATCH", "Семантика: описание относится к другому ресурсу"),
    "2": ("UNIT_MULTIPLICITY", "Единицы измерения или кратность (например, 100 м)"),
    "3": ("PARSER_ERROR", "Ошибка парсера: данные извлечены неверно"),
    "4": ("OTHER", "Иное (уточните комментарием)"),
}
VALID_COMMANDS: Final[frozenset[str]] = frozenset(
    {"1", "2", "3", "A", "R", "C", "S", "Q"}
)
CLOSED_SESSION_STATUSES: Final[frozenset[str]] = frozenset({"REVIEWED", "EXPORTED"})

# Батч-запросы очереди: количество запросов не зависит от числа карточек.
_RD_ITEMS_SQL: Final[str] = "SELECT * FROM rd_items WHERE session_id = ? ORDER BY id ASC"
_SMETA_ITEMS_SQL: Final[str] = (
    "SELECT * FROM smeta_items WHERE session_id = ? ORDER BY id ASC"
)


@dataclass(frozen=True)
class CandidateMatch:
    """Кандидат РД для выбора сметчиком (score >= 40%)."""

    rd_item_id: int
    score: Decimal
    description: str
    unit: str
    quantity: Decimal


@dataclass(frozen=True)
class ReviewCard:
    """Данные карточки сравнения одной связи для терминального ревью."""

    match_id: int
    index: int
    total: int
    confidence: Decimal
    conversion_ratio: Decimal
    rd_code: Optional[str]
    rd_description: Optional[str]
    rd_unit: Optional[str]
    rd_quantity: Optional[Decimal]
    smeta_code: Optional[str]
    smeta_description: Optional[str]
    smeta_unit: Optional[str]
    smeta_quantity: Optional[Decimal]
    unit_price: Optional[Decimal]
    total_cost: Optional[Decimal]
    candidates: tuple[CandidateMatch, ...] = ()

    @property
    def is_paired(self) -> bool:
        """Признак наличия обеих сторон связи (РД и позиция сметы)."""
        return self.rd_quantity is not None and self.smeta_quantity is not None

    @property
    def waste_coefficient(self) -> Decimal:
        """Норматив отходов из единственного источника — ``DiscrepancyDetector``."""
        return DiscrepancyDetector.waste_coefficient(self.smeta_description)

    def adjusted_rd_volume(self, ratio: Optional[Decimal] = None) -> Optional[Decimal]:
        """Объем РД с учетом коэффициента приведения и норматива отходов."""
        if self.rd_quantity is None:
            return None
        effective_ratio = self.conversion_ratio if ratio is None else ratio
        return self.rd_quantity * effective_ratio * (ONE + self.waste_coefficient)

    def delta_volume(self, ratio: Optional[Decimal] = None) -> Optional[Decimal]:
        """Дельта объема: смета минус приведенный объем РД (>0 — завышение)."""
        adjusted = self.adjusted_rd_volume(ratio)
        if adjusted is None or self.smeta_quantity is None:
            return None
        return self.smeta_quantity - adjusted

    def allowed_delta(self, ratio: Optional[Decimal] = None) -> Optional[Decimal]:
        """Допустимое расхождение по регламенту ``DiscrepancyDetector``."""
        adjusted = self.adjusted_rd_volume(ratio)
        if adjusted is None:
            return None
        return (
            DiscrepancyDetector.ABS_TOLERANCE
            + DiscrepancyDetector.REL_TOLERANCE * adjusted
        )

    def delta_cost(self, ratio: Optional[Decimal] = None) -> Optional[Decimal]:
        """Стоимостная оценка расхождения, если в смете заполнена цена."""
        if self.unit_price is None or self.unit_price == 0:
            return None
        delta = self.delta_volume(ratio)
        if delta is None:
            return None
        return delta * self.unit_price


@dataclass(frozen=True)
class ReviewOutcome:
    """Итоги сеанса ревью."""

    session_id: str
    total: int
    approved: int
    rejected: int
    adjusted: int
    skipped: int
    remaining: int
    session_status: str

    @property
    def processed(self) -> int:
        """Количество записей, по которым сохранено решение."""
        return self.approved + self.rejected + self.adjusted


# ----------------------------------------------------------------------
# 1. Приведение типов и форматирование (без UI-зависимостей)
# ----------------------------------------------------------------------


def _to_decimal(value: object) -> Decimal:
    """Приводит значение SQLite (REAL/TEXT) к Decimal без потери точности."""
    if value is None:
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def _optional_decimal(value: object) -> Optional[Decimal]:
    """Возвращает Decimal либо None (для отсутствующей стороны связи)."""
    if value is None:
        return None
    return _to_decimal(value)


def _text_or_none(value: object) -> Optional[str]:
    """Нормализует текстовое поле: пустые строки превращает в None."""
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def format_quantity(value: Optional[Decimal]) -> str:
    """Форматирует количество без экспоненты и лишних нулей."""
    if value is None:
        return "—"
    try:
        quantized = value.quantize(Decimal("0.001"))
    except InvalidOperation:
        return f"{value:f}"
    if quantized == 0:
        return "0"
    text = f"{quantized:f}"
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def format_ratio(value: Decimal) -> str:
    """Форматирует коэффициент приведения с четырьмя знаками."""
    return f"{value:.4f}"


def format_money(value: Optional[Decimal]) -> str:
    """Форматирует денежное значение."""
    if value is None:
        return "—"
    return f"{value:,.2f}"


def format_percent(value: Optional[Decimal], *, signed: bool = False) -> str:
    """Форматирует процент."""
    if value is None:
        return "—"
    return f"{value:+.2f}%" if signed else f"{value:.2f}%"


def format_delta(value: Optional[Decimal]) -> str:
    """Форматирует дельту объема со знаком."""
    if value is None:
        return "—"
    return f"{value:+.3f}"


def verdict(card: ReviewCard, ratio: Optional[Decimal] = None) -> tuple[str, str]:
    """Текст и стиль вердикта по количественному расхождению."""
    delta = card.delta_volume(ratio)
    allowed = card.allowed_delta(ratio)
    if delta is None or allowed is None:
        return "Расчет недоступен: отсутствует пара РД/смета", "dim"
    if abs(delta) > allowed:
        if delta > 0:
            return "Завышение объема в смете", "red"
        return "Занижение объема относительно РД", "yellow"
    return "В пределах допуска", "green"


def _payload_from_row(row: sqlite3.Row) -> dict[str, object]:
    """Читает аудиторский JSON-след позиции, устойчиво к старым записям."""
    raw = row["raw_payload"]
    if not raw:
        return {}
    try:
        payload = json.loads(raw)
    except (TypeError, ValueError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _normalized_from_row(
    row: sqlite3.Row,
    *,
    is_smeta: bool,
) -> NormalizedItem:
    """Восстанавливает ``NormalizedItem`` из строки SSOT для пересчета score."""
    payload = _payload_from_row(row)
    quantity = _to_decimal(row["quantity"]) or Decimal("0")
    raw_unit = str(payload.get("raw_unit") or row["unit"] or "")
    raw_title = str(payload.get("raw_title") or row["description"] or "")

    canonical_value = payload.get("canonical_unit")
    try:
        canonical_unit = UnitClass(str(canonical_value))
    except ValueError:
        canonical_unit, _, _ = UnitNormalizer.normalize(raw_unit, quantity)

    multiplier_value = payload.get("unit_multiplier")
    if multiplier_value is not None:
        try:
            unit_multiplier = Decimal(str(multiplier_value))
        except InvalidOperation:
            unit_multiplier = Decimal("1.0")
    else:
        _, _, unit_multiplier = UnitNormalizer.normalize(raw_unit, quantity)

    dimensions_raw = payload.get("dimensions")
    dimensions = (
        {str(key): str(value) for key, value in dimensions_raw.items()}
        if isinstance(dimensions_raw, dict)
        else DimensionLockExtractor.extract(raw_title)
    )

    pointer_raw = payload.get("pointer")
    pointer_payload = pointer_raw if isinstance(pointer_raw, dict) else {}
    pointer = CellPointer(
        doc_type=str(
            pointer_payload.get("doc_type") or ("LSR" if is_smeta else "SPECIFICATION")
        ),
        file_name=str(pointer_payload.get("file_name") or ""),
        sheet_name=str(pointer_payload.get("sheet_name") or ""),
        row_idx=int(pointer_payload.get("row_idx") or 1),
        col_letter=str(pointer_payload.get("col_letter") or "A"),
        coordinate=str(pointer_payload.get("coordinate") or "—"),
    )

    unit_price = (
        _optional_decimal(row["unit_price"])
        if is_smeta and "unit_price" in row.keys()
        else None
    )
    total_cost = (
        _optional_decimal(row["total_cost"])
        if is_smeta and "total_cost" in row.keys()
        else None
    )

    return NormalizedItem(
        raw_id=str(payload.get("raw_id") or f"DB_{row['id']}"),
        code=_text_or_none(row["code"]),
        raw_title=raw_title,
        clean_title=str(payload.get("clean_title") or raw_title.lower()),
        raw_unit=raw_unit,
        canonical_unit=canonical_unit,
        unit_multiplier=unit_multiplier,
        volume=quantity,
        dimensions=dimensions,
        context_section=str(payload.get("context_section") or "Default_Section"),
        pointer=pointer,
        is_resource=bool(payload.get("is_resource", is_smeta)),
        parent_id=(
            str(payload["parent_id"]) if payload.get("parent_id") is not None else None
        ),
        unit_price=unit_price,
        total_cost=total_cost,
    )


# ----------------------------------------------------------------------
# 2. Загрузка очереди HITL без N+1
# ----------------------------------------------------------------------


def _traced_execute(
    conn: sqlite3.Connection,
    sql: str,
    params: Sequence[object],
    trace_callback: Optional[Callable[[str], None]],
) -> sqlite3.Cursor:
    """Выполняет запрос и (опционально) сообщает о нем диагностическому хуку.

    Хук используется тестами и профилировщиком для контроля числа обращений
    к SSOT: ``load_pending_cards`` обязан укладываться в фиксированное число
    запросов независимо от размера очереди.
    """
    if trace_callback is not None:
        trace_callback(sql)
    return conn.execute(sql, params)


def _score_candidates(
    smeta_item: NormalizedItem,
    rd_cache: Mapping[int, sqlite3.Row],
    rd_items: Mapping[int, NormalizedItem],
) -> tuple[CandidateMatch, ...]:
    """Возвращает top-3 кандидатов РД со score >= ``MIN_CANDIDATE_SCORE``."""
    scored: list[tuple[int, Decimal]] = []
    for rd_id, rd_item in rd_items.items():
        score = Decimal(str(CascadeMatcher.calculate_score(smeta_item, rd_item)))
        if score >= MIN_CANDIDATE_SCORE:
            scored.append((rd_id, score))

    # Сортировка стабильна: при равных score сохраняется порядок id ASC.
    scored.sort(key=lambda entry: entry[1], reverse=True)
    return tuple(
        CandidateMatch(
            rd_item_id=rd_id,
            score=score,
            description=str(rd_cache[rd_id]["description"] or ""),
            unit=str(rd_cache[rd_id]["unit"] or ""),
            quantity=_to_decimal(rd_cache[rd_id]["quantity"]) or Decimal("0"),
        )
        for rd_id, score in scored[:MAX_CANDIDATES]
    )


def _card_from_row(
    row: sqlite3.Row,
    *,
    index: int,
    total: int,
    candidates: Sequence[CandidateMatch] = (),
) -> ReviewCard:
    """Собирает карточку из строки ``db.PENDING_REVIEWS_SQL``."""
    conversion = row["conversion_ratio"]
    return ReviewCard(
        match_id=int(row["match_id"]),
        index=index,
        total=total,
        confidence=_to_decimal(row["confidence_score"]),
        conversion_ratio=(ONE if conversion is None else _to_decimal(conversion)),
        rd_code=_text_or_none(row["rd_code"]),
        rd_description=_text_or_none(row["rd_description"]),
        rd_unit=_text_or_none(row["rd_unit"]),
        rd_quantity=_optional_decimal(row["rd_quantity"]),
        smeta_code=_text_or_none(row["smeta_code"]),
        smeta_description=_text_or_none(row["smeta_description"]),
        smeta_unit=_text_or_none(row["smeta_unit"]),
        smeta_quantity=_optional_decimal(row["smeta_quantity"]),
        unit_price=_optional_decimal(row["unit_price"]),
        total_cost=_optional_decimal(row["total_cost"]),
        candidates=tuple(candidates),
    )


def load_pending_cards(
    conn: sqlite3.Connection,
    session_id: str,
    *,
    trace_callback: Optional[Callable[[str], None]] = None,
) -> list[ReviewCard]:
    """Загружает очередь HITL сессии, выполняя ровно три SQL-запроса.

    Раньше кандидаты считались отдельными запросами на каждую карточку (N+1).
    Теперь позиции РД и сметы читаются батчем, нормализация и скоринг
    выполняются в памяти, а результат кэшируется по ``smeta_item_id``: число
    запросов постоянно (очередь + rd_items + smeta_items) при любой длине
    очереди.

    :param trace_callback: диагностический хук, получающий текст каждого
        выполненного SQL-запроса.
    """
    rows = _traced_execute(
        conn, db.PENDING_REVIEWS_SQL, (session_id,), trace_callback
    ).fetchall()
    if not rows:
        return []

    rd_rows = _traced_execute(
        conn, _RD_ITEMS_SQL, (session_id,), trace_callback
    ).fetchall()
    smeta_rows = _traced_execute(
        conn, _SMETA_ITEMS_SQL, (session_id,), trace_callback
    ).fetchall()

    rd_cache: dict[int, sqlite3.Row] = {int(row["id"]): row for row in rd_rows}
    rd_items: dict[int, NormalizedItem] = {
        rd_id: _normalized_from_row(row, is_smeta=False)
        for rd_id, row in rd_cache.items()
    }
    smeta_cache: dict[int, sqlite3.Row] = {int(row["id"]): row for row in smeta_rows}
    candidates_cache: dict[int, tuple[CandidateMatch, ...]] = {}

    total = len(rows)
    cards: list[ReviewCard] = []
    for index, row in enumerate(rows, start=1):
        candidates: tuple[CandidateMatch, ...] = ()
        smeta_item_id = row["smeta_item_id"]
        if smeta_item_id is not None:
            key = int(smeta_item_id)
            if key not in candidates_cache:
                smeta_row = smeta_cache.get(key)
                candidates_cache[key] = (
                    _score_candidates(
                        _normalized_from_row(smeta_row, is_smeta=True),
                        rd_cache,
                        rd_items,
                    )
                    if smeta_row is not None
                    else ()
                )
            candidates = candidates_cache[key]
        cards.append(
            _card_from_row(row, index=index, total=total, candidates=candidates)
        )
    return cards


def candidates_for_row(
    conn: sqlite3.Connection,
    row: sqlite3.Row,
) -> tuple[CandidateMatch, ...]:
    """Совместимость: кандидаты для одной строки очереди (2 SQL-запроса).

    Используется только внешними вызовами; пакетная загрузка очереди идет
    через ``load_pending_cards`` и не деградирует до N+1.
    """
    smeta_item_id = row["smeta_item_id"]
    if smeta_item_id is None:
        return ()

    smeta_row = conn.execute(
        "SELECT * FROM smeta_items WHERE id = ?",
        (int(smeta_item_id),),
    ).fetchone()
    if smeta_row is None:
        return ()

    rd_rows = conn.execute(_RD_ITEMS_SQL, (str(row["session_id"]),)).fetchall()
    if not rd_rows:
        return ()

    rd_cache: dict[int, sqlite3.Row] = {int(rd_row["id"]): rd_row for rd_row in rd_rows}
    rd_items: dict[int, NormalizedItem] = {
        rd_id: _normalized_from_row(rd_row, is_smeta=False)
        for rd_id, rd_row in rd_cache.items()
    }
    return _score_candidates(
        _normalized_from_row(smeta_row, is_smeta=True),
        rd_cache,
        rd_items,
    )


# ----------------------------------------------------------------------
# 3. Запись решений в SSOT (арбитраж + append-only журнал)
# ----------------------------------------------------------------------


def apply_approval(
    conn: sqlite3.Connection,
    card: ReviewCard,
    reviewer: str,
) -> None:
    """Фиксирует подтверждение связи (для записи без пары — отсутствие пары)."""
    comment = None if card.is_paired else "Подтверждено отсутствие пары РД/смета"
    db.record_expert_review(
        conn,
        card.match_id,
        reviewer,
        "APPROVED",
        comment=comment,
    )


def apply_candidate_selection(
    conn: sqlite3.Connection,
    card: ReviewCard,
    candidate: CandidateMatch,
    position: int,
    reviewer: str,
) -> None:
    """Фиксирует выбор сметчиком одного из top-3 кандидатов РД.

    Уточнение связи и экспертное решение пишутся в одной транзакции: журнал
    арбитража не может содержать выбор без итогового решения.
    """
    comment = f"Выбран кандидат №{position}: {candidate.description}"
    with db.transaction(conn):
        db.update_match_link(
            conn,
            card.match_id,
            candidate.rd_item_id,
            candidate.score,
            actor=reviewer,
            comment=comment,
        )
        db.record_expert_review(
            conn,
            card.match_id,
            reviewer,
            "APPROVED",
            comment=comment,
        )


def apply_rejection(
    conn: sqlite3.Connection,
    card: ReviewCard,
    reviewer: str,
    reason_code: str,
    comment: Optional[str],
) -> None:
    """Фиксирует отклонение связи, сохраняя лучшего кандидата (top-1).

    Связь с РД не обнуляется: эксперт фиксирует ошибочность сопоставления,
    но ``rd_item_id`` остается для трассируемости и последующего разбора.
    """
    if card.candidates:
        top = card.candidates[0]
        with db.transaction(conn):
            db.update_match_link(
                conn,
                card.match_id,
                top.rd_item_id,
                top.score,
                actor=reviewer,
                comment=f"Отклонено ({reason_code}); лучший кандидат сохранен",
            )
            db.record_expert_review(
                conn,
                card.match_id,
                reviewer,
                "REJECTED",
                reason_code=reason_code,
                comment=comment,
            )
        return
    db.record_expert_review(
        conn,
        card.match_id,
        reviewer,
        "REJECTED",
        reason_code=reason_code,
        comment=comment,
    )


def apply_adjustment(
    conn: sqlite3.Connection,
    card: ReviewCard,
    reviewer: str,
    ratio: Decimal,
) -> Optional[Decimal]:
    """Сохраняет скорректированный коэффициент и новую дельту в аудит."""
    delta = card.delta_volume(ratio)
    comment = (
        f"Коэффициент приведения: {format_ratio(card.conversion_ratio)} -> "
        f"{format_ratio(ratio)}; дельта объема: {format_delta(delta)}"
    )
    db.record_expert_review(
        conn,
        card.match_id,
        reviewer,
        "ADJUSTED",
        adjusted_ratio=ratio,
        comment=comment,
    )
    return delta


# ----------------------------------------------------------------------
# 4. Сессия и финализация
# ----------------------------------------------------------------------


def resolve_session_id(conn: sqlite3.Connection, session: str) -> str:
    """Возвращает ID сессии; ``latest`` — последняя созданная сессия."""
    value = (session or "").strip()
    if value.lower() == LATEST_SESSION:
        row = conn.execute(
            "SELECT id FROM sessions ORDER BY created_at DESC, rowid DESC LIMIT 1"
        ).fetchone()
        if row is None:
            raise LookupError("В базе нет ни одной сессии аудита.")
        return str(row["id"])

    row = conn.execute("SELECT id FROM sessions WHERE id = ?", (value,)).fetchone()
    if row is None:
        raise LookupError(f"Сессия '{value}' не найдена.")
    return str(row["id"])


def pending_count(conn: sqlite3.Connection, session_id: str) -> int:
    """Число связей, ожидающих решения (для итоговой статистики)."""
    return len(db.get_pending_reviews(conn, session_id))


def finalize_session(conn: sqlite3.Connection, session_id: str) -> str:
    """Переводит сессию в ``REVIEWED``, если очередь HITL полностью разобрана."""
    row = conn.execute(
        "SELECT status FROM sessions WHERE id = ?", (session_id,)
    ).fetchone()
    if row is None:
        raise LookupError(f"Сессия '{session_id}' не найдена.")
    current_status = str(row["status"])

    total_matches = conn.execute(
        "SELECT COUNT(*) AS c FROM match_records WHERE session_id = ?",
        (session_id,),
    ).fetchone()["c"]
    open_matches = conn.execute(
        """
        SELECT COUNT(*) AS c FROM match_records
        WHERE session_id = ? AND status NOT IN ('AUTO', 'RESOLVED')
        """,
        (session_id,),
    ).fetchone()["c"]

    if (
        int(total_matches) > 0
        and int(open_matches) == 0
        and current_status not in CLOSED_SESSION_STATUSES
    ):
        db.update_session_status(conn, session_id, "REVIEWED")
        return "REVIEWED"
    return current_status
