"""Интерактивный терминальный интерфейс сметчика (HITL) на базе Rich.

Модуль читает очередь ``HITL_PENDING`` из SQLite (SSOT), показывает карточку
сравнения РД и сметы и фиксирует экспертное решение в ``expert_reviews``.
После разбора очереди сессия переводится в статус ``REVIEWED``.

Запуск::

    python cli_review.py --db sst_audit.db --session latest --reviewer ivanov
"""

from __future__ import annotations

import argparse
from collections.abc import Sequence
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import json
from pathlib import Path
import sqlite3
from typing import Final, Optional

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

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
        """Норматив технологических отходов по наименованию позиции сметы."""
        title = (self.smeta_description or "").lower()
        coefficients = DiscrepancyDetector.WASTE_COEFFICIENTS
        if "арматур" in title:
            return coefficients["REBAR"]
        if "бетон" in title:
            return coefficients["CONCRETE"]
        if "кабел" in title:
            return coefficients["CABLE"]
        return coefficients["DEFAULT"]

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
# 1. Приведение типов и форматирование
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


def _format_quantity(value: Optional[Decimal]) -> str:
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


def _format_ratio(value: Decimal) -> str:
    """Форматирует коэффициент приведения с четырьмя знаками."""
    return f"{value:.4f}"


def _format_money(value: Optional[Decimal]) -> str:
    """Форматирует денежное значение."""
    if value is None:
        return "—"
    return f"{value:,.2f}"


def _format_percent(value: Optional[Decimal], *, signed: bool = False) -> str:
    """Форматирует процент."""
    if value is None:
        return "—"
    return f"{value:+.2f}%" if signed else f"{value:.2f}%"


def _format_delta(value: Optional[Decimal]) -> str:
    """Форматирует дельту объема со знаком."""
    if value is None:
        return "—"
    return f"{value:+.3f}"


def _confidence_style(score: Decimal) -> str:
    """Цвет заголовка карточки: красный для низкой уверенности."""
    return "red" if score < CONFIDENCE_WARN_THRESHOLD else "yellow"


def _verdict(card: ReviewCard, ratio: Optional[Decimal] = None) -> tuple[str, str]:
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


def _candidates_for_row(
    conn: sqlite3.Connection,
    row: sqlite3.Row,
) -> tuple[CandidateMatch, ...]:
    """Возвращает top-3 кандидатов РД со score >= ``MIN_CANDIDATE_SCORE``."""
    smeta_item_id = row["smeta_item_id"]
    if smeta_item_id is None:
        return ()

    smeta_row = conn.execute(
        "SELECT * FROM smeta_items WHERE id = ?",
        (int(smeta_item_id),),
    ).fetchone()
    if smeta_row is None:
        return ()

    rd_rows = conn.execute(
        "SELECT * FROM rd_items WHERE session_id = ? ORDER BY id ASC",
        (str(row["session_id"]),),
    ).fetchall()
    if not rd_rows:
        return ()

    smeta_item = _normalized_from_row(smeta_row, is_smeta=True)
    scored: list[tuple[NormalizedItem, Decimal, sqlite3.Row]] = []
    for rd_row in rd_rows:
        rd_item = _normalized_from_row(rd_row, is_smeta=False)
        score = Decimal(str(CascadeMatcher.calculate_score(smeta_item, rd_item)))
        if score >= MIN_CANDIDATE_SCORE:
            scored.append((rd_item, score, rd_row))

    scored.sort(key=lambda entry: entry[1], reverse=True)
    return tuple(
        CandidateMatch(
            rd_item_id=int(rd_row["id"]),
            score=score,
            description=str(rd_row["description"] or ""),
            unit=str(rd_row["unit"] or ""),
            quantity=_to_decimal(rd_row["quantity"]) or Decimal("0"),
        )
        for _, score, rd_row in scored[:MAX_CANDIDATES]
    )


def _card_from_row(
    row: sqlite3.Row,
    *,
    index: int,
    total: int,
    candidates: Sequence[CandidateMatch] = (),
) -> ReviewCard:
    """Собирает карточку из строки ``PENDING_REVIEWS_SQL``."""
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


# ----------------------------------------------------------------------
# 2. Отрисовка карточки сравнения
# ----------------------------------------------------------------------


def _field_table() -> Table:
    """Создает сетку «поле — значение» для блоков карточки."""
    table = Table.grid(padding=(0, 2))
    table.add_column(style="bold cyan", no_wrap=True)
    table.add_column(overflow="fold")
    return table


def _render_rd_block(card: ReviewCard) -> Panel:
    """Блок рабочей документации (спецификации)."""
    if card.rd_description is None:
        return Panel(
            Text("Позиция РД отсутствует: сметная позиция без пары", style="yellow"),
            title="РД (спецификация)",
            border_style="yellow",
            padding=(0, 1),
        )
    table = _field_table()
    table.add_row("Наименование", card.rd_description)
    table.add_row("Шифр", card.rd_code or "—")
    table.add_row(
        "Объем",
        f"{_format_quantity(card.rd_quantity)} {card.rd_unit or ''}".strip(),
    )
    return Panel(
        table, title="РД (спецификация)", border_style="cyan", padding=(0, 1)
    )


def _render_smeta_block(card: ReviewCard) -> Panel:
    """Блок позиции локальной сметы."""
    if card.smeta_description is None:
        return Panel(
            Text("Позиция сметы отсутствует: обратный проход по РД", style="yellow"),
            title="Смета",
            border_style="yellow",
            padding=(0, 1),
        )
    table = _field_table()
    table.add_row("Шифр расценки", card.smeta_code or "—")
    table.add_row("Наименование", card.smeta_description)
    table.add_row(
        "Объем",
        f"{_format_quantity(card.smeta_quantity)} {card.smeta_unit or ''}".strip(),
    )
    table.add_row("Цена за единицу", _format_money(card.unit_price))
    table.add_row("Стоимость", _format_money(card.total_cost))
    return Panel(table, title="Смета", border_style="green", padding=(0, 1))


def _render_calculation_block(card: ReviewCard) -> Panel:
    """Расчетный блок: коэффициент приведения и дельта объема."""
    table = _field_table()
    table.add_row("Коэффициент приведения", _format_ratio(card.conversion_ratio))
    if not card.is_paired:
        table.add_row("Дельта объема", "—")
        text, style = _verdict(card)
        table.add_row("Вердикт", Text(text, style=style))
    else:
        table.add_row(
            "Норматив отходов",
            _format_percent(card.waste_coefficient * PERCENT),
        )
        table.add_row(
            "Объем РД с отходами",
            _format_quantity(card.adjusted_rd_volume()),
        )
        table.add_row("Дельта объема", _format_delta(card.delta_volume()))
        table.add_row("Допуск", _format_quantity(card.allowed_delta()))
        text, style = _verdict(card)
        table.add_row("Вердикт", Text(text, style=style))
        delta_cost = card.delta_cost()
        if delta_cost is not None:
            table.add_row("Дельта стоимости", _format_money(delta_cost))
    return Panel(table, title="Расчет", border_style="magenta", padding=(0, 1))


def _render_candidates_block(card: ReviewCard) -> Optional[Panel]:
    """Блок top-3 кандидатов РД с оценкой сходства."""
    if not card.candidates:
        return None
    lines = Text()
    for position, candidate in enumerate(card.candidates, start=1):
        if position > 1:
            lines.append("\n")
        lines.append(f"[{position}] ", style="bold cyan")
        lines.append(f"{candidate.score:.1f}%", style="green")
        lines.append(
            f" | {candidate.description} | "
            f"{_format_quantity(candidate.quantity)} {candidate.unit}".rstrip()
        )
    return Panel(
        lines,
        title=f"Кандидаты РД (top-{len(card.candidates)})",
        border_style="blue",
        padding=(0, 1),
    )


def render_card(console: Console, card: ReviewCard) -> None:
    """Выводит карточку сравнения с цветовой градацией уверенности."""
    style = _confidence_style(card.confidence)
    title = Text.assemble(
        (f"[{card.index}/{card.total}] ", "bold white"),
        (f"Связь #{card.match_id}", "bold"),
        ("  ·  Confidence ", "white"),
        (f"{card.confidence:.1f}%", style),
    )
    body = Table.grid(padding=(0, 0))
    body.add_column()
    body.add_row(_render_rd_block(card))
    body.add_row(_render_smeta_block(card))
    candidates_block = _render_candidates_block(card)
    if candidates_block is not None:
        body.add_row(candidates_block)
    body.add_row(_render_calculation_block(card))
    console.print(Panel(body, title=title, border_style=style, padding=(0, 1)))
    if card.candidates:
        console.print(
            "[dim]1-3 / Enter — выбрать кандидата (Enter = №1) · "
            "A — принять текущую связь · R — отклонить (сохранив №1) · "
            "C — коэффициент · S — пропустить · Q — выход[/dim]"
        )
    elif card.is_paired:
        console.print(
            "[dim]A — принять · R — отклонить · C — коэффициент · "
            "S — пропустить · Q — выход[/dim]"
        )
    else:
        console.print(
            "[dim]A — подтвердить отсутствие пары · R — отклонить · "
            "S — пропустить · Q — выход[/dim]"
        )


# ----------------------------------------------------------------------
# 3. Диалог со сметчиком
# ----------------------------------------------------------------------


def _prompt_command(console: Console, *, candidate_count: int = 0) -> str:
    """Запрашивает команду до получения допустимого значения."""
    if candidate_count:
        prompt = (
            f"Команда [1-{candidate_count}/A/R/C/S/Q, Enter = 1]: "
        )
    else:
        prompt = "Команда [A/R/C/S/Q]: "
    while True:
        try:
            raw = input(prompt)
        except EOFError:
            return "Q"
        command = raw.strip().upper()
        if not command and candidate_count:
            return "1"
        if command in VALID_COMMANDS:
            return command
        console.print(
            "[red]Некорректная команда. Допустимо: "
            + ("1-3, " if candidate_count else "")
            + "A, R, C, S, Q.[/red]"
        )


def _prompt_reject(console: Console) -> Optional[tuple[str, Optional[str]]]:
    """Подменю отклонения: код причины и опциональный комментарий."""
    table = Table(title="Причина отклонения", header_style="bold")
    table.add_column("Код", justify="center", width=4)
    table.add_column("Причина", overflow="fold")
    for key, (_, label) in REJECT_REASONS.items():
        table.add_row(key, label)
    console.print(table)

    while True:
        try:
            raw = input("Код причины [1-4, Enter — отмена]: ").strip()
        except EOFError:
            return None
        if not raw:
            return None
        if raw in REJECT_REASONS:
            break
        console.print("[red]Некорректный код: введите 1, 2, 3 или 4.[/red]")

    try:
        comment = input("Комментарий (опционально): ").strip()
    except EOFError:
        comment = ""
    return REJECT_REASONS[raw][0], comment or None


def _prompt_ratio(console: Console) -> Optional[Decimal]:
    """Запрашивает новый коэффициент приведения (> 0)."""
    while True:
        try:
            raw = input("Новый коэффициент приведения (> 0, Enter — отмена): ")
        except EOFError:
            return None
        raw = raw.strip().replace(",", ".")
        if not raw:
            return None
        try:
            ratio = Decimal(raw)
        except InvalidOperation:
            console.print("[red]Не число. Пример: 1.25[/red]")
            continue
        if not ratio.is_finite() or ratio <= 0:
            console.print("[red]Коэффициент должен быть положительным числом.[/red]")
            continue
        return ratio


# ----------------------------------------------------------------------
# 4. Сохранение решений в SSOT
# ----------------------------------------------------------------------


def _apply_approval(
    conn: sqlite3.Connection,
    card: ReviewCard,
    reviewer: str,
) -> None:
    """Фиксирует подтверждение связи (для записи без пары — отсутствие пары)."""
    comment = None if card.is_paired else "Подтверждено отсутствие пары РД/смета"
    with db.transaction(conn):
        db.record_expert_review(
            conn,
            card.match_id,
            reviewer,
            "APPROVED",
            comment=comment,
        )


def _apply_candidate_selection(
    conn: sqlite3.Connection,
    card: ReviewCard,
    candidate: CandidateMatch,
    position: int,
    reviewer: str,
) -> None:
    """Фиксирует выбор сметчиком одного из top-3 кандидатов РД."""
    comment = f"Выбран кандидат №{position}: {candidate.description}"
    with db.transaction(conn):
        db.update_match_link(
            conn,
            card.match_id,
            candidate.rd_item_id,
            candidate.score,
        )
        db.record_expert_review(
            conn,
            card.match_id,
            reviewer,
            "APPROVED",
            comment=comment,
        )


def _apply_rejection(
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
    with db.transaction(conn):
        if card.candidates:
            top = card.candidates[0]
            db.update_match_link(
                conn,
                card.match_id,
                top.rd_item_id,
                top.score,
            )
        db.record_expert_review(
            conn,
            card.match_id,
            reviewer,
            "REJECTED",
            reason_code=reason_code,
            comment=comment,
        )


def _apply_adjustment(
    conn: sqlite3.Connection,
    card: ReviewCard,
    reviewer: str,
    ratio: Decimal,
) -> Optional[Decimal]:
    """Сохраняет скорректированный коэффициент и новую дельту в аудит."""
    delta = card.delta_volume(ratio)
    comment = (
        f"Коэффициент приведения: {_format_ratio(card.conversion_ratio)} -> "
        f"{_format_ratio(ratio)}; дельта объема: {_format_delta(delta)}"
    )
    with db.transaction(conn):
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
# 5. Сессия, очередь и финализация
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


def _finalize_session(conn: sqlite3.Connection, session_id: str) -> str:
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


def _render_summary(console: Console, outcome: ReviewOutcome) -> None:
    """Выводит итоговую панель сеанса ревью."""
    table = Table.grid(padding=(0, 2))
    table.add_column(style="bold cyan")
    table.add_column(justify="right")
    table.add_row("Всего в очереди", str(outcome.total))
    table.add_row("Принято", str(outcome.approved))
    table.add_row("Отклонено", str(outcome.rejected))
    table.add_row("Скорректировано", str(outcome.adjusted))
    table.add_row("Пропущено", str(outcome.skipped))
    table.add_row("Осталось в очереди", str(outcome.remaining))
    table.add_row("Статус сессии", outcome.session_status)
    style = "green" if outcome.remaining == 0 else "yellow"
    console.print(
        Panel(table, title="Ревью завершено", border_style=style, padding=(0, 1))
    )


def run_review(
    conn: sqlite3.Connection,
    session_id: str,
    reviewer: str = DEFAULT_REVIEWER,
    *,
    console: Optional[Console] = None,
) -> ReviewOutcome:
    """Обрабатывает очередь HITL и фиксирует решения сметчика в SSOT.

    Очередь фиксируется на входе: пропущенные (``S``) и необработанные из-за
    выхода (``Q``) записи остаются в статусе ``HITL_PENDING``. Сессия
    переводится в ``REVIEWED`` только когда незакрытых связей не осталось.
    """
    if not reviewer or not reviewer.strip():
        raise ValueError("Имя сметчика не может быть пустым")
    reviewer = reviewer.strip()
    console = console or Console()

    pending = db.get_pending_reviews(conn, session_id)
    if not pending:
        console.print(
            Panel(
                "Очередь HITL пуста. Все позиции верифицированы",
                title="Готово",
                border_style="green",
                padding=(0, 1),
            )
        )
        session_status = _finalize_session(conn, session_id)
        return ReviewOutcome(
            session_id=session_id,
            total=0,
            approved=0,
            rejected=0,
            adjusted=0,
            skipped=0,
            remaining=0,
            session_status=session_status,
        )

    total = len(pending)
    approved = 0
    rejected = 0
    adjusted = 0
    skipped = 0

    for index, row in enumerate(pending, start=1):
        card = _card_from_row(
            row,
            index=index,
            total=total,
            candidates=_candidates_for_row(conn, row),
        )
        render_card(console, card)
        quit_requested = False

        while True:
            command = _prompt_command(console, candidate_count=len(card.candidates))
            if command == "Q":
                quit_requested = True
                break
            if command == "S":
                skipped += 1
                console.print("[yellow]Пропущено: статус записи не изменен.[/yellow]")
                break
            if command == "A":
                _apply_approval(conn, card, reviewer)
                approved += 1
                console.print("[green]Принято: связь зафиксирована.[/green]")
                break
            if command == "R":
                rejection = _prompt_reject(console)
                if rejection is None:
                    console.print("[yellow]Отклонение отменено.[/yellow]")
                    continue
                reason_code, comment = rejection
                _apply_rejection(conn, card, reviewer, reason_code, comment)
                rejected += 1
                console.print(
                    "[red]Отклонено: причина сохранена, лучший кандидат РД "
                    "оставлен для трассируемости.[/red]"
                )
                break
            if command in {"1", "2", "3"}:
                position = int(command)
                if position > len(card.candidates):
                    console.print(
                        f"[red]Кандидат №{position} отсутствует: "
                        f"доступно {len(card.candidates)}.[/red]"
                    )
                    continue
                _apply_candidate_selection(
                    conn,
                    card,
                    card.candidates[position - 1],
                    position,
                    reviewer,
                )
                approved += 1
                console.print(
                    f"[green]Принят кандидат №{position}: связь зафиксирована.[/green]"
                )
                break

            # Осталась только команда C (корректировка коэффициента).
            if not card.is_paired:
                console.print(
                    "[yellow]Корректировка недоступна: нет пары РД/смета.[/yellow]"
                )
                continue
            new_ratio = _prompt_ratio(console)
            if new_ratio is None:
                console.print("[yellow]Корректировка отменена.[/yellow]")
                continue
            new_delta = _apply_adjustment(conn, card, reviewer, new_ratio)
            adjusted += 1
            console.print(
                "[magenta]Коэффициент сохранен, новая дельта: "
                f"{_format_delta(new_delta)}[/magenta]"
            )
            break

        if quit_requested:
            console.print(
                Panel(
                    "Выход. Прогресс сохранен.",
                    title="Выход",
                    border_style="yellow",
                    padding=(0, 1),
                )
            )
            break

    remaining = len(db.get_pending_reviews(conn, session_id))
    session_status = _finalize_session(conn, session_id)
    outcome = ReviewOutcome(
        session_id=session_id,
        total=total,
        approved=approved,
        rejected=rejected,
        adjusted=adjusted,
        skipped=skipped,
        remaining=remaining,
        session_status=session_status,
    )
    _render_summary(console, outcome)
    return outcome


# ----------------------------------------------------------------------
# 6. Точка входа CLI
# ----------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    """Создает парсер аргументов командной строки."""
    parser = argparse.ArgumentParser(
        prog="cli_review.py",
        description=(
            "Интерактивное ревью очереди HITL для сметчика "
            "(SQLite как единственный источник истины)"
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--db",
        dest="db_path",
        default=db.DEFAULT_DB_PATH,
        help="Путь к файлу SQLite (SSOT)",
    )
    parser.add_argument(
        "--session",
        default=LATEST_SESSION,
        help="Идентификатор сессии либо 'latest' для последней созданной",
    )
    parser.add_argument(
        "--reviewer",
        default=DEFAULT_REVIEWER,
        help="Имя или табельный номер сметчика для expert_reviews",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Запускает интерактивное ревью; возвращает код выхода процесса."""
    args = build_parser().parse_args(argv)
    console = Console()

    if not Path(args.db_path).is_file():
        console.print(
            Panel(
                f"Файл БД не найден: {args.db_path}",
                title="Сбой",
                border_style="red",
            )
        )
        return 2

    conn: Optional[sqlite3.Connection] = None

    try:
        conn = db.get_connection(args.db_path)
        db.init_db(conn)
        try:
            session_id = resolve_session_id(conn, args.session)
        except LookupError as exc:
            console.print(
                Panel(str(exc), title="Сессия не найдена", border_style="red")
            )
            return 2
        run_review(conn, session_id, reviewer=args.reviewer, console=console)
    except sqlite3.Error as exc:
        console.print(Panel(f"Ошибка SQLite: {exc}", title="Сбой", border_style="red"))
        return 1
    finally:
        if conn is not None:
            conn.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())





