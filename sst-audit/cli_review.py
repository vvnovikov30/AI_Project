"""Интерактивный терминальный интерфейс сметчика (HITL) на базе Rich.

Тонкий слой представления: вся доменная логика (очередь, расчет дельт,
запись решений и журнала арбитража) живет в ``review_service``. Модуль
отвечает только за отрисовку карточки, диалог с оператором и CLI-аргументы.

Запуск::

    python cli_review.py --db sst_audit.db --session latest --reviewer ivanov
"""

from __future__ import annotations

import argparse
from collections.abc import Sequence
from decimal import Decimal, InvalidOperation
from pathlib import Path
import sqlite3
from typing import Final, Optional

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

import db
import review_service
from review_service import (
    CONFIDENCE_WARN_THRESHOLD,
    DEFAULT_REVIEWER,
    LATEST_SESSION,
    PERCENT,
    REJECT_REASONS,
    VALID_COMMANDS,
    CandidateMatch,
    ReviewCard,
    ReviewOutcome,
)

# Константы и модели домена реэкспортируются из review_service: внешний код
# (тесты, скрипты) продолжает импортировать их из cli_review без изменений.
CLOSED_SESSION_STATUSES: Final[frozenset[str]] = (
    review_service.CLOSED_SESSION_STATUSES
)


# ----------------------------------------------------------------------
# 1. Приведение типов и форматирование
# ----------------------------------------------------------------------


def _confidence_style(score: Decimal) -> str:
    """Цвет заголовка карточки: красный для низкой уверенности."""
    return "red" if score < CONFIDENCE_WARN_THRESHOLD else "yellow"


# Доменные хелперы — единые реализации из review_service (без дублирования).
_to_decimal = review_service._to_decimal
_optional_decimal = review_service._optional_decimal
_text_or_none = review_service._text_or_none
_format_quantity = review_service.format_quantity
_format_ratio = review_service.format_ratio
_format_money = review_service.format_money
_format_percent = review_service.format_percent
_format_delta = review_service.format_delta
_verdict = review_service.verdict
_payload_from_row = review_service._payload_from_row
_normalized_from_row = review_service._normalized_from_row
_candidates_for_row = review_service.candidates_for_row
_card_from_row = review_service._card_from_row


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


_apply_approval = review_service.apply_approval
_apply_candidate_selection = review_service.apply_candidate_selection
_apply_rejection = review_service.apply_rejection
_apply_adjustment = review_service.apply_adjustment


# ----------------------------------------------------------------------
# 5. Сессия, очередь и финализация
# ----------------------------------------------------------------------


resolve_session_id = review_service.resolve_session_id
_finalize_session = review_service.finalize_session


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

    # Пакетная загрузка: 3 SQL-запроса независимо от длины очереди (без N+1).
    cards = review_service.load_pending_cards(conn, session_id)
    if not cards:
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

    total = len(cards)
    approved = 0
    rejected = 0
    adjusted = 0
    skipped = 0

    for card in cards:
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

    remaining = review_service.pending_count(conn, session_id)
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





