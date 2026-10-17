"""Оркестратор sst-audit: SQLite как единственный источник истины (SSOT).

Конвейер:
1. INGEST   — парсинг РД (PDF) и сметы (XLSX), нормализация и атомарная
              запись позиций в ``rd_items`` / ``smeta_items``.
2. MATCHING — каскадное сопоставление, запись связей в ``match_records``
              (``AUTO`` при уверенности >= ``AUTO_MATCH_THRESHOLD`` и
              отсутствии неоднозначности, иначе ``HITL_PENDING``).
3. RULES    — оценка количественных расхождений в ``rule_evaluations``.
4. STATUS   — сессия переводится в ``MATCHED`` (есть очередь HITL) либо
              ``REVIEWED`` (очередь пуста).

XML-отчет не формируется: единственный источник истины — SQLite, выгрузка
отчета выполняется из представления ``v_audit_export``.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
import sqlite3
from typing import Any, Final, Mapping, Optional, Sequence
import uuid

import db
from evaluator import DiscrepancyDetector
from matcher import CascadeMatcher
from models import CollisionSeverity, NormalizedItem
from parsers import LSRParser, PDFSpecParser

# Порог автоматического принятия связи (шкала матчера 0-100).
AUTO_MATCH_THRESHOLD: Final[Decimal] = Decimal("70.0")
# Минимальный отрыв лидера от второго кандидата; иначе связь неоднозначна.
AMBIGUITY_MARGIN: Final[Decimal] = Decimal("5.0")
DEFAULT_CONVERSION_RATIO: Final[Decimal] = db.DEFAULT_CONVERSION_RATIO

MATCH_METHOD_CASCADE: Final[str] = "FUZZY_CASCADE"
MATCH_METHOD_NO_CANDIDATE: Final[str] = "NO_CANDIDATE"
MATCH_METHOD_REVERSE: Final[str] = "REVERSE_PASS"

RULE_BY_SEVERITY: Final[dict[CollisionSeverity, str]] = {
    CollisionSeverity.CRITICAL: "VOLUME_DELTA_CRITICAL",
    CollisionSeverity.SUSPECT: "VOLUME_DELTA_SUSPECT",
    CollisionSeverity.INCOMPATIBLE: "UNIT_INCOMPATIBLE",
    CollisionSeverity.INFO: "VOLUME_WITHIN_TOLERANCE",
}


@dataclass(frozen=True)
class SessionConfig:
    """Конфигурация сессии аудита."""

    project_name: str
    rd_path: str
    smeta_path: str
    db_path: str = db.DEFAULT_DB_PATH
    session_id: Optional[str] = None

    def resolved_session_id(self) -> str:
        """Возвращает заданный идентификатор либо генерирует новый."""
        return self.session_id or uuid.uuid4().hex


@dataclass(frozen=True)
class MatchDecision:
    """Решение матчера по одной позиции сметы либо РД."""

    smeta_item: Optional[NormalizedItem]
    rd_item: Optional[NormalizedItem]
    confidence: Decimal
    status: str
    method: str
    candidates: tuple[tuple[NormalizedItem, float], ...] = ()


@dataclass(frozen=True)
class PipelineResult:
    """Сводка выполненного прогона."""

    session_id: str
    session_status: str
    rd_saved: int
    smeta_saved: int
    matches_saved: int
    evaluations_saved: int
    auto_matches: int
    hitl_pending: int


def _candidates(
    smeta_item: NormalizedItem,
    rd_items: Sequence[NormalizedItem],
) -> list[tuple[NormalizedItem, float]]:
    """Возвращает кандидатов РД с ненулевым сходством по убыванию score."""
    scored = [
        (rd_item, CascadeMatcher.calculate_score(smeta_item, rd_item))
        for rd_item in rd_items
    ]
    return sorted(
        ((item, score) for item, score in scored if score > 0.0),
        key=lambda pair: pair[1],
        reverse=True,
    )


def match_items(
    rd_items: Sequence[NormalizedItem],
    smeta_items: Sequence[NormalizedItem],
) -> list[MatchDecision]:
    """Сопоставляет позиции сметы с РД и классифицирует связи.

    * ``AUTO`` — лидер >= порога и отрыв от второго кандидата не меньше
      ``AMBIGUITY_MARGIN``;
    * ``HITL_PENDING`` — низкая уверенность, неоднозначность или отсутствие
      кандидатов;
    * обратный проход добавляет ``HITL_PENDING`` для позиций РД, не нашедших
      отражения в смете.
    """
    decisions: list[MatchDecision] = []
    covered_rd: set[int] = set()

    for smeta_item in smeta_items:
        candidates = _candidates(smeta_item, rd_items)
        best_item: Optional[NormalizedItem] = None
        best_score = Decimal("0")
        ambiguous = False

        if candidates:
            best_item, raw_score = candidates[0]
            best_score = Decimal(str(raw_score))
            if len(candidates) > 1:
                runner_up = Decimal(str(candidates[1][1]))
                ambiguous = (best_score - runner_up) < AMBIGUITY_MARGIN

        is_auto = (
            best_item is not None
            and best_score >= AUTO_MATCH_THRESHOLD
            and not ambiguous
        )
        if best_item is not None:
            covered_rd.add(id(best_item))

        decisions.append(
            MatchDecision(
                smeta_item=smeta_item,
                rd_item=best_item,
                confidence=best_score,
                status="AUTO" if is_auto else "HITL_PENDING",
                method=(
                    MATCH_METHOD_CASCADE
                    if best_item is not None
                    else MATCH_METHOD_NO_CANDIDATE
                ),
                candidates=tuple(candidates[:3]),
            )
        )

    for rd_item in rd_items:
        if id(rd_item) in covered_rd:
            continue
        decisions.append(
            MatchDecision(
                smeta_item=None,
                rd_item=rd_item,
                confidence=Decimal("0"),
                status="HITL_PENDING",
                method=MATCH_METHOD_REVERSE,
            )
        )

    return decisions


def build_rule_evaluations(
    decisions: Sequence[MatchDecision],
    match_ids: Sequence[int],
) -> list[dict[str, Any]]:
    """Готовит записи ``rule_evaluations`` для сохранения в SSOT.

    Для связанных пар применяется ``DiscrepancyDetector``; для позиций без
    пары фиксируются правила ``RD_NOT_IN_SMETA`` / ``SMETA_NOT_IN_RD``.
    """
    if len(decisions) != len(match_ids):
        raise RuntimeError(
            f"Количество решений ({len(decisions)}) не совпадает с количеством "
            f"записей match_records ({len(match_ids)})"
        )

    evaluations: list[dict[str, Any]] = []
    for decision, match_id in zip(decisions, match_ids):
        if decision.smeta_item is not None and decision.rd_item is not None:
            collision = DiscrepancyDetector.evaluate(
                decision.smeta_item,
                decision.rd_item,
                decision.confidence,
            )
            evaluations.append(
                {
                    "match_record_id": match_id,
                    "rule_name": RULE_BY_SEVERITY[collision.severity],
                    "is_violation": collision.severity is not CollisionSeverity.INFO,
                    "delta_volume": collision.volume_delta,
                    "delta_cost": None,
                    "message": collision.description,
                }
            )
        elif decision.rd_item is not None:
            evaluations.append(
                {
                    "match_record_id": match_id,
                    "rule_name": "RD_NOT_IN_SMETA",
                    "is_violation": True,
                    "delta_volume": decision.rd_item.volume,
                    "delta_cost": None,
                    "message": "Позиция РД не найдена в смете",
                }
            )
        elif decision.smeta_item is not None:
            evaluations.append(
                {
                    "match_record_id": match_id,
                    "rule_name": "SMETA_NOT_IN_RD",
                    "is_violation": True,
                    "delta_volume": -decision.smeta_item.volume,
                    "delta_cost": None,
                    "message": "Позиция сметы не подтверждена РД",
                }
            )
    return evaluations


def _load_item_ids(
    conn: sqlite3.Connection,
    table: str,
    session_id: str,
    expected: int,
) -> list[int]:
    """Возвращает id вставленных позиций в порядке вставки (id ASC)."""
    rows = conn.execute(
        f"SELECT id FROM {table} WHERE session_id = ? ORDER BY id ASC",
        (session_id,),
    ).fetchall()
    ids = [int(row["id"]) for row in rows]
    if len(ids) != expected:
        raise RuntimeError(
            f"Ожидалось {expected} записей в {table} для сессии {session_id}, "
            f"найдено {len(ids)}"
        )
    return ids


def resolve_pending_reviews(
    conn: sqlite3.Connection,
    session_id: str,
    match_rows: Sequence[sqlite3.Row],
    decisions: Sequence[MatchDecision],
    rd_id_by_object: Mapping[int, int],
    reviewer: str,
) -> list[MatchDecision]:
    """Интерактивный HITL-цикл: фиксирует решения эксперта по связям.

    Возвращает решения в исходном порядке, обновленные выбранными
    кандидатами: последующая оценка правил выполняется уже по финальной паре.
    Записи без кандидатов остаются в очереди для экспертной проверки
    (в терминале выбирать не из чего).
    """
    index_by_match_id = {
        int(row["id"]): index for index, row in enumerate(match_rows)
    }
    resolved = list(decisions)

    for row in db.get_pending_reviews(conn, session_id):
        match_id = int(row["match_id"])
        index = index_by_match_id.get(match_id)
        if index is None:
            continue
        decision = decisions[index]
        if decision.smeta_item is None or not decision.candidates:
            continue

        selected = _prompt_operator_review(decision.smeta_item, decision.candidates)
        if selected is None:
            db.update_match_link(conn, match_id, None, Decimal("0"))
            resolved[index] = MatchDecision(
                smeta_item=decision.smeta_item,
                rd_item=None,
                confidence=Decimal("0"),
                status="RESOLVED",
                method=MATCH_METHOD_NO_CANDIDATE,
                candidates=decision.candidates,
            )
            db.record_expert_review(
                conn,
                match_id,
                reviewer,
                "REJECTED",
                reason_code="NO_MATCH_CONFIRMED",
                comment="Оператор отклонил все кандидаты",
            )
        else:
            selected_score = Decimal(
                str(
                    next(
                        score
                        for candidate, score in decision.candidates
                        if candidate is selected
                    )
                )
            )
            db.update_match_link(
                conn,
                match_id,
                rd_id_by_object.get(id(selected)),
                selected_score,
            )
            resolved[index] = MatchDecision(
                smeta_item=decision.smeta_item,
                rd_item=selected,
                confidence=selected_score,
                status="RESOLVED",
                method=MATCH_METHOD_CASCADE,
                candidates=decision.candidates,
            )
            db.record_expert_review(
                conn,
                match_id,
                reviewer,
                "APPROVED",
                comment=f"Подтверждена позиция {selected.raw_id}",
            )

    return resolved


def _prompt_operator_review(
    smeta_item: NormalizedItem,
    candidates: Sequence[tuple[NormalizedItem, float]],
) -> Optional[NormalizedItem]:
    """Интерактивный Human-in-the-Loop интерфейс в терминале."""
    print("\n" + "=" * 70)
    print("🔍 [HUMAN-IN-THE-LOOP] ТРЕБУЕТСЯ ВЕРИФИКАЦИЯ СОПОСТАВЛЕНИЯ")
    print("=" * 70)
    print("📌 Позиция сметы:")
    print(f"   • Наименование : {smeta_item.raw_title}")
    print(f"   • Шифр/Код     : {smeta_item.code or '—'}")
    print(
        f"   • Объем        : {smeta_item.volume} "
        f"{smeta_item.canonical_unit.value} "
        f"(ячейка {smeta_item.pointer.coordinate})"
    )
    print("\n🎯 Найденные кандидаты из РД:")

    for idx, (rd_item, score) in enumerate(candidates, start=1):
        print(f"   [{idx}] Сходство: {score:.1f}%")
        print(f"       • РД    : {rd_item.raw_title}")
        print(f"       • Код   : {rd_item.code or '—'}")
        print(
            f"       • Объем : {rd_item.volume} {rd_item.canonical_unit.value} "
            f"({rd_item.pointer.coordinate})"
        )

    print("   [0] Ни один вариант не подходит (Отклонить все связи)")

    while True:
        choice = input(
            "\n👉 Выберите номер верного соответствия [0-N] (Enter = 1): "
        ).strip()
        if choice == "" and candidates:
            choice = "1"

        if choice == "0":
            print("❌ Связь отклонена оператором.")
            return None
        if choice.isdigit() and 1 <= int(choice) <= len(candidates):
            selected = candidates[int(choice) - 1][0]
            print(
                "✅ Подтверждена связь с позицией: "
                f"{selected.code or selected.raw_title[:40]}"
            )
            return selected

        print("⚠️ Некорректный ввод, попробуйте еще раз.")


def run_pipeline(
    conn: sqlite3.Connection,
    config: SessionConfig,
    rd_items: Optional[Sequence[NormalizedItem]] = None,
    smeta_items: Optional[Sequence[NormalizedItem]] = None,
    *,
    interactive: bool = False,
    reviewer: str = "Оператор ПТО",
) -> PipelineResult:
    """Выполняет полный цикл ingest -> matching -> rules -> status.

    :param conn: соединение SQLite; схема инициализируется идемпотентно.
    :param config: конфигурация сессии (проект, пути, db_path).
    :param rd_items: заранее подготовленные позиции РД (иначе парсятся из PDF).
    :param smeta_items: заранее подготовленные позиции сметы (иначе из XLSX).
    :param interactive: запускать терминальный HITL-опрос по очереди.
    :param reviewer: имя эксперта для записей ``expert_reviews``.
    """
    db.init_db(conn)
    session_id = config.resolved_session_id()

    if rd_items is None:
        rd_items = PDFSpecParser(config.rd_path).extract_items()
    if smeta_items is None:
        smeta_items = LSRParser(config.smeta_path).parse()

    rd_items = list(rd_items)
    smeta_items = list(smeta_items)

    # --- ЭТАП INGEST: сессия и позиции сохраняются атомарно -----------------
    with db.transaction(conn):
        db.create_session(
            conn,
            session_id,
            config.project_name,
            config.rd_path,
            config.smeta_path,
        )
        rd_saved = db.save_rd_items(
            conn,
            session_id,
            [item.to_rd_record() for item in rd_items],
        )
        smeta_saved = db.save_smeta_items(
            conn,
            session_id,
            [item.to_smeta_record() for item in smeta_items],
        )

    rd_ids = _load_item_ids(conn, "rd_items", session_id, rd_saved)
    smeta_ids = _load_item_ids(conn, "smeta_items", session_id, smeta_saved)
    rd_id_by_object = {id(item): item_id for item, item_id in zip(rd_items, rd_ids)}
    smeta_id_by_object = {
        id(item): item_id for item, item_id in zip(smeta_items, smeta_ids)
    }

    # --- ЭТАП MATCHING ------------------------------------------------------
    decisions = match_items(rd_items, smeta_items)
    db.save_matches(
        conn,
        session_id,
        [
            {
                "rd_item_id": (
                    rd_id_by_object.get(id(decision.rd_item))
                    if decision.rd_item is not None
                    else None
                ),
                "smeta_item_id": (
                    smeta_id_by_object.get(id(decision.smeta_item))
                    if decision.smeta_item is not None
                    else None
                ),
                "confidence_score": decision.confidence,
                "conversion_ratio": DEFAULT_CONVERSION_RATIO,
                "match_method": decision.method,
                "status": decision.status,
            }
            for decision in decisions
        ],
    )

    match_rows = db.get_match_records(conn, session_id)
    if len(match_rows) != len(decisions):
        raise RuntimeError(
            f"Сохранено {len(match_rows)} связей, ожидалось {len(decisions)}"
        )

    # --- ЭТАП HITL (опционально) -------------------------------------------
    if interactive:
        decisions = resolve_pending_reviews(
            conn,
            session_id,
            match_rows,
            decisions,
            rd_id_by_object,
            reviewer,
        )

    # --- ЭТАП RULES ---------------------------------------------------------
    match_ids = [int(row["id"]) for row in match_rows]
    evaluations = build_rule_evaluations(decisions, match_ids)
    evaluations_saved = db.save_rule_evaluations(conn, evaluations)

    # --- ЭТАП STATUS --------------------------------------------------------
    pending = db.get_pending_reviews(conn, session_id)
    session_status = "MATCHED" if pending else "REVIEWED"
    db.update_session_status(conn, session_id, session_status)

    return PipelineResult(
        session_id=session_id,
        session_status=session_status,
        rd_saved=rd_saved,
        smeta_saved=smeta_saved,
        matches_saved=len(match_rows),
        evaluations_saved=evaluations_saved,
        auto_matches=sum(1 for d in decisions if d.status == "AUTO"),
        hitl_pending=len(pending),
    )


def export_audit_view(
    conn: sqlite3.Connection,
    session_id: str,
    output_path: str | Path,
) -> int:
    """Выгружает отчет из ``v_audit_export`` в Excel.

    Возвращает количество выгруженных строк. XML-отчет не формируется.
    """
    import pandas as pd

    frame = pd.read_sql_query(
        "SELECT * FROM v_audit_export WHERE session_id = ? ORDER BY match_id",
        conn,
        params=(session_id,),
    )
    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    frame.to_excel(target, index=False)
    return int(len(frame))


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Точка входа CLI: конфигурация сессии и запуск пайплайна."""
    parser = argparse.ArgumentParser(
        description="Аудит смет (Excel) и РД (PDF) с SQLite как SSOT"
    )
    parser.add_argument(
        "--project",
        "--project-name",
        dest="project_name",
        required=True,
        help="Наименование проекта/объекта",
    )
    parser.add_argument(
        "--rd",
        required=True,
        help="Путь к рабочей документации / спецификации (.pdf)",
    )
    parser.add_argument(
        "--smeta",
        required=True,
        help="Путь к локальной смете (.xlsx)",
    )
    parser.add_argument(
        "--db",
        dest="db_path",
        default=db.DEFAULT_DB_PATH,
        help="Путь к файлу SQLite (SSOT)",
    )
    parser.add_argument("--session-id", default=None, help="Идентификатор сессии")
    parser.add_argument(
        "--out",
        default=None,
        help="Опциональная Excel-выгрузка из v_audit_export",
    )
    parser.add_argument(
        "--interactive",
        action="store_true",
        help="Включить интерактивный HITL-опрос оператора",
    )
    parser.add_argument(
        "--reviewer",
        default="Оператор ПТО",
        help="Имя эксперта для записей expert_reviews",
    )
    args = parser.parse_args(argv)

    for flag, path in (("--rd", args.rd), ("--smeta", args.smeta)):
        if not Path(path).is_file():
            parser.error(f"Файл не найден: {flag} {path}")

    config = SessionConfig(
        project_name=args.project_name,
        rd_path=args.rd,
        smeta_path=args.smeta,
        db_path=args.db_path,
        session_id=args.session_id,
    )

    conn = db.get_connection(args.db_path)
    try:
        result = run_pipeline(
            conn,
            config,
            interactive=args.interactive,
            reviewer=args.reviewer,
        )
        print(
            f"[+] Сессия {result.session_id} завершена "
            f"со статусом {result.session_status}"
        )
        print(
            f"    РД: {result.rd_saved}, смета: {result.smeta_saved}, "
            f"связей: {result.matches_saved}, "
            f"правил: {result.evaluations_saved}"
        )
        print(
            f"    AUTO: {result.auto_matches}, "
            f"HITL_PENDING: {result.hitl_pending}"
        )
        if args.out:
            exported = export_audit_view(conn, result.session_id, args.out)
            print(f"    Excel-выгрузка: {args.out} (строк: {exported})")
    finally:
        conn.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
