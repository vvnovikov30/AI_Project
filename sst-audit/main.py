import os
import argparse
from decimal import Decimal
from typing import List, Optional, Set, Tuple, Dict
import pandas as pd

from models import NormalizedItem, CollisionRecord, CollisionSeverity
from parsers import LSRParser, PDFSpecParser
from matcher import CascadeMatcher
from evaluator import DiscrepancyDetector


class AuditOrchestrator:
    """Координатор конвейера сверки сметы и проектной спецификации с поддержкой двухфазной реконсиляции."""

    FUZZY_ACCEPTANCE_THRESHOLD = 85.0
    FUZZY_SUSPECT_THRESHOLD = 50.0

    def __init__(self, lsr_path: str, spec_path: str, interactive: bool = True):
        self.lsr_path = lsr_path
        self.spec_path = spec_path
        self.interactive = interactive
        self.lsr_parser = LSRParser(lsr_path)
        self.spec_parser = PDFSpecParser(spec_path)

    def _prompt_operator_review(
        self, 
        lsr: NormalizedItem, 
        candidates: List[Tuple[NormalizedItem, float]]
    ) -> Optional[NormalizedItem]:
        """Интерактивный Human-in-the-Loop интерфейс в терминале."""
        print("\n" + "=" * 70)
        print("🔍 [HUMAN-IN-THE-LOOP] ТРЕБУЕТСЯ ВЕРИФИКАЦИЯ СОПОСТАВЛЕНИЯ")
        print("=" * 70)
        print(f"📌 Позиция сметы:")
        print(f"   • Наименование : {lsr.raw_title}")
        print(f"   • Шифр/Код     : {lsr.code or '—'}")
        print(f"   • Объем        : {lsr.volume} {lsr.canonical_unit.value} (ячейка {lsr.pointer.coordinate})")
        print("\n🎯 Найденные кандидаты из проекта:")

        for idx, (spec, score) in enumerate(candidates, start=1):
            print(f"   [{idx}] Сходство: {score:.1f}%")
            print(f"       • Проект: {spec.raw_title}")
            print(f"       • Код   : {spec.code or '—'}")
            print(f"       • Объем : {spec.volume} {spec.canonical_unit.value} ({spec.pointer.coordinate})")

        print("   [0] Ни один вариант не подходит (Отклонить все связи)")

        while True:
            choice = input("\n👉 Выберите номер верного соответствия [0-N] (Enter = 1): ").strip()
            if choice == "" and len(candidates) > 0:
                choice = "1"
            
            if choice == "0":
                print("❌ Связь отклонена оператором.")
                return None
            elif choice.isdigit() and 1 <= int(choice) <= len(candidates):
                selected = candidates[int(choice) - 1][0]
                print(f"✅ Подтверждена связь с позицией: {selected.code or selected.raw_title[:40]}")
                return selected
            else:
                print("⚠️ Некорректный ввод, попробуйте еще раз.")

    def run_audit(self, output_excel: str):
        print(f"[*] Старт парсинга сметы: {self.lsr_path}")
        lsr_items = self.lsr_parser.parse()
        print(f"    Найдено позиций в смете: {len(lsr_items)}")

        print(f"[*] Старт извлечения спецификации: {self.spec_path}")
        spec_items = self.spec_parser.extract_items()
        print(f"    Найдено позиций в спецификации: {len(spec_items)}")

        # Разделение сметы на материальные ресурсы и расценки работ
        resource_items = [it for it in lsr_items if it.is_resource]
        work_items = [it for it in lsr_items if not it.is_resource]

        matched_spec_ids: Set[str] = set()
        # Реестр подтвержденных ресурсов: {lsr_raw_id: (spec_item, collision_record)}
        matched_resources: Dict[str, Tuple[NormalizedItem, CollisionRecord]] = {}
        collisions: List[CollisionRecord] = []

        # =========================================================================
        # ФАЗА 1: ФИЗИЧЕСКИЙ АУДИТ РЕСУРСОВ (Оборудование / Материалы <-> Проект)
        # =========================================================================
        print(f"\n[*] Фаза 1: Сверка материальных ресурсов ({len(resource_items)} позиций)...")
        for lsr in resource_items:
            scored_candidates: List[Tuple[NormalizedItem, float]] = []
            for spec in spec_items:
                score = CascadeMatcher.calculate_score(lsr, spec)
                if score >= self.FUZZY_SUSPECT_THRESHOLD:
                    scored_candidates.append((spec, score))

            scored_candidates.sort(key=lambda x: x[1], reverse=True)

            best_match: Optional[NormalizedItem] = scored_candidates[0][0] if scored_candidates else None
            best_score: float = scored_candidates[0][1] if scored_candidates else 0.0

            # Автоматическое принятие (>= 85.0%)
            if best_score >= self.FUZZY_ACCEPTANCE_THRESHOLD and best_match:
                matched_spec_ids.add(best_match.raw_id)
                record = DiscrepancyDetector.evaluate(lsr, best_match, Decimal(str(best_score)))
                matched_resources[lsr.raw_id] = (best_match, record)
                collisions.append(record)

            # Серая зона (70.0% - 84.9%) -> Human-in-the-Loop
            elif self.FUZZY_SUSPECT_THRESHOLD <= best_score < self.FUZZY_ACCEPTANCE_THRESHOLD and best_match:
                confirmed_spec: Optional[NormalizedItem] = None
                if self.interactive:
                    confirmed_spec = self._prompt_operator_review(lsr, scored_candidates[:3])

                if confirmed_spec:
                    matched_spec_ids.add(confirmed_spec.raw_id)
                    record = DiscrepancyDetector.evaluate(lsr, confirmed_spec, Decimal(str(best_score)))
                    record.description = f"[HITL Подтверждено] {record.description}"
                    matched_resources[lsr.raw_id] = (confirmed_spec, record)
                    collisions.append(record)
                else:
                    collisions.append(CollisionRecord(
                        lsr_item=lsr,
                        spec_item=best_match if not self.interactive else None,
                        similarity_score=Decimal(str(best_score)),
                        volume_delta=lsr.volume,
                        allowed_delta=Decimal("0"),
                        severity=CollisionSeverity.SUSPECT if not self.interactive else CollisionSeverity.CRITICAL,
                        description=(
                            f"Низкая уверенность совпадения ({best_score:.1f}%). Требуется ручная проверка"
                            if not self.interactive else
                            f"Отклонено оператором вручную (сходство было {best_score:.1f}%)"
                        )
                    ))

            # Полное несовпадение (< 70.0%)
            else:
                collisions.append(CollisionRecord(
                    lsr_item=lsr,
                    spec_item=None,
                    similarity_score=Decimal(str(best_score)),
                    volume_delta=lsr.volume,
                    allowed_delta=Decimal("0"),
                    severity=CollisionSeverity.CRITICAL,
                    description="Позиция сметы отсутствует в спецификации проекта (Необоснованный объем)"
                ))

        # =========================================================================
        # ФАЗА 2: ТОПОЛОГИЧЕСКАЯ ВАЛИДАЦИЯ РАСЦЕНОК РАБОТ (Anchor Resource Pattern)
        # =========================================================================
        print(f"[*] Фаза 2: Валидация монтажных расценок ({len(work_items)} позиций)...")
        for work in work_items:
            # Находим дочерние ресурсы текущей работы по parent_id
            children = [r for r in resource_items if r.parent_id == work.raw_id]

            if children:
                # Фильтруем подтвержденные в Фазе 1 ресурсы
                verified_children = [c for c in children if c.raw_id in matched_resources]

                if verified_children:
                    # Выбор ресурса-якоря (Anchor):
                    # Приоритет: ресурс с той же физической группой единиц (COUNT или LENGTH)
                    anchor = next(
                        (c for c in verified_children if c.canonical_unit == work.canonical_unit), 
                        verified_children[0]
                    )
                    spec_item, _ = matched_resources[anchor.raw_id]

                    # Проверка кратности объема монтажа и оборудования
                    if work.canonical_unit == spec_item.canonical_unit:
                        delta = work.volume - spec_item.volume
                        if delta == Decimal("0"):
                            severity = CollisionSeverity.INFO
                            desc = f"[Монтаж подтвержден] Привязано оборудование: {spec_item.clean_title[:45]}"
                        elif delta > Decimal("0"):
                            severity = CollisionSeverity.SUSPECT
                            desc = f"[Монтаж: завышение на +{delta} {work.canonical_unit.value}] Оборудование: {spec_item.clean_title[:35]}"
                        else:
                            severity = CollisionSeverity.SUSPECT
                            desc = f"[Монтаж: занижение на {delta} {work.canonical_unit.value}] Оборудование: {spec_item.clean_title[:35]}"
                    else:
                        # Разные размерности (применительные расценки: перемычки, комплекты)
                        delta = Decimal("0")
                        severity = CollisionSeverity.INFO
                        desc = f"[Монтаж подтвержден (применительная норма)] По ресурсу: {spec_item.clean_title[:40]}"

                    collisions.append(CollisionRecord(
                        lsr_item=work,
                        spec_item=spec_item,
                        similarity_score=Decimal("100.0"),
                        volume_delta=delta,
                        allowed_delta=Decimal("0"),
                        severity=severity,
                        description=desc
                    ))
                else:
                    # Дочерние ресурсы есть, но ни один не подтвержден проектом
                    collisions.append(CollisionRecord(
                        lsr_item=work,
                        spec_item=None,
                        similarity_score=Decimal("0.0"),
                        volume_delta=work.volume,
                        allowed_delta=Decimal("0"),
                        severity=CollisionSeverity.CRITICAL,
                        description="Расценка монтажа не подтверждена проектом (дочерние ресурсы отсутствуют в спецификации)"
                    ))
            else:
                # Автономная строительная расценка без дочерних ресурсов
                collisions.append(CollisionRecord(
                    lsr_item=work,
                    spec_item=None,
                    similarity_score=Decimal("0.0"),
                    volume_delta=work.volume,
                    allowed_delta=Decimal("0"),
                    severity=CollisionSeverity.SUSPECT,
                    description="Автономная расценка без дочерних ресурсов (требуется проверка по ведомости объемов работ / ВОР)"
                ))

        # =========================================================================
        # ОБРАТНЫЙ ПРОХОД: НЕОСМЕЧЕННЫЕ ПОЗИЦИИ ПРОЕКТА
        # =========================================================================
        print("[*] Анализ неосмеченных позиций спецификации...")
        for spec in spec_items:
            if spec.raw_id not in matched_spec_ids:
                collisions.append(CollisionRecord(
                    lsr_item=None,
                    spec_item=spec,
                    similarity_score=Decimal("0.0"),
                    volume_delta=-spec.volume,
                    allowed_delta=Decimal("0"),
                    severity=CollisionSeverity.CRITICAL,
                    description="Проектная позиция не осмечена (Пропущенный объем)"
                ))

        # Сортировка: сметные позиции в исходном порядке сметы (по номеру строки Excel),
        # неосмеченные проектные позиции — в конце таблицы
        collisions.sort(key=lambda c: c.lsr_item.pointer.row_idx if c.lsr_item else 999999)

        print(f"\n[*] Формирование отчета: {output_excel}")
        self._export_report(collisions, output_excel)
        print("[+] Аудит успешно завершен!")

    def _export_report(self, collisions: List[CollisionRecord], output_path: str):
        rows = []
        for c in collisions:
            rows.append({
                "Статус": c.severity.value,
                "Уверенность скоринга (%)": float(round(c.similarity_score, 1)),
                "Шифр сметы": c.lsr_item.code if c.lsr_item else "—",
                "Наименование в смете": c.lsr_item.raw_title if c.lsr_item else "—",
                "Координата ячейки сметы": c.lsr_item.pointer.coordinate if c.lsr_item else "—",
                "Позиция проекта": c.spec_item.code if c.spec_item else "—",
                "Наименование в проекте": c.spec_item.raw_title if c.spec_item else "—",
                "Область в проекте": c.spec_item.pointer.coordinate if c.spec_item else "—",
                "Ед. изм.": (c.lsr_item.canonical_unit.value if c.lsr_item else c.spec_item.canonical_unit.value) if (c.lsr_item or c.spec_item) else "—",
                "Объем по смете": float(c.lsr_item.volume) if c.lsr_item else 0.0,
                "Объем по проекту": float(c.spec_item.volume) if c.spec_item else 0.0,
                "Расхождение (Дельта)": float(c.volume_delta),
                "Допустимая погрешность": float(c.allowed_delta),
                "Описание коллизии": c.description
            })

        df = pd.DataFrame(rows)
        output_dir = os.path.dirname(os.path.abspath(output_path))
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        df.to_excel(output_path, index=False)


def main():
    parser = argparse.ArgumentParser(description="Автоматизированный аудит смет (Excel) и спецификаций (PDF)")
    parser.add_argument("--lsr", default="data/сот.xlsx", help="Путь к локальной смете (.xlsx)")
    parser.add_argument("--spec", default="data/сот2.pdf", help="Путь к спецификации проекта (.pdf)")
    parser.add_argument("--out", default="data/audit_clean.xlsx", help="Путь для сохранения отчета (.xlsx)")
    parser.add_argument("--no-interactive", action="store_true", help="Отключить интерактивный опрос оператора (пакетный режим)")

    args = parser.parse_args()

    orchestrator = AuditOrchestrator(
        lsr_path=args.lsr, 
        spec_path=args.spec,
        interactive=not args.no_interactive
    )
    orchestrator.run_audit(output_excel=args.out)


if __name__ == "__main__":
    main()