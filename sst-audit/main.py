import os
import argparse
from decimal import Decimal
from typing import List, Optional, Set, Tuple
import pandas as pd

from models import NormalizedItem, CollisionRecord, CollisionSeverity
from parsers import LSRParser, PDFSpecParser
from matcher import CascadeMatcher
from evaluator import DiscrepancyDetector


class AuditOrchestrator:
    """Координатор конвейера сверки сметы и проектной спецификации с поддержкой HITL."""

    FUZZY_ACCEPTANCE_THRESHOLD = 85.0
    FUZZY_SUSPECT_THRESHOLD = 70.0

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

        collisions: List[CollisionRecord] = []
        matched_spec_ids: Set[str] = set()

        print("[*] Сопоставление номенклатуры и анализ коллизий...")

        # 1. Прямой проход: сверка сметных позиций с проектом
        for lsr in lsr_items:
            # Собираем скоринг по всем позициям проекта
            scored_candidates: List[Tuple[NormalizedItem, float]] = []
            for spec in spec_items:
                score = CascadeMatcher.calculate_score(lsr, spec)
                if score >= self.FUZZY_SUSPECT_THRESHOLD:
                    scored_candidates.append((spec, score))

            # Сортируем кандидатов по убыванию сходства
            scored_candidates.sort(key=lambda x: x[1], reverse=True)

            best_match: Optional[NormalizedItem] = scored_candidates[0][0] if scored_candidates else None
            best_score: float = scored_candidates[0][1] if scored_candidates else 0.0

            # Сценарий А: Автоматическое принятие (>= 85.0%)
            if best_score >= self.FUZZY_ACCEPTANCE_THRESHOLD and best_match:
                matched_spec_ids.add(best_match.raw_id)
                record = DiscrepancyDetector.evaluate(lsr, best_match, Decimal(str(best_score)))
                collisions.append(record)

            # Сценарий Б: Серая зона (70.0% - 84.9%) -> Human-in-the-Loop
            elif self.FUZZY_SUSPECT_THRESHOLD <= best_score < self.FUZZY_ACCEPTANCE_THRESHOLD and best_match:
                confirmed_spec: Optional[NormalizedItem] = None

                if self.interactive:
                    # Показываем оператору топ-3 лучших кандидатов
                    top_candidates = scored_candidates[:3]
                    confirmed_spec = self._prompt_operator_review(lsr, top_candidates)
                
                if confirmed_spec:
                    matched_spec_ids.add(confirmed_spec.raw_id)
                    # После подтверждения человеком отправляем на расчет расхождения
                    record = DiscrepancyDetector.evaluate(lsr, confirmed_spec, Decimal(str(best_score)))
                    record.description = f"[HITL Подтверждено] {record.description}"
                    collisions.append(record)
                else:
                    # Оператор отклонил сопоставление (или режим non-interactive)
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

            # Сценарий В: Полное несовпадение (< 70.0%)
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

        # 2. Обратный проход: поиск неосмеченных проектных спецификаций
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
    parser.add_argument("--lsr", default="data/estimate.xlsx", help="Путь к локальной смете (.xlsx)")
    parser.add_argument("--spec", default="data/specification.pdf", help="Путь к спецификации проекта (.pdf)")
    parser.add_argument("--out", default="data/audit_report.xlsx", help="Путь для сохранения отчета (.xlsx)")
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