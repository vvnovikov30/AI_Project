import os
import argparse
from decimal import Decimal
from typing import List, Optional, Set
import pandas as pd

from models import NormalizedItem, CollisionRecord, CollisionSeverity
from parsers import LSRParser, PDFSpecParser
from matcher import CascadeMatcher
from evaluator import DiscrepancyDetector

class AuditOrchestrator:
    """Координатор конвейера сверки сметы и проектной спецификации."""

    FUZZY_ACCEPTANCE_THRESHOLD = 85.0
    FUZZY_SUSPECT_THRESHOLD = 70.0

    def __init__(self, lsr_path: str, spec_path: str):
        self.lsr_path = lsr_path
        self.spec_path = spec_path
        self.lsr_parser = LSRParser(lsr_path)
        self.spec_parser = PDFSpecParser(spec_path)

    def run_audit(self, output_excel: str):
        print(f"[*] Старт парсинга сметы: {self.lsr_path}")
        lsr_items = self.lsr_parser.parse()
        print(f"    Найдено позиций в смете: {len(lsr_items)}")

        print(f"[*] Старт извлечения спецификации: {self.spec_path}")
        spec_items = self.spec_parser.extract_items()
        print(f"    Найдено позиций в спецификации: {len(spec_items)}")

        collisions: List[CollisionRecord] = []
        matched_spec_ids: Set[str] = set()

        print("[*] Сопоставление номенклатуры и поиск коллизий...")
        # 1. Прямой проход: сверка сметных позиций с проектом
        for lsr in lsr_items:
            best_match: Optional[NormalizedItem] = None
            best_score = 0.0

            for spec in spec_items:
                score = CascadeMatcher.calculate_score(lsr, spec)
                if score > best_score:
                    best_score = score
                    best_match = spec

            if best_score >= self.FUZZY_ACCEPTANCE_THRESHOLD and best_match:
                matched_spec_ids.add(best_match.raw_id)
                record = DiscrepancyDetector.evaluate(lsr, best_match, best_score)
                collisions.append(record)
            elif self.FUZZY_SUSPECT_THRESHOLD <= best_score < self.FUZZY_ACCEPTANCE_THRESHOLD and best_match:
                matched_spec_ids.add(best_match.raw_id)
                collisions.append(CollisionRecord(
                    lsr_item=lsr,
                    spec_item=best_match,
                    similarity_score=best_score,
                    volume_delta=lsr.volume - best_match.volume,
                    allowed_delta=Decimal("0"),
                    severity=CollisionSeverity.SUSPECT,
                    description=f"Низкая уверенность совпадения ({best_score:.1f}%). Требуется ручная проверка"
                ))
            else:
                collisions.append(CollisionRecord(
                    lsr_item=lsr,
                    spec_item=None,
                    similarity_score=best_score,
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
                    similarity_score=0.0,
                    volume_delta=-spec.volume,
                    allowed_delta=Decimal("0"),
                    severity=CollisionSeverity.CRITICAL,
                    description="Проектная позиция не осмечена (Пропущенный объем)"
                ))

        print(f"[*] Формирование отчета: {output_excel}")
        self._export_report(collisions, output_excel)
        print("[+] Аудит успешно завершен!")

    def _export_report(self, collisions: List[CollisionRecord], output_path: str):
        rows = []
        for c in collisions:
            rows.append({
                "Статус": c.severity.value,
                "Уверенность скоринга (%)": round(c.similarity_score, 1),
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

    args = parser.parse_args()

    orchestrator = AuditOrchestrator(lsr_path=args.lsr, spec_path=args.spec)
    orchestrator.run_audit(output_excel=args.out)


if __name__ == "__main__":
    main()