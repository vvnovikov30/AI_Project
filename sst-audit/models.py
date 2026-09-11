from dataclasses import dataclass
from decimal import Decimal
from typing import Optional, Dict, Any
from enum import Enum

class UnitClass(Enum):
    AREA = "м2"
    VOLUME = "м3"
    MASS = "т"
    LENGTH = "м"
    COUNT = "шт"
    UNKNOWN = "unknown"

class CollisionSeverity(Enum):
    CRITICAL = "CRITICAL"          # MULTIPLIER_ERROR, расхождение объемов выше лимита
    SUSPECT = "SUSPECT"            # В диапазоне 70.0 - 84.9 (ручная проверка)
    INCOMPATIBLE = "INCOMPATIBLE"  # Конфликт размерностей
    INFO = "INFO"                  # Корректно учтенный технологический отход

@dataclass(frozen=True)
class CellPointer:
    doc_type: str        # 'LSR' | 'SPECIFICATION' | 'VOR'
    file_name: str
    sheet_name: str      # Имя листа Excel или 'Page N' для PDF
    row_idx: int
    col_letter: str
    coordinate: str      # Например, 'C24' или 'p. 12 [bbox]'

@dataclass
class NormalizedItem:
    raw_id: str
    code: Optional[str]              # Шифр расценки/ГОСТ/ФЕР/ТЕР
    raw_title: str
    clean_title: str                 # Очищенный лексический вектор
    raw_unit: str
    canonical_unit: UnitClass
    unit_multiplier: Decimal         # 1.0, 100.0, 1000.0
    volume: Decimal                  # Базовый нормализованный объем
    dimensions: Dict[str, Any]       # RegEx-метки: {'DN': 50, 'PN': 16, 'concrete': 'B25'}
    context_section: str             # Номер тома/раздел сметы
    pointer: CellPointer
    is_resource: bool = False        # Дочерний неучтенный ресурс к расценке
    parent_id: Optional[str] = None

@dataclass
class CollisionRecord:
    lsr_item: Optional[NormalizedItem]
    spec_item: Optional[NormalizedItem]
    similarity_score: Decimal
    volume_delta: Decimal
    allowed_delta: Decimal
    severity: CollisionSeverity
    description: str