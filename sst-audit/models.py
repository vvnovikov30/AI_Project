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
    unit_price: Optional[Decimal] = None   # Сметная цена за единицу (если извлечена)
    total_cost: Optional[Decimal] = None   # Сметная стоимость позиции (если извлечена)

    def _persistence_payload(self) -> Dict[str, Any]:
        """Аудиторский след позиции для JSON-колонки ``raw_payload``.

        Все числовые значения сериализуются строками, чтобы не терять
        точность ``Decimal`` при JSON-кодировании.
        """
        return {
            "raw_id": self.raw_id,
            "raw_title": self.raw_title,
            "clean_title": self.clean_title,
            "raw_unit": self.raw_unit,
            "canonical_unit": self.canonical_unit.value,
            "unit_multiplier": str(self.unit_multiplier),
            "volume": str(self.volume),
            "dimensions": dict(self.dimensions),
            "context_section": self.context_section,
            "is_resource": self.is_resource,
            "parent_id": self.parent_id,
            "pointer": {
                "doc_type": self.pointer.doc_type,
                "file_name": self.pointer.file_name,
                "sheet_name": self.pointer.sheet_name,
                "row_idx": self.pointer.row_idx,
                "col_letter": self.pointer.col_letter,
                "coordinate": self.pointer.coordinate,
            },
        }

    def to_rd_record(self) -> Dict[str, Any]:
        """Экспорт позиции РД в запись для ``db.save_rd_items()``.

        :return: словарь с ключами row_idx, code, description, unit,
            quantity, raw_payload. ``quantity`` остается ``Decimal`` —
            конвертация в тип SQLite выполняется на границе ``db.py``.
        """
        return {
            "row_idx": self.pointer.row_idx,
            "code": self.code,
            "description": self.raw_title,
            "unit": self.raw_unit or self.canonical_unit.value,
            "quantity": self.volume,
            "raw_payload": self._persistence_payload(),
        }

    def to_smeta_record(self) -> Dict[str, Any]:
        """Экспорт позиции сметы в запись для ``db.save_smeta_items()``.

        :return: словарь с ключами row_idx, code, description, unit,
            quantity, unit_price, total_cost, raw_payload.

        Если парсер не извлек цену, ``unit_price``/``total_cost``
        подставляются нулями (схема SSOT требует NOT NULL); при заданной
        цене и отсутствующей стоимости стоимость рассчитывается как
        ``unit_price * volume``.
        """
        unit_price = self.unit_price if self.unit_price is not None else Decimal("0")
        if self.total_cost is not None:
            total_cost = self.total_cost
        elif self.unit_price is not None:
            total_cost = unit_price * self.volume
        else:
            total_cost = Decimal("0")

        return {
            "row_idx": self.pointer.row_idx,
            "code": self.code or "",
            "description": self.raw_title,
            "unit": self.raw_unit or self.canonical_unit.value,
            "quantity": self.volume,
            "unit_price": unit_price,
            "total_cost": total_cost,
            "raw_payload": self._persistence_payload(),
        }

@dataclass
class CollisionRecord:
    lsr_item: Optional[NormalizedItem]
    spec_item: Optional[NormalizedItem]
    similarity_score: Decimal
    volume_delta: Decimal
    allowed_delta: Decimal
    severity: CollisionSeverity
    description: str
