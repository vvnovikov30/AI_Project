from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from typing import ClassVar, Optional, Dict, Any
from enum import Enum

class UnitClass(Enum):
    AREA = "м2"
    VOLUME = "м3"
    MASS = "т"
    LENGTH = "м"
    COUNT = "шт"
    UNKNOWN = "unknown"

class DimensionFamily(Enum):
    """Семейства физических размерностей для предиката R05_UNT (Unit Gate)."""

    LENGTH = "LENGTH"
    AREA = "AREA"
    VOLUME = "VOLUME"
    MASS = "MASS"
    COUNT = "COUNT"
    ELECTRICAL = "ELECTRICAL"   # точечные/системные единицы электромонтажа
    UNKNOWN = "UNKNOWN"

@dataclass(frozen=True)
class UnitDimension:
    """Канонизированное описание единицы измерения.

    :param family: семейство физической размерности;
    :param base_unit: базовая единица семейства (м, м2, м3, кг, шт, ...);
    :param multiplier: кратность относительно базовой единицы (Decimal,
        например 100 для "100 м" и 0.001 для "мм").
    """

    family: DimensionFamily
    base_unit: str
    multiplier: Decimal

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


# Статусы предикатов аудита (rule_evaluations). Единый перечень для слоя
# данных и расчетного ядра: PASSED / WARNING / FAILED / SKIPPED.
RULE_EVALUATION_STATUSES: frozenset[str] = frozenset(
    {"PASSED", "WARNING", "FAILED", "SKIPPED"}
)


def _coerce_rule_decimal(value: object, field_name: str) -> Decimal:
    """Приводит числовой атрибут предиката к ``Decimal``.

    :raises TypeError: если передан ``float``/``bool`` (Zero Float Policy).
    """
    if isinstance(value, bool) or isinstance(value, float):
        raise TypeError(
            f"{field_name} должен быть Decimal: float запрещен Zero Float Policy"
        )
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


@dataclass
class RuleEvaluation:
    """Результат проверки одного предиката аудита.

    :param rule_code: код правила (например, ``R01_CAT``, ``R_CMP_VOLUME_DEFICIT``);
    :param status: один из ``RULE_EVALUATION_STATUSES``;
    :param weight: вес предиката в композитной уверенности (строго ``Decimal``);
    :param score: нормированная оценка предиката ``0..1`` (строго ``Decimal``);
    :param message: человекочитаемое пояснение результата;
    :param session_id: сессия аудита (SSOT-скоуп; может быть пустой до привязки);
    :param id: идентификатор записи в ``rule_evaluations`` (``None`` до вставки);
    :param match_record_id: связь ``match_records``, если правило привязано к паре;
    :param details: JSON аудиторского следа для SSOT;
    :param score_weight: legacy-вес правила для колонки ``rule_evaluations``;
    :param created_at: отметка времени из SQLite (``CURRENT_TIMESTAMP``).
    """

    rule_code: str
    status: str
    weight: Decimal = Decimal("0")
    score: Decimal = Decimal("0")
    message: str = ""
    session_id: str = ""
    id: Optional[int] = None
    match_record_id: Optional[int] = None
    details: Optional[str] = None
    score_weight: Optional[Decimal] = None
    created_at: Optional[str] = None

    def __post_init__(self) -> None:
        if not self.rule_code or not self.rule_code.strip():
            raise ValueError("Поле 'rule_code' должно быть непустой строкой")
        if self.status not in RULE_EVALUATION_STATUSES:
            raise ValueError(
                f"Недопустимый статус правила '{self.status}'. "
                f"Разрешены: {sorted(RULE_EVALUATION_STATUSES)}"
            )
        self.weight = _coerce_rule_decimal(self.weight, "weight")
        self.score = _coerce_rule_decimal(self.score, "score")
        if self.score_weight is None:
            self.score_weight = self.weight
        else:
            self.score_weight = _coerce_rule_decimal(
                self.score_weight, "score_weight"
            )
        # Legacy-совместимость: вес из SSOT подхватывается, если не задан явно.
        if self.weight == Decimal("0") and self.score_weight != Decimal("0"):
            self.weight = self.score_weight
        if not self.message and self.details:
            self.message = str(self.details)
        elif self.details is None and self.message:
            self.details = self.message


@dataclass
class RuleBreakdown:
    """Агрегированная сводка предикатов R01-R05 по одной связи аудита.

    Веса предикатов фиксированы и в сумме дают ``1.0000``:

    * ``R01_CAT`` — совпадение категории (работа/материал) — ``0.20``;
    * ``R02_DIM`` — блокировка технических параметров — ``0.30``;
    * ``R03_STD`` — совпадение шифра/ГОСТ/ТУ — ``0.15``;
    * ``R04_MAT`` — лексическое сходство наименований — ``0.15``;
    * ``R05_UNT`` — шлюз размерности единиц измерения — ``0.20``.
    """

    rule_evaluations: list[RuleEvaluation] = field(default_factory=list)
    details_summary: str = ""

    R01_CAT: ClassVar[Decimal] = Decimal("0.20")
    R02_DIM: ClassVar[Decimal] = Decimal("0.30")
    R03_STD: ClassVar[Decimal] = Decimal("0.15")
    R04_MAT: ClassVar[Decimal] = Decimal("0.15")
    R05_UNT: ClassVar[Decimal] = Decimal("0.20")

    WEIGHTS: ClassVar[Dict[str, Decimal]] = {
        "R01_CAT": R01_CAT,
        "R02_DIM": R02_DIM,
        "R03_STD": R03_STD,
        "R04_MAT": R04_MAT,
        "R05_UNT": R05_UNT,
    }
    CONFIDENCE_QUANTUM: ClassVar[Decimal] = Decimal("0.0001")

    def add(self, evaluation: RuleEvaluation) -> "RuleBreakdown":
        """Добавляет результат предиката и возвращает ``self`` для цепочки."""
        self.rule_evaluations.append(evaluation)
        return self

    def get(self, rule_code: str) -> Optional[RuleEvaluation]:
        """Возвращает результат предиката по коду либо ``None``."""
        for evaluation in self.rule_evaluations:
            if evaluation.rule_code == rule_code:
                return evaluation
        return None

    @property
    def is_valid(self) -> bool:
        """``False``, если хотя бы один предикат завершился со статусом FAILED."""
        return all(
            evaluation.status != "FAILED"
            for evaluation in self.rule_evaluations
        )

    @property
    def is_blocked(self) -> bool:
        """Жесткий шлюз R05_UNT: несовместимые размерности блокируют автопринятие."""
        return any(
            evaluation.rule_code == "R05_UNT"
            and evaluation.status == "FAILED"
            for evaluation in self.rule_evaluations
        )

    @property
    def weight_sum(self) -> Decimal:
        """Сумма весов всех предикатов R01-R05 (контроль SSOT-инварианта)."""
        return sum(self.WEIGHTS.values(), Decimal("0"))

    @property
    def total_confidence(self) -> Decimal:
        """Суммарная уверенность (строго ``Decimal``, квант ``0.0001``)."""
        return self.calculate_total_confidence()

    def calculate_total_confidence(self) -> Decimal:
        """Считает взвешенную сумму предикатов без float-арифметики.

        Используются фиксированные веса R01-R05; предикаты вне этого набора
        (например, ``R_CMP_VOLUME_DEFICIT``) в композитную уверенность не входят.
        """
        total = Decimal("0")
        for rule_code, weight in self.WEIGHTS.items():
            evaluation = self.get(rule_code)
            if evaluation is None:
                continue
            total += weight * evaluation.score
        return total.quantize(
            self.CONFIDENCE_QUANTUM, rounding=ROUND_HALF_UP
        )


# Статусы композитной записи сопоставления: BLOCKED — жесткий шлюз R05_UNT.
MATCH_RECORD_STATUSES: frozenset[str] = frozenset(
    {"AUTO", "HITL_PENDING", "BLOCKED", "NO_MATCH"}
)


@dataclass
class MatchRecord:
    """Композитная связь «позиция РД ↔ узел сметы» с аудитом предикатов.

    :param rd_item: позиция спецификации/РД;
    :param smeta_item: сопоставленный узел сметы (работа либо дочерний ресурс);
    :param confidence: композитная уверенность ``0..1`` (строго ``Decimal``);
    :param status: один из ``MATCH_RECORD_STATUSES``;
    :param method: метод сопоставления (``COMPOSITE_CASCADE`` по умолчанию);
    :param parent_work_id: ``raw_id`` родительской расценки монтажа, если
        сопоставлен дочерний ресурс (родитель считается покрытым);
    :param evaluations: результаты всех предикатов, включая объемные проверки;
    :param breakdown: сводка R01-R05 с весами композитной уверенности.

    ``BLOCKED``/``NO_MATCH`` — статусы in-memory аудита; при записи в SSOT
    они маппятся на допустимые статусы ``match_records`` (``AUTO`` /
    ``HITL_PENDING`` / ``RESOLVED``).
    """

    rd_item: NormalizedItem
    smeta_item: NormalizedItem
    confidence: Decimal
    status: str
    method: str = "COMPOSITE_CASCADE"
    parent_work_id: Optional[str] = None
    evaluations: list[RuleEvaluation] = field(default_factory=list)
    breakdown: Optional[RuleBreakdown] = None

    def __post_init__(self) -> None:
        self.confidence = _coerce_rule_decimal(self.confidence, "confidence")
        if self.status not in MATCH_RECORD_STATUSES:
            raise ValueError(
                f"Недопустимый статус связи '{self.status}'. "
                f"Разрешены: {sorted(MATCH_RECORD_STATUSES)}"
            )
        if not self.evaluations and self.breakdown is not None:
            self.evaluations = list(self.breakdown.rule_evaluations)

    @property
    def is_auto_confirmed(self) -> bool:
        """Связь автоматически подтверждена (не требует HITL)."""
        return self.status == "AUTO"

    @property
    def is_resource_covered(self) -> bool:
        """Связь закрывает дочерний ресурс, а значит и родительскую расценку."""
        return self.parent_work_id is not None

    @property
    def is_unit_blocked(self) -> bool:
        """Связь заблокирована шлюзом размерностей R05_UNT."""
        if self.breakdown is not None and self.breakdown.is_blocked:
            return True
        return any(
            evaluation.rule_code == RuleBreakdown.R05_UNT
            and evaluation.status == "FAILED"
            for evaluation in self.evaluations
        )

    @property
    def score_100(self) -> Decimal:
        """Композитная уверенность на шкале ``0..100`` (квант ``0.0001``)."""
        return (self.confidence * Decimal("100")).quantize(
            Decimal("0.0001"), rounding=ROUND_HALF_UP
        )


@dataclass(frozen=True)
class SmetaItem:
    """Сметная строка, прочитанная из SSOT (``smeta_items``).

    :param parent_id: идентификатор родительской работы в той же сессии
        (``None`` для корневых разделов). Позволяет матчеру получать
        родительскую позицию за O(1) из ``get_smeta_items_map()``.
    """

    id: int
    session_id: str
    code: str
    description: str
    unit: str
    quantity: Decimal
    unit_price: Decimal
    total_cost: Decimal
    row_idx: Optional[int] = None
    parent_id: Optional[int] = None
    raw_payload: Optional[Dict[str, Any]] = None
