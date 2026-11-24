"""Pydantic-схемы веб-слоя 0.1.0.

Zero Float Policy: все расчетные числа (объемы, коэффициенты, дельты,
уверенность) передаются строками. ``int`` остается только там, где это
идентификатор или счетчик. Булевы значения сохраняются булевыми.
"""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator, model_validator


def normalize_numeric_strings(value: Any) -> Any:
    """Рекурсивно переводит числа в строки, не искажая строки и bool.

    * ``bool`` проверяется первым: в Python ``bool`` — подкласс ``int``;
    * ``int``/``float``/``Decimal`` превращаются в ``str(Decimal)``, чтобы
      не потерять точность и единый формат (``842.0``, ``1E-5``);
    * строки возвращаются без изменений: ведущие нули (``"00123"``) и
      артикулы сохраняются;
    * списки и словари обходятся рекурсивно.
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, Decimal)):
        try:
            return str(Decimal(str(value)))
        except (InvalidOperation, ValueError):
            return str(value)
    if isinstance(value, dict):
        return {str(key): normalize_numeric_strings(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [normalize_numeric_strings(item) for item in value]
    return value


def normalize_page_meta(meta: Any) -> dict[str, Any] | None:
    """Нормализует словарь ``page_meta`` (размеры, поворот, масштаб).

    :return: словарь со строковыми числами либо ``None``, если вход не словарь.
    """
    if isinstance(meta, dict):
        return {
            str(key): normalize_numeric_strings(item) for key, item in meta.items()
        }
    return None


class ConsentRequest(BaseModel):
    """Запрос фиксации согласия субъекта ПДн."""

    login: str = Field(min_length=1, max_length=320)
    consent_type: str = Field(default="PDN_PROCESSING", min_length=1, max_length=64)
    doc_version: str = Field(min_length=1, max_length=64)
    accepted: bool = True
    client_meta: dict[str, Any] | None = None


class ConsentResponse(BaseModel):
    """Ответ с обезличенным профилем и идентификатором согласия."""

    user_id: str
    contact_mask: str | None = None
    consent_id: int
    accepted: bool


class SessionCreateResponse(BaseModel):
    """Итог создания сессии из загруженных файлов."""

    session_id: str
    status: str
    pending_count: int
    rd_file: str
    lsr_file: str


class CandidateSchema(BaseModel):
    """Кандидат РД для выбора сметчиком (top-3)."""

    rd_item_id: int
    score: str
    description: str
    unit: str
    quantity: str

class ReviewCardSchema(BaseModel):
    """Карточка HITL: все расчетные величины — строки."""

    match_id: int
    index: int
    total: int
    confidence: str
    conversion_ratio: str
    adjusted_ratio: str | None = None
    rd_code: str | None = None
    rd_description: str | None = None
    rd_unit: str | None = None
    rd_quantity: str | None = None
    smeta_code: str | None = None
    smeta_description: str | None = None
    smeta_unit: str | None = None
    smeta_quantity: str | None = None
    unit_price: str | None = None
    total_cost: str | None = None
    waste_coefficient: str
    delta_volume: str | None = None
    bbox: tuple[str, str, str, str] | None = None
    page_meta: dict[str, Any] | None = None
    candidates: list[CandidateSchema] = Field(default_factory=list)


DecisionAction = Literal["APPROVE", "SELECT_CANDIDATE", "REJECT", "ADJUST"]


class DecisionRequest(BaseModel):
    """Экспертное решение по карточке HITL."""

    match_id: int
    action: DecisionAction
    reviewer: str = Field(min_length=1, max_length=200)
    candidate_rd_id: int | None = None
    reason_code: str | None = Field(default=None, max_length=64)
    adjusted_ratio: str | None = None
    comment: str | None = Field(default=None, max_length=2000)

    @field_validator("adjusted_ratio")
    @classmethod
    def _validate_ratio(cls, value: str | None) -> str | None:
        """Проверяет коэффициент через ``Decimal`` и возвращает канон строки."""
        if value is None:
            return None
        try:
            ratio = Decimal(value)
        except (InvalidOperation, ValueError) as exc:
            raise ValueError("adjusted_ratio должен быть числом в строке") from exc
        if not ratio.is_finite() or ratio <= 0:
            raise ValueError("adjusted_ratio должен быть положительным числом")
        return str(ratio)

    @model_validator(mode="after")
    def _validate_action_payload(self) -> "DecisionRequest":
        """Проверяет обязательные поля выбранного действия."""
        if self.action == "ADJUST" and self.adjusted_ratio is None:
            raise ValueError("ADJUST требует adjusted_ratio")
        if self.action == "SELECT_CANDIDATE" and self.candidate_rd_id is None:
            raise ValueError("SELECT_CANDIDATE требует candidate_rd_id")
        if self.action == "REJECT" and not (self.reason_code or "").strip():
            raise ValueError("REJECT требует reason_code")
        return self


class DecisionResponse(BaseModel):
    """Результат применения экспертного решения."""

    match_id: int
    decision: str
    status: str
    reviewer: str
    adjusted_ratio: str | None = None
    delta_volume: str | None = None
    pending_count: int


class FinalizeResponse(BaseModel):
    """Итог финализации сессии."""

    session_id: str
    status: str
    pending_count: int


class ArbitrationEventSchema(BaseModel):
    """Запись append-only журнала арбитража."""

    id: int
    session_id: str
    match_record_id: int
    actor: str
    action: str
    old_payload: dict[str, Any] | None = None
    new_payload: dict[str, Any] | None = None
    comment: str | None = None
    created_at: str | None = None
