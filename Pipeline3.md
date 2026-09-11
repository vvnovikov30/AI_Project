# Архитектурная спецификация конвейера аудита ПСД (Pipeline Specification)

* **Идентификатор:** `SPEC-ARCH-001`
* **Версия документа:** `1.4.1 (Merged)`
* **Дата фиксации:** `2026-09-08`
* **Статус:** `Approved Architecture Baseline (Includes MEP/Engineering Systems)`
* **Базовый стек:** Python 3.12, Pydantic v2, openpyxl, pdfplumber, rapidfuzz, SQLite
* **Нормативная база:** 
  * Приказ Минстроя России № 421/пр (в ред. приказа № 557/пр).
  * ГОСТ Р 21.101-2020 (СПДС).
  * Сборники ГЭСНм/ФЕРм (Часть 8. Электротехнические установки; Часть 10. Оборудование связи).

---

## 1. Журнал изменений (Revision History Delta)

| Ревизия | Дата | Автор | Суть изменений и статус инженерных гипотез |
| :--- | :--- | :--- | :--- |
| **v1.2** | 2026-09-03 | Systems Architect | Базовая фиксация монолитного пайплайна сбора данных из Excel/PDF. |
| **v1.3** | 2026-09-08 | Lead Systems Architect | Утвержден 5-фазный ETL, переход на `decimal.Decimal`, внедрен каскадный шлюз размерностей (Unit Gating). |
| **v1.4** | 2026-09-08 | Lead Systems Architect | **Интеграция инженерных систем (ЭОМ, СС, АК):**<br>1. Добавлены регулярные выражения для детекции сечения кабелей.<br>2. Расширен словарь технологических глаголов из монтажных сборников.<br>3. Внедрен алгоритм выявления расхождений объемов, вызванных нормативным запасом кабеля. |

---

## 2. Назначение и границы системы

### 2.1. Цель разработки
Автоматизация выявления пространственно-количественных, номенклатурных и нормативных расхождений между проектной частью (ведомости объемов работ — ВОР, спецификации чертежей) и сметной частью (локальные сметные расчеты — ЛСР) объектов капитального строительства, с углубленным анализом внутренних инженерных систем (электроснабжение, слаботочные системы). 

### 2.2. Границы ответственности модуля
* **Входит в скоуп:**
  * Парсинг табличных форм ЛСР по 421/пр и извлечение проектных спецификаций (`.xlsx`, `.pdf`).
  * Сопоставление кабельной продукции с учетом жильности, сечения и типа изоляции.
  * Блокирующая валидация физических размерностей и расчет коэффициента подобия.
  * Разделение оборудования и материалов, заложенных в расценке (гофротруба, крепеж).
* **Не входит в скоуп:**
  * Электротехнические расчеты (токи короткого замыкания). Прямой парсинг бинарных файлов CAD/BIM (`.dwg`, `.ifc`). Пересчет индексов.

---

## 3. Архитектура 5-фазного конвейера ETL

    [Фаза 1: Ingestion & Topology Extraction]
      ├── openpyxl -> ЛСР (421/пр) и ВОР (Excel)
      └── pdfplumber -> Спецификации оборудования и кабельные журналы (PDF)
                            │
                            ▼
    [Фаза 2: Normalization & Canonicalization]
      ├── Топологическая привязка (CellPointer)
      ├── Лексическая чистка (включая MEP стоп-слова)
      └── Канонизация размерностей в Decimal
                            │
                            ▼
    [Фаза 3: Intermediate Checkpoint (SQLite)]
      └── Таблица audit_staging (Human-in-the-Loop аудит промежуточных распарсенных строк)
                            │
                            ▼
    [Фаза 4: Gated Hybrid Matching Engine]
      ├── Unit Gating (I_unit) -> запрет сравнения несовместимых размерностей
      ├── Dimensional Lock (I_dim) -> проверка сечений и габаритов
      └── Вычисление скоринга: w1*Code + w2*Fuzzy + w3*Context
                            │
                            ▼
    [Фаза 5: Collision Analysis & Reporting]
      ├── Детекция скрытых сдвигов кратности (Multiplier Shift 10x, 100x, 1000x)
      ├── Анализ запаса кабеля (Cable Slack Tolerance)
      └── Экспорт DiscrepancyReportItem в JSON / Excel

---

## 4. Схемы данных (Pydantic v2 Models)

    from decimal import Decimal
    from typing import Literal, Optional
    from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator
    
    class CellPointer(BaseModel):
        """Топологический указатель источника данных в первичном документе (Audit Trail)."""
        model_config = ConfigDict(frozen=True)
        doc_type: Literal["SPECIFICATION", "VOR", "LSR", "SSR"]
        file_name: str
        sheet_name: str
        row_idx: int = Field(..., ge=1)
        coordinate: str
    
    class AuditLineItem(BaseModel):
        """Канонизированная строка строительной ведомости/сметы."""
        model_config = ConfigDict(arbitrary_types_allowed=True)
        pointer: CellPointer
        item_code: Optional[str] = Field(None, description="Шифр расценки или артикул")
        raw_title: str
        clean_title: str
        
        quantity: Decimal
        raw_unit: str
        canonical_unit: str
        unit_multiplier: Decimal = Field(default=Decimal("1.0"))
        base_price: Optional[Decimal] = None
    
        @field_validator("quantity", "unit_multiplier", "base_price", mode="before")
        @classmethod
        def parse_strict_decimal(cls, value: object) -> Optional[Decimal]:
            """Защита от неявного приведения через float. Принимает int, str или Decimal."""
            if value is None:
                return None
            if isinstance(value, float):
                raise TypeError("Недопустимо использовать тип float для объемов/цен сметы.")
            if isinstance(value, (int, str)):
                sanitized = str(value).replace(" ", "").replace(",", ".").strip()
                return Decimal(sanitized)
            if isinstance(value, Decimal):
                return value
            raise ValueError(f"Невозможно привести {value} к Decimal.")
    
    class DiscrepancyReportItem(BaseModel):
        """Модель зафиксированной стоимостной или объемной коллизии."""
        model_config = ConfigDict(arbitrary_types_allowed=True)
    
        collision_type: Literal[
            "VOLUME_OVERRUN",      # Превышение объема сметы
            "MULTIPLIER_ERROR",    # Ошибка кратности
            "INCOMPATIBLE_UNITS",  # Конфликт размерностей
            "DUPLICATE_RESOURCE",  # Задвоение ресурса
            "CABLE_SLACK_ANOMALY"  # Аномальный запас кабеля (>8%)
        ]
        severity: Literal["CRITICAL", "WARNING", "INFO"]
        spec_pointer: CellPointer
        estimate_pointer: CellPointer
        
        spec_qty: Decimal
        normalized_est_qty: Decimal
        delta: Decimal
        confidence_score: float = Field(default=100.0, ge=0.0, le=100.0)
        message: str
    
        @model_validator(mode="after")
        def compute_severity(self) -> "DiscrepancyReportItem":
            if self.collision_type in ("MULTIPLIER_ERROR", "INCOMPATIBLE_UNITS"):
                self.severity = "CRITICAL"
            return self

---

## 5. Алгоритмический конвейер сопоставления (Matcher Tuning)

### 5.1. Весовая формула подобия позиций
Итоговый показатель сходства строк $S_{\text{total}} \in [0.0; 100.0]$ вычисляется по каскадной формуле:

$$S_{\text{total}} = \mathbb{I}_{\text{unit}} \cdot \mathbb{I}_{\text{dim}} \cdot \left( w_1 \cdot S_{\text{code}} + w_2 \cdot S_{\text{fuzzy}} + w_3 \cdot S_{\text{context}} \right)$$

**Весовые коэффициенты и пороги:**
* $w_1$ (Шифр/Код) = **0.45**
* $w_2$ (Fuzzy Title) = **0.40**
* $w_3$ (Контекст раздела) = **0.15**
* `FUZZY_ACCEPTANCE_THRESHOLD = 85.0`
* `FUZZY_SUSPECT_THRESHOLD = 70.0`

### 5.2. Инженерный фильтр габаритов (Dimensional Lock)
* **Паттерн геометрический:** `(?:\b(?:d|ду|диам|ø)\s*[:=]?\s*(\d+)\b)`, `\b(?:толщ?|s)\s*[:=]?\s*(\d+)\b)`
* **Паттерн сечения кабеля:** `(?:\d+\s*[хxX*]\s*\d+(?:[.,]\d+)?)` (Извлекает `3х1.5`, `5x6`).
* **Паттерн автоматов:** `(?:\b\d+\s*(?:А|A)\b)` (например, `16А`).
Если в строке спецификации выделен типоразмер, а в смете присутствует иной числовой параметр, $\mathbb{I}_{\text{dim}}$ принудительно устанавливается в $0$.

### 5.3. Стоп-слова и аббревиатуры

    ABBREVIATIONS_MAP = {
        r"\bж/?б\b": "железобетонный",
        r"\bду\b": "диаметр условный",
        r"\bухл\d*\b": "",               # MEP: Удаляем климатическое исполнение
        r"\bнг(?:\(a\))?-ls\b": "нг-ls", # MEP: Нормализация индексов
        r"\bв/в\b": "высоковольтный",
        r"\bщр\b": "щит распределительный",
    }
    
    CONSTRUCTION_STOP_WORDS = {
        "устройство", "монтаж", "прокладка", "разборка", "вручную",
        "затягивание", "оконцевание", "присоединение", "прозвонка", # Глаголы ФЕРм
        "в лотках", "в гофрированной трубе", "без учета"
    }

---

## 6. Выявленные краевые случаи (Edge Cases)

| Краевой случай / Аномалия | В чем выражается физически | Способ программной обработки |
| :--- | :--- | :--- |
| **Топология `merged_cells`** | Шапки разделов объединены в Excel. | Топологический генератор транслирует левую верхнюю ячейку на весь прямоугольник. |
| **Сдвиг кратности (Multiplier)** | Объем из чертежа внесен без деления на 100. | Детектор отношения в коридорах 10x, 100x. Статус `MULTIPLIER_ERROR`. |
| **Разрыв «ВОР vs ЛСР»** | Смета урезана после экспертизы, а ВОР старый. | Внутренний кросс-чек книги Excel до матчинга со спецификацией PDF. |
| **Нормативный запас кабеля** | В смете заложен запас кабеля на изгибы (102–106 м на 100 м проекта). | Вычисляется отношение. Если $R \in [1.02; 1.06]$ — `INFO`. Если $R > 1.06$ — `CABLE_SLACK_ANOMALY`. |
| **Комплектация vs Розница** | Проект: «Шкаф (компл)». Смета: детали (стойка, ИБП). | Запуск *Grouped Matching*: суммарная стоимость деталей сопоставляется с лимитом на комплект. |
| **Задвоение (Скрытая гофра)** | Материал сидит внутри комплексной расценки ФЕР, но вынесен отдельно. | Векторная проверка ресурсной части. Отдельная строка маркируется `DUPLICATE_RESOURCE`. |

---

## 7. Метрики эффективности конвейера (для ВКР и публикаций ВАК)

Для валидации эффективности на тестовой выборке реальных строительных объектов применяются следующие метрики:

1. **Метрики качества сопоставления (Matching Performance):**
   * $\text{Precision} = \frac{TP}{TP + FP}$ (доля корректно связанных пар «чертеж — смета»).
   * $\text{Recall} = \frac{TP}{TP + FN}$ (доля найденных связей от общего числа реальных пересечений).
   * $\text{F1-Score} = 2 \cdot \frac{\text{Precision} \cdot \text{Recall}}{\text{Precision} + \text{Recall}}$.
2. **Метрика выявления коллизий:**
   * **Collision Recall Rate (CRR):** Процент выявленных реальных сметных ошибок, подтвержденных независимой строительно-технической экспертизой.
3. **Экономический показатель (Budget Risk Value):**
   * Суммарный предотвращенный перерасход:
     $$\Delta B = \sum_{i \in \text{OVERRUN}} (V_{\text{est}, i} - V_{\text{spec}, i}) \cdot C_{\text{unit}, i} \quad [\text{вычисление строго в Decimal}]$$