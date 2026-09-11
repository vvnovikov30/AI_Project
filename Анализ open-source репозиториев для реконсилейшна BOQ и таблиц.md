# **Исследование открытых ConTech и Data Reconciliation библиотек для кросс-аудита строительных спецификаций и смет**

## **Блок 1: Поисковые запросы (GitHub Search Queries)**

Разработка компактного и надежного сервиса для стоимостного аудита проектно-сметной документации требует интеграции практик из двух смежных технологических областей: прикладного строительного инжиниринга (ConTech / BOQ / Quantity Takeoff) и прецизионного анализа полуструктурированных табличных массивов (Data Reconciliation)1. В отличие от учетных монолитов и закрытых сметных калькуляторов, аналитический конвейер независимого аудитора должен обрабатывать неструктурированные таблицы коммерческих расчетов, спецификации в рабочих чертежах и ведомости объемов работ (ВОР), сохраняя при этом детерминированность вычислений и прозрачный математический аудит2.  
Для формирования пула исходных открытых компонентов был выполнен таргетированный поиск по двум специализированным направлениям запросов.

### **Направление A: Международный ConTech (BOQ & Quantity Takeoff)**

Целевой задачей поиска являлось выявление открытых утилит и предметных модулей, решающих задачи сопоставления сметных ведомостей и спецификаций, анализа расхождений объемов и парсинга строительных классификаторов без привязки к проприетарным облачным платформам:

* "bill of quantities" AND (reconciliation OR comparison OR diff) language:python — поиск специализированных скриптов сопоставления и выявления дельт между плановыми и фактическими ведомостями объемов работ5;  
* ("quantity takeoff" OR QTO) AND (matcher OR extract) language:python — идентификация модулей извлечения и связывания физических объемов из проектных спецификаций и моделей7;  
* "BOQ" AND ("specification" OR "drawing") AND (parser OR validator) language:python — отбор синтаксических анализаторов проектных спецификаций и валидаторов соответствия документации рабочей стадии6;  
* construction "cost estimating" (diff OR variance OR audit) language:python — выявление расчетных алгоритмов стоимостного контроля, дисперсионного анализа и финансового аудита сметных расценок6.

### **Направление B: Табличный реконсилейшн и нечеткий мэтчинг (Data Reconciliation)**

Фокус поиска был сосредоточен на библиотеках алгоритмической сверки таблиц произвольной структуры, способных сопоставлять разнородные схемы данных, компенсировать опечатки в наименованиях ресурсов и оперировать порогами погрешностей:

* ("table diff" OR "data reconciliation") AND (tolerance OR threshold) language:python — нахождение движков сверки табличных данных с математическим аппаратом допустимых абсолютных и относительных отклонений2;  
* fuzzy (entity OR record) matching excel (openpyxl OR pandas) language:python — выявление легковесных пайплайнов нечеткого сопоставления сущностей в электронных таблицах Excel14;  
* "fuzzy matching" AND ("unit conversion" OR "exact match") language:python — поиск комбинированных стратегий сопоставления, объединяющих строгую нормализацию физических единиц с нечетким лексическим скорингом2;  
* reconciliation engine (discrepancy OR variance OR delta) language:python — изучение архитектур движков выявления расхождений и генерации дифференциальных отчетов2.

## **Блок 2: Критерии фильтрации и отбора целевых репозиториев**

Первичный массив поисковой выдачи содержал сотни репозиториев различной степени проработки, однако подавляющее большинство не удовлетворяло жестким требованиям коммерческой аудиторской практики. Из рассмотрения были исключены крупные комплексные платформы управления строительством (например, OpenConstructionERP)11, поскольку их монолитная архитектура, включающая сотни взаимосвязанных модулей, контейнеры баз данных и веб\-интерфейсы, делает невозможным извлечение обособленной алгоритмической логики.  
Аналогичным образом были отклонены решения на базе Apache Spark или облачных хранилищ данных (Snowflake, Databricks)20, а также проекты, опирающиеся на внешние облачные API машинного зрения и распознавания образов21. Подобные зависимости нарушают базовое требование автономности (Air-Gapped On-Premise) и несут прямые риски компрометации коммерческой тайны строительных корпораций2.  
Отбор 5 наиболее зрелых репозиториев производился на основе трех строгих инженерных критериев:

* **Прагматичный и минималистичный стек:** библиотека должна функционировать локально на чистом Python с опорой на стандартные структуры данных (dataclasses), легковесные табличные интерфейсы (openpyxl, pandas), векторные строковые компараторы (rapidfuzz) и библиотеки точной арифметики (decimal)2.  
* **Математический аппарат допусков и конвертации величин:** архитектура решения должна содержать встроенные механизмы абсолютных и относительных погрешностей (abs\_tol, rel\_tol), а также позволять внедрять правила нормализации единиц измерения строительных классификаторов2.  
* **Сквозная трассируемость (Audit Trail):** ядро алгоритма обязано фиксировать и передавать на уровень отчета метаданные об источнике каждой коллизии — файл, лист, строку и буквенно-цифровой индекс исходной ячейки Excel2.

| Репозиторий | Основной технологический стек | Назначение и функциональный фокус | Ключевая ценность для адаптации в сметный MVP |
| :---- | :---- | :---- | :---- |
| **testuteab/reconlify-cli** \[cite: 2\] | Python 3.11+, PyYAML, Dataclasses | Локальный декларативный движок сверки таблиц по правилам | Декларативная YAML-конфигурация колонок, встроенный чистый Decimal, поколоночные допуски2 |
| **capitalone/datacompy** \[cite: 20, 23\] | Python, Pandas, Polars | Промышленное поколоночное и построчное сравнение датафреймов | Архитектурный паттерн Comparator Strategy, расчет относительных и абсолютных дельт20 |
| **paulfitz/daff** \[cite: 26\] | Python, C++ core / CLI | Топологический дифференциальный анализ табличных данных | Выравнивание неупорядоченных строк без ключей, наглядная нотация изменений ячеек (a \-\> b)24 |
| **kondakrindirahul/BOQ-Matching** \[cite: 5\] | Python, Pandas, Scikit-learn, NLTK | Семантическое сопоставление позиций строительных ведомостей | Предметная фильтрация несовместимых физических размерностей (Unit Gating) перед fuzzy-поиском5 |
| **nerd-oreo/bom-comparison-v3** \[cite: 15\] | Python, Openpyxl, RapidFuzz / FuzzyWuzzy | Сверка конструкторских и сметных спецификаций материалов | Развертывание объединенных диапазонов Excel, сохранение буквенно-цифровых координат ячеек15 |

## **Блок 3: Детальный архитектурный разбор отобранных репозиториев**

### **testuteab/reconlify-cli**

Репозиторий представляет собой компактную консольную утилиту на Python 3.11+, созданную для детерминированного семантического сопоставления экспортированных табличных наборов данных без передачи информации во внешнюю сеть2. Архитектура ориентирована на строго локальное исполнение и формирование детальных отчетов об отклонениях2.

#### **Архитектурная модель элемента ведомости**

Модель конфигурации и результатов сопоставления опирается на строгую типизацию стандартными средствами языка, изолируя правила преобразования данных от исполнительного механизма2:

Python  
from dataclasses import dataclass, field  
from decimal import Decimal  
from typing import Any, Dict, List, Optional

@dataclass(frozen=True)  
class ColumnConfig:  
    name: str  
    target\_name: Optional\[str\] \= None  
    tolerance: Optional\[Decimal\] \= None  
    trim\_whitespace: bool \= True  
    case\_sensitive: bool \= False  
    regex\_extract: Optional\[str\] \= None  
    null\_equivalents: List\[str\] \= field(  
        default\_factory=lambda: \["", "NULL", "None", "N/A", "-", "0.00"\]  
    )

@dataclass  
class DiscrepancyRecord:  
    key\_value: str  
    column\_name: str  
    source\_value: Any  
    target\_value: Any  
    delta: Optional\[Decimal\] \= None  
    source\_row\_index: int \= 0  
    target\_row\_index: int \= 0  
    mismatch\_type: str \= "VALUE\_DIVERGENCE"

#### **Алгоритм сопоставления (Matching Strategy)**

Стратегия сопоставления организована в виде линейного детерминированного конвейера из четырех этапов2:

> 1. **Построение индексов по первичному ключу:** строки обеих таблиц считываются в хеш-таблицы, где ключом выступает нормализованный составной или одиночный идентификатор строки; элементы, присутствующие только в одной таблице, сразу отсекаются в отчет об односторонних коллизиях2.  
> 2. **Семантическая очистка строк:** текстовые поля очищаются от пробельных символов и невидимых управляющих кодов, переводятся в унифицированный регистр и пропускаются через регулярные выражения для извлечения значимой части2.  
> 3. **Числовое приведение через Decimal:** строковые представления объемов и сумм валидируются и конвертируются в экземпляры класса Decimal, предотвращая накопление погрешностей вычислений с плавающей запятой2.  
> 4. **Поколоночная оценка толерантности:** для числовых величин вычисляется модуль разности ![][image1], который сопоставляется с индивидуально настроенным порогом допуска (tolerance); при превышении порога инцидент регистрируется с указанием точной дельты2.

#### **Практическая ценность для сметного MVP**

Из проекта следует перенять компонент декларативного описания полей через конфигурационные классы и принцип тотального отказа от стандартного типа float в пользу Decimal при расчете любых физических объемов и стоимостей2. Это защищает аудитора от математических артефактов при проверке смет с тысячами позиций17.

### **capitalone/datacompy**

Промышленная библиотека с открытым исходным кодом, созданная для всесторонней сверки табличных данных с акцентом на выявление отклонений в схемах, строках и значениях ячеек13.

#### **Архитектурная модель элемента ведомости**

Внутренняя модель библиотеки структурирована вокруг паттерна «Стратегия» (Strategy Pattern), где логика сопоставления конкретной колонки инкапсулирована в обособленные компараторы20:

Python  
from abc import ABC, abstractmethod  
from dataclasses import dataclass  
from decimal import Decimal  
import pandas as pd

@dataclass  
class ColumnComparisonSummary:  
    column\_name: str  
    matched\_records: int  
    divergent\_records: int  
    max\_absolute\_delta: Decimal  
    max\_relative\_delta: Decimal  
    divergent\_samples: pd.DataFrame

class BaseColumnComparator(ABC):  
    @abstractmethod  
    def compare(self, series\_a: pd.Series, series\_b: pd.Series) \-\> ColumnComparisonSummary:  
        pass

class NumericToleranceComparator(BaseColumnComparator):  
    def \_\_init\_\_(self, abs\_tol: Decimal \= Decimal("0"), rel\_tol: Decimal \= Decimal("0")):  
        self.abs\_tol \= abs\_tol  
        self.rel\_tol \= rel\_tol

    def compare(self, series\_a: pd.Series, series\_b: pd.Series) \-\> ColumnComparisonSummary:  
        \# Векторизованная проверка коридора допустимых погрешностей  
        pass

#### **Алгоритм сопоставления (Matching Strategy)**

Процесс сопоставления табличных наборов данных структурирован по пятифазной схеме20:

> 1. **Инспекция пересечения схем:** выявление идентичных названий колонок, а также столбцов, существующих исключительно в исходном или целевом файле25.  
> 2. **Внешнее объединение (Outer Merge) по набору ключей:** объединение датафреймов с одновременной сегментацией на общие строки и уникальные записи25.  
> 3. **Диспетчеризация компараторов:** сопоставление каждой пары колонок соответствующим компаратором в зависимости от физического смысла данных20.  
> 4. **Комбинированная оценка погрешности:** числовые данные признаются совпадающими, только если модуль разницы значений не превышает сумму абсолютного допуска и взвешенного относительного допуска:

![][image2]

> 5. **Изоляция расходящегося подмножества (Diverging Subset):** строки, содержащие хотя бы одну несовпадающую ячейку за пределами коридора допуска, извлекаются в единый датасет с фиксацией исходных значений25.

#### **Практическая ценность для сметного MVP**

Паттерн *Comparator Strategy* представляет максимальную ценность для аудитора: он позволяет подключать к шифрам расценок точный текстовый компаратор, к наименованиям работ — нечеткий семантический компаратор, а к объемам — компаратор с допуском на строительные нормы погрешностей (![][image3])20.

### **paulfitz/daff**

Библиотека daff реализует дифференциальный анализ табличных структур с упором на отслеживание перемещений строк, сдвигов колонок и модификаций на уровне отдельных ячеек в условиях отсутствия строгих первичных ключей26.

#### **Архитектурная модель элемента ведомости**

Модель данных опирается на табличную сетку TableView и реестр сопоставления индексов с типизацией операций26:

Python  
from dataclasses import dataclass  
from typing import List, Optional, Tuple

@dataclass  
class CellMutation:  
    row\_source: Optional\[int\]  
    row\_target: Optional\[int\]  
    column\_identifier: str  
    value\_before: str  
    value\_after: str  
    action\_type: str  \# '+++' (вставка), '---' (удаление), '-\>' (изменение), ' ' (без изменений)

@dataclass  
class MatrixAlignment:  
    row\_pairs: List\[Tuple\[int, int\]\]  
    orphan\_source\_rows: List\[int\]  
    orphan\_target\_rows: List\[int\]  
    aligned\_headers: List\[Tuple\[str, str\]\]

#### **Алгоритм сопоставления (Matching Strategy)**

Алгоритм ориентирован на сопоставление таблиц со смещенным порядком следования строк и граф26:

> 1. **Эвристическое совмещение колонок:** сравнение заголовков и плотности типов данных для определения эквивалентных столбцов26.  
> 2. **Топологическое выравнивание строк:** применение алгоритма наибольшей общей подпоследовательности (LCS) для нахождения оптимальной траектории связывания строк по совокупной схожести их содержимого26.  
> 3. **Фильтрация числового микро-дрейфа:** использование эвристики \--ignore-epsilon для подавления ложных срабатываний, вызванных аппаратным округлением дробных чисел24.  
> 4. **Сборка дифференциальной матрицы:** генерация финальной таблицы, где в измененных ячейках наглядно фиксируется вектор модификации через строковый литерал исходное \-\> новое25.

#### **Практическая ценность для сметного MVP**

Синтаксис отображения изменений old\_val \-\> new\_val является отраслевым эталоном для формирования понятных дефектных ведомостей25. Кроме того, механизм топологического выравнивания незаменим, когда подрядчик перегруппировал разделы сметы и изменил исходный порядок позиций ведомости объемов работ26.

### **kondakrindirahul/BOQ-Matching**

Проект решает прикладную задачу автоматизированного связывания позиций проектных ведомостей объемов работ (Project BOQ) с корпоративным справочником единичных расценок (Master BOQ) с помощью статистического анализа текста5.

#### **Архитектурная модель элемента ведомости**

Модель обособляет исходное технологическое описание строительного процесса от его нормализованного инженерного профиля и физических параметров5:

Python  
from dataclasses import dataclass  
from decimal import Decimal  
from typing import List, Optional

@dataclass  
class ConstructionBOQItem:  
    source\_line: int  
    raw\_work\_description: str  
    lemmatized\_tokens: List\[str\]  
    physical\_quantity: Decimal  
    raw\_unit\_of\_measure: str  
    dimension\_class: str  \# 'AREA', 'VOLUME', 'MASS', 'LENGTH', 'PIECE'  
    item\_code: Optional\[str\] \= None

#### **Алгоритм сопоставления (Matching Strategy)**

Последовательность проверки ориентирована на минимизацию дорогостоящих семантических вычислений за счет предварительной жесткой фильтрации5:

> 1. **Лингвистическая предобработка строительных формулировок:** очистка от шума, раскрытие профессиональных сокращений («ж/б» ![][image4] «железобетонный», «тр-р» ![][image4] «трансформатор», «диам.» ![][image4] «диаметр») и удаление второстепенных предлогов5.  
> 2. **Шлюз совместимости физических размерностей (Unit Gating):** проверка принадлежности единиц измерения сравниваемых позиций к единому классу размерности; если одна строка описывает квадратные метры, а вторая — кубические метры или штуки, процедура сравнения немедленно прерывается с нулевым коэффициентом соответствия2.  
> 3. **Многомерный текстовый скоринг:** вычисление косинусного сходства на матрицах TF-IDF в комбинации с посимвольным расстоянием модифицированного алгоритма Левенштейна5.  
> 4. **Ранжирование и выбор соответствия:** позиция связывается с наилучшим кандидатом из справочника при условии превышения минимального порога доверия (Confidence Score ![][image5])5.

#### **Практическая ценность для сметного MVP**

Идея предиктивного *шлюза размерностей (Unit Gating)* критически важна: она исключает грубые семантические галлюцинации алгоритма, когда похожие по описанию материалы ошибочно мэтчатся при принципиально разных физических единицах (например, погонаж труб и штучные задвижки)2.

### **nerd-oreo/bom-comparison-v3 (в связке с bintang)**

Связка подходов из утилиты сверки конструкторских спецификаций материалов bom-comparison-v315 и движка in-memory обработки таблиц bintang16 демонстрирует решение проблемы парсинга реальных инженерных таблиц Excel, содержащих объединения ячеек, с последующим быстрым сопоставлением через C-модуль rapidfuzz14.

#### **Архитектурная модель элемента ведомости**

Модель ориентирована на обеспечение полной трассируемости (Audit Trail), сохраняя адресные координаты ячейки вплоть до имени вкладки и буквенного индекса столбца книги Excel6:

Python  
from dataclasses import dataclass  
from decimal import Decimal  
from typing import Optional

@dataclass(frozen=True)  
class CellOrigin:  
    workbook\_name: str  
    sheet\_name: str  
    row\_number: int  
    column\_letter: str  
    cell\_address: str  \# Например, 'C27'

@dataclass  
class SheetTakeoffRow:  
    origin: CellOrigin  
    raw\_title: str  
    normalized\_title: str  
    unit: str  
    volume: Decimal  
    is\_span\_root: bool \= False  
    parent\_section: Optional\[str\] \= None

#### **Алгоритм сопоставления (Matching Strategy)**

Архитектурный конвейер объединяет низкоуровневую работу с координатной сеткой и скоростное нечеткое сопоставление15:

> 1. **Топологическое развертывание объединенных диапазонов (Unmerge & Propagate):** итерация по коллекции sheet.merged\_cells.ranges библиотеки openpyxl; скалярное значение из крайней левой верхней ячейки тиражируется на всю виртуальную сетку объединения, предотвращая генерацию пустых значений None в последующих строках15.  
> 2. **Токенизация и лексическая нормализация:** удаление знаков препинания и нормализация порядка следования слов в наименовании ресурса.  
> 3. **Двухвекторный нечеткий скоринг через RapidFuzz:**  
   * Проверка устойчивости к перестановке слов через rapidfuzz.fuzz.token\_sort\_ratio16;  
   * Проверка вхождения укороченных наименований через rapidfuzz.fuzz.partial\_ratio16.  
> 4. **Сквозная фиксация аудиторского следа:** при подтверждении соответствия позиций генерируется отчетная запись, содержащая прямые указатели на ячейки сравниваемых файлов (Спецификация\!C15 \<-\> Смета\!E48)6.

#### **Практическая ценность для сметного MVP**

Извлечение значений из объединенных ячеек с сохранением их физических адресов решает ключевую техническую проблему парсинга смет: российские сметчики повсеместно объединяют ячейки подразделов, заголовков и итогов, что ломает наивные табличные парсеры3. Применение rapidfuzz обеспечивает 20–30-кратное ускорение по сравнению с устаревшими аналогами вроде fuzzywuzzy16.

## **Блок 4: Синтез архитектуры и алгоритмических компонентов MVP сметного аудитора**

Обобщение исследованных решений позволяет спроектировать легковесный, локально развертываемый программный комплекс для аудита проектно-сметной документации. Архитектурный конвейер системы строится на трех последовательных технологических слоях:  
Первый слой отвечает за топологический парсинг файлов Excel и таблиц PDF. На этом этапе выполняется развертывание объединенных диапазонов ячеек с трансляцией контекста разделов и сквозной фиксацией физических координат каждого элемента документа.  
Второй слой производит нормализацию и подготовку данных: строковые представления чисел переводятся в точный тип Decimal, строительные формулировки очищаются от типовых аббревиатур через лексические словари, а единицы измерения приводятся к каноническим размерностям с расчетом скрытых коэффициентов кратности (100 м², 1 000 шт., тонны)17.  
Третий слой реализует каскадный алгоритм реконсилейшна: последовательно отрабатывают точный фильтр по артикулам и кодам ГОСТ, шлюз проверки совместимости физических размерностей, нечеткое семантическое сопоставление наименований работ через библиотеку rapidfuzz и вычисление дельт объемов с учетом настраиваемых допусков2. Финальные отклонения транслируются в матрицу коллизий с прямыми ссылками на координаты исходных ячеек2.  
Взаимосвязь между этапами обработки, входными данными и получаемыми результатами формализована в структуре конвейера аудита.

| Этап конвейера | Входные данные | Применяемые библиотеки | Выходной артефакт этапа |
| :---- | :---- | :---- | :---- |
| **1\. Извлечение сетки** | Файлы РД, ВОР и смет (XLSX, PDF) | openpyxl, специализированные PDF-парсеры | Плоская матрица данных с адресами ячеек (CellPointer)15 |
| **2\. Нормализация** | Сырые строки и числовые ячейки | Встроенные модули decimal, re | Экземпляры AuditLineItem с нормализованными единицами2 |
| **3\. Шлюз размерностей** | Пары единиц измерения (РД vs Смета) | Реестр переводных коэффициентов | Допуск пары к расчету либо отсечение по несовместимости2 |
| **4\. Семантический мэтчинг** | Наименования работ и материалов | rapidfuzz.fuzz (token\_sort\_ratio) | Связанные пары записей с оценкой схожести ![][image6] \[cite: 16\] |
| **5\. Расчет дельт и допусков** | Числовые объемы связанных пар | Арифметика Decimal, параметры abs\_tol / rel\_tol | Зафиксированные превышения объемов и ошибки кратности2 |
| **6\. Генерация отчета** | Реестр подтвержденных коллизий | openpyxl.styles, форматирование отчетов | Итоговый аудит-отчет с трассировкой до ячеек6 |

### **Модуль нормализации и топологического парсинга Excel**

Ниже представлен модуль развертывания объединенных диапазонов ячеек и безопасного детерминированного приведения чисел к классу Decimal:

Python  
from decimal import Decimal, InvalidOperation  
from typing import Any, Dict, Generator, Optional, Tuple  
import openpyxl  
from openpyxl.worksheet.worksheet import Worksheet

def parse\_sheet\_with\_coordinates(sheet: Worksheet) \-\> Generator\[Tuple\[int, Dict\[str, Any\]\], None, None\]:  
    """Разворачивает объединенные диапазоны ячеек Excel и транслирует значение

    верхней левой ячейки на весь диапазон, сохраняя точные координаты.  
    """  
    merged\_spans: Dict\[Tuple\[int, int\], Any\] \= {}  
      
    for span in sheet.merged\_cells.ranges:  
        root\_value \= sheet.cell(row=span.min\_row, column=span.min\_col).value  
        for r in range(span.min\_row, span.max\_row \+ 1):  
            for c in range(span.min\_col, span.max\_col \+ 1):  
                merged\_spans\[(r, c)\] \= root\_value

    for row\_idx in range(1, sheet.max\_row \+ 1):  
        row\_payload: Dict\[str, Any\] \= {}  
        for col\_idx in range(1, sheet.max\_column \+ 1):  
            cell \= sheet.cell(row=row\_idx, column=col\_idx)  
            resolved\_value \= merged\_spans.get((row\_idx, col\_idx), cell.value)  
            col\_letter \= openpyxl.utils.get\_column\_letter(col\_idx)  
            row\_payload\[col\_letter\] \= {  
                "val": resolved\_value,  
                "coordinate": f"{col\_letter}{row\_idx}",  
                "data\_type": cell.data\_type  
            }  
        yield row\_idx, row\_payload

def parse\_decimal\_strict(raw\_val: Any, fallback: Decimal \= Decimal("0.00")) \-\> Decimal:  
    """Обеспечивает строгое преобразование произвольных табличных значений

    в Decimal с очисткой от неразрывных пробелов и артефактов форматирования.  
    """  
    if raw\_val is None:  
        return fallback  
    if isinstance(raw\_val, (int, float)):  
        return Decimal(str(raw\_val))  
    if isinstance(raw\_val, Decimal):  
        return raw\_val  
      
    clean\_str \= (  
        str(raw\_val)  
        .strip()  
        .replace("\\u00a0", "")  
        .replace(" ", "")  
        .replace(",", ".")  
    )  
    try:  
        return Decimal(clean\_str)  
    except (InvalidOperation, ValueError):  
        return fallback

### **Унифицированные структуры данных аудитора**

Структуры данных спроектированы на основе dataclasses для обеспечения прозрачного аудиторского следа:

Python  
from dataclasses import dataclass  
from decimal import Decimal  
from typing import Optional

@dataclass(frozen=True)  
class CellPointer:  
    doc\_type: str  \# 'SPECIFICATION', 'BOQ', 'ESTIMATE'  
    file\_name: str  
    sheet\_name: str  
    cell\_coordinate: str  
    row\_number: int

@dataclass  
class AuditLineItem:  
    pointer: CellPointer  
    catalog\_code: Optional\[str\]  
    raw\_description: str  
    clean\_description: str  
    quantity: Decimal  
    raw\_unit: str  
    canonical\_unit: str  
    multiplier\_to\_base: Decimal \= Decimal("1.0")

@dataclass  
class AuditCollision:  
    collision\_type: str  \# 'VOLUME\_OVERRUN', 'MULTIPLIER\_ERROR', 'UNMATCHED\_SPEC'  
    severity: str        \# 'CRITICAL', 'WARNING', 'INFO'  
    spec\_ref: CellPointer  
    estimate\_ref: CellPointer  
    spec\_volume: Decimal  
    estimate\_volume\_normalized: Decimal  
    volume\_difference: Decimal  
    confidence\_score: float  
    description: str

### **Алгоритмический компаратор с допуском и контролем кратностей**

Функция сопоставления оценивает физические объемы, проверяет допустимые погрешности и диагностирует системные ошибки кратности строительных расценок:

Python  
from decimal import Decimal  
from typing import Optional  
from rapidfuzz import fuzz

def evaluate\_line\_items(  
    spec\_item: AuditLineItem,  
    est\_item: AuditLineItem,  
    abs\_tolerance: Decimal \= Decimal("0.001"),  
    rel\_tolerance: Decimal \= Decimal("0.01"),  
    fuzzy\_threshold: float \= 85.0  
) \-\> Optional\[AuditCollision\]:  
    """Выполняет каскадную проверку позиций: точный шифр \-\> fuzzy-скоринг \-\>

    совместимость размерностей \-\> допуск абсолютных и относительных дельт.  
    """  
    is\_code\_matched \= bool(  
        spec\_item.catalog\_code   
        and est\_item.catalog\_code   
        and spec\_item.catalog\_code.strip() \== est\_item.catalog\_code.strip()  
    )  
      
    score \= 100.0 if is\_code\_matched else float(  
        fuzz.token\_sort\_ratio(spec\_item.clean\_description, est\_item.clean\_description)  
    )  
      
    if score \< fuzzy\_threshold:  
        return None  \# Позиции семантически не соответствуют друг другу

    \# Шлюз физической совместимости единиц измерения  
    if spec\_item.canonical\_unit \!= est\_item.canonical\_unit:  
        return AuditCollision(  
            collision\_type="INCOMPATIBLE\_UNITS",  
            severity="CRITICAL",  
            spec\_ref=spec\_item.pointer,  
            estimate\_ref=est\_item.pointer,  
            spec\_volume=spec\_item.quantity,  
            estimate\_volume\_normalized=est\_item.quantity,  
            volume\_difference=Decimal("0.00"),  
            confidence\_score=score,  
            description=(  
                f"Коллизия размерностей: РД указывает \[{spec\_item.raw\_unit}\], "  
                f"а смета составлена в \[{est\_item.raw\_unit}\]."  
            )  
        )

    \# Приведение объема сметы с учетом коэффициента кратности расценки  
    normalized\_est\_qty \= est\_item.quantity \* est\_item.multiplier\_to\_base  
    delta \= normalized\_est\_qty \- spec\_item.quantity  
      
    \# Расчет допустимого диапазона отклонения  
    allowed\_boundary \= abs\_tolerance \+ (rel\_tolerance \* spec\_item.quantity)  
      
    if delta \> allowed\_boundary:  
        \# Диагностика типовой сметной ошибки неприменения делителя расценки (100 м2, 1000 шт)  
        is\_multiplier\_error \= (  
            normalized\_est\_qty \>= spec\_item.quantity \* Decimal("98")  
            and normalized\_est\_qty \<= spec\_item.quantity \* Decimal("102")  
        )  
          
        detected\_type \= "MULTIPLIER\_ERROR" if is\_multiplier\_error else "VOLUME\_OVERRUN"  
        criticality \= "CRITICAL" if (is\_multiplier\_error or delta \> spec\_item.quantity \* Decimal("0.10")) else "WARNING"  
          
        return AuditCollision(  
            collision\_type=detected\_type,  
            severity=criticality,  
            spec\_ref=spec\_item.pointer,  
            estimate\_ref=est\_item.pointer,  
            spec\_volume=spec\_item.quantity,  
            estimate\_volume\_normalized=normalized\_est\_qty,  
            volume\_difference=delta,  
            confidence\_score=score,  
            description=(  
                f"Обнаружено превышение объема ({detected\_type}). "  
                f"РД \[{spec\_item.pointer.sheet\_name}\!{spec\_item.pointer.cell\_coordinate}\]: {spec\_item.quantity} {spec\_item.canonical\_unit} \<-\> "  
                f"Смета \[{est\_item.pointer.sheet\_name}\!{est\_item.pointer.cell\_coordinate}\]: {normalized\_est\_qty} {est\_item.canonical\_unit}. "  
                f"Завышение: \+{delta} {spec\_item.canonical\_unit}."  
            )  
        )  
          
    return None

### **Генерация структурированного отчета о расхождениях (Audit Trail Report)**

Итоговый результат работы алгоритма транслируется в сводную матрицу выявленных дефектов, формируя готовую доказательную базу для ведения предтендерных переговоров или выставления официальных мотивированных отказов подрядным организациям:

| Координаты в РД / ВОР | Значение РД | Единица РД | Координаты в смете | Значение сметы (факт) | Нормализованное значение | Зафиксированная дельта | Классификация дефекта |
| :---- | :---- | :---- | :---- | :---- | :---- | :---- | :---- |
| АР\_Спец\!D18 | 1 420,00 | м² | Смета\_Кровля\!F44 | 14,20 | 1 420,00 | 0,00 | В пределах нормы (кратность 100 м² учтена) |
| ЭОМ\_Спец\!C31 | 450,00 | м | ЛСР\_Сети\!G89 | 450,00 | 45 000,00 | \+44 550,00 (+9900%) | Критическая ошибка кратности (100 м) |
| ОВ\_ВОР\!E12 | 240,00 | шт. | ЛСР\_Вент\!G15 | 285,00 | 285,00 | \+45,00 (+18,7%) | Необоснованное завышение объема30 |
| ВК\_Спец\!B67 | 18,00 | шт. | ЛСР\_ВК\!G92 | 18,00 | 18,00 | 0,00 (Дублирование) | Задвоение: ресурс включен в комплексную расценку31 |

Подобный подход к построению аналитического ядра обеспечивает 100% автономную локальную работу без обращения к внешним веб\-серверам22, соблюдает предельную математическую точность расчетов за счет типизации Decimal17 и гарантирует полную воспроизводимость результатов аудита с точностью до конкретной ячейки исходного проектно-сметного тома2.

#### **Works cited**

> 1. Росстат: с начала 2025 года объем работ по строительству, [https://np-mcc.ru/chlenam-sro/news/rosstat-s-nachala-2025-goda-obem-rabot-po-stroitelstvu-sostavil-1-8538-mln-rubley/](https://np-mcc.ru/chlenam-sro/news/rosstat-s-nachala-2025-goda-obem-rabot-po-stroitelstvu-sostavil-1-8538-mln-rubley/)  
> 2. testuteab/reconlify-cli \- GitHub, [https://github.com/testuteab/reconlify-cli](https://github.com/testuteab/reconlify-cli)  
> 3. Спор о стоимости работ по договору подряда \- завышение, [https://vitvet.com/articles/stroitelnyj\_podryad/spor\_o\_stoimosti\_vypolnennyh\_rabot/](https://vitvet.com/articles/stroitelnyj_podryad/spor_o_stoimosti_vypolnennyh_rabot/)  
> 4. Как проверить, не завышена ли сметная стоимость подрядчика, [https://www.smeta-expert.ru/kak\_proveritj\_ne\_zavyshena\_li\_smetnaya\_stoimostj\_podryadchika.shtml](https://www.smeta-expert.ru/kak_proveritj_ne_zavyshena_li_smetnaya_stoimostj_podryadchika.shtml)  
> 5. kondakrindirahul/BOQ-Matching: Python script to match the ... \- GitHub, [https://github.com/kondakrindirahul/BOQ-Matching](https://github.com/kondakrindirahul/BOQ-Matching)  
> 6. quantity-surveying · GitHub Topics, [https://github.com/topics/quantity-surveying](https://github.com/topics/quantity-surveying)  
> 7. GitHub \- datadrivenconstruction/QuantityTakeoff-Python: Quantity, [https://github.com/datadrivenconstruction/QuantityTakeoff-Python](https://github.com/datadrivenconstruction/QuantityTakeoff-Python)  
> 8. infra-plan/ifc-material-qto \- GitHub, [https://github.com/infra-plan/ifc-material-qto](https://github.com/infra-plan/ifc-material-qto)  
> 9. GitHub \- simondilhas/qto\_buccaneer: Quantity Takeoff tools for data, [https://github.com/simondilhas/qto\_buccaneer](https://github.com/simondilhas/qto_buccaneer)  
> 10. bill-of-quantities · GitHub Topics, [https://github.com/topics/bill-of-quantities](https://github.com/topics/bill-of-quantities)  
> 11. masterformat · GitHub Topics, [https://github.com/topics/masterformat](https://github.com/topics/masterformat)  
> 12. sergiomontey/Data-Validation-Reconciliation-Tool \- GitHub, [https://github.com/sergiomontey/Data-Validation-Reconciliation-Tool](https://github.com/sergiomontey/Data-Validation-Reconciliation-Tool)  
> 13. awesome-data-engineering.md \- GitHub, [https://github.com/icopy-site/awesome-cn/blob/master/docs/awesome/awesome-data-engineering.md](https://github.com/icopy-site/awesome-cn/blob/master/docs/awesome/awesome-data-engineering.md)  
> 14. simple-diy-electronics-inventory/README.md at main \- GitHub, [https://github.com/PleatherStarfish/simple-diy-electronics-inventory/blob/main/README.md](https://github.com/PleatherStarfish/simple-diy-electronics-inventory/blob/main/README.md)  
> 15. A BOM (Bill of Materials) comparison tool using Python and Flask, [https://github.com/nerd-oreo/bom-comparison-v3](https://github.com/nerd-oreo/bom-comparison-v3)  
> 16. GitHub \- tomexiskandar/bintang: A tiny and temporary db for quick, [https://github.com/tomexiskandar/bintang](https://github.com/tomexiskandar/bintang)  
> 17. Fariz Mohamed Fariz7154 \- GitHub, [https://github.com/Fariz7154](https://github.com/Fariz7154)  
> 18. data-reconciliation · GitHub Topics, [https://github.com/topics/data-reconciliation](https://github.com/topics/data-reconciliation)  
> 19. datadrivenconstruction/OpenConstructionERP: Open ... \- GitHub, [https://github.com/datadrivenconstruction/OpenConstructionERP](https://github.com/datadrivenconstruction/OpenConstructionERP)  
> 20. Releases · capitalone/datacompy \- GitHub, [https://github.com/capitalone/datacompy/releases](https://github.com/capitalone/datacompy/releases)  
> 21. quantity-takeoff · GitHub Topics, [https://github.com/topics/quantity-takeoff](https://github.com/topics/quantity-takeoff)  
> 22. Отделка офисов в Москве под ключ недорого \- Ремонт квартир, [https://mir-rem.ru/otdelka-ofisov/](https://mir-rem.ru/otdelka-ofisov/)  
> 23. datacompy\_develop/CLAUDE.md at develop · capitalone ... \- GitHub, [https://github.com/capitalone-contributions/datacompy\_develop/blob/develop/CLAUDE.md](https://github.com/capitalone-contributions/datacompy_develop/blob/develop/CLAUDE.md)  
> 24. question: Is it possible to specify a tolerance value for floating point, [https://github.com/paulfitz/daff/issues/155](https://github.com/paulfitz/daff/issues/155)  
> 25. erich-hs/tabularcompare: Tabular data comparison tool \- GitHub, [https://github.com/erich-hs/tabularcompare](https://github.com/erich-hs/tabularcompare)  
> 26. paulfitz/daff: align and compare tables \- GitHub, [https://github.com/paulfitz/daff](https://github.com/paulfitz/daff)  
> 27. Comparison fails on dataframes with a single column \#253 \- GitHub, [https://github.com/capitalone/datacompy/issues/253](https://github.com/capitalone/datacompy/issues/253)  
> 28. Comma: A Python CSV Library for Humans \- GitHub, [https://github.com/jlumbroso/comma](https://github.com/jlumbroso/comma)  
> 29. Releases · paulfitz/daff \- GitHub, [https://github.com/paulfitz/daff/releases](https://github.com/paulfitz/daff/releases)  
> 30. Влияние ошибок в смете на итоговую стоимость строительства, [https://ekspertiza-smetnoy-dokumentatsii.ru/vliyanie\_oshibok\_v\_smete\_na\_itogovuyu\_stoimostj\_stroiteljstva.dhtml](https://ekspertiza-smetnoy-dokumentatsii.ru/vliyanie_oshibok_v_smete_na_itogovuyu_stoimostj_stroiteljstva.dhtml)  
> 31. Проверка сметы на завышение объемов и расценок, [https://smety-v-stroitelstve.ru/proverka\_smety\_na\_zavyshenie\_objemov\_i\_rascenok\_prakticheskie\_primery.html](https://smety-v-stroitelstve.ru/proverka_smety_na_zavyshenie_objemov_i_rascenok_prakticheskie_primery.html)

[image1]: <data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAH0AAAAaCAYAAACacVPHAAADGklEQVR4Xu2Yv69MQRTHjyBRSJBXiFCIRjQikdCQvEJLo9VTUwgteqETiSg0opCoRLNRCQX/gASNRCIqGgnmY+437+zZH9eN2b37dueTnLw7Z3buzvl+d+bOfWaVSqVSWWJ2pLie4kDsqGwaTsZEGztTPE1x3OXOpPgd4oPrh/tNXvFruHvpoL6oyf6hT4xqgo7z4GxMtDHOdDiW4ovlye9LsXW4+2+bon5YHsuOscxssawJejyzrElEmry2UT1nSTHT96R4Y7nIbaFPXEjxKiaXGDRBjyc2XZO1mJwxxUyHh7ax0iMPUty2vAJWCfT4aKOabLcNTeZNUdM54FHkkdiReGurefhDj682qsk560+ToqZftFzkqZCnsPMhtyqgx08b1QTD+9KkqOnkvqe46nKnU7xz7VUDPTA+anLQtedNUdMPpfic4k7T5oDCwY3DyqqCHpgeNemToqbvTfE+xSPLB7ZrKR5bPrQsOpyuP3WI3XlYK+iB6VGTPilqOnCCp9DDKa6EvlVEbzRo8sIWQ5Pipt+yvKXds/xZj575iIAYuiaERNJrjvoZS5u8no/cX/3+lUg5H9y3D9CD70eT5zZeE1+j10eoX2O9LgrG8RhVO74ieoqbftnyl36LHQ2Mk2n+4Md9B5aLlbm0mSD3A+XjNXAfzYkx+o54sJw30oM4EfqE/yH7uqSJv5bxN921Foiu2+otbjr5lyl2xY4GJieTL9nGJIlBkx8HovkVPrDh1cA1OfqYw3qTx/TORRZEP2w0mcQk08nFFa1appmuz/q8p7MebaavpTgakw4mcdfya5zfrolJkwQmKgPbTOfHpPn1bTp6YNA0TaaZPmiuI9QooumM87tlpLMebaa3ocnw15uubdxPaL35K0FkPGCyxFFbn/OroG/T/4VJpqOx/x8HdaiWNtNBCynSWY//MV1bHeZiUjzIaQXT1jPa9zPGb3MUqjZjJZjfDhWMXUR8jT68vr5machfX288yBEDG/+4nKvps8avEg+/+EU1fRKznO/SmX4jJm1yfpHx23dpOpteqVQqlUplc/MHNHESGQbxFYoAAAAASUVORK5CYII=>

[image2]: <data:image/png;base64,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>

[image3]: <data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD8AAAAZCAYAAACGqvb0AAACPElEQVR4Xu2WTyhnURTHj1CKQqaklLJQU1IaqVkpWbCY2WhKjZ2FPUnJRpqFrVKSEhuRJQuNxW85sbCalJrFTGSpKZT/vl/3Xb/7Du96v4WX+t1Pfbu/e979nXfOO/efSCAQCORZ1oYEuqFjaC1qu+KPvWxDl9AedAGNxR8/ktZ/CdQG5aCq+KPCSZP8V+gWKo/6bK+gnqcRyTRCh1BT1O+HrsUkYUnjv0LMRzuD7iWj5OugXehA2Y+gX8qmYYLTUK9jK4NWoc+Ozee/Ourb5DugScko+VkxX1qPY592H6zcHdSg7F/ELIHaqO/z/0PZybhklPyO+INjJZMYETPmpeRPoGaoUvz+N5SdZJY8K8QgOANcbHC+AFg1jvmg7Ez+HPokpvo+/zllJwUn/y9B3Fi0jZqDSqG/YoLgC11scLqqLkkfiMnTzpb/9/nn+zUFJ5/Ea5Uv6uR/iz84XwB2s9Rj3GnPJeHzn1N2klnyi2KC0ONscD765OXZweT/QPVR3+ef79dklvx3MUFsKbs9BVy4eXGfsHA3567+0bERngI86+1J4fM/oOwks+TJqJjzuiTqs2WfdguryGApVtzSLiYRe1nphE4lHnga/6RGzE1xXcxewCsxZ5W9GSZiA0urJclXhs5noBUxM4Etb1nuS23AN2Iq6/If2hfzH/7mx3BJ45/YpaDFvePNYSUWoCH9wGECGlY2XmS+QVNQi8Tv9S5p/L9bWqFNyU/xouInNKiNxcK8PF+ngUAgkBkP9pPER/W96T0AAAAASUVORK5CYII=>

[image4]: <data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABUAAAAZCAYAAADe1WXtAAAAWUlEQVR4XmNgGAWjYFCAOegC1AARQMyCLkgpEANiY3RBaoBZUMyKLkEJ4Abi+UC8G10CBhqB+D8FOIeBSoAm3gd5mx9dkBJAkyRF9cTPCMRC6IKjYBTQEAAAJusWzDE3kc4AAAAASUVORK5CYII=>

[image5]: <data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADoAAAAZCAYAAABggz2wAAACm0lEQVR4Xu2XT6hNURTGl1BECQMD+ZuepEhGlBQKiQmixICBiQwoJkoGylASZfIyIkUZiCROhihDBkzJABFF/n6/t8929tnvnPP2O/f2CudXX+/etff93l53773WuWYdHR1/A5ukOXGwR8ZJn6Tn0nXpkTRQmlHPWumF9C7XB2lXaYbz32uF/3tL8MfkjXTXnEE/OCods8LvnvRaWvJnRj3M22jus+OlA9K30gzn/9MK/2mW6D9R2i09ldaZ+wdtWSS9kiYHsVXSV+lKEKuCBR+KYiRzXpoVxPC/E7wH7z8hitdCkqvN7fJ+Ky84hVvSryg2Vcoq4jErpfvSjChO8gvz15PM+RwvhofI8vieKD4iy8wdZ87/lGisiZc2PKHURGebm8Md5WTBTHN33O8UO9uU6Okonsxi6aN0Ih6ogbsSJ5SaKMeUOV5LpYfS7WAOO8vYwSAGWR6/HMVHxQJpUDpjIx/lzzY8oTBRXjcxV3pg5YSvBeMcb2JbgxhkeZy/rRjtjvaaKG3plDTdyslSqKDviba9o2+tOdEmVkhnrWgbrOGZuc/5akwLaUo06ehuMddiMKfltIGeGSeUkqifU7XjFKnM3BhfOj51xehIFC/Rzz5K1YsTolJSjel1IRxPv3tNiRKjbdFaAP9zxfAQ3p+eXcl2c/ewX9AGLprz9VySfkjbgthycwvm/vMa5kmPpfn5e6Cn0lv9HQX88fOwUbH/mMCivpj71mngvN5n5UdMFsdj3HdpcxCnh7Lom7kYvxGMA/5XrfCnlsT+Y8aAuep5weqLGQsblDYEMa4NlZXP8ryMT1UCxLz/zmisEh73diRqvbUvUlXwy4kqmfx82guHzd2hFJ206mLRlifSmjj4LxL+lOvo6PhP+Q3bW6kqZH8uFgAAAABJRU5ErkJggg==>

[image6]: <data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACUAAAAWCAYAAABHcFUAAAAAm0lEQVR4XmNgGAWjgDbAA4hl0QUHGoQD8Usg3gXEjGhyAwpYgTgCiM8DsRMQM6NKDzwAOciKARJ6SUDMiSo9sECPARKl74GYG01uwIE6EH8C4hog5kOTG1CgCMTzgfgbwyCJzkERUqBiwZxhECV0bwZIsXCdAVJMDCgYdOVUMAMk3YyCIQlAVUoICZguGSAPiGeRgHkg2kbBwAIAJuMcqnDkhCoAAAAASUVORK5CYII=>