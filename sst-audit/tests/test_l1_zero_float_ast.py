"""L1 Static Gate: Статический AST-анализатор соблюдения Zero Float Policy.

Выполняет жесткую инспекцию синтаксического дерева Python (AST) в
продакшн-модулях проекта:
1. Запрет литералов чисел с плавающей точкой (например, 0.05, 1.0).
2. Запрет явных вызовов конструктора float(...).
3. Запрет типовых аннотаций параметров (: float) и возвратов (-> float).
4. Запрет псевдонимов типов PEP 695 с использованием float.

Политика допуска:
* ядро (``models``/``normalizer``/``evaluator``/``db``) и все новые модули —
  строгий запрет, файлы находятся автосканированием дерева проекта;
* пакет ``parsers/`` — ``float`` разрешен только внутри геометрических
  конструкций (bbox ячеек, отступы, границы графов, линии таблицы, поворот
  листа): такие значения приходят из pdfplumber и сразу сериализуются строками;
* легаси-модули из ``QUARANTINED_MODULES`` временно исключены из проверки, но
  ratchet-тесты не дают карантину разрастаться и требуют выводить модуль из
  карантина, как только он станет чистым.
"""

import ast
from dataclasses import dataclass
from pathlib import Path
from typing import Final, TypeAlias
import pytest

# ----------------------------------------------------------------------
# 1. Структуры данных и AST-посетитель
# ----------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class ZeroFloatViolation:
    """Запись о выявленном нарушении Zero Float Policy в исходном коде."""

    file_path: str
    line_number: int
    column_offset: int
    violation_type: str
    code_snippet: str
    message: str


ViolationList: TypeAlias = list[ZeroFloatViolation]

# Ключевые слова геометрии листа: bbox ячеек, отступы, границы графов,
# линии таблицы, поворот страницы. Только внутри этих конструкций пакет
# ``parsers/`` вправе работать с координатами PDF в float: значения приходят
# из pdfplumber и сразу сериализуются строками (``_coordinate_text``).
_GEOMETRY_TOKENS: Final[tuple[str, ...]] = (
    "bbox",
    "box",
    "crop",
    "x0",
    "x1",
    "top",
    "bottom",
    "left",
    "right",
    "width",
    "height",
    "pt",
    "point",
    "tolerance",
    "snap",
    "column",
    "col",
    "bound",
    "edge",
    "line",
    "word",
    "layout",
    "rotation",
    "stamp",
    "margin",
    "padding",
    "gap",
    "offset",
    "coord",
    "page_meta",
    "scale",
    "geometry",
    "header",
    "row",
    "cell",
)

GeometrySpan: TypeAlias = tuple[int, int]


def _is_geometry_name(name: str) -> bool:
    """Проверяет, относится ли имя конструкции к геометрии листа."""
    lowered = name.lower()
    return any(token in lowered for token in _GEOMETRY_TOKENS)


def _node_names(node: ast.AST) -> tuple[str, ...]:
    """Возвращает имена, которыми конструкция называет себя в коде."""
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
        return (node.name,)
    if isinstance(node, ast.Assign):
        return tuple(
            target.id for target in node.targets if isinstance(target, ast.Name)
        )
    if isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
        return (node.target.id,)
    if isinstance(node, ast.arg):
        return (node.arg,)
    if isinstance(node, ast.Dict):
        return tuple(
            key.value
            for key in node.keys
            if isinstance(key, ast.Constant) and isinstance(key.value, str)
        )
    if isinstance(node, ast.keyword) and node.arg:
        return (node.arg,)
    return ()


def geometry_spans(tree: ast.AST) -> tuple[GeometrySpan, ...]:
    """Диапазоны строк геометрических конструкций (функции, поля, словари).

    Разрешение на float действует только внутри этих диапазонов, поэтому
    бизнес-логика парсеров (нормализация объема, расчет стоимости) остается
    под полным запретом ``float``.
    """
    spans: list[GeometrySpan] = []
    for node in ast.walk(tree):
        if not any(_is_geometry_name(name) for name in _node_names(node)):
            continue
        end_line = getattr(node, "end_lineno", None)
        if end_line is None:
            continue
        spans.append((node.lineno, end_line))
    return tuple(spans)


class ZeroFloatASTVisitor(ast.NodeVisitor):
    """AST-посетитель для детекции конструкций float в коде расчетного ядра."""

    def __init__(
        self,
        file_path: str,
        source_lines: list[str],
        geometry_spans: tuple[GeometrySpan, ...] = (),
    ) -> None:
        self.file_path: Final[str] = file_path
        self.source_lines: Final[list[str]] = source_lines
        # Диапазоны строк, где пакет parsers вправе держать координаты в float.
        self.geometry_spans: Final[tuple[GeometrySpan, ...]] = geometry_spans
        self.violations: list[ZeroFloatViolation] = []

    def allowed_by_geometry(self, line_number: int) -> bool:
        """Строка попадает в разрешенную геометрическую конструкцию."""
        return any(
            start <= line_number <= end for start, end in self.geometry_spans
        )

    def _get_snippet(self, node: ast.AST) -> str:
        """Извлекает строку исходного кода, соответствующую узлу AST."""
        lineno = getattr(node, "lineno", None)
        if lineno is not None and 1 <= lineno <= len(self.source_lines):
            return self.source_lines[lineno - 1].strip()
        return "<unknown snippet>"

    def _contains_float_type(self, node: ast.AST | None) -> bool:
        """Рекурсивно проверяет, содержит ли выражение аннотации ссылку на float."""
        if node is None:
            return False
        for child in ast.walk(node):
            if isinstance(child, ast.Name) and child.id == "float":
                return True
        return False

    def visit_Constant(self, node: ast.Constant) -> None:
        """Перехватывает литералы float стандарта IEEE 754."""
        if isinstance(node.value, float):
            self.violations.append(
                ZeroFloatViolation(
                    file_path=self.file_path,
                    line_number=node.lineno,
                    column_offset=node.col_offset,
                    violation_type="FLOAT_LITERAL",
                    code_snippet=self._get_snippet(node),
                    message=(
                        f"Обнаружен литерал типа float ({node.value!r}). "
                        f"Используйте строковый литерал в Decimal('...')."
                    ),
                )
            )
        self.generic_visit(node)

    def visit_Call(self, node: ast.Call) -> None:
        """Перехватывает функциональные вызовы float(...)."""
        if isinstance(node.func, ast.Name) and node.func.id == "float":
            self.violations.append(
                ZeroFloatViolation(
                    file_path=self.file_path,
                    line_number=node.lineno,
                    column_offset=node.col_offset,
                    violation_type="FLOAT_CALL",
                    code_snippet=self._get_snippet(node),
                    message="Запрещен прямой вызов функции float(). Преобразуйте данные в Decimal.",
                )
            )
        self.generic_visit(node)

    def visit_AnnAssign(self, node: ast.AnnAssign) -> None:
        """Перехватывает аннотации переменных: x: float = ..."""
        if self._contains_float_type(node.annotation):
            self.violations.append(
                ZeroFloatViolation(
                    file_path=self.file_path,
                    line_number=node.lineno,
                    column_offset=node.col_offset,
                    violation_type="FLOAT_VAR_ANNOTATION",
                    code_snippet=self._get_snippet(node),
                    message="Запрещено объявление переменной с типом float. Используйте Decimal.",
                )
            )
        self.generic_visit(node)

    def visit_FunctionDef(self, node: ast.FunctionDef) -> None:
        """Перехватывает сигнатуры функций с аннотациями float (аргументы и возвраты)."""
        # Проверка возвращаемого значения: def f() -> float:
        if node.returns and self._contains_float_type(node.returns):
            self.violations.append(
                ZeroFloatViolation(
                    file_path=self.file_path,
                    line_number=node.returns.lineno,
                    column_offset=node.returns.col_offset,
                    violation_type="FLOAT_RETURN_ANNOTATION",
                    code_snippet=self._get_snippet(node.returns),
                    message=f"Функция '{node.name}' объявляет возвращаемый тип float. Замените на Decimal.",
                )
            )

        # Проверка аргументов: def f(x: float):
        all_args = (
            node.args.posonlyargs
            + node.args.args
            + node.args.kwonlyargs
            + ([node.args.vararg] if node.args.vararg else [])
            + ([node.args.kwarg] if node.args.kwarg else [])
        )
        for arg in all_args:
            if arg.annotation and self._contains_float_type(arg.annotation):
                self.violations.append(
                    ZeroFloatViolation(
                        file_path=self.file_path,
                        line_number=arg.lineno,
                        column_offset=arg.col_offset,
                        violation_type="FLOAT_ARG_ANNOTATION",
                        code_snippet=self._get_snippet(arg),
                        message=f"Аргумент '{arg.arg}' функции '{node.name}' типизирован как float.",
                    )
                )

        self.generic_visit(node)

    def visit_TypeAlias(self, node: ast.TypeAlias) -> None:
        """Перехватывает синтаксис PEP 695: type MyNumber = float."""
        if self._contains_float_type(node.value):
            self.violations.append(
                ZeroFloatViolation(
                    file_path=self.file_path,
                    line_number=node.lineno,
                    column_offset=node.col_offset,
                    violation_type="PEP695_FLOAT_TYPE_ALIAS",
                    code_snippet=self._get_snippet(node),
                    message=f"Псевдоним типа PEP 695 '{node.name.id}' ссылается на float.",
                )
            )
        self.generic_visit(node)


# ----------------------------------------------------------------------
# 2. Вспомогательные сервисные функции анализа
# ----------------------------------------------------------------------


def scan_source_string(
    code: str,
    virtual_name: str = "<memory>",
    *,
    allow_geometry: bool = False,
) -> ViolationList:
    """Парсит и валидирует переданную строку кода на предмет нарушений.

    :param code: исходный текст модуля.
    :param virtual_name: имя модуля в сообщениях об ошибках.
    :param allow_geometry: разрешить ``float`` внутри геометрических
        конструкций (только для пакета ``parsers``: координаты PDF).
    :return: список нарушений Zero Float Policy.
    """
    tree = ast.parse(code, filename=virtual_name)
    lines = code.splitlines()
    spans = geometry_spans(tree) if allow_geometry else ()
    visitor = ZeroFloatASTVisitor(virtual_name, lines, spans)
    visitor.visit(tree)
    if not spans:
        return visitor.violations
    return [
        violation
        for violation in visitor.violations
        if not visitor.allowed_by_geometry(violation.line_number)
    ]


def scan_python_file(
    file_path: Path,
    *,
    allow_geometry: bool = False,
) -> ViolationList:
    """Выполняет статический анализ содержимого физического файла Python."""
    content = file_path.read_text(encoding="utf-8")
    return scan_source_string(
        content,
        virtual_name=str(file_path),
        allow_geometry=allow_geometry,
    )


# ----------------------------------------------------------------------
# 3. Модульные тесты корректности самого AST-посетителя
# ----------------------------------------------------------------------


class TestZeroFloatASTVisitorLogic:
    """Набор тестов, подтверждающий чувствительность анализатора к нарушениям."""

    def test_detects_float_literal(self) -> None:
        """Проверяет детектирование литералов 0.05 и 10.0."""
        code = "waste_ratio = 0.05\nbase_mult = 10.0\n"
        violations = scan_source_string(code)
        assert len(violations) == 2
        assert violations[0].violation_type == "FLOAT_LITERAL"
        assert "0.05" in violations[0].message
        assert violations[1].violation_type == "FLOAT_LITERAL"
        assert "10.0" in violations[1].message

    def test_detects_float_call(self) -> None:
        """Проверяет детектирование вызова float('1.25')."""
        code = "parsed_volume = float('1.25')\n"
        violations = scan_source_string(code)
        assert len(violations) == 1
        assert violations[0].violation_type == "FLOAT_CALL"

    def test_detects_variable_annotation(self) -> None:
        """Проверяет запрет аннотаций переменных типом float."""
        code = "from decimal import Decimal\ntotal_sum: float = Decimal('100.0')\n"
        violations = scan_source_string(code)
        assert len(violations) == 1
        assert violations[0].violation_type == "FLOAT_VAR_ANNOTATION"

    def test_detects_function_argument_and_return_types(self) -> None:
        """Проверяет запрет float в сигнатурах функций."""
        code = "def calc_cost(volume: float, rate: float | None = None) -> float:\n    pass\n"
        violations = scan_source_string(code)
        types = [v.violation_type for v in violations]
        assert "FLOAT_RETURN_ANNOTATION" in types
        assert types.count("FLOAT_ARG_ANNOTATION") == 2

    def test_detects_pep695_type_alias_violation(self) -> None:
        """Проверяет запрет использования float в синтаксисе PEP 695 (Python 3.12)."""
        code = "type RateValue = float\n"
        violations = scan_source_string(code)
        assert len(violations) == 1
        assert violations[0].violation_type == "PEP695_FLOAT_TYPE_ALIAS"

    def test_allows_isinstance_guard_clauses(self) -> None:
        """Защитные проверки типов isinstance(x, float) не должны считаться нарушением."""
        code = (
            "def validate(x):\n"
            "    if isinstance(x, float):\n"
            "        raise TypeError('Floats forbidden')\n"
            "    return x\n"
        )
        violations = scan_source_string(code)
        assert len(violations) == 0

    def test_allows_clean_decimal_code(self) -> None:
        """Детерминированный сметный код на Decimal обязан проходить без замечаний."""
        code = (
            "from decimal import Decimal\n"
            "type Volume = Decimal\n"
            "def calculate_total(vol: Volume, waste: Decimal) -> Decimal:\n"
            "    return vol * (Decimal('1.0') + waste)\n"
        )
        violations = scan_source_string(code)
        assert len(violations) == 0


# ----------------------------------------------------------------------
# 4. Политика допуска: пакет parsers и карантин легаси-модулей
# ----------------------------------------------------------------------

PROJECT_ROOT: Final[Path] = Path(__file__).resolve().parent.parent
# Каталоги, которые не являются продакшн-кодом конвейера.
_EXCLUDED_DIRS: Final[frozenset[str]] = frozenset(
    {"tests", ".venv", "venv", "__pycache__", "scripts", "backups", "reports"}
)
# Пакет, которому разрешены координаты PDF в float — строго внутри
# геометрических конструкций (см. ``geometry_spans``).
GEOMETRY_PACKAGE: Final[str] = "parsers"
# Легаси-модули, еще не переведенные на Decimal. Список закрыт: новый код
# нельзя прикрыть карантином (test_quarantine_does_not_cover_parser_package),
# а очистившийся модуль обязан быть из него выведен (ratchet-тест).
QUARANTINED_MODULES: Final[tuple[str, ...]] = (
    "dispatcher.py",
    "exporter.py",
    "main.py",
    "matcher.py",
)


def _production_modules() -> list[str]:
    """Относительные пути всех продакшн-модулей проекта без тестов и скриптов."""
    modules: list[str] = []
    for path in sorted(PROJECT_ROOT.rglob("*.py")):
        relative = path.relative_to(PROJECT_ROOT)
        if _EXCLUDED_DIRS.intersection(relative.parts[:-1]):
            continue
        modules.append(relative.as_posix())
    return modules


def _allows_geometry(relative_path: str) -> bool:
    """Геометрический допуск действует только внутри пакета ``parsers``."""
    return relative_path.split("/")[0] == GEOMETRY_PACKAGE


def _scan_module(relative_path: str) -> ViolationList:
    """Сканирует продакшн-модуль с учетом области действия допуска."""
    return scan_python_file(
        PROJECT_ROOT / relative_path,
        allow_geometry=_allows_geometry(relative_path),
    )


def _format_violations(title: str, violations: ViolationList) -> str:
    """Формирует читаемый отчет о нарушениях для ``pytest.fail``."""
    lines = [f"{title} ({len(violations)} нарушений):"]
    for violation in violations:
        lines.append(
            f"  [{violation.violation_type}] Строка "
            f"{violation.line_number}:{violation.column_offset} -> "
            f"{violation.message}\n      Фрагмент: {violation.code_snippet}"
        )
    return "\n".join(lines)


@pytest.mark.parametrize(
    "relative_module_path",
    [
        "models.py",
        "normalizer.py",
        "evaluator.py",
        "db.py",
    ],
)
def test_production_core_modules_strictly_adhere_to_zero_float(
    relative_module_path: str,
) -> None:
    """Гарантирует, что ключевые расчетные файлы ядра не содержат типов и литералов float."""
    target_file = PROJECT_ROOT / relative_module_path

    assert target_file.is_file(), (
        f"Критический модуль ядра {relative_module_path} не найден по пути {target_file}"
    )

    violations = scan_python_file(target_file)

    if violations:
        pytest.fail(
            _format_violations(
                f"Файл {relative_module_path} нарушает Zero Float Policy", violations
            )
        )


def test_all_production_modules_adhere_to_zero_float() -> None:
    """Автосканирование всех продакшн-модулей вне карантина.

    Проверка не опирается на ручной список файлов: любой новый модуль ядра
    или парсеров попадает под запрет ``float`` автоматически.
    """
    scanned: list[str] = []
    offenders: list[str] = []

    for relative_path in _production_modules():
        scanned.append(relative_path)
        if relative_path in QUARANTINED_MODULES:
            continue
        violations = _scan_module(relative_path)
        if violations:
            offenders.append(_format_violations(relative_path, violations))

    # Защита от «зеленого» теста, который ничего не сканирует.
    assert "models.py" in scanned, scanned
    assert f"{GEOMETRY_PACKAGE}/common.py" in scanned, scanned

    if offenders:
        pytest.fail("\n\n".join(offenders))


def test_parser_package_still_needs_geometry_allowance() -> None:
    """Ratchet: если геометрия парсеров стала целочисленной, допуск пора убирать."""
    strict_violations = sum(
        len(scan_python_file(PROJECT_ROOT / relative_path))
        for relative_path in _production_modules()
        if _allows_geometry(relative_path)
    )
    assert strict_violations > 0, (
        "Пакет parsers больше не содержит float даже без геометрического допуска: "
        "исключение _GEOMETRY_TOKENS можно удалить."
    )


@pytest.mark.parametrize("relative_module_path", QUARANTINED_MODULES)
def test_quarantined_module_is_still_dirty(relative_module_path: str) -> None:
    """Ratchet: очищенный легаси-модуль обязан быть выведен из карантина."""
    target_file = PROJECT_ROOT / relative_module_path
    assert target_file.is_file(), f"Карантинный модуль {relative_module_path} не найден"

    violations = _scan_module(relative_module_path)
    assert violations, (
        f"Модуль {relative_module_path} больше не содержит float и должен быть "
        f"исключен из QUARANTINED_MODULES: карантин не может быть вечным."
    )


def test_quarantine_does_not_cover_parser_package() -> None:
    """Новый код парсеров нельзя прикрыть карантином."""
    assert not [path for path in QUARANTINED_MODULES if _allows_geometry(path)]


def test_geometry_allowance_is_enabled_only_for_parser_package() -> None:
    """Допуск включается исключительно для модулей пакета ``parsers``."""
    assert _allows_geometry(f"{GEOMETRY_PACKAGE}/lsr_pdf.py")
    assert not _allows_geometry("models.py")
    assert not _allows_geometry("main.py")


class TestGeometryAllowancePolicy:
    """Допуск покрывает только геометрию листа и не ослабляет бизнес-логику."""

    def test_geometry_floats_are_allowed_only_with_allowance(self) -> None:
        """Координаты CropBox: запрещены строго, разрешены в пакете parsers."""
        code = (
            "def _get_crop_box(page):\n"
            "    return (15.0, 15.0, page.width - 15.0, 0.0)\n"
        )
        strict = scan_source_string(code)
        assert len(strict) == 4
        assert {v.violation_type for v in strict} == {"FLOAT_LITERAL"}
        assert scan_source_string(code, allow_geometry=True) == []

    def test_business_logic_float_stays_forbidden_with_allowance(self) -> None:
        """Расчетная арифметика остается под запретом даже в пакете parsers."""
        code = "def calculate_total_cost(volume):\n    return volume * 1.18\n"
        assert len(scan_source_string(code)) == 1
        assert len(scan_source_string(code, allow_geometry=True)) == 1

    def test_module_level_float_outside_geometry_names_is_forbidden(self) -> None:
        """Литерал вне геометрической конструкции не покрывается допуском."""
        code = "WASTE_RATIO = 0.05\nTOLERANCE_RATIO = 0.1\n"
        assert len(scan_source_string(code)) == 2
        relaxed = scan_source_string(code, allow_geometry=True)
        # Допуск снимает только TOLERANCE_RATIO: имя относится к геометрии.
        assert [v.line_number for v in relaxed] == [1]

    def test_geometry_spans_cover_only_named_constructions(self) -> None:
        """Диапазоны допуска строятся по именам функций, полей и словарей."""
        assert geometry_spans(ast.parse("CROP_LEFT_PT = 15.5\n")) == ((1, 1),)
        assert geometry_spans(ast.parse("WASTE_RATIO = 0.05\n")) == ()
        spans = geometry_spans(
            ast.parse("def _page_layout(page):\n    box = 1.5\n    return box\n")
        )
        assert spans == ((1, 3), (2, 2))

    def test_float_annotation_allowed_inside_geometry_field(self) -> None:
        """Аннотация ``float`` допустима для поля геометрии и запрещена вне неё."""
        code = "class _Layout:\n    bounds: tuple[float, ...] = ()\n"
        assert len(scan_source_string(code)) == 1
        assert scan_source_string(code, allow_geometry=True) == []
