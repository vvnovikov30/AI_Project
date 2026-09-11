"""L1 Static Gate: Статический AST-анализатор соблюдения Zero Float Policy.

Выполняет жесткую инспекцию синтаксического дерева Python (AST) в расчетных
модулях ETL-конвейера:
1. Запрет литералов чисел с плавающей точкой (например, 0.05, 1.0).
2. Запрет явных вызовов конструктора float(...).
3. Запрет типовых аннотаций параметров (: float) и возвратов (-> float).
4. Запрет псевдонимов типов PEP 695 с использованием float.
"""

import ast
from dataclasses import dataclass
from pathlib import Path
from typing import Final
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


type ViolationList = list[ZeroFloatViolation]


class ZeroFloatASTVisitor(ast.NodeVisitor):
    """AST-посетитель для детекции конструкций float в коде расчетного ядра."""

    def __init__(self, file_path: str, source_lines: list[str]) -> None:
        self.file_path: Final[str] = file_path
        self.source_lines: Final[list[str]] = source_lines
        self.violations: list[ZeroFloatViolation] = []

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


def scan_source_string(code: str, virtual_name: str = "<memory>") -> ViolationList:
    """Парсит и валидирует переданную строку кода на предмет нарушений."""
    tree = ast.parse(code, filename=virtual_name)
    lines = code.splitlines()
    visitor = ZeroFloatASTVisitor(virtual_name, lines)
    visitor.visit(tree)
    return visitor.violations


def scan_python_file(file_path: Path) -> ViolationList:
    """Выполняет статический анализ содержимого физического файла Python."""
    content = file_path.read_text(encoding="utf-8")
    return scan_source_string(content, virtual_name=str(file_path))


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
# 4. Проверка реальных модулей расчетного ядра проекта
# ----------------------------------------------------------------------


@pytest.mark.parametrize(
    "relative_module_path",
    [
        "models.py",
        "normalizer.py",
        "evaluator.py",
    ],
)
def test_production_core_modules_strictly_adhere_to_zero_float(
    relative_module_path: str,
) -> None:
    """Гарантирует, что ключевые расчетные файлы ядра не содержат типов и литералов float."""
    # Поиск корня проекта относительно текущего тестового каталога
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    target_file = project_root / relative_module_path

    assert target_file.is_file(), f"Критический модуль ядра {relative_module_path} не найден по пути {target_file}"

    violations = scan_python_file(target_file)

    if violations:
        error_lines = [
            f"Файл {relative_module_path} нарушает Zero Float Policy ({len(violations)} нарушений):"
        ]
        for v in violations:
            error_lines.append(
                f"  [{v.violation_type}] Строка {v.line_number}:{v.column_offset} -> {v.message}\n"
                f"      Фрагмент: {v.code_snippet}"
            )
        pytest.fail("\n".join(error_lines))