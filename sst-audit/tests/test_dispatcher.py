"""Тесты автоматического сопоставления томов РД и листов многолистовой сметы."""

from pathlib import Path

from openpyxl import Workbook
import pytest

from dispatcher import MAPPING_RULES, ProjectDispatcher, score_pair


def _make_smeta(path: Path) -> Path:
    """Создает смету с двумя разделами ЛС и служебными листами."""
    workbook = Workbook()

    service = workbook.active
    service.title = "ССРСС"
    service["A8"] = "Сводный сметный расчет (АПС)"

    aps = workbook.create_sheet("ЛС 02-01-08")
    aps["A8"] = "Автоматическая пожарная сигнализация (АПС)"

    sot = workbook.create_sheet("ЛС 02-01-14")
    # Проверяем чтение заголовка из A9, если A8 пуст.
    sot["A9"] = "Система охранного телевидения (СОТ)"

    other = workbook.create_sheet("ОСР")
    other["A8"] = "Объектная смета (СОТ)"

    workbook.save(path)
    return path


def test_mapping_rules_cover_required_sections() -> None:
    """Словарь содержит все обязательные аббревиатуры разделов."""
    assert {
        "АПС",
        "СОУЭ",
        "РТС",
        "СКТ",
        "СКС",
        "СОТ",
        "СКУД",
        "СДС",
        "ЭОМ",
        "ОВ",
        "ВК",
        "АР",
    } <= set(MAPPING_RULES)


def test_match_volumes_maps_pdf_to_ls_sheet(tmp_path: Path) -> None:
    """PDF с аббревиатурой раздела привязывается к профильному листу ЛС."""
    rd_dir = tmp_path / "rd"
    rd_dir.mkdir()
    (rd_dir / "Том 1 (АПС) 08.04.2026.pdf").write_bytes(b"")
    (rd_dir / "Том 2 (СОТ) 08.04.2026.pdf").write_bytes(b"")
    (rd_dir / "Пояснительная записка.pdf").write_bytes(b"")
    (rd_dir / "readme.txt").write_text("not a pdf", encoding="utf-8")

    smeta_path = _make_smeta(tmp_path / "smeta.xlsx")
    pairs = ProjectDispatcher().match_volumes(rd_dir, smeta_path)

    assert [(pdf.name, sheet) for pdf, sheet, _score in pairs] == [
        ("Том 1 (АПС) 08.04.2026.pdf", "ЛС 02-01-08"),
        ("Том 2 (СОТ) 08.04.2026.pdf", "ЛС 02-01-14"),
    ]
    assert all(score >= ProjectDispatcher.MIN_CONFIDENCE for _, _, score in pairs)


def test_match_volumes_ignores_service_sheets_and_unrelated_pdf(
    tmp_path: Path,
) -> None:
    """Служебные листы и PDF без подтвержденного раздела не попадают в пары."""
    rd_dir = tmp_path / "rd"
    rd_dir.mkdir()
    (rd_dir / "Разное.pdf").write_bytes(b"")

    smeta_path = _make_smeta(tmp_path / "smeta.xlsx")
    sections = ProjectDispatcher().read_sections(smeta_path)

    assert [section.name for section in sections] == [
        "ЛС 02-01-08",
        "ЛС 02-01-14",
    ]
    assert ProjectDispatcher().match_volumes(rd_dir, smeta_path) == []


def test_score_pair_requires_section_evidence() -> None:
    """Высокая лексическая близость без подтверждения раздела не проходит порог."""
    no_evidence = score_pair("Том 5", "ЛС 02-01-14", "Система охраны")
    confirmed = score_pair(
        "Том 5 (СОТ)",
        "ЛС 02-01-14",
        "Система охранного телевидения (СОТ)",
    )

    assert no_evidence < ProjectDispatcher.MIN_CONFIDENCE
    assert confirmed >= ProjectDispatcher.MIN_CONFIDENCE
    assert confirmed > no_evidence


def test_duplicate_section_pdf_does_not_reuse_sheet(tmp_path: Path) -> None:
    """Один лист ЛС не может быть назначен двум разным PDF-томам."""
    rd_dir = tmp_path / "rd"
    rd_dir.mkdir()
    (rd_dir / "Том 1 (АПС).pdf").write_bytes(b"")
    (rd_dir / "Том 2 (АПС).pdf").write_bytes(b"")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "ЛС 02-01-08"
    worksheet["A8"] = "Автоматическая пожарная сигнализация (АПС)"
    smeta_path = tmp_path / "smeta.xlsx"
    workbook.save(smeta_path)

    pairs = ProjectDispatcher().match_volumes(rd_dir, smeta_path)

    assert len(pairs) == 1
    assert pairs[0][1] == "ЛС 02-01-08"


def test_match_volumes_rejects_missing_input(tmp_path: Path) -> None:
    """Отсутствие папки РД или сметы дает явную ошибку."""
    with pytest.raises(FileNotFoundError):
        ProjectDispatcher().match_volumes(tmp_path / "missing", tmp_path / "x.xlsx")

    rd_dir = tmp_path / "rd"
    rd_dir.mkdir()
    with pytest.raises(FileNotFoundError):
        ProjectDispatcher().match_volumes(rd_dir, tmp_path / "missing.xlsx")
