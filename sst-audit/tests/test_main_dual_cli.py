"""Тесты двухпапочного режима CLI: опрос путей и изоляция сессий ВКЛ/ИСКЛ."""

from pathlib import Path
import re
from types import SimpleNamespace

import pytest

import main


class _FakeConnection:
    """Заглушка соединения SQLite: проверяется только корректное закрытие."""

    def __init__(self) -> None:
        self.closed = False

    def close(self) -> None:
        self.closed = True


def test_headless_requires_explicit_directories(
    capsys: pytest.CaptureFixture[str],
) -> None:
    """Без путей и без терминала двухпапочный режим завершается кодом 2."""
    assert main.main(["--headless"]) == 2
    assert "--rd-dir" in capsys.readouterr().err


def test_interactive_prompts_strip_quotes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Кавычки из проводника снимаются, пути проверяются через is_dir()."""
    rd_dir = tmp_path / "ВОР"
    smeta_dir = tmp_path / "ЛСР"
    rd_dir.mkdir()
    smeta_dir.mkdir()

    answers = iter([f'"{rd_dir}"', f"'{smeta_dir}'"])
    monkeypatch.setattr(main.Prompt, "ask", staticmethod(lambda label: next(answers)))
    monkeypatch.setattr(main, "_create_startup_backup", lambda args: None)

    connection = _FakeConnection()
    monkeypatch.setattr(main.db, "get_connection", lambda path: connection)

    captured: dict[str, object] = {}

    def fake_dual(conn: object, **kwargs: object) -> main.BatchSummary:
        captured.update(kwargs)
        return main.BatchSummary(total=1, succeeded=1, failed=0, hitl_pending=0)

    monkeypatch.setattr(main, "run_dual_batch_pipeline", fake_dual)

    assert main.main([]) == 0
    assert connection.closed is True
    assert captured["rd_dir"] == rd_dir
    assert captured["smeta_dir"] == smeta_dir
    assert captured["headless"] is False
    assert captured["project_name"] == tmp_path.name
    assert str(captured["output_path"]).endswith("_Сводный.xlsx")


def test_dual_pipeline_isolates_vkl_and_iskl(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """ВКЛ и ИСКЛ получают разные session_id и разные файлы смет."""
    rd_dir = tmp_path / "ВОР"
    smeta_dir = tmp_path / "ЛСР"
    rd_dir.mkdir()
    smeta_dir.mkdir()
    (rd_dir / "02-01-01_ВОР.pdf").write_bytes(b"")
    for name in (
        "02-01-01_ЛСР_включение.xlsx",
        "02-01-01_ЛСР_исключение.xlsx",
    ):
        (smeta_dir / name).write_bytes(b"")

    configs: list[main.SessionConfig] = []

    def fake_run(
        conn: object,
        config: main.SessionConfig,
        *,
        interactive: bool,
        reviewer: str,
    ) -> SimpleNamespace:
        configs.append(config)
        return SimpleNamespace(session_id=config.session_id, hitl_pending=0)

    monkeypatch.setattr(main, "run_pipeline", fake_run)

    exported: dict[str, object] = {}

    def fake_export(
        conn: object,
        session_ids: object,
        output_path: object,
        *,
        section_hints: object,
    ) -> Path:
        exported["session_ids"] = list(session_ids)  # type: ignore[arg-type]
        exported["section_hints"] = dict(section_hints)  # type: ignore[arg-type]
        return Path(output_path)

    monkeypatch.setattr(main.exporter, "export_sessions_to_excel", fake_export)

    summary = main.run_dual_batch_pipeline(
        object(),  # type: ignore[arg-type]
        project_name="Объект",
        rd_dir=rd_dir,
        smeta_dir=smeta_dir,
        output_path=tmp_path / "out.xlsx",
    )

    assert summary.total == 2
    assert summary.succeeded == 2
    assert len({config.session_id for config in configs}) == 2
    for config in configs:
        assert re.fullmatch(
            r"Объект_02-01-01_(ВКЛ|ИСКЛ)_[0-9a-f]{6}",
            config.session_id or "",
        )
        assert config.smeta_sheet is None
    assert {Path(config.smeta_path).name for config in configs} == {
        "02-01-01_ЛСР_включение.xlsx",
        "02-01-01_ЛСР_исключение.xlsx",
    }
    assert exported["session_ids"] == [config.session_id for config in configs]
    assert exported["section_hints"] == {
        config.session_id: ("02-01-01", config.session_id.split("_")[2])
        for config in configs
    }