"""Tests for the CLI."""

import subprocess
import sys

import pytest

from tech6_team_composition.cli import main


def test_cli_version(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["--version"])
    assert exc_info.value.code == 0
    assert "2.0.0" in capsys.readouterr().out


def test_cli_generate(examples_dir, tmp_path, capsys):
    assert main(["generate", str(examples_dir / "minimal.json"), "--output-dir", str(tmp_path)]) == 0
    assert (tmp_path / "tech6_minimal_example.xlsx").exists()
    assert "Generated:" in capsys.readouterr().out


def test_cli_generate_missing_config(tmp_path, capsys):
    assert main(["generate", "nope.json", "--output-dir", str(tmp_path)]) == 2
    assert "Error:" in capsys.readouterr().err


def test_cli_validate(examples_dir, capsys):
    assert main(["validate", str(examples_dir / "validation.json")]) == 0
    assert "Valid:" in capsys.readouterr().out


def test_cli_validate_bad_config(tmp_path, capsys):
    bad = tmp_path / "bad.json"
    bad.write_text('{"settings": {}}', encoding="utf-8")
    assert main(["validate", str(bad)]) == 2
    assert "Error:" in capsys.readouterr().err


def test_cli_format_matrix(examples_dir, tmp_path, capsys):
    assert (
        main(
            [
                "format-matrix",
                "--input",
                str(examples_dir / "qualifications_matrix.xlsx"),
                "--output-dir",
                str(tmp_path),
                "--docx",
                str(examples_dir / "qualifications_styles.docx"),
            ]
        )
        == 0
    )
    assert (tmp_path / "qualifications_matrix_Formatted.xlsx").exists()
    assert "Formatted:" in capsys.readouterr().out


def test_cli_format_matrix_missing_input(tmp_path, capsys):
    assert main(["format-matrix", "--input", str(tmp_path / "missing.xlsx"), "--output-dir", str(tmp_path)]) == 2
    assert "Error:" in capsys.readouterr().err


def test_cli_menu_eof_exits_cleanly(examples_dir, tmp_path, monkeypatch, capsys):
    monkeypatch.setattr("builtins.input", lambda _: (_ for _ in ()).throw(EOFError))
    assert main(["menu", "--configs-dir", str(examples_dir), "--output-dir", str(tmp_path)]) == 0


def test_cli_menu_quit(examples_dir, tmp_path, monkeypatch):
    monkeypatch.setattr("builtins.input", lambda _: "q")
    assert main(["menu", "--configs-dir", str(examples_dir), "--output-dir", str(tmp_path)]) == 0


def test_cli_menu_selects_and_generates(examples_dir, tmp_path, monkeypatch):
    answers = iter(["3", "n"])
    monkeypatch.setattr("builtins.input", lambda _: next(answers))
    # Sorted example configs: basic, full, minimal, phase2, validation.
    # "3" is minimal.json.
    assert main(["menu", "--configs-dir", str(examples_dir), "--output-dir", str(tmp_path)]) == 0
    assert (tmp_path / "tech6_minimal_example.xlsx").exists()


def test_module_entry_point(examples_dir, tmp_path):
    result = subprocess.run(
        [sys.executable, "-m", "tech6_team_composition", "validate", str(examples_dir / "minimal.json")],
        capture_output=True,
        text=True,
        cwd=tmp_path,
    )
    assert result.returncode == 0, result.stderr
    assert "Valid:" in result.stdout
