"""Tests for the CLI."""

import subprocess
import sys

import pytest

from tech6_team_composition.cli import main


def test_cli_version(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["--version"])
    assert exc_info.value.code == 0
    assert "2.1.0" in capsys.readouterr().out


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


def _bad_sheet_config(tmp_path, examples_dir):
    import json
    import shutil

    config_path = tmp_path / "bad_sheet.json"
    shutil.copy(examples_dir / "minimal.json", config_path)
    data = json.loads(config_path.read_text(encoding="utf-8"))
    data["project_info"]["sheet_name"] = "Bad/Name?"
    config_path.write_text(json.dumps(data), encoding="utf-8")
    return config_path


def test_cli_generate_bad_sheet_name_is_clean_error(tmp_path, examples_dir, capsys):
    config_path = _bad_sheet_config(tmp_path, examples_dir)
    assert main(["generate", str(config_path), "--output-dir", str(tmp_path)]) == 2
    err = capsys.readouterr().err
    assert "Error:" in err
    assert "Traceback" not in err


def test_cli_validate_bad_sheet_name_is_clean_error(tmp_path, examples_dir, capsys):
    config_path = _bad_sheet_config(tmp_path, examples_dir)
    assert main(["validate", str(config_path)]) == 2
    err = capsys.readouterr().err
    assert "Error:" in err
    assert "Traceback" not in err


def test_cli_generate_directory_as_config_is_clean_error(tmp_path, capsys):
    assert main(["generate", str(tmp_path), "--output-dir", str(tmp_path)]) == 2
    err = capsys.readouterr().err
    assert "Error:" in err
    assert "Traceback" not in err


def test_cli_format_matrix_missing_sheet_is_clean_error(examples_dir, tmp_path, capsys):
    assert (
        main(
            [
                "format-matrix",
                "--input",
                str(examples_dir / "qualifications_matrix.xlsx"),
                "--output-dir",
                str(tmp_path),
                "--sheet",
                "Nope",
            ]
        )
        == 2
    )
    err = capsys.readouterr().err
    assert "Error:" in err
    assert "Traceback" not in err


def test_cli_format_matrix_missing_input_creates_no_output_dir(tmp_path, capsys):
    out_dir = tmp_path / "newdir"
    assert main(["format-matrix", "--input", str(tmp_path / "missing.xlsx"), "--output-dir", str(out_dir)]) == 2
    assert "Error:" in capsys.readouterr().err
    assert not out_dir.exists()


def test_cli_menu_without_bundled_examples_is_clean_error(tmp_path, monkeypatch, capsys):
    import tech6_team_composition.cli as cli_module

    monkeypatch.setattr(cli_module, "_default_configs_dir", lambda: tmp_path / "nope")
    assert main(["menu"]) == 2
    err = capsys.readouterr().err
    assert "Bundled examples are not installed" in err
    assert "Traceback" not in err
