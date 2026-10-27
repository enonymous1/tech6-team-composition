# Changelog

## [2.0.0] - 2026-10-27

Modernization of the original single-commit `mp-tech6` tool (2026-07-19),
rebuilt as a sanitized public export under the name `tech6-team-composition`.
All sample data is fictional; the original configs and reference files
carrying real proposal personnel and client details were not carried over.

### Added

- Installable `tech6_team_composition` package with a `tech6-team-composition`
  CLI: `generate`, `validate`, `format-matrix`, and an interactive `menu`.
- Config schema validation with named errors (`ConfigError`): missing keys,
  wrong types, `h_in`/`f_in` length mismatches, and non-plain output names
  are rejected before anything is written.
- 74 pytest tests, ruff + mypy gates, and GitHub Actions CI (Python 3.11–3.13).
- MIT LICENSE, landing-page README, and this changelog.
- Freeze panes below the headers, landscape fit-to-width page setup, and a
  `fitToPage` flag (v1 set `fitToWidth` without it, so Excel could ignore it).

### Fixed

- **Configs can no longer overwrite their own inputs.** v1 configs pointed
  at a template/reference workbook and wrote the output over it; v2
  always builds on a fresh sheet and never loads a template.
- **Stale conditional formatting no longer leaks into outputs.** v1
  cleared cell values but not formatting, so old rules accumulated in
  the committed templates (2–3 identical copies). v2 builds each
  workbook's rules from scratch; repeated runs produce identical rule sets.
- **Section validation direction.** v1's section check used `greaterThan`
  (flagging totals *above* the threshold) while per-expert and grand-total
  checks used `lessThan` (flagging totals *below* a minimum). All checks
  are now minimum checks (`lessThan`).
- **Formulas are plain `SUM`s.** v1 stored bracketed strings like
  `"[2 month]"` and parsed them back with a `SUBSTITUTE`/`VALUE`/`SEARCH`
  formula. v2 parses the leading number when the config loads, so cells
  hold numbers and formulas stay readable.
- **Ghost rows/columns removed.** v1 outputs kept the template's
  70–75-row × 25-column footprint. v2 outputs end at the legend
  (e.g. the validation example is A1:K37).
- **`format_matrix` paths are arguments, not hardcoded
  Windows paths from the original machine.** It now raises on a missing input instead
  of returning `None`, and the bare `except` around rich-text writes is
  a logged, specific fallback.
- **`run_menu` replaced by the `menu` subcommand**: sorted config list,
  in-process generation (no subprocess, no Windows-only venv path),
  clean exit on EOF/Ctrl-C, and a non-zero exit on invalid configs
  instead of a traceback.
- **Missing template no longer silently falls back to a blank workbook.**
  `template_path` is rejected outright; generation never used template
  content in the first place.
- `verify_outputs.py` (which compared two intentionally different
  outputs and exited 0 when files were missing) is replaced by pytest
  tests, including LibreOffice-recalculated totals checked during
  development (validation example: 9.1 months, matching v1).

### Changed

- `template_path` is a breaking removal: configs that set it must drop
  the key. `output_directory` is still honoured.
- Dependencies are declared and pinned by range in `pyproject.toml`
  (`openpyxl`, `XlsxWriter`, `python-docx`) instead of unpinned
  `pip install` instructions in the README.
- All files are LF with a `.gitattributes`; the original tree was CRLF
  on disk vs LF in git, so every file showed as modified on checkout.
- The `TECH-6_Qual_Text_Styling.docx` / matrix reference files and the
  official FORM TECH-6 PDF are not shipped; the bundled matrix example
  and styling reference are fictional fixtures generated for this repo.

## [1.0.0] - 2026-07-19

Original `mp-tech6`: config-driven TECH-6 Excel generator (`generator.py`),
qualifications matrix formatter (`format_matrix.py`), interactive menu
(`run_menu.py`), and output verifier (`tests/verify_outputs.py`).
