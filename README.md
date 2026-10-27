# tech6-team-composition

Build a World Bank FORM TECH-6 workbook — team composition, assignment, and key experts' inputs — from a JSON config, with live Excel formulas instead of hand-typed totals.

Define the experts, their home/field inputs per deliverable, and the deliverable legend once in JSON. The generator writes the formatted sheet, sums each expert's inputs, subtotals each section, and flags totals below a required minimum in red. No template workbook, no manual row insertion, no stale formatting carried over from a previous run.

A second command reformats a qualifications summary matrix, re-applying bold phrases from a reference Word document as rich text.

## Quick start

```bash
pip install .
tech6-team-composition generate examples/validation.json
# -> output/tech6_validation_example.xlsx
```

Or pick a bundled example interactively:

```bash
tech6-team-composition menu
```

Validate a config without writing anything:

```bash
tech6-team-composition validate examples/full.json
```

## Commands

| Command | What it does |
|---|---|
| `generate CONFIG [--output-dir DIR] [--output-name NAME]` | Build the TECH-6 workbook. Default output directory is `./output`. |
| `validate CONFIG` | Check a config's schema and report its sections, experts, and deliverables. |
| `format-matrix --input FILE [--input FILE ...] [--docx FILE] [--output-dir DIR]` | Reformat matrix workbook(s), writing `<name>_Formatted.xlsx`. |
| `menu [--configs-dir DIR]` | Interactively choose a config and generate it. |

## Config format

```json
{
  "project_info": {
    "name": "Validation Example",
    "titles": ["Section 3. Technical Proposal – Standard Forms",
               "FORM TECH-6 (FOR FTP AND STP)",
               "TEAM COMPOSITION, ASSIGNMENT, AND KEY EXPERTS’ INPUTS"],
    "sheet_name": "Validation Test",
    "output_filename": "tech6_validation_example.xlsx"
  },
  "settings": {
    "validation_threshold": 6,
    "deliverable_labels": ["D-1", "D-2", "D-3", "D-4"],
    "time_input_requirements": {"total_minimum_months": 13.5, "field_minimum_months": 7.5}
  },
  "sections": [
    {"header": "KEY EXPERTS", "experts": [
      {"n": "K-1", "name": "Mr. Team Leader", "position": "Team Leader",
       "h_in": ["0.5", "0.5", "0", ""],
       "f_in": ["0.5", "0.5", "0.1", ""],
       "total_required_input": "2.0 (1.0 Field / 1.0 Remote)"}
    ]}
  ],
  "deliverable_details": [
    {"id": "D-1", "title": "Inception Report", "subtitle": "Week 2",
     "subtext": "Methodology, work plan, stakeholder meetings, and mobilization."}
  ]
}
```

Rules the loader enforces:

- `h_in` / `f_in` lists must match the number of `deliverable_labels`. Omit them for blank inputs.
- Input values may be numbers, or bracketed proposal shorthand such as `"[2 month]"` or `"[0.5 m]"` — the leading number is used. Non-numeric text is kept as text; `SUM` ignores it.
- `total_required_input` is a per-expert minimum, e.g. `"2.0 (1.0 Field / 1.0 Remote)"` — only the first number is the requirement. Experts without one fall back to the section `validation_threshold` (KEY EXPERTS only, default 6).
- Grand totals are checked against `time_input_requirements` when present. If only total and field minimums are given, the home minimum is derived as total minus field.
- All three checks are minimum checks: totals *below* the requirement are flagged red. (v1's section check flagged totals *above* the threshold; see CHANGELOG.)
- `output_directory` in a config is honoured when `--output-dir` is not given. `template_path` is rejected: v2 always starts from a clean sheet, so no template file is used or overwritten.

## Format a qualifications matrix

```bash
tech6-team-composition format-matrix \
  --input examples/qualifications_matrix.xlsx \
  --docx examples/qualifications_styles.docx
```

The source matrix uses literal `<br>` separators, which become line breaks, and education/experience columns get bullet prefixes. Bold runs in the reference document's first table are re-applied to matching cells (matched by expert name and column). Without `--docx`, the matrix is reformatted without bold phrases.

## Bundled examples

All examples use fictional people and generic deliverables.

| File | Shape |
|---|---|
| `examples/basic.json` | 7 deliverables, placeholder inputs |
| `examples/minimal.json` | 1 deliverable, 2 experts |
| `examples/full.json` | 12 experts, bracketed shorthand inputs |
| `examples/phase2.json` | 3 deliverables, higher section threshold |
| `examples/validation.json` | Per-expert requirements + grand-total minimums |
| `examples/qualifications_matrix.xlsx`, `examples/qualifications_styles.docx` | Fictional matrix + bold-phrase reference for `format-matrix` |

## About TECH-6

FORM TECH-6 (for FTP and STP) is a standard form used in multilateral development bank procurement (World Bank, ADB, JICA, and others) for technical proposals. This tool is unaffiliated with any of those institutions; it only automates filling in the form's layout. No copy of the official form is shipped in this repo.

## Development

```bash
pip install -e ".[dev]"
pytest -q
ruff check src tests && ruff format --check src tests
mypy src
```

CI runs the same gates on Python 3.11–3.13.

## License

MIT — see [LICENSE](LICENSE).
