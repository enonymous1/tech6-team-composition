"""Tests for config loading and validation."""

import pytest

from tech6_team_composition.config import ConfigError, load_config, parse_config

EXAMPLES = ["basic.json", "full.json", "minimal.json", "phase2.json", "validation.json"]


def _base_config() -> dict:
    return {
        "project_info": {
            "name": "Example",
            "titles": ["T1", "T2", "T3"],
            "sheet_name": "Team Input",
            "output_filename": "out.xlsx",
        },
        "settings": {"deliverable_labels": ["D-1", "D-2"], "validation_threshold": 6},
        "sections": [
            {
                "header": "KEY EXPERTS",
                "experts": [{"n": "K-1", "name": "Jane Q. Expert", "position": "Team Leader"}],
            }
        ],
        "deliverable_details": [
            {"id": "D-1", "title": "One", "subtitle": "Week 2", "subtext": "Text"},
            {"id": "D-2", "title": "Two", "subtitle": "Week 4", "subtext": "Text"},
        ],
    }


@pytest.mark.parametrize("name", EXAMPLES)
def test_examples_load(examples_dir, name):
    config = load_config(examples_dir / name)
    assert config.project_info.output_filename.endswith(".xlsx")
    assert config.settings.deliverable_labels
    assert config.deliverables
    assert config.sections


def test_missing_file():
    with pytest.raises(ConfigError, match="not found"):
        load_config("no-such-config.json")


def test_invalid_json(tmp_path):
    bad = tmp_path / "bad.json"
    bad.write_text("{not json", encoding="utf-8")
    with pytest.raises(ConfigError, match="not valid JSON"):
        load_config(bad)


def test_missing_project_info():
    data = _base_config()
    del data["project_info"]
    with pytest.raises(ConfigError, match="project_info"):
        parse_config(data)


def test_missing_settings_key():
    data = _base_config()
    del data["settings"]["deliverable_labels"]
    with pytest.raises(ConfigError, match="deliverable_labels"):
        parse_config(data)


def test_h_in_length_must_match_labels():
    data = _base_config()
    data["sections"][0]["experts"][0]["h_in"] = ["1.0"]
    with pytest.raises(ConfigError, match="lengths must match"):
        parse_config(data)


def test_f_in_length_must_match_labels():
    data = _base_config()
    data["sections"][0]["experts"][0]["f_in"] = ["1.0", "2.0", "3.0"]
    with pytest.raises(ConfigError, match="lengths must match"):
        parse_config(data)


def test_expert_inputs_default_to_blank():
    config = parse_config(_base_config())
    expert = config.sections[0].experts[0]
    assert expert.h_in == ()
    assert expert.f_in == ()


def test_template_path_rejected():
    data = _base_config()
    data["template_path"] = "templates/master.xlsx"
    with pytest.raises(ConfigError, match="template_path is no longer supported"):
        parse_config(data)


def test_output_filename_must_be_plain_name():
    data = _base_config()
    data["project_info"]["output_filename"] = "sub/dir/out.xlsx"
    with pytest.raises(ConfigError, match="plain file name"):
        parse_config(data)


def test_defaults():
    data = _base_config()
    del data["settings"]["validation_threshold"]
    config = parse_config(data)
    assert config.settings.validation_threshold == 6.0
    assert config.settings.total_minimum_months is None


def test_time_input_requirements_parsed():
    data = _base_config()
    data["settings"]["time_input_requirements"] = {
        "total_minimum_months": 13.5,
        "field_minimum_months": 7.5,
    }
    config = parse_config(data)
    assert config.settings.total_minimum_months == 13.5
    assert config.settings.field_minimum_months == 7.5
    assert config.settings.home_minimum_months is None


def test_non_numeric_threshold_rejected():
    data = _base_config()
    data["settings"]["validation_threshold"] = "six"
    with pytest.raises(ConfigError, match="must be a number"):
        parse_config(data)


def test_empty_sections_rejected():
    data = _base_config()
    data["sections"] = []
    with pytest.raises(ConfigError, match="non-empty"):
        parse_config(data)


def test_deliverable_requires_id_and_title():
    data = _base_config()
    del data["deliverable_details"][0]["title"]
    with pytest.raises(ConfigError, match="title"):
        parse_config(data)


def test_extra_project_info_fields_kept():
    data = _base_config()
    data["project_info"]["client"] = "Fictional Client"
    data["project_info"]["rfp_reference"] = "{Insert RFP reference}"
    config = parse_config(data)
    assert config.project_info.client == "Fictional Client"
    assert config.project_info.rfp_reference == "{Insert RFP reference}"


def test_sheet_name_forbidden_characters_rejected():
    data = _base_config()
    data["project_info"]["sheet_name"] = "Bad/Name"
    with pytest.raises(ConfigError, match="sheet_name"):
        parse_config(data)


def test_sheet_name_too_long_rejected():
    data = _base_config()
    data["project_info"]["sheet_name"] = "x" * 32
    with pytest.raises(ConfigError, match="at most 31"):
        parse_config(data)


def test_deliverable_count_mismatch_rejected():
    data = _base_config()
    data["settings"]["deliverable_labels"] = ["D-1", "D-2", "D-3"]
    with pytest.raises(ConfigError, match="counts must match"):
        parse_config(data)


def test_deliverable_id_mismatch_rejected():
    data = _base_config()
    data["deliverable_details"][1]["id"] = "D-9"
    with pytest.raises(ConfigError, match="does not match"):
        parse_config(data)


def test_duplicate_deliverable_label_rejected():
    data = _base_config()
    data["settings"]["deliverable_labels"] = ["D-1", "D-1"]
    with pytest.raises(ConfigError, match="Duplicate"):
        parse_config(data)


def test_duplicate_expert_number_rejected():
    data = _base_config()
    data["sections"][0]["experts"].append({"n": "K-1", "name": "Second Person", "position": "Deputy"})
    with pytest.raises(ConfigError, match="Duplicate"):
        parse_config(data)


def test_negative_threshold_rejected():
    data = _base_config()
    data["settings"]["validation_threshold"] = -1
    with pytest.raises(ConfigError, match="must not be negative"):
        parse_config(data)


def test_negative_minimum_rejected():
    data = _base_config()
    data["settings"]["time_input_requirements"] = {"total_minimum_months": -13.5}
    with pytest.raises(ConfigError, match="must not be negative"):
        parse_config(data)


def test_negative_input_rejected():
    data = _base_config()
    data["sections"][0]["experts"][0]["h_in"] = [-1, 0]
    with pytest.raises(ConfigError, match="must not be negative"):
        parse_config(data)


def test_bool_input_rejected():
    data = _base_config()
    data["sections"][0]["experts"][0]["h_in"] = [True, 0]
    with pytest.raises(ConfigError, match="must be a string, number, or null"):
        parse_config(data)


def test_dict_input_rejected():
    data = _base_config()
    data["sections"][0]["experts"][0]["h_in"] = [{"months": 1}, 0]
    with pytest.raises(ConfigError, match="must be a string, number, or null"):
        parse_config(data)


def test_non_finite_threshold_rejected():
    data = _base_config()
    data["settings"]["validation_threshold"] = float("nan")
    with pytest.raises(ConfigError, match="finite"):
        parse_config(data)


def test_blank_expert_name_allowed():
    # Bundled examples use blank names for placeholder experts; a blank
    # name is a legitimate placeholder, a non-string name is not.
    data = _base_config()
    data["sections"][0]["experts"][0]["name"] = ""
    config = parse_config(data)
    assert config.sections[0].experts[0].name == ""
    data["sections"][0]["experts"][0]["name"] = 42
    with pytest.raises(ConfigError, match="'name' must be a string"):
        parse_config(data)


def test_key_experts_header_case_insensitive():
    data = _base_config()
    data["sections"][0]["header"] = "Key Experts"
    config = parse_config(data)
    assert config.sections[0].is_key_experts is True
    data["sections"][0]["header"] = "NON-KEY EXPERTS"
    config = parse_config(data)
    assert config.sections[0].is_key_experts is False


def test_load_directory_is_config_error(tmp_path):
    with pytest.raises(ConfigError, match="Could not read"):
        load_config(tmp_path)


def test_load_non_utf8_config_is_config_error(tmp_path):
    bad = tmp_path / "bad.json"
    bad.write_bytes(b"\xff\xfe{not json")
    with pytest.raises(ConfigError, match="not valid UTF-8"):
        load_config(bad)
