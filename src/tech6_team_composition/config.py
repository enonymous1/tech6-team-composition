"""Configuration loading and validation for TECH-6 generation.

The JSON schema is intentionally close to the v1 configs: a project block,
settings, sections of experts, and deliverable details. The one breaking
change from v1 is ``template_path``: v1 loaded a template workbook and
cleared it in place, which leaked stale conditional formatting and ghost
rows into every output and let some configs overwrite the template itself.
v2 always starts from a clean sheet, so ``template_path`` is rejected with
a clear error instead of being silently honoured or silently ignored.
"""

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any


class ConfigError(ValueError):
    """Raised when a config file is missing, malformed, or invalid."""


@dataclass(frozen=True)
class Deliverable:
    id: str
    title: str
    subtitle: str
    subtext: str


@dataclass(frozen=True)
class Expert:
    n: str
    name: str
    position: str
    h_in: tuple[Any, ...] = ()
    f_in: tuple[Any, ...] = ()
    total_required_input: str | None = None


# The section whose experts fall back to the section validation threshold
# when they have no individual requirement. Compared case-insensitively
# at load time (see parse_config): a config headed "Key Experts" must not
# silently lose its section check because of capitalisation.
KEY_EXPERTS_HEADER = "KEY EXPERTS"


@dataclass(frozen=True)
class Section:
    header: str
    experts: tuple[Expert, ...]
    is_key_experts: bool = False


@dataclass(frozen=True)
class Settings:
    deliverable_labels: tuple[str, ...]
    validation_threshold: float = 6.0
    total_minimum_months: float | None = None
    field_minimum_months: float | None = None
    home_minimum_months: float | None = None


@dataclass(frozen=True)
class ProjectInfo:
    name: str
    titles: tuple[str, ...]
    sheet_name: str
    output_filename: str
    client: str | None = None
    rfp_reference: str | None = None


@dataclass(frozen=True)
class Tech6Config:
    project_info: ProjectInfo
    settings: Settings
    sections: tuple[Section, ...]
    deliverables: tuple[Deliverable, ...]
    output_directory: Path | None = None
    source_path: Path | None = None


def _require_mapping(value: Any, what: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ConfigError(f"{what} must be an object, got {type(value).__name__}")
    return value


def _require_str(mapping: dict[str, Any], key: str, what: str) -> str:
    value = mapping.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"{what} requires a non-empty string {key!r}")
    return value


def _optional_str(mapping: dict[str, Any], key: str) -> str | None:
    value = mapping.get(key)
    if value is None:
        return None
    if not isinstance(value, str):
        raise ConfigError(f"{key!r} must be a string, got {type(value).__name__}")
    return value


def _optional_number(mapping: dict[str, Any], key: str) -> float | None:
    value = mapping.get(key)
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ConfigError(f"{key!r} must be a number, got {value!r}")
    if not math.isfinite(value):
        raise ConfigError(f"{key!r} must be a finite number, got {value!r}")
    return float(value)


def _str_list(mapping: dict[str, Any], key: str, what: str) -> tuple[str, ...]:
    value = mapping.get(key)
    if not isinstance(value, list) or not value:
        raise ConfigError(f"{what} requires a non-empty list {key!r}")
    items: list[str] = []
    for item in value:
        if not isinstance(item, str) or not item.strip():
            raise ConfigError(f"{what} {key!r} must contain only non-empty strings")
        items.append(item)
    return tuple(items)


def _input_list(mapping: dict[str, Any], key: str, label_count: int, what: str) -> tuple[Any, ...]:
    """Validate one expert's home/field input list.

    Entries must be strings, numbers, or null. Anything else — booleans,
    objects, nested lists — used to reach the generator, where it was
    stringified into text that Excel's ``SUM`` silently ignores, so the
    expert's total came out low with no warning. Numbers must be finite
    and non-negative: months of input cannot be negative, and a negative
    input would silently reduce a total instead of flagging an error.
    """
    value = mapping.get(key)
    if value is None:
        return ()
    if not isinstance(value, list):
        raise ConfigError(f"{what} {key!r} must be a list, got {type(value).__name__}")
    if len(value) != label_count:
        raise ConfigError(
            f"{what} {key!r} has {len(value)} entries but there are "
            f"{label_count} deliverable labels; lengths must match"
        )
    for index, item in enumerate(value):
        if isinstance(item, bool) or not isinstance(item, (str, int, float, type(None))):
            raise ConfigError(f"{what} {key!r}[{index}] must be a string, number, or null, got {item!r}")
        if isinstance(item, (int, float)):
            if not math.isfinite(item):
                raise ConfigError(f"{what} {key!r}[{index}] must be finite, got {item!r}")
            if item < 0:
                raise ConfigError(f"{what} {key!r}[{index}] must not be negative, got {item!r}")
    return tuple(value)


def _check_duplicates(items: list[str], what: str) -> None:
    seen: set[str] = set()
    for item in items:
        if item in seen:
            raise ConfigError(f"Duplicate {what}: {item!r}")
        seen.add(item)


_SHEET_NAME_FORBIDDEN = set("[]:*?/\\")


def _validate_sheet_name(sheet_name: str) -> None:
    """Reject sheet names Excel (and openpyxl) cannot build.

    Previously this surfaced only at generation time, as an openpyxl
    ``ValueError`` traceback, so ``validate`` passed configs that
    ``generate`` then rejected. Checking here makes the failure a
    clean config error at load time.
    """
    if len(sheet_name) > 31:
        raise ConfigError(f"project_info sheet_name must be at most 31 characters, got {len(sheet_name)}")
    bad = sorted(set(sheet_name) & _SHEET_NAME_FORBIDDEN)
    if bad:
        raise ConfigError(f"project_info sheet_name must not contain any of {bad}, got {sheet_name!r}")


def load_config(path: str | Path) -> Tech6Config:
    """Load and validate a TECH-6 JSON config file."""
    config_path = Path(path)
    if not config_path.exists():
        raise ConfigError(f"Config file not found: {config_path}")
    try:
        text = config_path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ConfigError(f"Could not read config file: {config_path}: {exc}") from exc
    except UnicodeDecodeError as exc:
        raise ConfigError(f"Config file is not valid UTF-8: {config_path}: {exc}") from exc
    try:
        raw = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ConfigError(f"Config file is not valid JSON: {config_path}: {exc}") from exc
    return parse_config(raw, source_path=config_path)


def parse_config(raw: Any, source_path: Path | None = None) -> Tech6Config:
    """Validate a decoded JSON document and build a Tech6Config."""
    data = _require_mapping(raw, "Config")

    if data.get("template_path"):
        raise ConfigError(
            "template_path is no longer supported in v2: generation always "
            "starts from a clean sheet, so a template workbook is not used. "
            "Remove template_path from the config."
        )

    project_raw = _require_mapping(data.get("project_info"), "project_info")
    sheet_name = _require_str(project_raw, "sheet_name", "project_info")
    _validate_sheet_name(sheet_name)
    project_info = ProjectInfo(
        name=_require_str(project_raw, "name", "project_info"),
        titles=_str_list(project_raw, "titles", "project_info"),
        sheet_name=sheet_name,
        output_filename=_require_str(project_raw, "output_filename", "project_info"),
        client=_optional_str(project_raw, "client"),
        rfp_reference=_optional_str(project_raw, "rfp_reference"),
    )
    if Path(project_info.output_filename).name != project_info.output_filename:
        raise ConfigError(
            f"project_info output_filename must be a plain file name, got {project_info.output_filename!r}"
        )

    settings_raw = _require_mapping(data.get("settings"), "settings")
    labels = _str_list(settings_raw, "deliverable_labels", "settings")
    _check_duplicates(list(labels), "deliverable label")
    threshold = _optional_number(settings_raw, "validation_threshold")
    requirements_raw = settings_raw.get("time_input_requirements")
    if requirements_raw is None:
        requirements_raw = {}
    requirements = _require_mapping(requirements_raw, "time_input_requirements")
    total_min = _optional_number(requirements, "total_minimum_months")
    field_min = _optional_number(requirements, "field_minimum_months")
    home_min = _optional_number(requirements, "home_minimum_months")
    for key, minimum in (
        ("validation_threshold", threshold),
        ("total_minimum_months", total_min),
        ("field_minimum_months", field_min),
        ("home_minimum_months", home_min),
    ):
        # A negative minimum makes its check vacuously pass (no total
        # can fall below it), which is worse than having no check.
        if minimum is not None and minimum < 0:
            raise ConfigError(f"settings {key!r} must not be negative, got {minimum!r}")
    settings = Settings(
        deliverable_labels=labels,
        validation_threshold=6.0 if threshold is None else threshold,
        total_minimum_months=total_min,
        field_minimum_months=field_min,
        home_minimum_months=home_min,
    )

    sections_raw = data.get("sections")
    if not isinstance(sections_raw, list) or not sections_raw:
        raise ConfigError("Config requires a non-empty list 'sections'")
    sections: list[Section] = []
    expert_numbers: list[str] = []
    for index, section_raw in enumerate(sections_raw):
        section_map = _require_mapping(section_raw, f"sections[{index}]")
        header = _require_str(section_map, "header", f"sections[{index}]")
        experts_raw = section_map.get("experts")
        if not isinstance(experts_raw, list) or not experts_raw:
            raise ConfigError(f"sections[{index}] requires a non-empty list 'experts'")
        experts: list[Expert] = []
        for expert_index, expert_raw in enumerate(experts_raw):
            expert_map = _require_mapping(expert_raw, f"sections[{index}].experts[{expert_index}]")
            what = f"sections[{index}].experts[{expert_index}]"
            # A blank name is allowed: the bundled examples use blank
            # names for placeholder experts. A non-string name is not.
            name = _optional_str(expert_map, "name")
            experts.append(
                Expert(
                    n=_require_str(expert_map, "n", what),
                    name="" if name is None else name,
                    position=_require_str(expert_map, "position", what),
                    h_in=_input_list(expert_map, "h_in", len(labels), what),
                    f_in=_input_list(expert_map, "f_in", len(labels), what),
                    total_required_input=_optional_str(expert_map, "total_required_input"),
                )
            )
            expert_numbers.append(experts[-1].n)
        sections.append(
            Section(
                header=header,
                experts=tuple(experts),
                is_key_experts=header.strip().upper() == KEY_EXPERTS_HEADER,
            )
        )
    _check_duplicates(expert_numbers, "expert number")

    deliverables_raw = data.get("deliverable_details")
    if not isinstance(deliverables_raw, list) or not deliverables_raw:
        raise ConfigError("Config requires a non-empty list 'deliverable_details'")
    deliverables: list[Deliverable] = []
    for index, deliverable_raw in enumerate(deliverables_raw):
        deliverable_map = _require_mapping(deliverable_raw, f"deliverable_details[{index}]")
        what = f"deliverable_details[{index}]"
        subtitle = _optional_str(deliverable_map, "subtitle")
        subtext = _optional_str(deliverable_map, "subtext")
        deliverables.append(
            Deliverable(
                id=_require_str(deliverable_map, "id", what),
                title=_require_str(deliverable_map, "title", what),
                subtitle="" if subtitle is None else subtitle,
                subtext="" if subtext is None else subtext,
            )
        )
    # The legend and the input columns describe the same deliverables:
    # a detail entry must exist for every label, in the same order, or
    # the table gains a column with no legend entry (or vice versa).
    if len(deliverables) != len(labels):
        raise ConfigError(
            f"deliverable_details has {len(deliverables)} entries but there are "
            f"{len(labels)} deliverable labels; counts must match"
        )
    for index, (deliverable, label) in enumerate(zip(deliverables, labels)):
        if deliverable.id != label:
            raise ConfigError(
                f"deliverable_details[{index}].id {deliverable.id!r} does not match "
                f"deliverable_labels[{index}] {label!r}"
            )

    output_directory = data.get("output_directory")
    if output_directory is not None and not isinstance(output_directory, str):
        raise ConfigError("output_directory must be a string")

    return Tech6Config(
        project_info=project_info,
        settings=settings,
        sections=tuple(sections),
        deliverables=tuple(deliverables),
        output_directory=Path(output_directory) if output_directory else None,
        source_path=source_path,
    )
