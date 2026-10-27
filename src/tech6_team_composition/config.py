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


@dataclass(frozen=True)
class Section:
    header: str
    experts: tuple[Expert, ...]


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
    return tuple(value)


def load_config(path: str | Path) -> Tech6Config:
    """Load and validate a TECH-6 JSON config file."""
    config_path = Path(path)
    if not config_path.exists():
        raise ConfigError(f"Config file not found: {config_path}")
    try:
        raw = json.loads(config_path.read_text(encoding="utf-8"))
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
    project_info = ProjectInfo(
        name=_require_str(project_raw, "name", "project_info"),
        titles=_str_list(project_raw, "titles", "project_info"),
        sheet_name=_require_str(project_raw, "sheet_name", "project_info"),
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
    threshold = _optional_number(settings_raw, "validation_threshold")
    requirements_raw = settings_raw.get("time_input_requirements")
    if requirements_raw is None:
        requirements_raw = {}
    requirements = _require_mapping(requirements_raw, "time_input_requirements")
    settings = Settings(
        deliverable_labels=labels,
        validation_threshold=6.0 if threshold is None else threshold,
        total_minimum_months=_optional_number(requirements, "total_minimum_months"),
        field_minimum_months=_optional_number(requirements, "field_minimum_months"),
        home_minimum_months=_optional_number(requirements, "home_minimum_months"),
    )

    sections_raw = data.get("sections")
    if not isinstance(sections_raw, list) or not sections_raw:
        raise ConfigError("Config requires a non-empty list 'sections'")
    sections: list[Section] = []
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
            experts.append(
                Expert(
                    n=_require_str(expert_map, "n", what),
                    name=expert_map.get("name", ""),
                    position=_require_str(expert_map, "position", what),
                    h_in=_input_list(expert_map, "h_in", len(labels), what),
                    f_in=_input_list(expert_map, "f_in", len(labels), what),
                    total_required_input=_optional_str(expert_map, "total_required_input"),
                )
            )
            if not isinstance(experts[-1].name, str):
                raise ConfigError(f"{what} 'name' must be a string")
        sections.append(Section(header=header, experts=tuple(experts)))

    deliverables_raw = data.get("deliverable_details")
    if not isinstance(deliverables_raw, list) or not deliverables_raw:
        raise ConfigError("Config requires a non-empty list 'deliverable_details'")
    deliverables: list[Deliverable] = []
    for index, deliverable_raw in enumerate(deliverables_raw):
        deliverable_map = _require_mapping(deliverable_raw, f"deliverable_details[{index}]")
        what = f"deliverable_details[{index}]"
        deliverables.append(
            Deliverable(
                id=_require_str(deliverable_map, "id", what),
                title=_require_str(deliverable_map, "title", what),
                subtitle=deliverable_map.get("subtitle", ""),
                subtext=deliverable_map.get("subtext", ""),
            )
        )
        for field_name in ("subtitle", "subtext"):
            if not isinstance(getattr(deliverables[-1], field_name), str):
                raise ConfigError(f"{what} {field_name!r} must be a string")

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
