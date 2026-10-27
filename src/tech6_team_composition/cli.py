"""Command-line interface for tech6-team-composition."""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from . import __version__
from .config import ConfigError, load_config
from .format_matrix import format_matrix
from .generator import Tech6Generator

logger = logging.getLogger(__name__)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="tech6-team-composition",
        description="Generate FORM TECH-6 Excel workbooks and format qualifications matrices.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    subparsers = parser.add_subparsers(dest="command", required=True)

    gen = subparsers.add_parser("generate", help="Generate a TECH-6 workbook from a JSON config.")
    gen.add_argument("config", help="Path to a JSON configuration file")
    gen.add_argument("--output-dir", default=None, help="Output directory (default: ./output)")
    gen.add_argument("--output-name", default=None, help="Override the output file name from the config")

    val = subparsers.add_parser("validate", help="Validate a JSON config without writing output.")
    val.add_argument("config", help="Path to a JSON configuration file")

    fmt = subparsers.add_parser("format-matrix", help="Reformat qualifications summary matrix workbook(s).")
    fmt.add_argument("--input", dest="inputs", action="append", required=True, help="Source .xlsx file (repeatable)")
    fmt.add_argument("--output-dir", default="output", help="Output directory (default: output)")
    fmt.add_argument("--docx", default=None, help="Reference .docx whose bold runs supply bold phrases")
    fmt.add_argument("--sheet", default=None, help="Worksheet name to read (default: active sheet)")

    menu = subparsers.add_parser("menu", help="Interactively pick a bundled example config and generate it.")
    menu.add_argument(
        "--configs-dir",
        default=None,
        help="Directory of JSON configs to offer (default: bundled examples)",
    )
    menu.add_argument("--output-dir", default="output", help="Output directory (default: output)")
    return parser


def _cmd_generate(args: argparse.Namespace) -> int:
    try:
        generator = Tech6Generator.from_config_file(args.config)
        out_path = generator.generate(output_dir=args.output_dir, output_filename=args.output_name)
    except (ConfigError, FileNotFoundError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    print(f"Generated: {out_path}")
    return 0


def _cmd_validate(args: argparse.Namespace) -> int:
    try:
        config = load_config(args.config)
    except (ConfigError, FileNotFoundError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    expert_count = sum(len(s.experts) for s in config.sections)
    print(f"Valid: {config.project_info.name}")
    print(f"  Sections: {len(config.sections)} ({expert_count} experts)")
    print(f"  Deliverables: {len(config.settings.deliverable_labels)} labels, {len(config.deliverables)} details")
    print(f"  Output: {config.project_info.output_filename} (sheet: {config.project_info.sheet_name})")
    return 0


def _cmd_format_matrix(args: argparse.Namespace) -> int:
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    for input_path in args.inputs:
        source = Path(input_path)
        out_path = output_dir / f"{source.stem}_Formatted.xlsx"
        try:
            result = format_matrix(source, out_path, docx_path=args.docx, sheet_name=args.sheet)
        except (FileNotFoundError, ValueError) as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 2
        print(f"Formatted: {result}")
    return 0


def _cmd_menu(args: argparse.Namespace) -> int:
    configs_dir = (
        Path(args.configs_dir) if args.configs_dir else Path(__file__).resolve().parent.parent.parent / "examples"
    )
    if not configs_dir.exists():
        print(f"Error: configs directory not found: {configs_dir}", file=sys.stderr)
        return 2
    config_files = sorted(configs_dir.rglob("*.json"))
    if not config_files:
        print(f"No configuration files found in {configs_dir}")
        return 0
    while True:
        try:
            print("\n--- TECH-6: World Bank Form Generator ---")
            print("Available configurations:")
            for i, config in enumerate(config_files, 1):
                print(f"{i}. {config.relative_to(configs_dir)}")
            print("q. Quit")
            choice = input("\nSelect a configuration to run (number): ").strip().lower()
            if choice == "q":
                return 0
            idx = int(choice) - 1
            if not 0 <= idx < len(config_files):
                print("Invalid selection.")
                continue
            selected = config_files[idx]
            print(f"\nRunning generator for: {selected}")
            try:
                generator = Tech6Generator.from_config_file(selected)
                out_path = generator.generate(output_dir=args.output_dir)
                print(f"Generated: {out_path}")
            except (ConfigError, FileNotFoundError) as exc:
                print(f"Error: {exc}", file=sys.stderr)
            cont = input("Run another? (y/n): ").strip().lower()
            if cont != "y":
                return 0
        except (EOFError, KeyboardInterrupt):
            print()
            return 0
        except ValueError:
            print("Please enter a number.")


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    parser = _build_parser()
    args = parser.parse_args(argv)
    if args.command == "generate":
        return _cmd_generate(args)
    if args.command == "validate":
        return _cmd_validate(args)
    if args.command == "format-matrix":
        return _cmd_format_matrix(args)
    if args.command == "menu":
        return _cmd_menu(args)
    parser.error(f"Unknown command: {args.command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
