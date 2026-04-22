"""Command-line entry point for excel-semantic-md."""

from __future__ import annotations

import argparse
import importlib
import importlib.metadata
import importlib.util
import json
import platform
import shutil
import subprocess
import tempfile
import zipfile
from collections.abc import Sequence
from pathlib import Path
from xml.etree import ElementTree


SUPPORTED_WORKBOOK_EXTENSIONS = {".xlsx", ".xlsm"}
INPUT_ERROR_EXIT_CODE = 2
NOT_IMPLEMENTED_EXIT_CODE = 1
SUCCESS_EXIT_CODE = 0


def _non_negative_int(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("must be a non-negative integer") from exc
    if parsed < 0:
        raise argparse.ArgumentTypeError("must be a non-negative integer")
    return parsed


def _print_not_implemented(command: str) -> int:
    print(f"excel-semantic-md {command}: downstream Phase 1 processing is not implemented yet.")
    return NOT_IMPLEMENTED_EXIT_CODE


def _validate_input_workbook(parser: argparse.ArgumentParser, raw_path: str) -> Path:
    path = Path(raw_path)
    if not path.exists():
        parser.error(f"input workbook does not exist: {raw_path}")
    if not path.is_file():
        parser.error(f"input workbook is not a file: {raw_path}")
    if path.suffix.lower() not in SUPPORTED_WORKBOOK_EXTENSIONS:
        extensions = " / ".join(sorted(SUPPORTED_WORKBOOK_EXTENSIONS))
        parser.error(f"input workbook must be {extensions}: {raw_path}")
    return path


def _ensure_output_directory(parser: argparse.ArgumentParser, raw_path: str) -> Path:
    path = Path(raw_path)
    if path.exists() and not path.is_dir():
        parser.error(f"output path is not a directory: {raw_path}")
    try:
        path.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        parser.error(f"failed to create output directory: {raw_path}: {exc}")
    return path


def _validate_sheet_name(parser: argparse.ArgumentParser, raw_sheet: str) -> str:
    if not raw_sheet.strip():
        parser.error("sheet name must not be empty")
    return raw_sheet


def _find_console_entry_point() -> str | None:
    try:
        entry_points = importlib.metadata.entry_points()
    except importlib.metadata.PackageNotFoundError:
        return None

    if hasattr(entry_points, "select"):
        matches = entry_points.select(group="console_scripts", name="excel-semantic-md")
    else:
        matches = [
            entry_point
            for entry_point in entry_points.get("console_scripts", [])  # type: ignore[attr-defined]
            if entry_point.name == "excel-semantic-md"
        ]

    for entry_point in matches:
        return str(entry_point.value)
    return None


def _has_importable_module(module_name: str) -> bool:
    try:
        return importlib.util.find_spec(module_name) is not None
    except ModuleNotFoundError:
        return False


def _check_output_directory(raw_path: str) -> tuple[bool, str]:
    path = Path(raw_path)
    created_dirs: list[Path] = []
    temp_path: Path | None = None

    try:
        if path.exists() and not path.is_dir():
            return False, f"not a directory: {raw_path}"

        probe = path
        while not probe.exists():
            parent = probe.parent
            if parent == probe:
                return False, f"not writable: no existing parent directory for {raw_path}"
            created_dirs.append(probe)
            probe = parent

        path.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            prefix=".excel-semantic-md-setup-",
            suffix=".tmp",
            dir=path,
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)
            temp_file.write(b"ok")
        return True, f"writable: {path}"
    except OSError as exc:
        return False, f"not writable: {raw_path}: {exc}"
    finally:
        if temp_path is not None:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
        for created_dir in created_dirs:
            try:
                created_dir.rmdir()
            except OSError:
                pass


def _run_command(command: Sequence[str], timeout_seconds: float = 3.0) -> subprocess.CompletedProcess[str] | None:
    try:
        return subprocess.run(
            list(command),
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return None


def _is_gh_copilot_available() -> bool:
    if shutil.which("gh") is None:
        return False
    result = _run_command(["gh", "copilot", "--help"])
    return result is not None and result.returncode == 0


def _describe_gh_auth_status() -> str:
    if shutil.which("gh") is None:
        return "not checked: gh executable was not found"

    result = _run_command(["gh", "auth", "status"])
    if result is None:
        return "unknown: gh auth status could not be checked"
    if result.returncode == 0:
        return "signed in or usable according to gh auth status"
    return "not signed in or unavailable according to gh auth status"


def _handle_setup(args: argparse.Namespace) -> int:
    report: list[tuple[str, str]] = []

    package = importlib.import_module("excel_semantic_md")
    package_path = getattr(package, "__file__", "unknown")
    report.append(("Python package", f"ok: excel_semantic_md imported from {package_path}"))

    entry_point = _find_console_entry_point()
    if entry_point is None:
        report.append(("CLI entry point", "not found in installed console_scripts metadata"))
    else:
        report.append(("CLI entry point", f"ok: excel-semantic-md -> {entry_point}"))

    is_windows = platform.system() == "Windows"
    report.append(("Windows environment", "ok" if is_windows else f"not Windows: {platform.system()}"))

    excel_com_available = is_windows and _has_importable_module("pythoncom") and _has_importable_module("win32com.client")
    if excel_com_available:
        report.append(("Excel COM", "available candidate: pywin32 modules are importable"))
    else:
        report.append(("Excel COM", "not available or not confirmed: Windows and pywin32 modules are required"))

    copilot_executable = shutil.which("copilot")
    gh_copilot_available = _is_gh_copilot_available()
    copilot_candidates: list[str] = []
    if copilot_executable is not None:
        copilot_candidates.append(f"copilot ({copilot_executable})")
    if gh_copilot_available:
        copilot_candidates.append("gh copilot")
    if copilot_candidates:
        report.append(("Copilot CLI", "available candidate(s): " + ", ".join(copilot_candidates)))
    else:
        report.append(("Copilot CLI", "not found: checked copilot and gh copilot"))

    if gh_copilot_available:
        report.append(("Copilot sign-in", _describe_gh_auth_status()))
    else:
        report.append(("Copilot sign-in", "not checked: gh copilot was not available"))

    skill_launcher = Path.cwd() / "skills" / "excel-semantic-markdown" / "run_excel_semantic_md.ps1"
    if skill_launcher.is_file():
        report.append(("Skill launcher", f"available: {skill_launcher}"))
    else:
        report.append(("Skill launcher", f"not found at expected path: {skill_launcher}"))

    if args.out is not None:
        ok, message = _check_output_directory(args.out)
        report.append(("Output directory", ("ok: " if ok else "not confirmed: ") + message))
    else:
        report.append(("Output directory", "not checked: pass --out to test a target directory"))

    print("excel-semantic-md setup diagnostics")
    print("This command does not install external tools, store credentials, or open workbooks.")
    for label, message in report:
        print(f"- {label}: {message}")
    print("Setup diagnostics do not guarantee end-to-end workbook conversion.")
    return SUCCESS_EXIT_CODE


def _handle_convert(args: argparse.Namespace, parser: argparse.ArgumentParser) -> int:
    _validate_input_workbook(parser, args.input)
    _ensure_output_directory(parser, args.out)
    return _print_not_implemented("convert")


def _handle_inspect(args: argparse.Namespace, parser: argparse.ArgumentParser) -> int:
    input_path = _validate_input_workbook(parser, args.input)
    from excel_semantic_md.excel import detect_blocks, read_visual_metadata, read_workbook
    from openpyxl.utils.exceptions import InvalidFileException

    try:
        result = read_workbook(input_path)
        block_model = detect_blocks(result)
        visual_model = read_visual_metadata(input_path)
    except (OSError, InvalidFileException, zipfile.BadZipFile, ElementTree.ParseError, KeyError, ValueError) as exc:
        parser.error(f"failed to read input workbook: {args.input}: {exc}")
    payload = result.to_dict()
    for sheet_payload, block_sheet, visual_sheet in zip(payload["sheets"], block_model.sheets, visual_model, strict=True):
        sheet_payload["blocks"] = [block.to_dict() for block in block_sheet.blocks]
        sheet_payload["visuals"] = [visual.to_dict() for visual in visual_sheet.visuals]
        sheet_payload["warnings"].extend(warning.to_dict() for warning in visual_sheet.warnings)
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return SUCCESS_EXIT_CODE


def _handle_render(args: argparse.Namespace, parser: argparse.ArgumentParser) -> int:
    _validate_input_workbook(parser, args.input)
    _validate_sheet_name(parser, args.sheet)
    return _print_not_implemented("render")


def _add_input_option(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--input",
        required=True,
        help="Input .xlsx or .xlsm workbook path.",
    )


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-semantic-md",
        description="Convert Excel workbooks to semantic Markdown.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    setup_parser = subparsers.add_parser(
        "setup",
        help="Check local prerequisites.",
    )
    setup_parser.add_argument(
        "--out",
        help="Optional output directory to check for writability.",
    )
    setup_parser.set_defaults(handler=_handle_setup)

    convert_parser = subparsers.add_parser(
        "convert",
        help="Convert a workbook to semantic Markdown.",
    )
    _add_input_option(convert_parser)
    convert_parser.add_argument(
        "--out",
        required=True,
        help="Output directory.",
    )
    convert_parser.add_argument(
        "--model",
        help="Text model passed through to the Copilot CLI or SDK.",
    )
    convert_parser.add_argument(
        "--vision-model",
        help="Vision model passed through to the Copilot CLI or SDK.",
    )
    convert_parser.add_argument(
        "--max-images-per-sheet",
        type=_non_negative_int,
        help="Maximum image attachments sent for each sheet.",
    )
    convert_parser.add_argument(
        "--save-debug-json",
        action="store_true",
        help="Save intermediate JSON files under debug/.",
    )
    convert_parser.add_argument(
        "--save-render-artifacts",
        action="store_true",
        help="Save rendering artifacts used as LLM context.",
    )
    convert_parser.add_argument(
        "--strict",
        action="store_true",
        help="Treat sheet-level failures as a final CLI failure.",
    )
    convert_parser.set_defaults(
        handler=lambda args: _handle_convert(args, convert_parser)
    )

    inspect_parser = subparsers.add_parser(
        "inspect",
        help="Inspect workbook structure as JSON.",
    )
    _add_input_option(inspect_parser)
    inspect_parser.set_defaults(
        handler=lambda args: _handle_inspect(args, inspect_parser)
    )

    render_parser = subparsers.add_parser(
        "render",
        help="Render a workbook sheet for local confirmation.",
    )
    _add_input_option(render_parser)
    render_parser.add_argument(
        "--sheet",
        required=True,
        help="Sheet name to render.",
    )
    render_parser.set_defaults(handler=lambda args: _handle_render(args, render_parser))

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    return int(args.handler(args))


if __name__ == "__main__":
    raise SystemExit(main())
