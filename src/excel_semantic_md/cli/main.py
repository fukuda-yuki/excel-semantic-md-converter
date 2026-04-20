"""Command-line entry point for excel-semantic-md."""

from __future__ import annotations

import argparse
from collections.abc import Sequence


def _print_not_implemented(command: str) -> int:
    print(
        f"excel-semantic-md {command}: not implemented in the phase1 skeleton yet."
    )
    return 1


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
    setup_parser.set_defaults(handler=lambda _args: _print_not_implemented("setup"))

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
        type=int,
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
        handler=lambda _args: _print_not_implemented("convert")
    )

    inspect_parser = subparsers.add_parser(
        "inspect",
        help="Inspect workbook structure as JSON.",
    )
    _add_input_option(inspect_parser)
    inspect_parser.set_defaults(
        handler=lambda _args: _print_not_implemented("inspect")
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
    render_parser.set_defaults(handler=lambda _args: _print_not_implemented("render"))

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    return int(args.handler(args))


if __name__ == "__main__":
    raise SystemExit(main())
