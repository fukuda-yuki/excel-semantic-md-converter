from __future__ import annotations

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import excel_semantic_md.cli.main as cli_main


main = cli_main.main


class CliTests(unittest.TestCase):
    def _run_with_output(self, argv: list[str]) -> tuple[int | str | None, str, str]:
        stdout = io.StringIO()
        stderr = io.StringIO()
        with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
            try:
                code = main(argv)
            except SystemExit as exc:
                code = exc.code
        return code, stdout.getvalue(), stderr.getvalue()

    def test_top_level_help_lists_expected_commands(self) -> None:
        code, stdout, _stderr = self._run_with_output(["--help"])

        self.assertEqual(code, 0)
        self.assertIn("setup", stdout)
        self.assertIn("convert", stdout)
        self.assertIn("inspect", stdout)
        self.assertIn("render", stdout)
        self.assertNotIn("resume", stdout)

    def test_convert_help_lists_phase1_options(self) -> None:
        code, stdout, _stderr = self._run_with_output(["convert", "--help"])

        self.assertEqual(code, 0)
        self.assertIn("--input", stdout)
        self.assertIn("--out", stdout)
        self.assertIn("--model", stdout)
        self.assertIn("--vision-model", stdout)
        self.assertIn("--max-images-per-sheet", stdout)
        self.assertIn("--save-debug-json", stdout)
        self.assertIn("--save-render-artifacts", stdout)
        self.assertIn("--strict", stdout)
        self.assertNotIn("--resume", stdout)

    def test_setup_help_lists_out_option(self) -> None:
        code, stdout, _stderr = self._run_with_output(["setup", "--help"])

        self.assertEqual(code, 0)
        self.assertIn("--out", stdout)

    def test_setup_reports_diagnostics_without_external_side_effects(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            out_dir = Path(temp_dir) / "setup-out"
            with (
                mock.patch.object(cli_main, "_find_console_entry_point", return_value=None),
                mock.patch.object(cli_main.platform, "system", return_value="Windows"),
                mock.patch.object(cli_main, "_has_importable_module", return_value=False),
                mock.patch.object(cli_main.shutil, "which", return_value=None),
                mock.patch.object(cli_main, "_is_gh_copilot_available", return_value=False),
            ):
                code, stdout, stderr = self._run_with_output(["setup", "--out", str(out_dir)])

            self.assertEqual(code, 0)
            self.assertEqual(stderr, "")
            self.assertIn("setup diagnostics", stdout)
            self.assertIn("Python package: ok", stdout)
            self.assertIn("CLI entry point: not found", stdout)
            self.assertIn("Excel COM: not available", stdout)
            self.assertIn("Copilot CLI: not found", stdout)
            self.assertIn("Output directory: ok: writable", stdout)
            self.assertIn("does not install external tools", stdout)
            self.assertFalse(out_dir.exists())
            self.assertEqual(list(Path(temp_dir).rglob(".excel-semantic-md-setup-*.tmp")), [])

    def test_setup_out_preserves_existing_output_directory_and_removes_probe_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            out_dir = Path(temp_dir) / "existing-out"
            out_dir.mkdir()
            with (
                mock.patch.object(cli_main, "_find_console_entry_point", return_value=None),
                mock.patch.object(cli_main.platform, "system", return_value="Windows"),
                mock.patch.object(cli_main, "_has_importable_module", return_value=False),
                mock.patch.object(cli_main.shutil, "which", return_value=None),
                mock.patch.object(cli_main, "_is_gh_copilot_available", return_value=False),
            ):
                code, stdout, stderr = self._run_with_output(["setup", "--out", str(out_dir)])

            self.assertEqual(code, 0)
            self.assertEqual(stderr, "")
            self.assertIn("Output directory: ok: writable", stdout)
            self.assertTrue(out_dir.is_dir())
            self.assertEqual(list(out_dir.glob(".excel-semantic-md-setup-*.tmp")), [])

    def test_setup_out_does_not_loop_when_no_parent_exists(self) -> None:
        with mock.patch.object(cli_main.Path, "exists", return_value=False):
            ok, message = cli_main._check_output_directory("missing-root")

        self.assertFalse(ok)
        self.assertIn("no existing parent directory", message)

    def test_resume_command_is_not_available(self) -> None:
        code, _stdout, stderr = self._run_with_output(["resume"])

        self.assertNotEqual(code, 0)
        self.assertIn("invalid choice", stderr)

    def test_convert_accepts_options_and_runs_pipeline(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "sample.xlsx"
            out_path = temp_path / "out"
            input_path.write_bytes(b"synthetic workbook placeholder")
            fake_result = mock.Mock(has_failures=False)
            fake_output_files = mock.Mock(
                result_markdown=out_path / "result.md",
                manifest_json=out_path / "manifest.json",
                debug_dir=out_path / "debug",
            )
            with (
                mock.patch("excel_semantic_md.app.run_convert_pipeline", return_value=fake_result) as run_pipeline,
                mock.patch("excel_semantic_md.output.write_convert_outputs", return_value=fake_output_files) as write_outputs,
                mock.patch("excel_semantic_md.app.cleanup_convert_result") as cleanup_result,
            ):
                code, stdout, _stderr = self._run_with_output(
                    [
                        "convert",
                        "--input",
                        str(input_path),
                        "--out",
                        str(out_path),
                        "--model",
                        "text-model",
                        "--vision-model",
                        "vision-model",
                        "--max-images-per-sheet",
                        "3",
                        "--save-debug-json",
                        "--save-render-artifacts",
                        "--strict",
                    ]
                )

            self.assertEqual(code, 0)
            self.assertIn("result.md:", stdout)
            self.assertIn("manifest.json:", stdout)
            self.assertIn("debug/:", stdout)
            self.assertTrue(out_path.is_dir())
            run_pipeline.assert_called_once()
            write_outputs.assert_called_once_with(fake_result)
            cleanup_result.assert_called_once_with(fake_result)

    def test_convert_rejects_missing_input_with_clear_error(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            missing_path = Path(temp_dir) / "missing.xlsx"
            out_path = Path(temp_dir) / "out"

            code, _stdout, stderr = self._run_with_output(
                ["convert", "--input", str(missing_path), "--out", str(out_path)]
            )

        self.assertEqual(code, 2)
        self.assertIn("input workbook does not exist", stderr)

    def test_convert_rejects_unsupported_extension(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "sample.csv"
            out_path = Path(temp_dir) / "out"
            input_path.write_text("a,b\n1,2\n", encoding="utf-8")

            code, _stdout, stderr = self._run_with_output(
                ["convert", "--input", str(input_path), "--out", str(out_path)]
            )

        self.assertEqual(code, 2)
        self.assertIn("input workbook must be", stderr)

    def test_convert_rejects_negative_max_images_per_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "sample.xlsx"
            out_path = Path(temp_dir) / "out"
            input_path.write_bytes(b"synthetic workbook placeholder")

            code, _stdout, stderr = self._run_with_output(
                [
                    "convert",
                    "--input",
                    str(input_path),
                    "--out",
                    str(out_path),
                    "--max-images-per-sheet",
                    "-1",
                ]
            )

        self.assertEqual(code, 2)
        self.assertIn("must be a non-negative integer", stderr)

    def test_inspect_rejects_invalid_workbook_content(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "sample.xlsm"
            input_path.write_bytes(b"synthetic workbook placeholder")

            code, stdout, stderr = self._run_with_output(["inspect", "--input", str(input_path)])

        self.assertEqual(code, 2)
        self.assertEqual(stdout, "")
        self.assertIn("failed to read input workbook", stderr)

    def test_render_rejects_invalid_workbook_content(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "sample.xlsx"
            input_path.write_bytes(b"synthetic workbook placeholder")

            code, stdout, stderr = self._run_with_output(
                ["render", "--input", str(input_path), "--sheet", "Sheet1"]
            )

        self.assertEqual(code, 2)
        self.assertEqual(stdout, "")
        self.assertIn("failed to read input workbook", stderr)

    def test_render_rejects_empty_sheet_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "sample.xlsx"
            input_path.write_bytes(b"synthetic workbook placeholder")

            code, _stdout, stderr = self._run_with_output(
                ["render", "--input", str(input_path), "--sheet", "   "]
            )

        self.assertEqual(code, 2)
        self.assertIn("sheet name must not be empty", stderr)


if __name__ == "__main__":
    unittest.main()
