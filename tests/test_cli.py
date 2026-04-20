from __future__ import annotations

import contextlib
import io
import unittest

from excel_semantic_md.cli.main import main


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

    def test_resume_command_is_not_available(self) -> None:
        code, _stdout, stderr = self._run_with_output(["resume"])

        self.assertNotEqual(code, 0)
        self.assertIn("invalid choice", stderr)

    def test_convert_accepts_options_and_fails_as_unimplemented(self) -> None:
        code, stdout, _stderr = self._run_with_output(
            [
                "convert",
                "--input",
                "sample.xlsx",
                "--out",
                "out",
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

        self.assertEqual(code, 1)
        self.assertIn("not implemented", stdout)


if __name__ == "__main__":
    unittest.main()
