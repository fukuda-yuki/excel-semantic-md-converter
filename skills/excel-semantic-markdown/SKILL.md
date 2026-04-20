---
name: excel-semantic-markdown
description: Launch the local excel-semantic-md CLI for Excel workbook conversion.
allowed-tools: Shell
---

# Excel Semantic Markdown

Use this skill only as a thin launcher for the local `excel-semantic-md` CLI.

## Inputs

- `input`: path to a `.xlsx` or `.xlsm` workbook.
- `out`: output directory for the CLI run.

## Workflow

1. Confirm that the input workbook path exists.
2. Confirm that the output directory exists or can be created.
3. Run `run_excel_semantic_md.ps1` with the input and output paths.
4. Report where the CLI wrote its outputs.

This skill must not contain prompt text, LLM response contracts, workbook conversion logic, or Excel block-detection logic. Those responsibilities belong to the Python package.
