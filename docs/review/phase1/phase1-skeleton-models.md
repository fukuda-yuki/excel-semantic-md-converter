# phase1-skeleton-models review

## 2026-04-21

### Review scope

- Implement `docs/phase1/task.md` sections `## 1. プロジェクト骨格` and `## 2. 共通モデル`.
- Align Source of Truth wording so `docs/phase1/spec.md` is authoritative.
- Add package metadata, CLI skeleton, thin skill launcher, common models, and focused tests.

### Changed files under review

- `.gitignore`
- `README.md`
- `pyproject.toml`
- `docs/phase1/spec.md`
- `docs/phase1/plan.md`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`
- `src/excel_semantic_md/**`
- `skills/excel-semantic-markdown/**`
- `tests/test_cli.py`
- `tests/test_models.py`

### Subagents used

- `019dacb2-ef46-7942-aa2f-d7ebdd29b649`: Worker A, project skeleton / CLI / skill / CLI tests. Launch status: success.
- `019dacb3-1e45-7692-9261-0b46cfa0c806`: Worker B, common models / model tests. Launch status: success.
- `019dacb9-0b21-7ce2-9d4a-2abcd7cb826f`: Spec compliance and functional correctness reviewer. Launch status: success.
- `019dacb9-321a-7920-a00c-9bb2758c73da`: Tests, edge cases, and regression risk reviewer. Launch status: success.

### Raw findings summary

1. `make_asset_path()` does not accept or encode asset kind explicitly, even though `docs/phase1/spec.md` requires asset naming to include sheet index, block id, asset kind, and sequence number.
2. Stub CLI commands print "not implemented" but return exit code 0, which makes unsupported behavior look successful to callers.
3. `make_block_id()` accepts arbitrary strings and can generate IDs with non-Phase-1 block kinds.
4. `TableBlock.header_rows` and `header_cols` allow non-integer values such as `bool`.

### MainAgent validation

1. Valid. The final plan used a shorter asset path, but `docs/phase1/spec.md` is authoritative and explicitly requires asset kind in the naming rule.
2. Valid. The plan only required help/argument surfaces for skeleton commands; returning success for unimplemented conversion can mislead automation.
3. Valid. `BlockKind` is closed by the Phase 1 spec and ID generation should reject invalid kinds.
4. Valid. Header counts are numeric counters and should use strict integer validation.

### Response plan

- Update asset path generation to take `asset_kind` and include it in the filename when the block ID does not already end with the same kind.
- Change unimplemented CLI command stubs to return non-zero and update tests accordingly.
- Validate block kind through `BlockKind`.
- Add non-negative integer validation for `TableBlock.header_rows` and `header_cols`.
- Re-run editable install and pytest after fixes.

### Fixes applied

- Updated `make_asset_path()` to accept `asset_kind` and include asset kind in generated paths when it is not already present at the end of the block ID.
- Changed skeleton command handlers to return a non-zero status when the command body is not implemented.
- Changed `make_block_id()` to validate `kind` through `BlockKind`.
- Added strict non-negative integer validation for `TableBlock.header_rows` and `header_cols`.
- Added regression tests for invalid block kinds, asset kind paths, and table header count validation.
- Added minimal skill frontmatter with `allowed-tools: Shell`.

### Remaining risks

- Excel COM, Copilot CLI, workbook extraction, rendering, and LLM integration are still out of scope for this skeleton milestone and require later implementation plus live confirmation where specified.
- `excel-semantic-md.exe` was installed successfully, but this machine's user script directory is not on `PATH`; verification used the installed script by full path and module execution.

### Deferred items or open questions

- Excel COM, Copilot CLI, workbook extraction, rendering, and LLM integration remain out of scope for this skeleton milestone.

### Post-fix validation

- `python -m pip install -e .`: passed.
- `python -m pip install -e ".[test]"`: passed and installed pytest for this environment.
- `python -m pytest`: passed, 12 tests.
- `python -m excel_semantic_md.cli.main --help`: passed.
- Installed `excel-semantic-md.exe --help` by full path: passed.
- Installed `excel-semantic-md.exe convert --input sample.xlsx --out out` by full path: returned non-zero with an explicit skeleton not implemented message.
