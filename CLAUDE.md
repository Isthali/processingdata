# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Pipeline for generating PDF reports from mechanical testing lab data (concrete cores, panels, beams, manhole covers). The flow is always the same: read raw curves from Excel/CSV → compute characteristic points (max load, toughness, interpolated values at standard-defined deflection/CMOD points) → write results into a pre-existing Excel template → convert that template to PDF → merge with matplotlib plot pages → overlay an accredited/non-accredited header/footer template.

## Common Commands

Single CLI entry point (`unified_report.py`) with one subcommand per test type:

```bash
python unified_report.py cores           --infle 336-24 --subinfle S --standard CORES      --empresa PRODIMIN --n 6
python unified_report.py panels          --infle 111-25 --subinfle C --standard EFNARC1996 --empresa PRODIMIN --n 3
python unified_report.py panels_residual --infle 111-25 --subinfle C --standard EN14488    --empresa PRODIMIN --n 3
python unified_report.py beams_residual  --infle 222-25 --subinfle B --standard ASTMC1609  --empresa PRODIMIN --n 3
python unified_report.py tapas           --infle 222-25 --subinfle A --standard NTP339.111 --empresa PRODIMIN --n 3
python unified_report.py generic         --infle 336-24 --subinfle S --standard DM         --empresa EXC
```

Use `--ids 3 4 7` instead of `--n` to pick non-consecutive sample IDs. `--offset N` shifts the starting ID when using `--n`. `-v`/`--verbose` switches logging to DEBUG. See `UNIFIED_REPORT_GUIDE.md` for the full argument reference per subcommand.

Install deps: `pip install -r requirements.txt`. No test suite is configured — the `test_` prefix in `test_data.py` / `test_ledi.py` is domain naming ("test" = mechanical ensayo), not pytest.

## Platform Constraint

`edit_pdfs.convert_excel_to_pdf` drives Excel via COM (`win32com.client`), so **any subcommand that produces a PDF requires Windows with Excel installed**. On Linux, import paths and non-PDF code (data processing, plotting) work, but report generation fails at the Excel→PDF step. The default `--base-dir` is `D:/`, reflecting this Windows-only runtime.

## Architecture

Four modules, layered bottom-up:

- **`test_data.py`** — low-level I/O. Reads Excel/text into DataFrames (with chardet-based encoding detection and retry), and writes individual values into specific Excel cells by `(row, column)` using openpyxl. All report results are written cell-by-cell into a pre-configured template workbook.

- **`edit_pdfs.py`** — PDF operations: Excel→PDF via COM, `merge_pdfs`, `normalize_pdf_orientation`, and `apply_header_footer_pdf` (overlays a template PDF from `formatos/` onto every page of the report).

- **`test_ledi.py`** — the domain core. Contains two parallel class hierarchies that must be read together:
  - **Test classes** (one instance per specimen): `Mechanical_test` → `Resistance_mechanical_test` / `Toughness_mechanical_test` → concrete tests (`Axial_compression_test`, `Panels_toughness_test`, `Panel_Beam_residual_strength_test`, `Beam_residual_strength_test`, `Tapa_buzon_flexion_test`). They load a single data file, compute max/min load, first peak, cumulative toughness, and interpolated characteristic points.
  - **Report classes** (one per run): `Test_report` → `Panel_toughness_test_report`, `Panel_Beam_residual_strength_test_report`, `Beam_residual_strength_test_report`, `Axial_compression_test_report`, `Tapa_buzon_flexion_test_report`, `Generate_test_report`. Each orchestrates: `add_tests()` (build test instances from sample IDs, loading files with a naming convention specific to that test type), `preprocess_data()` per test, `write_report()` (push results into hardcoded cells of a specific sheet), plot generation via `make_plot_report`, and the final convert/merge/overlay pipeline in `make_report_file()`.

- **`unified_report.py` + `report_helpers.py`** — CLI layer. `REPORT_CONFIGS` in `unified_report.py` maps each subcommand name to its report class, default standard/client, and sample-count defaults. Adding a new test type = add a new report class in `test_ledi.py`, then add one entry to `REPORT_CONFIGS`.

## Required Inputs on Disk

The report classes **do not generate the Excel template or the raw data files** — they expect them to already exist under `{base-dir}/{infle}/`:

- A pre-configured template workbook named `INFLE_{infle}[-{subinfle}]_{standard}_{client}[_{n}].xlsm` (or `.xlsx` for `generic`) with the sheets and cells the report writes into (`Resultados`, `Cores`, etc. — see each report's `write_report`).
- Per-sample raw data files with type-specific naming:
  - `cores`: `{infle}-d{id}.xlsx` with columns Time/Displacement/Load
  - `panels`: `Losa P{id}.xlsx` with Time/Load/Deflection/Displacement
  - `panels_residual` / `beams_residual`: expect CMOD or Deflection columns depending on standard
  - `tapas`: `tapas_{id}.xlsx` with Time/Load

If data files or template cells are missing, failures surface at `add_tests()` or `write_report()` — not at CLI parse time.

## x_col Convention (Residual Strength)

Two superficially similar report classes use different x-axis columns for interpolation:

- `Panel_Beam_residual_strength_test_report` → `x_col='CMOD'` (EN 14651 / EN 14488)
- `Beam_residual_strength_test_report` → `x_col='Deflection'` (ASTM C1609)

This is set internally in each report's `add_tests()` via `Toughness_mechanical_test.preprocess_data(x_col=...)`. CLI users never pass it. If you add a new residual-strength variant, pick the correct `x_col` — interpolated values go into the Excel template at fixed cells that assume a specific x-axis.

## Formatos (Header/Footer Templates)

`formatos/formato_acreditado.pdf` and `formato_no_acreditado.pdf` are overlaid on every page as the last step. Each report class hardcodes which one to use (accredited tests → `formato_acreditado.pdf`). The path is relative (`./formatos/...`), so reports must be run from the repo root.
