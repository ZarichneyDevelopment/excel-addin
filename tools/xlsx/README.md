# XLSX Tooling

This folder contains helper scripts for inspecting and validating the budget workbook.

## Scripts
- `unpack_xlsx.py`: Export workbook structure and sparse cell data to TSV/JSON for LLM-friendly review.
- `recalc.py`: Recalculate formulas with LibreOffice and scan for formula errors.
- `inspect_workbook.py`: Quick workbook inspection (sheets, ranges, tables).
- `create_last_update_range.py`: Creates the `LastRolloverUpdate` named range if missing.
- `create_ambiguous_table.py`: Creates the `AmbiguousItems` table if missing.
- `setup_macro.py`: Sets up LibreOffice macro for `recalc.py`.

## Usage

Unpack workbook:
```bash
python3 tools/xlsx/unpack_xlsx.py Budget.backup.xlsx unpacked/budget-backup
```

Recalculate formulas:
```bash
python3 tools/xlsx/recalc.py Budget.backup.xlsx
```

## Requirements
- Python 3
- LibreOffice installed and available via `flatpak` for `recalc.py`
