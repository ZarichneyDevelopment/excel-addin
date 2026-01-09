# XLSX Unpacking

This repo includes a script to unpack a workbook into an LLM-friendly, text-based format.

## Usage

```bash
python3 tools/xlsx/unpack_xlsx.py Budget.backup.xlsx unpacked/budget-backup
```

## Output

- `summary.json`: high-level workbook summary, sheets, named ranges, and tables
- `named_ranges.json`: defined names only
- `tables.json`: table definitions only
- `sheets/<sheet_id>_<name>/meta.json`: per-sheet metadata
- `sheets/<sheet_id>_<name>/cells.tsv`: sparse cell list with values and formulas
- `sheets/<sheet_id>_<name>/merges.tsv`: merged-cell ranges

Notes:
- `cells.tsv` includes raw values and formulas from the XLSX. Cached values are included when present.
- Large sheets are represented sparsely (only non-empty cells).
