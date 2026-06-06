# AGENTS.md

## Project overview

**HMMR** is a single-file Python batch tool (`mrp_v10.py`) for automotive Material Requirements Planning (MRP). It reads Excel inputs from the script directory, computes 3-month material demand and delivery schedules, and writes `MRP_System_v9.xlsx` plus optional per-supplier exports under `Графики_поставщиков/`.

There is no web server, database, or Docker stack. Development means running the Python script with the correct `.xlsx` inputs in `/workspace`.

## Cursor Cloud specific instructions

### Dependencies

Install Python packages from `requirements.txt` (handled by the VM update script on startup).

### Run the MRP calculation

```bash
cd /workspace
python3 mrp_v10.py
```

Place input Excel files in the same directory as `mrp_v10.py`. The script auto-discovers the newest file matching each pattern.

**Required inputs (script crashes without these):**

| Pattern | Purpose |
|---------|---------|
| `Master_BOM*.xlsx` | Sheet `BOM_Детальный` — bill of materials |
| `Упаковка локала*.xlsx` | Sheet `Haval_Stock` — packaging quantities |

**Expected for realistic output (optional but recommended):**

| Pattern | Purpose |
|---------|---------|
| `*Plan-Fact*.xlsx` | Production plan (sheets: `WS_in (5)`, `WS_in_H`, `PS_in`, etc.) |
| Stock `*.xlsx` files | Auto-scanned by filename keywords (`warehouse`, `aat`, `lear`, etc.) |

**Outputs:** `MRP_System_v9.xlsx` (16+ sheets) and `Графики_поставщиков/*.xlsx`.

### Lint and tests

No linter or test suite is configured in this repository.

- **Syntax check:** `python3 -m py_compile mrp_v10.py`
- **Smoke test:** run `python3 mrp_v10.py` with production Excel files present; expect exit code 0 and `✅ Готово! Файл сохранён: .../MRP_System_v9.xlsx`

### Gotchas

- Input and output `.xlsx` files are **not** committed to the repo; they must be supplied locally.
- The script embeds extensive reference data (624 part mappings). A minimal BOM fixture alone may not drive demand; production Plan-Fact and stock files are needed for meaningful calculations.
- `MRP_System_v9.xlsx` in the workspace enables **LIVE mode**: BOM edits on sheet `BOM_Детальный` are preserved on the next run.
- Generated artifacts (`MRP_System_v9.xlsx`, `Графики_поставщиков/`) should not be committed.
