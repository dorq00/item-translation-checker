# Item Translation Checker

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-green)

Validates French customs translations on import invoices against a historical translation database. Each item code has an established official translation — this tool flags anything that changed, anything missing, and anything never seen before.

Built to replace a manual line-by-line review process that runs on every incoming shipment. A wrong or inconsistent translation on a customs declaration causes clearance delays.

## What It Does

1. Loads all `.xlsx` files in `db/` as translation history — deduplicates by item code, latest file wins
2. Parses the incoming invoice — auto-detects format (raw portal export, PO Master, or flat)
3. Joins invoice items against the DB on item code
4. Classifies every line and writes a color-coded Excel report

## Status Codes

| Status | Meaning |
|---|---|
| ✅ NO CHANGE | Item in DB — translation matches |
| 📋 IN DB | Item in DB — invoice has no translation (auto-filled from DB) |
| 🛑 CHANGED | Item in DB — translation differs from history |
| 🆕 NEW ITEM | Item not in DB — needs translation |

## Quick Start

```
1. Drop invoice .xlsx into   invoices/incoming/
2. Double-click              START_CHECKER.bat
3. Output opens in           output/YYYY-MM/
```

**Requirements:**
```bash
pip install pandas openpyxl
```

## Output

Single workbook, 3 sheets:

| Sheet | Contents |
|---|---|
| Status Check | All rows — full classification |
| Changed | Conflicts only — new vs old translation side by side |
| New Items | Items not in DB — ready to paste into translation workflow |

## Invoice Formats Supported

| Format | Detected by | Notes |
|---|---|---|
| Raw portal export (MEINV) | `DATA DETAILS` sheet | Has existing translations |
| PO Master | `HS Code` sheet | No translation column — all known items show as `IN DB` |
| Flat / translated | First sheet | Columns: `class`, `Item`, `translation`, `Designation` |

## DB Structure

All `.xlsx` files in `db/` are loaded and merged on each run. To add new translations:

1. Prepare a file with columns: `class` · `Item` · `Designation` · `translation`
2. Drop it in `db/`
3. Next run picks it up automatically

Files are processed in alphabetical order — later files override earlier ones for the same item code.

## Translating New Items

When Sheet 3 has new items needing translation, paste the rows into an LLM with this prompt:

```
You are a French customs translation assistant for electronics spare parts imports.
Translation = official French term for Algerian customs declarations (DUM).

Classes: RF=Refrigerator, WM=Washing Machine, AC=Air Conditioner, TV=Television

Rules:
- Short French noun/phrase (1–4 words), no articles
- Use standard customs vocabulary: MOTEUR, CARTE ELECTRONIQUE, JOINT, RESISTANCE,
  CONDENSATEUR, COMPRESSEUR, CAPTEUR, VENTILATEUR, POMPE, THERMOSTAT, FILTRE...
- Flag obvious class mismatches (e.g. designation doesn't match the class assigned)

Input: class | Item | Designation
Output table: class | Item | Designation | translation | confidence | notes
confidence: HIGH / MEDIUM / UNCERTAIN
```

Save the result as `translations_YYYY-MM.xlsx` and drop it in `db/`.
