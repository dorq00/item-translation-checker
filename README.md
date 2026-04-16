# Translation Checker

Checks French customs translations on import invoices against a historical translation database.

## How to use

1. Drop your invoice `.xlsx` into `invoices/incoming/`
2. Double-click `START_CHECKER.bat`
3. Output opens automatically in `output/YYYY-MM/`

## Requirements

```
pip install pandas openpyxl
```

## Output

Single workbook, 3 sheets:

| Sheet | Contents |
|---|---|
| Status Check | All rows with classification |
| Changed | Items where translation differs from DB |
| New Items | Items not in DB — need translation |

## Status codes

| Status | Meaning |
|---|---|
| ✅ NO CHANGE | Item in DB, translation matches |
| 📋 IN DB | Item in DB, invoice has no translation (auto-filled) |
| 🛑 CHANGED | Item in DB, translation differs |
| 🆕 NEW ITEM | Item not in DB |

## DB

All `.xlsx` files in `db/` are loaded as translation history. To add new translations, drop a file with columns `class`, `Item`, `Designation`, `translation` into `db/`.
