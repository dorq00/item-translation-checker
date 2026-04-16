"""
translation_checker — run.py
=============================
Drop your invoice in  invoices/incoming/
Double-click          START_CHECKER.bat
Done — output lands in output/YYYY-MM/

FOLDER STRUCTURE:
  translation_checker/
  ├── run.py                   ← this file
  ├── START_CHECKER.bat        ← double-click launcher
  ├── db/                      ← past MEINV invoices go here (history)
  │   ├── MEINV00084440.xlsx
  │   ├── MEINV00085066.xlsx
  │   └── ...
  ├── invoices/
  │   └── incoming/            ← drop the NEW invoice here before running
  └── output/
      └── 2026-04/             ← auto-created, organized by month
          └── check_<invoice>_<datetime>.xlsx   ← single workbook, 3 sheets:
              Sheet 1 — Status Check  (all rows)
              Sheet 2 — Changed       (translation conflicts)
              Sheet 3 — New Items     (not in DB, need translation)

HOW THE DB WORKS:
  All .xlsx files in db/ are loaded and combined as the translation history.
  After reviewing a new invoice, move it from invoices/incoming/ into db/
  so future runs treat it as history.
"""

import os
import sys
from datetime import datetime
from pathlib import Path

# ── Dependencies ──────────────────────────────────────────────────────────────

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas not installed.  Run:  pip install pandas openpyxl")
    try:
        input("\nPress Enter to close...")
    except (EOFError, KeyboardInterrupt):
        pass
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl not installed.  Run:  pip install openpyxl")
    try:
        input("\nPress Enter to close...")
    except (EOFError, KeyboardInterrupt):
        pass
    sys.exit(1)

# ── Paths ─────────────────────────────────────────────────────────────────────

ROOT       = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
DB_DIR     = ROOT / "db"
INCOMING   = ROOT / "invoices" / "incoming"
OUTPUT_DIR = ROOT / "output"

# ── Status labels ─────────────────────────────────────────────────────────────

NO_CHANGE = "✅ NO CHANGE"
CHANGED   = "🛑 CHANGED"
NEW_ITEM  = "🆕 NEW ITEM"
IN_DB     = "📋 IN DB"

# ── Colors ────────────────────────────────────────────────────────────────────

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

GREEN_FILL  = _fill("C6EFCE")
RED_FILL    = _fill("FFC7CE")
ORANGE_FILL = _fill("FFEB9C")
GREY_FILL   = _fill("F2F2F2")

GREEN_FONT  = Font(color="006100", bold=True)
RED_FONT    = Font(color="9C0006", bold=True)
ORANGE_FONT = Font(color="9C5700", bold=True)
WHITE_FONT  = Font(color="FFFFFF", bold=True)

THIN = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

STATUS_STYLE = {
    NO_CHANGE: (GREEN_FILL,  GREEN_FONT),
    IN_DB:     (GREEN_FILL,  GREEN_FONT),
    CHANGED:   (RED_FILL,    RED_FONT),
    NEW_ITEM:  (ORANGE_FILL, ORANGE_FONT),
}

# ── Loaders ───────────────────────────────────────────────────────────────────

def _clean_str(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none") else s


# LG Korea division codes → Algeria customs class codes
_DIV_MAP = {
    "W/M": "WM", "RAC": "AC", "CAC": "AC",
    "REF": "RF", "LTV": "TV", "MNT": "TV",
}

def _parse_invoice_file(path: Path) -> pd.DataFrame:
    """
    Parse a single invoice file into canonical columns:
      class, Item, translation, Designation
    translation is "" when the file has no translation column (PO format).

    Handles three formats:
      - Raw MEINV from LG Korea portal (has DATA DETAILS sheet):
          Product → class | Item → Item | Designation → translation | Item Desc → Designation
      - PO Master (has HS Code sheet):
          Div Name → class | Part Number → Item | Product Description → Designation
          (no translation — filled as "")
      - Flat/transformed format (first sheet):
          class | Item | translation | Designation
    """
    xl = pd.ExcelFile(path)

    if "DATA DETAILS" in xl.sheet_names:
        df = xl.parse("DATA DETAILS", dtype=str).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        df = df.rename(columns={
            "Product":     "class",
            "Item":        "Item",
            "Designation": "translation",
            "Item Desc":   "Designation",
        })

    elif "HS Code" in xl.sheet_names:
        df = xl.parse("HS Code", dtype=str).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        df = df.rename(columns={
            "Div Name":           "class",
            "Part Number":        "Item",
            "Product Description": "Designation",
        })
        df["class"] = df["class"].map(lambda v: _DIV_MAP.get(v, v))
        df["translation"] = ""

    else:
        df = xl.parse(xl.sheet_names[0], dtype=str).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        col_map = {}
        for c in df.columns:
            lc = c.lower()
            if lc == "class":            col_map[c] = "class"
            elif lc == "item":           col_map[c] = "Item"
            elif lc == "translation":    col_map[c] = "translation"
            elif lc == "designation":    col_map[c] = "Designation"
        df = df.rename(columns=col_map)
        if "translation" not in df.columns:
            df["translation"] = ""

    required = {"class", "Item", "Designation"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"{path.name}: missing columns {missing}\n"
            f"  Found: {list(df.columns)}"
        )

    df = df[["class", "Item", "translation", "Designation"]].copy()
    for col in df.columns:
        df[col] = df[col].apply(_clean_str)
    return df[df["Item"].str.strip() != ""].reset_index(drop=True)


def load_invoice(path: Path) -> pd.DataFrame:
    """Load and parse the single incoming invoice."""
    df = _parse_invoice_file(path)
    print(f"  Invoice  : {len(df)} rows  ({path.name})")
    return df


def load_db() -> pd.DataFrame:
    """
    Load translation history from all .xlsx files in db/.
    Files are sorted by name (MEINV numbers are chronological),
    so drop_duplicates keep='last' keeps the most recent translation per Item.
    """
    db_files = sorted(
        [f for f in DB_DIR.glob("*.xlsx") if not f.name.startswith("~$")],
        key=lambda f: f.name,
    )
    if not db_files:
        return pd.DataFrame(columns=["class", "Item", "translation", "Designation"])

    frames = []
    for f in db_files:
        try:
            frames.append(_parse_invoice_file(f))
        except Exception as e:
            print(f"  WARNING: skipping {f.name} — {e}")

    if not frames:
        return pd.DataFrame(columns=["class", "Item", "translation", "Designation"])

    df = pd.concat(frames, ignore_index=True)
    df = df[df["Item"].str.strip() != ""]
    df = df.drop_duplicates(subset="Item", keep="last").reset_index(drop=True)
    print(f"  DB       : {len(df)} unique items  ({len(db_files)} invoice files)")
    return df

# ── Check logic ───────────────────────────────────────────────────────────────

def classify(df_invoice: pd.DataFrame, df_db: pd.DataFrame) -> pd.DataFrame:
    """
    LEFT JOIN invoice vs DB on Item.
    Adds db_translation column and Status.
    """
    db_lookup = df_db.set_index("Item")[["translation"]].rename(
        columns={"translation": "db_translation"}
    )
    df = df_invoice.merge(db_lookup, on="Item", how="left")
    df["db_translation"] = df["db_translation"].fillna("")

    def _status(row):
        if row["db_translation"] == "":
            return NEW_ITEM
        if row["translation"] == "":          # PO format — no translation submitted
            return IN_DB
        if row["translation"].strip().upper() == row["db_translation"].strip().upper():
            return NO_CHANGE
        return CHANGED

    df["Status"] = df.apply(_status, axis=1)
    return df

# ── Excel output helpers ──────────────────────────────────────────────────────

def _write_header(ws, columns, fill):
    for ci, col in enumerate(columns, 1):
        c = ws.cell(1, ci, col)
        c.fill = fill
        c.font = WHITE_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN
    ws.row_dimensions[1].height = 20


def _write_rows(ws, df):
    for ri, (_, row) in enumerate(df.iterrows(), 2):
        for ci, val in enumerate(row, 1):
            c = ws.cell(ri, ci, "" if pd.isna(val) else val)
            c.alignment = Alignment(vertical="center")
            c.border = THIN
            if ri % 2 == 0:
                c.fill = GREY_FILL


def _color_status_col(ws, df, col_name):
    if col_name not in df.columns:
        return
    ci = list(df.columns).index(col_name) + 1
    for ri, val in enumerate(df[col_name], 2):
        style = STATUS_STYLE.get(str(val))
        if style:
            ws.cell(ri, ci).fill = style[0]
            ws.cell(ri, ci).font = style[1]


def _autofit(ws):
    for col in ws.columns:
        w = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w + 2, 10), 48)


def _setup(ws):
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

# ── Output writer ─────────────────────────────────────────────────────────────

def write_output(path: Path, df: pd.DataFrame):
    """Single workbook with 3 sheets: Status Check, Changed, New Items."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Sheet 1: Status Check ─────────────────────────────────────
    out = df[["class", "Item", "translation", "db_translation", "Designation", "Status"]].copy()
    out = out.rename(columns={"translation": "Invoice Translation", "db_translation": "DB Translation"})
    ws = wb.create_sheet("Status Check")
    ws.sheet_properties.tabColor = "2E4057"
    _setup(ws)
    _write_header(ws, list(out.columns), _fill("2E4057"))
    _write_rows(ws, out)
    _color_status_col(ws, out, "Status")
    _autofit(ws)

    # ── Sheet 2: Changed ──────────────────────────────────────────
    rows = df[df["Status"] == CHANGED].copy()
    out2 = pd.DataFrame({
        "class":           rows["class"].values,
        "Item":            rows["Item"].values,
        "Designation":     rows["Designation"].values,
        "New Translation": rows["translation"].values,
        "Old Translation": rows["db_translation"].values,
        "Status":          rows["Status"].values,
    })
    ws2 = wb.create_sheet("Changed")
    ws2.sheet_properties.tabColor = "C55A11"
    _setup(ws2)
    _write_header(ws2, list(out2.columns), _fill("C55A11"))
    _write_rows(ws2, out2)
    _color_status_col(ws2, out2, "Status")
    new_ci = list(out2.columns).index("New Translation") + 1
    old_ci = list(out2.columns).index("Old Translation") + 1
    for ri in range(2, len(out2) + 2):
        ws2.cell(ri, new_ci).fill = _fill("FCE4D6")
        ws2.cell(ri, old_ci).fill = _fill("DDEBF7")
    _autofit(ws2)

    # ── Sheet 3: New Items ────────────────────────────────────────
    rows3 = df[df["Status"] == NEW_ITEM].copy()
    out3 = pd.DataFrame({
        "class":       rows3["class"].values,
        "Item":        rows3["Item"].values,
        "Designation": rows3["Designation"].values,
        "Status":      rows3["Status"].values,
    })
    ws3 = wb.create_sheet("New Items")
    ws3.sheet_properties.tabColor = "375623"
    _setup(ws3)
    _write_header(ws3, list(out3.columns), _fill("375623"))
    _write_rows(ws3, out3)
    _color_status_col(ws3, out3, "Status")
    _autofit(ws3)

    wb.save(path)

# ── Main ──────────────────────────────────────────────────────────────────────

def abort(msg: str):
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Translation Checker — Error", msg)
        root.destroy()
    except Exception:
        print(f"\nERROR: {msg}")
    try:
        input("\nPress Enter to close...")
    except (EOFError, KeyboardInterrupt):
        pass
    sys.exit(1)


def main():
    print("=" * 60)
    print("  LG ALGERIA — TRANSLATION CHECKER")
    print("=" * 60)

    # ── Ensure folders exist ──────────────────────────────────────
    INCOMING.mkdir(parents=True, exist_ok=True)
    DB_DIR.mkdir(parents=True, exist_ok=True)

    # ── Find invoice in incoming/ ─────────────────────────────────
    invoices = sorted(
        [f for f in INCOMING.iterdir() if f.suffix.lower() == ".xlsx" and not f.name.startswith("~$")],
        key=lambda f: f.stat().st_mtime,
        reverse=True,
    )
    if not invoices:
        abort(
            f"No invoice found in:\n  {INCOMING}\n\n"
            f"Drop your .xlsx invoice there and try again."
        )

    invoice_path = invoices[0]
    if len(invoices) > 1:
        print(f"\n  NOTE: {len(invoices)} invoices in incoming/ — using most recent.")
        print(f"  Others: {[f.name for f in invoices[1:]]}")

    print(f"\n  Invoice  : {invoice_path.name}")
    print()

    # ── Load ──────────────────────────────────────────────────────
    print("  Loading data...")
    try:
        df_invoice = load_invoice(invoice_path)
        df_db      = load_db()
    except Exception as e:
        abort(str(e))

    if len(df_db) == 0:
        print("  WARNING: DB is empty — all items will be classified as NEW ITEM.")

    # ── Classify ──────────────────────────────────────────────────
    print("\n  Classifying rows...")
    df_classified = classify(df_invoice, df_db)

    n_no_change = int((df_classified["Status"] == NO_CHANGE).sum())
    n_in_db     = int((df_classified["Status"] == IN_DB).sum())
    n_changed   = int((df_classified["Status"] == CHANGED).sum())
    n_new       = int((df_classified["Status"] == NEW_ITEM).sum())
    total       = len(df_classified)

    print(f"\n  {'─' * 44}")
    print(f"  Total    : {total} lines")
    if n_no_change: print(f"  {NO_CHANGE}  : {n_no_change}")
    if n_in_db:     print(f"  {IN_DB}       : {n_in_db}  (translation auto-filled from DB)")
    if n_changed:   print(f"  {CHANGED}    : {n_changed}")
    print(f"  {NEW_ITEM}   : {n_new}  (need translation)")
    print(f"  {'─' * 44}")

    # ── Write output ──────────────────────────────────────────────
    month_dir = OUTPUT_DIR / datetime.now().strftime("%Y-%m")
    month_dir.mkdir(parents=True, exist_ok=True)

    out_name = f"check_{invoice_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    out_path = month_dir / out_name

    print(f"\n  Writing → output/{datetime.now().strftime('%Y-%m')}/{out_name}")
    write_output(out_path, df_classified)
    print(f"    Sheet 1 — Status Check  ({total} rows)")
    print(f"    Sheet 2 — Changed       ({n_changed} rows)")
    print(f"    Sheet 3 — New Items     ({n_new} rows)")

    print(f"\n  [DONE]  {out_path}")
    print(f"\n  TIP: Once reviewed, move {invoice_path.name} → db/ to add it to history.")

    try:
        os.startfile(out_path)
    except Exception:
        pass

    try:
        input("\nPress Enter to close...")
    except (EOFError, KeyboardInterrupt):
        pass


if __name__ == "__main__":
    main()
