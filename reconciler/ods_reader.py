from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

from dateutil import parser as dateparser
from odf import opendocument
from odf.table import Table, TableRow, TableCell
from odf.namespaces import TABLENS, OFFICENS

from reconciler.models import SpreadsheetRow
from reconciler.settings import load_settings


def _cell_text(cell) -> str:
    texts = []
    for p in cell.getElementsByType(__import__("odf.text", fromlist=["P"]).P):
        if p.firstChild:
            texts.append(str(p.firstChild))
    return "".join(texts).strip()


def _cell_date_value(cell) -> str | None:
    return cell.attributes.get((OFFICENS, "date-value"))


def _expand_row(tr) -> list:
    cells = []
    for cell in tr.childNodes:
        if cell.nodeType != cell.ELEMENT_NODE:
            continue
        try:
            repeat = int(cell.attributes.get((TABLENS, "number-columns-repeated"), 1))
        except (TypeError, ValueError):
            repeat = 1
        for _ in range(repeat):
            cells.append(cell)
    return cells


def _cell_numeric_value(cell) -> str | None:
    return cell.attributes.get((OFFICENS, "value"))


def _parse_decimal(s: str) -> Decimal | None:
    s = s.strip().replace(",", "")
    if not s:
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def _cell_decimal(cell) -> Decimal | None:
    v = _cell_numeric_value(cell)
    if v is not None:
        result = _parse_decimal(v)
        if result is not None:
            return result
    return _parse_decimal(_cell_text(cell))


def _parse_date(cell) -> date | None:
    dv = _cell_date_value(cell)
    if dv:
        try:
            return dateparser.parse(dv).date()
        except Exception:
            pass
    text = _cell_text(cell)
    if text:
        try:
            return dateparser.parse(text).date()
        except Exception:
            pass
    return None


def list_sheet_names(ods_path: Path) -> list[str]:
    """Return all sheet names in the workbook."""
    doc = opendocument.load(str(ods_path))
    return [
        t.attributes.get((TABLENS, "name"))
        for t in doc.spreadsheet.getElementsByType(Table)
    ]


def load_sheet(ods_path: Path, settings: dict | None = None):
    """Load the target sheet and return (doc, rows).

    If the workbook has only one sheet it is used automatically regardless of
    the sheet_name setting.  Returns the raw odfpy document so the writer can
    mutate it without reloading from disk.
    """
    if settings is None:
        settings = load_settings()

    s = settings  # shorthand

    doc = opendocument.load(str(ods_path))
    all_tables = doc.spreadsheet.getElementsByType(Table)

    # Auto-detect: if there is exactly one sheet, use it
    if len(all_tables) == 1:
        target_table = all_tables[0]
    else:
        sheet_name = s["sheet_name"]
        target_table = None
        for table in all_tables:
            if table.attributes.get((TABLENS, "name")) == sheet_name:
                target_table = table
                break
        if target_table is None:
            names = [t.attributes.get((TABLENS, "name")) for t in all_tables]
            raise ValueError(
                f"Sheet '{sheet_name}' not found. Available sheets: {names}"
            )

    rows: list[SpreadsheetRow] = []

    for row_index, tr in enumerate(target_table.getElementsByType(TableRow)):
        cells = _expand_row(tr)

        if len(cells) <= s["col_amount"]:
            continue

        amount = _cell_decimal(cells[s["col_amount"]])
        if amount is None:
            continue

        tx_date = _parse_date(cells[s["col_date"]])
        if tx_date is None:
            continue

        def safe_text(idx):
            return _cell_text(cells[idx]) if idx < len(cells) else ""

        def safe_decimal(idx):
            return _cell_decimal(cells[idx]) if idx < len(cells) else None

        rows.append(SpreadsheetRow(
            row_index=row_index,
            date=tx_date,
            check_number=safe_text(s["col_check"]) if s["col_check"] is not None else "",
            description=safe_text(s["col_desc"]),
            receipt=safe_text(s["col_receipt"]) if s["col_receipt"] is not None else "",
            amount=amount,
            balance=safe_decimal(s["col_balance"]) or Decimal("0"),
            reconciled_date=safe_text(s["col_reconciled"]),
            reconciled_balance=safe_decimal(s["col_reconciled_bal"]),
            not_reconciled=safe_decimal(s["col_not_reconciled"]),
        ))

    return doc, rows
