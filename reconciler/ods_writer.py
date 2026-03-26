from datetime import datetime
from decimal import Decimal
from pathlib import Path

from odf.table import Table, TableRow, TableCell
from odf.namespaces import TABLENS, OFFICENS
from odf import text as odftext

from reconciler.models import BankTransaction, SpreadsheetRow
from reconciler.settings import load_settings


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


def _set_cell_text(cell, value: str) -> None:
    cell.attributes[(OFFICENS, "value-type")] = "string"
    for p in list(cell.getElementsByType(odftext.P)):
        cell.removeChild(p)
    cell.addElement(odftext.P(text=value))


def _set_cell_number(cell, value: Decimal) -> None:
    cell.attributes[(OFFICENS, "value-type")] = "float"
    cell.attributes[(OFFICENS, "value")] = str(value)
    for p in list(cell.getElementsByType(odftext.P)):
        cell.removeChild(p)
    cell.addElement(odftext.P(text=f"{value:,.2f}"))


def _get_table(doc, settings: dict) -> Table:
    all_tables = doc.spreadsheet.getElementsByType(Table)
    if len(all_tables) == 1:
        return all_tables[0]
    sheet_name = settings["sheet_name"]
    for table in all_tables:
        if table.attributes.get((TABLENS, "name")) == sheet_name:
            return table
    raise ValueError(f"Sheet '{sheet_name}' not found in document")


def _cell_text_content(cell) -> str:
    texts = []
    for p in cell.getElementsByType(odftext.P):
        if p.firstChild:
            texts.append(str(p.firstChild))
    return "".join(texts).strip()


def _is_row_available(tr, s: dict) -> bool:
    cells = _expand_row(tr)
    for idx in (s["col_date"], s["col_check"], s["col_desc"], s["col_amount"]):
        if idx is None:
            continue
        if idx < len(cells) and _cell_text_content(cells[idx]):
            return False
    return True


def _materialize_row(tr, num_cols: int) -> list:
    pairs = []
    for cell in list(tr.childNodes):
        if cell.nodeType != cell.ELEMENT_NODE:
            continue
        try:
            repeat = int(cell.attributes.get((TABLENS, "number-columns-repeated"), 1))
        except (TypeError, ValueError):
            repeat = 1
        pairs.append((cell, repeat))

    for cell, _ in pairs:
        tr.removeChild(cell)

    flat = []
    for orig_cell, repeat in pairs:
        for i in range(repeat):
            if i == 0:
                orig_cell.attributes.pop((TABLENS, "number-columns-repeated"), None)
                tr.addElement(orig_cell)
                flat.append(orig_cell)
            else:
                new_cell = TableCell()
                tr.addElement(new_cell)
                flat.append(new_cell)

    while len(flat) < num_cols:
        new_cell = TableCell()
        tr.addElement(new_cell)
        flat.append(new_cell)

    return flat


def write_reconciled_date(doc, row: SpreadsheetRow, recon_date: str,
                           settings: dict | None = None) -> None:
    if settings is None:
        settings = load_settings()
    table = _get_table(doc, settings)
    all_rows = table.getElementsByType(TableRow)
    tr = all_rows[row.row_index]
    cells = _expand_row(tr)
    if settings["col_reconciled"] < len(cells):
        _set_cell_text(cells[settings["col_reconciled"]], recon_date)


def insert_transaction(doc, tx: BankTransaction, recon_date: str, receipt_on: bool,
                       settings: dict | None = None) -> int:
    if settings is None:
        settings = load_settings()
    s = settings
    table = _get_table(doc, s)
    all_rows = table.getElementsByType(TableRow)

    target_tr = None
    target_index = None
    for row_index, tr in enumerate(all_rows):
        if row_index < 2:
            continue
        if _is_row_available(tr, s):
            target_tr = tr
            target_index = row_index
            break

    if target_tr is None:
        raise RuntimeError("No available blank row found in sheet to insert transaction.")

    max_col = max(c for c in (s["col_date"], s["col_check"], s["col_desc"],
                              s["col_amount"], s["col_receipt"], s["col_reconciled"])
                  if c is not None) + 1
    cells = _materialize_row(target_tr, max_col)

    _set_cell_text(cells[s["col_date"]],      tx.date.strftime(s["date_format"]))
    if s["col_check"] is not None:
        _set_cell_text(cells[s["col_check"]], tx.check_number)
    _set_cell_text(cells[s["col_desc"]],      tx.description)
    _set_cell_number(cells[s["col_amount"]], tx.amount)
    if s["col_receipt"] is not None:
        _set_cell_text(cells[s["col_receipt"]], s["checkmark_char"] if receipt_on else "")
    _set_cell_text(cells[s["col_reconciled"]], recon_date)

    return target_index


def update_row_fields(doc, row: SpreadsheetRow, date_str: str | None,
                      description: str | None, check_number: str | None,
                      receipt_on: bool, settings: dict | None = None) -> None:
    if settings is None:
        settings = load_settings()
    s = settings
    table = _get_table(doc, s)
    all_rows = table.getElementsByType(TableRow)
    tr = all_rows[row.row_index]
    cells = _expand_row(tr)
    if date_str and s["col_date"] < len(cells):
        _set_cell_text(cells[s["col_date"]], date_str)
    if check_number is not None and s["col_check"] is not None and s["col_check"] < len(cells):
        _set_cell_text(cells[s["col_check"]], check_number)
    if description and s["col_desc"] < len(cells):
        _set_cell_text(cells[s["col_desc"]], description)
    if s["col_receipt"] is not None and s["col_receipt"] < len(cells):
        _set_cell_text(cells[s["col_receipt"]], s["checkmark_char"] if receipt_on else "")


def save(doc, ods_path: Path) -> None:
    import shutil
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = ods_path.with_name(f"{ods_path.stem}_backup_{ts}{ods_path.suffix}")
    shutil.copy2(str(ods_path), str(backup_path))
    doc.save(str(ods_path))
