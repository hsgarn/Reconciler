"""User-configurable settings with defaults matching config.py."""
import json
from pathlib import Path

_SETTINGS_FILE = Path(__file__).parent.parent / "settings.json"

DEFAULTS = {
    # Spreadsheet
    "sheet_name":           "FHB Checking",
    "col_date":             0,
    "col_check":            1,
    "col_desc":             2,
    "col_receipt":          3,
    "col_amount":           4,
    "col_balance":          5,
    "col_reconciled":       7,
    "col_reconciled_bal":   8,
    "col_not_reconciled":   9,
    # Matching
    "date_window_days":     3,
    "oor_view_window_days": 60,
    # Formatting
    "date_format":          "%m/%d/%y",
    "checkmark_char":       "\u221a",  # √
    # CSV headers
    "csv_date":             "Date",
    "csv_description":      "Description",
    "csv_check":            "Check #",
    "csv_debit":            "Debit",
    "csv_credit":           "Credit",
}


def load_settings() -> dict:
    try:
        data = json.loads(_SETTINGS_FILE.read_text(encoding="utf-8"))
        merged = dict(DEFAULTS)
        merged.update(data)
        # Ensure integer column indices; col_check and col_receipt may be None (disabled)
        optional_cols = {"col_check", "col_receipt"}
        for key in ("col_date", "col_check", "col_desc", "col_receipt",
                    "col_amount", "col_balance", "col_reconciled",
                    "col_reconciled_bal", "col_not_reconciled",
                    "date_window_days", "oor_view_window_days"):
            val = merged.get(key)
            if val is None and key in optional_cols:
                merged[key] = None
            else:
                try:
                    merged[key] = int(val)
                except (TypeError, ValueError):
                    merged[key] = DEFAULTS[key]
        return merged
    except FileNotFoundError:
        return dict(DEFAULTS)
    except Exception:
        return dict(DEFAULTS)


def save_settings(settings: dict) -> None:
    _SETTINGS_FILE.write_text(
        json.dumps(settings, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def col_letter(idx: int) -> str:
    """Convert 0-based column index to spreadsheet letter (0→A, 1→B, …)."""
    return chr(65 + idx)


def col_index(letter: str) -> int:
    """Convert column letter to 0-based index ('A'→0, 'B'→1, …)."""
    return ord(letter.strip().upper()) - 65
