from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Literal


@dataclass
class SpreadsheetRow:
    row_index: int          # 0-based index into odfpy table rows
    date: date
    check_number: str
    description: str
    receipt: str            # empty string or checkmark character
    amount: Decimal
    balance: Decimal
    reconciled_date: str    # empty string or date string
    reconciled_balance: Decimal | None
    not_reconciled: Decimal | None


@dataclass
class BankTransaction:
    date: date
    description: str
    check_number: str
    amount: Decimal         # positive = debit, negative = credit (matches spreadsheet sign)


@dataclass
class MatchResult:
    bank_tx: BankTransaction
    matched_row: SpreadsheetRow | None = None
    candidates: list[SpreadsheetRow] = field(default_factory=list)
    out_of_range_candidates: list[SpreadsheetRow] = field(default_factory=list)
    status: Literal["auto", "multi", "out_of_range", "missing"] = "missing"
    # "auto"          = exactly 1 in-range match, auto-reconciled
    # "multi"         = 2+ in-range matches, user must choose
    # "out_of_range"  = 0 in-range matches but some beyond date window, user must choose
    # "missing"       = no matches at all (or user rejected all candidates)
