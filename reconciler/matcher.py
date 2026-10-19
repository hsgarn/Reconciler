from datetime import timedelta
from decimal import Decimal

from reconciler.models import BankTransaction, MatchResult, SpreadsheetRow


def match_all(
    bank_txs: list[BankTransaction],
    sheet_rows: list[SpreadsheetRow],
    date_window: int = 3,
) -> list[MatchResult]:
    consumed: set[int] = set()  # row_index values already matched

    results: list[MatchResult] = []

    for bank_tx in bank_txs:
        in_range: list[SpreadsheetRow] = []
        out_of_range: list[SpreadsheetRow] = []

        for row in sheet_rows:
            if row.reconciled_date:
                continue  # already reconciled in spreadsheet
            if row.row_index in consumed:
                continue  # already matched earlier in this run
            if row.amount != bank_tx.amount:
                continue

            day_diff = abs((bank_tx.date - row.date).days)
            if day_diff <= date_window:
                in_range.append(row)
            else:
                out_of_range.append(row)

        if len(in_range) == 1:
            result = MatchResult(
                bank_tx=bank_tx,
                matched_row=in_range[0],
                status="auto",
            )
            consumed.add(in_range[0].row_index)
        elif len(in_range) > 1:
            result = MatchResult(
                bank_tx=bank_tx,
                candidates=in_range,
                status="multi",
            )
        elif out_of_range:
            result = MatchResult(
                bank_tx=bank_tx,
                out_of_range_candidates=out_of_range,
                status="out_of_range",
            )
        else:
            result = MatchResult(
                bank_tx=bank_tx,
                status="missing",
            )

        results.append(result)

    # Post-process: remove consumed rows from OOR candidate lists.
    # A row consumed by an auto-match earlier (or later) in the loop must not
    # appear as a candidate for a different bank transaction.
    for result in results:
        if result.status == "out_of_range":
            result.out_of_range_candidates = [
                c for c in result.out_of_range_candidates
                if c.row_index not in consumed
            ]
            if not result.out_of_range_candidates:
                result.status = "missing"

    return results


def find_extras(
    bank_txs: list[BankTransaction],
    sheet_rows: list[SpreadsheetRow],
    results: list[MatchResult],
    date_window: int = 3,
) -> list[SpreadsheetRow]:
    """Unreconciled sheet rows inside the CSV date range (± date_window) that
    were neither matched nor offered as a candidate in this run."""
    if not bank_txs:
        return []

    used: set[int] = set()
    for r in results:
        if r.matched_row:
            used.add(r.matched_row.row_index)
        used.update(c.row_index for c in r.candidates)
        used.update(c.row_index for c in r.out_of_range_candidates)

    window = timedelta(days=date_window)
    lo = min(tx.date for tx in bank_txs) - window
    hi = max(tx.date for tx in bank_txs) + window

    return [
        row for row in sheet_rows
        if not row.reconciled_date
        and row.row_index not in used
        and lo <= row.date <= hi
    ]


_CENT = Decimal("0.01")
_CLOSE_AMOUNT = Decimal("1.00")


def classify_extra(
    tx: BankTransaction,
    row: SpreadsheetRow,
    date_window: int = 3,
) -> str | None:
    """Why an extra sheet row might be the entry for an unmatched bank tx.

    Returns "sign" (same amount, opposite sign), "close" (same sign and either
    within $1.00 or the same digits transposed), or None.
    """
    if abs((tx.date - row.date).days) > date_window:
        return None

    bank_amt = tx.amount.quantize(_CENT)
    sheet_amt = row.amount.quantize(_CENT)
    if bank_amt == 0 or sheet_amt == 0:
        return None

    if bank_amt == -sheet_amt:
        return "sign"

    if (bank_amt > 0) != (sheet_amt > 0):
        return None

    if abs(bank_amt - sheet_amt) <= _CLOSE_AMOUNT:
        return "close"

    def digits(d: Decimal) -> str:
        return f"{abs(d):.2f}".replace(".", "")

    a, b = digits(bank_amt), digits(sheet_amt)
    if sorted(a) == sorted(b) or _one_digit_off(a, b):
        return "close"

    return None


def _one_digit_off(a: str, b: str) -> bool:
    """True if the digit strings differ by exactly one substituted, inserted or
    deleted digit (5703/5803, 15703/5703, 15703/1703)."""
    if a == b:
        return False
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) == 1
    if abs(len(a) - len(b)) != 1:
        return False
    longer, shorter = (a, b) if len(a) > len(b) else (b, a)
    return any(longer[:i] + longer[i + 1:] == shorter for i in range(len(longer)))
