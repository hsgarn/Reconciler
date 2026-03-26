from datetime import timedelta

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
