import csv
from decimal import Decimal, InvalidOperation
from pathlib import Path

from dateutil import parser as dateparser

from reconciler.models import BankTransaction
from reconciler.settings import load_settings


def _parse_amount(value: str) -> Decimal:
    cleaned = value.strip().replace(",", "")
    if not cleaned:
        return Decimal("0")
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        raise ValueError(f"Cannot parse amount: {value!r}")


def load_csv(csv_path: Path, settings: dict | None = None) -> list[BankTransaction]:
    if settings is None:
        settings = load_settings()
    s = settings
    transactions: list[BankTransaction] = []

    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            date_str = row.get(s["csv_date"], "").strip()
            if not date_str:
                continue

            tx_date = dateparser.parse(date_str).date()

            credit_str = row.get(s["csv_credit"], "").strip()
            debit_str  = row.get(s["csv_debit"],  "").strip()

            if debit_str:
                amount = -_parse_amount(debit_str)
            elif credit_str:
                amount = -_parse_amount(credit_str)
            else:
                amount = Decimal("0")

            transactions.append(BankTransaction(
                date=tx_date,
                description=row.get(s["csv_description"], "").strip(),
                check_number=row.get(s["csv_check"], "").strip(),
                amount=amount,
            ))

    return transactions
