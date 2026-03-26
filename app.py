import json
import os
import tempfile
from datetime import date
from decimal import Decimal
from pathlib import Path

from flask import Flask, jsonify, render_template, request, session, redirect, url_for

from reconciler.csv_reader import load_csv
from reconciler.matcher import match_all
from reconciler.models import BankTransaction, MatchResult, SpreadsheetRow
from reconciler.nicknames import (
    load_nicknames,
    save_nicknames,
    match_description,
    next_id,
)
from reconciler.settings import (
    load_settings,
    save_settings as save_app_settings,
    DEFAULTS as SETTINGS_DEFAULTS,
    col_letter,
    col_index,
)
from reconciler.ods_reader import list_sheet_names
from reconciler.ods_reader import load_sheet
from reconciler.ods_writer import (
    insert_transaction,
    save,
    update_row_fields,
    write_reconciled_date,
)

_PREFS_FILE = Path(__file__).parent / "prefs.json"

def _load_prefs() -> dict:
    try:
        return json.loads(_PREFS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}

def _save_prefs(prefs: dict) -> None:
    try:
        _PREFS_FILE.write_text(json.dumps(prefs, indent=2), encoding="utf-8")
    except Exception:
        pass

def _get_secret_key():
    key = _load_prefs().get("secret_key")
    if not key:
        import secrets
        key = secrets.token_hex(32)
        prefs = _load_prefs()
        prefs["secret_key"] = key
        _save_prefs(prefs)
    return key

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET") or _get_secret_key()


# ---------------------------------------------------------------------------
# JSON serialization helpers
# ---------------------------------------------------------------------------

class ReconcilerEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, date):
            return {"__type__": "date", "value": obj.isoformat()}
        if isinstance(obj, Decimal):
            return {"__type__": "decimal", "value": str(obj)}
        return super().default(obj)


def reconciler_decoder(obj):
    if "__type__" in obj:
        if obj["__type__"] == "date":
            return date.fromisoformat(obj["value"])
        if obj["__type__"] == "decimal":
            return Decimal(obj["value"])
    return obj


def serialize_results(results):
    data = []
    for r in results:
        data.append({
            "bank_tx": _tx_to_dict(r.bank_tx),
            "matched_row": _row_to_dict(r.matched_row) if r.matched_row else None,
            "candidates": [_row_to_dict(c) for c in r.candidates],
            "out_of_range_candidates": [_row_to_dict(c) for c in r.out_of_range_candidates],
            "status": r.status,
        })
    return json.dumps(data, cls=ReconcilerEncoder)


def _tx_to_dict(tx):
    return {
        "date": tx.date,
        "description": tx.description,
        "check_number": tx.check_number,
        "amount": tx.amount,
    }


def _row_to_dict(row):
    return {
        "row_index": row.row_index,
        "date": row.date,
        "check_number": row.check_number,
        "description": row.description,
        "receipt": row.receipt,
        "amount": row.amount,
        "balance": row.balance,
        "reconciled_date": row.reconciled_date,
        "reconciled_balance": row.reconciled_balance,
        "not_reconciled": row.not_reconciled,
    }


def _dict_to_tx(d):
    return BankTransaction(
        date=d["date"], description=d["description"],
        check_number=d["check_number"], amount=d["amount"],
    )


def _dict_to_row(d):
    return SpreadsheetRow(
        row_index=d["row_index"], date=d["date"], check_number=d["check_number"],
        description=d["description"], receipt=d.get("receipt", ""),
        amount=d["amount"], balance=d.get("balance", Decimal("0")),
        reconciled_date=d.get("reconciled_date", ""),
        reconciled_balance=d.get("reconciled_balance"),
        not_reconciled=d.get("not_reconciled"),
    )


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html",
                           last_ods_path=session.get("ods_path", ""),
                           last_csv_path=session.get("csv_path", ""),
                           last_recon_date=session.get("recon_date", ""),
                           last_ending_bal=session.get("ending_bal", ""))


@app.route("/browse")
def browse():
    import subprocess
    kind = request.args.get("type", "ods")
    if kind == "csv":
        title = "Select Bank Transaction CSV"
        filt = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    else:
        title = "Select ODS Spreadsheet"
        filt = "ODS files (*.ods)|*.ods|All files (*.*)|*.*"
    prefs = _load_prefs()
    initial_dir = prefs.get(f"last_dir_{kind}", "")
    init_dir_ps = f"$d.InitialDirectory = '{initial_dir}';" if initial_dir else ""

    ps_script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "Add-Type -AssemblyName System.Drawing;"
        "$owner = New-Object System.Windows.Forms.Form;"
        "$owner.TopMost = $true;"
        "$owner.ShowInTaskbar = $false;"
        "$owner.Opacity = 0;"
        "$owner.StartPosition = 'CenterScreen';"
        "$owner.Size = New-Object System.Drawing.Size(1,1);"
        "$owner.Show();"
        "$d = New-Object System.Windows.Forms.OpenFileDialog;"
        f"$d.Title = '{title}';"
        f"$d.Filter = '{filt}';"
        f"{init_dir_ps}"
        "$d.Multiselect = $false;"
        "$r = $d.ShowDialog($owner);"
        "$owner.Dispose();"
        "if ($r -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $d.FileName } else { Write-Output '' }"
    )
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps_script],
        capture_output=True, text=True,
    )
    lines = [l.strip() for l in result.stdout.splitlines() if l.strip()]
    path = lines[-1] if lines else ""
    # Discard anything that doesn't look like a file path
    if path and not (path[1:3] == ":\\" or path.startswith("\\\\")):
        path = ""
    if path:
        prefs[f"last_dir_{kind}"] = str(Path(path).parent)
        _save_prefs(prefs)
    return jsonify(path=path)


@app.route("/process", methods=["POST"])
def process():
    from dateutil import parser as dateparser

    force = request.form.get("force") == "1"
    recon_date    = request.form.get("recon_date",  "").strip()
    ending_bal_str = request.form.get("ending_bal", "").strip()

    if force and session.get("ods_path") and session.get("csv_path"):
        # Reuse already-stored paths (user clicked "Continue Anyway")
        ods_path = Path(session["ods_path"])
        csv_path_str = session["csv_path"]
    else:
        # Fresh submission
        ods_path_str = request.form.get("ods_path", "").strip()
        csv_path_str = request.form.get("csv_path", "").strip()

        if not ods_path_str:
            return render_template("index.html", error="Please select the ODS spreadsheet.")
        if not csv_path_str:
            return render_template("index.html", error="Please select the bank transaction CSV.")

        ods_path = Path(ods_path_str)
        if not ods_path.exists():
            return render_template("index.html", error=f"ODS file not found: {ods_path_str}")
        if not Path(csv_path_str).exists():
            return render_template("index.html", error=f"CSV file not found: {csv_path_str}")

        session["ods_path"] = str(ods_path)
        session["csv_path"] = csv_path_str
        session["save_mode"] = request.form.get("save_mode", "overwrite")
        session["ods_original_name"] = ods_path.name

    session["recon_date"]  = recon_date
    session["ending_bal"]  = ending_bal_str

    settings = load_settings()

    try:
        _doc, sheet_rows = load_sheet(ods_path, settings)
        bank_txs = load_csv(Path(csv_path_str), settings)
    except Exception as e:
        return render_template("index.html", error=f"Error processing files: {e}")

    # ── Date conflict check ────────────────────────────────────────────────
    if not force and recon_date:
        try:
            recon_date_parsed = dateparser.parse(recon_date).date()
        except Exception:
            recon_date_parsed = None

        if recon_date_parsed:
            conflict_count = 0
            for row in sheet_rows:
                if not row.reconciled_date:
                    continue
                try:
                    if dateparser.parse(row.reconciled_date).date() == recon_date_parsed:
                        conflict_count += 1
                except Exception:
                    if row.reconciled_date.strip() == recon_date:
                        conflict_count += 1
            if conflict_count:
                return render_template("warn.html",
                                       conflict_count=conflict_count,
                                       csv_count=len(bank_txs),
                                       recon_date=recon_date,
                                       ending_bal=ending_bal_str)

    try:
        unreconciled = [r for r in sheet_rows if not r.reconciled_date]
        results = match_all(bank_txs, unreconciled, date_window=settings["date_window_days"])
    except Exception as e:
        return render_template("index.html", error=f"Error matching transactions: {e}")

    try:
        ending_bal = Decimal(ending_bal_str.replace(",", "")) if ending_bal_str else None
    except Exception:
        ending_bal = None

    last_recon_bal = next(
        (r.reconciled_balance for r in reversed(sheet_rows) if r.reconciled_balance is not None),
        None,
    )
    total_sum = sum(tx.amount for tx in bank_txs)
    computed_bal = (last_recon_bal - total_sum) if last_recon_bal is not None else None

    def fmt_bal(d):
        return f"{d:,.2f}" if d is not None else None

    balance_info = {
        "last_recon_bal": fmt_bal(last_recon_bal),
        "computed_bal": fmt_bal(computed_bal),
        "ending_bal": fmt_bal(ending_bal),
        "match": (ending_bal is not None and computed_bal is not None
                  and ending_bal == computed_bal),
    }

    # Build oor_all_matches now so it can be cached
    oor_all_matches, nickname_matches = _build_review_extras(
        results, sheet_rows, recon_date, settings
    )

    # Cache everything needed for /review (avoids re-loading ODS/CSV)
    cache = {
        "results": json.loads(serialize_results(results)),
        "balance_info": balance_info,
        "oor_all_matches": {str(k): v for k, v in oor_all_matches.items()},
        "recon_date": recon_date,
        "ending_bal_str": ending_bal_str,
    }
    tmp_json = tempfile.NamedTemporaryFile(
        delete=False, suffix=".json", prefix="reconciler_", mode="w", encoding="utf-8"
    )
    json.dump(cache, tmp_json)
    tmp_json.close()
    session["temp_json_key"] = tmp_json.name
    session["has_review"] = True

    return _render_review(results, oor_all_matches, nickname_matches,
                          balance_info, recon_date, ending_bal_str)


def _build_review_extras(results, sheet_rows, recon_date, settings=None):
    from datetime import timedelta
    from dateutil import parser as dateparser
    if settings is None:
        settings = load_settings()
    oor_view_days = settings.get("oor_view_window_days", 60)
    date_fmt = settings.get("date_format", "%m/%d/%y")
    oor_all_matches = {}
    try:
        recon_date_parsed = dateparser.parse(recon_date).date() if recon_date else None
    except Exception:
        recon_date_parsed = None
    for gi, r in enumerate(results):
        if r.status != "out_of_range":
            continue
        matches = []
        for row in sheet_rows:
            if row.amount != r.bank_tx.amount:
                continue
            if recon_date_parsed and abs((row.date - recon_date_parsed).days) > oor_view_days:
                continue
            matches.append({
                "date": row.date.strftime(date_fmt),
                "description": row.description,
                "check_number": row.check_number,
                "amount": f"{row.amount:,.2f}",
                "reconciled_date": row.reconciled_date or "",
                "row_index": row.row_index,
            })
        if matches:
            oor_all_matches[gi] = matches

    nicknames = load_nicknames()
    nickname_matches = {}
    for r in results:
        desc = r.bank_tx.description
        if desc and desc not in nickname_matches:
            nickname_matches[desc] = match_description(desc, nicknames)

    return oor_all_matches, nickname_matches


def _render_review(results, oor_all_matches, nickname_matches, balance_info,
                   recon_date, ending_bal_str):
    s = load_settings()
    return render_template(
        "review.html",
        results=results,
        results_json=serialize_results(results),
        balance_info=balance_info,
        recon_date=recon_date,
        ending_bal=ending_bal_str,
        oor_all_matches=oor_all_matches,
        nickname_matches=nickname_matches,
        has_check=s["col_check"] is not None,
        has_receipt=s["col_receipt"] is not None,
    )


@app.route("/review")
def review():
    """Re-render the review page from cache — no ODS/CSV reload needed."""
    temp_json_key = session.get("temp_json_key")
    if not temp_json_key or not os.path.exists(temp_json_key):
        return redirect(url_for("index"))

    try:
        with open(temp_json_key, encoding="utf-8") as f:
            cache = json.load(f, object_hook=reconciler_decoder)
    except Exception:
        return redirect(url_for("index"))

    results_raw    = cache.get("results", [])
    balance_info   = cache.get("balance_info", {})
    recon_date     = cache.get("recon_date", "")
    ending_bal_str = cache.get("ending_bal_str", "")
    oor_all_matches = {int(k): v for k, v in cache.get("oor_all_matches", {}).items()}

    results = [MatchResult(
        bank_tx=_dict_to_tx(r["bank_tx"]),
        matched_row=_dict_to_row(r["matched_row"]) if r.get("matched_row") else None,
        candidates=[_dict_to_row(c) for c in r.get("candidates", [])],
        out_of_range_candidates=[_dict_to_row(c) for c in r.get("out_of_range_candidates", [])],
        status=r["status"],
    ) for r in results_raw]

    # Reload nicknames fresh (user may have just added some)
    nicknames = load_nicknames()
    nickname_matches = {}
    for r in results:
        desc = r.bank_tx.description
        if desc and desc not in nickname_matches:
            nickname_matches[desc] = match_description(desc, nicknames)

    return _render_review(results, oor_all_matches, nickname_matches,
                          balance_info, recon_date, ending_bal_str)


@app.route("/confirm", methods=["POST"])
def confirm():
    from dateutil import parser as dateparser

    temp_json_key = session.get("temp_json_key")
    ods_path_str = session.get("ods_path")
    csv_path = session.get("csv_path")
    recon_date = session.get("recon_date", "")
    ending_bal_str = session.get("ending_bal", "")
    save_mode = session.get("save_mode", "overwrite")

    if not temp_json_key or not ods_path_str:
        return redirect(url_for("index"))

    ods_path = Path(ods_path_str)
    settings = load_settings()
    doc, sheet_rows = load_sheet(ods_path, settings)
    bank_txs = load_csv(Path(csv_path), settings)
    unreconciled = [r for r in sheet_rows if not r.reconciled_date]
    results = match_all(bank_txs, unreconciled, date_window=settings["date_window_days"])

    # Capture last reconciled balance NOW, before we write anything
    last_recon_bal = next(
        (r.reconciled_balance for r in reversed(sheet_rows) if r.reconciled_balance is not None),
        None,
    )

    form = request.form
    stats = {"auto": 0, "manual": 0, "inserted": 0, "skipped": 0}
    inserted_txs = []
    applied_sum = Decimal("0")   # running total of amounts actually written/inserted
    row_lookup = {r.row_index: r for r in sheet_rows}
    used_row_indices: set[int] = set()

    for i, result in enumerate(results):
        status = result.status

        # Collect any user edits for this result
        edit_date  = form.get(f"edit_{i}_date",  "").strip()
        edit_desc  = form.get(f"edit_{i}_desc",  "").strip()
        edit_check = form.get(f"edit_{i}_check", "").strip()
        edit_receipt = form.get(f"edit_{i}_receipt") == "on"

        if status == "auto":
            row = result.matched_row
            if row.row_index not in used_row_indices:
                write_reconciled_date(doc, row, recon_date, settings)
                update_row_fields(doc, row, edit_date or None, edit_desc or None,
                                  edit_check or None, edit_receipt, settings)
                used_row_indices.add(row.row_index)
                applied_sum += row.amount
                stats["auto"] += 1
            else:
                stats["skipped"] += 1

        elif status in ("multi", "out_of_range"):
            choice = form.get(f"choice_{i}", "skip")
            if choice == "skip":
                # Check if user chose to add this to the spreadsheet from the skip panel
                action = form.get(f"missing_{i}", "skip")
                if action == "add":
                    receipt_on = form.get(f"receipt_{i}") == "on"
                    tx = result.bank_tx
                    sk_date  = form.get(f"missing_edit_{i}_date",  "").strip()
                    sk_desc  = form.get(f"missing_edit_{i}_desc",  "").strip()
                    sk_check = form.get(f"missing_edit_{i}_check", "").strip()
                    insert_date = tx.date
                    if sk_date:
                        try:
                            insert_date = dateparser.parse(sk_date).date()
                        except Exception:
                            pass
                    modified_tx = BankTransaction(
                        date=insert_date,
                        description=sk_desc or tx.description,
                        check_number=sk_check or tx.check_number,
                        amount=tx.amount,
                    )
                    try:
                        insert_transaction(doc, modified_tx, recon_date, receipt_on, settings)
                        stats["inserted"] += 1
                        applied_sum += modified_tx.amount
                        inserted_txs.append({
                            "date": modified_tx.date.strftime(settings["date_format"]),
                            "description": modified_tx.description,
                            "amount": str(modified_tx.amount),
                            "receipt": receipt_on,
                        })
                    except RuntimeError:
                        stats["skipped"] += 1
                else:
                    stats["skipped"] += 1
            else:
                try:
                    chosen_row_index = int(choice)
                    if chosen_row_index in used_row_indices:
                        # Duplicate selection — treat as skipped
                        stats["skipped"] += 1
                    else:
                        chosen_row = row_lookup.get(chosen_row_index)
                        if chosen_row:
                            write_reconciled_date(doc, chosen_row, recon_date, settings)
                            update_row_fields(doc, chosen_row, edit_date or None,
                                              edit_desc or None, edit_check or None,
                                              edit_receipt, settings)
                            used_row_indices.add(chosen_row_index)
                            applied_sum += chosen_row.amount
                            stats["manual"] += 1
                        else:
                            stats["skipped"] += 1
                except (ValueError, TypeError):
                    stats["skipped"] += 1

        elif status == "missing":
            action = form.get(f"missing_{i}", "skip")
            if action == "add":
                receipt_on = form.get(f"receipt_{i}") == "on"
                tx = result.bank_tx
                # Apply any user edits to the inserted transaction
                insert_date = tx.date
                if edit_date:
                    try:
                        insert_date = dateparser.parse(edit_date).date()
                    except Exception:
                        pass
                modified_tx = BankTransaction(
                    date=insert_date,
                    description=edit_desc or tx.description,
                    check_number=edit_check or tx.check_number,
                    amount=tx.amount,
                )
                try:
                    insert_transaction(doc, modified_tx, recon_date, receipt_on, settings)
                    stats["inserted"] += 1
                    applied_sum += modified_tx.amount
                    inserted_txs.append({
                        "date": modified_tx.date.strftime(settings["date_format"]),
                        "description": modified_tx.description,
                        "amount": str(modified_tx.amount),
                        "receipt": receipt_on,
                    })
                except RuntimeError:
                    stats["skipped"] += 1
            else:
                stats["skipped"] += 1

    save_mode = session.get("save_mode", "overwrite")
    try:
        save(doc, ods_path)
        save_error = None
    except Exception as e:
        save_error = str(e)

    try:
        ending_bal = Decimal(ending_bal_str.replace(",", "")) if ending_bal_str else None
    except Exception:
        ending_bal = None

    # Compute expected ending balance from what was actually applied
    # (same sign logic: debits are positive → reduce balance, credits negative → increase balance)
    computed_ending = (last_recon_bal - applied_sum) if last_recon_bal is not None else None

    balance_match = (
        ending_bal is not None
        and computed_ending is not None
        and ending_bal == computed_ending
    )

    if temp_json_key and os.path.exists(temp_json_key):
        try:
            os.unlink(temp_json_key)
        except Exception:
            pass
    session.pop("temp_json_key", None)

    return render_template(
        "complete.html",
        stats=stats,
        inserted_txs=inserted_txs,
        ending_bal=f"{ending_bal:,.2f}" if ending_bal else ending_bal_str,
        last_recon_bal=f"{last_recon_bal:,.2f}" if last_recon_bal is not None else None,
        computed_ending=f"{computed_ending:,.2f}" if computed_ending is not None else None,
        balance_match=balance_match,
        save_error=save_error,
        recon_date=recon_date,
        save_mode=save_mode,
        saved_to_path=ods_path_str if save_mode == "overwrite" and not save_error else None,
    )


@app.route("/shutdown", methods=["POST"])
def shutdown():
    os._exit(0)


@app.route("/download")
def download():
    from flask import send_file
    ods_path = session.get("ods_path")
    original_name = session.get("ods_original_name", "reconcile.ods")
    if not ods_path or not os.path.exists(ods_path):
        return redirect(url_for("index"))
    return send_file(ods_path, as_attachment=True, download_name=original_name)


# ---------------------------------------------------------------------------
# Nickname management
# ---------------------------------------------------------------------------

@app.route("/nicknames")
def nicknames_page():
    return_to = url_for("review") if session.get("has_review") else url_for("index")
    return render_template("nicknames.html", nicknames=load_nicknames(), return_to=return_to)


@app.route("/nicknames/add", methods=["POST"])
def nicknames_add():
    pattern  = request.form.get("pattern", "").strip()
    nickname = request.form.get("nickname", "").strip()
    if pattern and nickname:
        nicks = load_nicknames()
        nicks.append({"id": next_id(nicks), "pattern": pattern, "nickname": nickname})
        save_nicknames(nicks)
    return redirect(url_for("nicknames_page"))


@app.route("/nicknames/update", methods=["POST"])
def nicknames_update():
    entry_id = request.form.get("id", type=int)
    pattern  = request.form.get("pattern", "").strip()
    nickname = request.form.get("nickname", "").strip()
    if pattern and nickname:
        nicks = load_nicknames()
        for entry in nicks:
            if entry.get("id") == entry_id:
                entry["pattern"] = pattern
                entry["nickname"] = nickname
                break
        save_nicknames(nicks)
    return redirect(url_for("nicknames_page"))


@app.route("/nicknames/delete", methods=["POST"])
def nicknames_delete():
    entry_id = request.form.get("id", type=int)
    nicks = [e for e in load_nicknames() if e.get("id") != entry_id]
    save_nicknames(nicks)
    return redirect(url_for("nicknames_page"))


@app.route("/api/nicknames/match")
def nicknames_match_api():
    desc = request.args.get("desc", "")
    matches = match_description(desc, load_nicknames())
    return jsonify(matches=matches)


@app.route("/nicknames/quick-add", methods=["POST"])
def nicknames_quick_add():
    """Called via fetch from the review page to add a nickname without leaving."""
    pattern  = request.form.get("pattern", "").strip()
    nickname = request.form.get("nickname", "").strip()
    if not pattern or not nickname:
        return jsonify(ok=False, error="Pattern and nickname are required."), 400
    import re as _re
    try:
        _re.compile(pattern)
    except _re.error as e:
        return jsonify(ok=False, error=f"Invalid regex: {e}"), 400
    nicks = load_nicknames()
    new_entry = {"id": next_id(nicks), "pattern": pattern, "nickname": nickname}
    nicks.append(new_entry)
    save_nicknames(nicks)
    return jsonify(ok=True, entry=new_entry)


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

@app.route("/settings", methods=["GET", "POST"])
def settings_page():
    ods_path_str = session.get("ods_path", "")
    sheet_names = []
    if ods_path_str and os.path.exists(ods_path_str):
        try:
            sheet_names = list_sheet_names(Path(ods_path_str))
        except Exception:
            sheet_names = []

    if request.args.get("reset") == "1":
        save_app_settings(dict(SETTINGS_DEFAULTS))
        return redirect(url_for("settings_page"))

    if request.method == "POST":
        s = load_settings()
        # General
        s["sheet_name"]           = request.form.get("sheet_name", s["sheet_name"]).strip()
        s["date_window_days"]     = int(request.form.get("date_window_days", s["date_window_days"]))
        s["oor_view_window_days"] = int(request.form.get("oor_view_window_days", s["oor_view_window_days"]))
        s["date_format"]          = request.form.get("date_format", s["date_format"]).strip()
        s["checkmark_char"]       = request.form.get("checkmark_char", s["checkmark_char"]).strip()
        # Spreadsheet columns (entered as letters, stored as 0-based indices)
        optional_cols = {"col_check", "col_receipt"}
        for key in ("col_date", "col_check", "col_desc", "col_receipt",
                    "col_amount", "col_balance", "col_reconciled",
                    "col_reconciled_bal", "col_not_reconciled"):
            letter = request.form.get(key, "").strip().upper()
            if letter and letter.isalpha() and len(letter) == 1:
                s[key] = col_index(letter)
            elif not letter and key in optional_cols:
                s[key] = None  # blank → disabled
        # CSV headers
        for key in ("csv_date", "csv_description", "csv_check", "csv_debit", "csv_credit"):
            val = request.form.get(key, "").strip()
            if val:
                s[key] = val
        save_app_settings(s)
        return redirect(url_for("settings_page"))

    s = load_settings()
    # Convert column indices to letters for display (None → "")
    col_letters = {
        key: (col_letter(s[key]) if s[key] is not None else "")
        for key in ("col_date", "col_check", "col_desc", "col_receipt",
                    "col_amount", "col_balance", "col_reconciled",
                    "col_reconciled_bal", "col_not_reconciled")
    }
    return render_template("settings.html",
                           s=s,
                           col_letters=col_letters,
                           sheet_names=sheet_names,
                           defaults=SETTINGS_DEFAULTS)


if __name__ == "__main__":
    app.run(debug=True)
