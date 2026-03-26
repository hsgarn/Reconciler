# Bank Reconciler

A local web application for reconciling bank transactions against an OpenOffice Calc spreadsheet (`.ods`). Runs as a Flask server accessed via Chrome — no internet connection required.

---

## Quick Start

Double-click **`launch.bat`** to start the server and open Chrome automatically. The batch file waits until Flask is ready before opening the browser, so there is no need to manually navigate to the URL.

To stop the server, click the **Shut Down** button in the top-right corner of any page.

---

## Requirements

- Python 3.11+
- Dependencies: `pip install -r requirements.txt`
- Google Chrome installed

---

## Workflow

### 1. Start Reconciliation (`/`)

- **ODS Spreadsheet** — Click **Browse…** to select your `reconcile.ods` file. The dialog opens to the last directory you used.
- **Bank Transaction CSV** — Click **Browse…** to select the CSV exported from your bank. Expected filename pattern: `Transactions-YYYY-MM-DD.csv`. The reconcile date is auto-filled from the filename.
- **Reconcile Date** — Auto-filled from the CSV filename; edit if needed. Format: `MM/DD/YY`.
- **Ending Reconciled Balance** — Enter the closing balance from your bank statement.
- **After Reconciling** — Choose **Overwrite original file** (default) or **Download updated file**.

All four fields are remembered across sessions (stored in `localStorage`) so returning next month pre-fills everything.

If the reconcile date already exists in the spreadsheet, a warning page is shown before proceeding.

---

### 2. Review Matches (`/review`)

Transactions are matched against unreconciled spreadsheet rows by **exact amount** and **date within 3 days**. Results are grouped into four categories:

| Category | Meaning |
|---|---|
| **Auto** | Exactly one match found — reconciled automatically |
| **Multi** | Two or more candidates — you select the correct row |
| **Out-of-Range** | Match found but date is more than 3 days off — you accept or skip |
| **Missing** | No match found — you add to the spreadsheet or skip |

#### Auto Matches
Collapsed by default. Expand to edit the sheet date, description, check number, or receipt flag before saving.

#### Multi / Out-of-Range
A candidate table shows all options. Select the correct row via the radio button. The edit panel below updates with that row's current values for review/editing. Selecting **Skip** moves the transaction to the missing panel.

The **View All** button on out-of-range cards shows every spreadsheet row with the same amount within ±2 months of the reconcile date (including already-reconciled rows) for reference.

#### Missing Transactions
Choose **Add to spreadsheet** or **Skip**. When adding, a description dropdown appears:
- If a nickname pattern matches, the nickname is pre-filled as the top option.
- The original CSV description is always available as an option.
- Click the **tag button** (🏷) to save a new nickname for this description.

Skipped transactions from Multi/Out-of-Range appear here with an **Undo** button to send them back.

#### Duplicate Prevention
The same spreadsheet row cannot be selected for two different bank transactions. Already-used rows are greyed out.

#### Balance Estimate
A running balance estimate is shown at the top of the review page based on the previous reconciled balance minus all CSV transaction amounts.

---

### 3. Confirm & Save (`/confirm`)

Clicking **Confirm & Save** applies all selections:
- Writes the reconcile date to matched rows
- Applies any edits (date, description, check number, receipt)
- Inserts new rows for transactions marked **Add to spreadsheet** (never inserted into rows 1 or 2)
- Saves the file (overwrite with timestamped backup, or offers download)

---

### 4. Reconciliation Complete

Shows a summary including:
- Number of auto-matched, manually matched, inserted, and skipped transactions
- **Previous Reconciled Balance** — last balance from the spreadsheet before this run
- **Computed Ending Balance** — previous balance minus the sum of all applied transactions
- **Entered Ending Balance** — what you typed on the start page
- A checkmark or warning indicating whether the balances match
- List of any newly inserted transactions

---

## Nickname System

The **Nicknames** page (`/nicknames`, linked in the navbar) maps regex patterns to short descriptions. This lets you translate verbose CSV descriptions like `POS DB WM SUPERCE 9040 03/06 105 CHICKASAW R4533 105` into `Walmart` automatically.

### Managing Nicknames

| Action | How |
|---|---|
| Add | Fill in the pattern and nickname fields at the top, click **Add** |
| Edit | Click the pencil icon on any row, edit inline, click the checkmark |
| Delete | Click the trash icon on any row |
| Test | Paste a CSV description in the **Test** field and click **Test** to see which nicknames match |

### Nickname Format

- **Pattern** — A regular expression (case-insensitive). Use `|` to match multiple variations: `WAL.MART|WALMART|WM SUPERCE`
- **Nickname** — The short text to use in the spreadsheet description field

### Adding from the Review Page

On any **Missing** or **Out-of-Range → Skip** transaction, click the 🏷 button next to the description dropdown to open a quick-add dialog. The CSV description is pre-filled as the pattern (escaped). Click **Test** to verify the pattern matches, enter the nickname, and click **Save & Apply** — the nickname is saved and immediately selected in the dropdown.

### Returning to Review

The **Back** button on the Nicknames page returns you to the Review Matches page with all your selections intact.

---

## Spreadsheet Structure

**Sheet name:** `FHB Checking`

| Column | Index | Contents |
|---|---|---|
| A | 0 | Date (`MM/DD/YY`) |
| B | 1 | Check number |
| C | 2 | Description |
| D | 3 | Receipt (checkmark or blank) |
| E | 4 | Amount (debits positive, credits negative) |
| F | 5 | Running balance |
| G | 6 | *(blank)* |
| H | 7 | Reconciled date — filled by this tool |
| I | 8 | Reconciled balance — formula column |
| J | 9 | Not-reconciled column |

**Sign convention:** Debits are stored as **positive** amounts; credits as **negative**. Bank CSV debit values (which the bank exports as negative) are negated during import.

---

## CSV Format

Expected columns (exported from FHB online banking):

| Column | Contents |
|---|---|
| `Date` | Transaction date |
| `Description` | Full bank description |
| `Check #` | Check number if applicable |
| `Debit` | Debit amount (exported as negative) |
| `Credit` | Credit amount (exported as positive) |

Expected filename pattern: `Transactions-YYYY-MM-DD.csv`

---

## Files

| File/Folder | Purpose |
|---|---|
| `app.py` | Flask application — all routes and business logic |
| `launch.bat` | Launcher — starts Flask and opens Chrome |
| `reconciler/` | Python package with domain logic |
| `reconciler/config.py` | Column indices, sheet name, constants |
| `reconciler/models.py` | Data classes (`BankTransaction`, `SpreadsheetRow`, `MatchResult`) |
| `reconciler/matcher.py` | Transaction matching algorithm |
| `reconciler/csv_reader.py` | CSV import and sign conversion |
| `reconciler/ods_reader.py` | ODS spreadsheet reader using odfpy |
| `reconciler/ods_writer.py` | ODS writer — reconcile dates, new rows, backup |
| `reconciler/nicknames.py` | Nickname pattern matching |
| `templates/` | Jinja2 HTML templates |
| `static/` | CSS and static assets |
| `prefs.json` | Persistent preferences (last directories, secret key) — auto-created |
| `nicknames.json` | Saved nickname patterns — auto-created |
| `ReconcileIcon.ico` | Application icon for the launcher shortcut |

---

## Matching Algorithm

1. For each bank transaction, scan all unreconciled spreadsheet rows.
2. A row is a candidate if the **amount matches exactly**.
3. Candidates are split by date proximity:
   - **In-range**: date within ±3 days → preferred
   - **Out-of-range**: date further than 3 days → fallback
4. A row consumed by one bank transaction is not offered as a candidate for any other.
5. After all matching, any out-of-range candidates that were consumed by auto-matches elsewhere are removed from their candidate lists.

---

## Balance Calculation

Because odfpy cannot recalculate spreadsheet formulas after writing, the ending balance is computed in Python:

```
Computed Ending Balance = Last Reconciled Balance − Sum of All Applied Transaction Amounts
```

The **Last Reconciled Balance** is read from column I of the last row that already has a reconcile date before this run begins.

---

## Data Persistence

| Data | Storage |
|---|---|
| ODS/CSV file paths, reconcile date, ending balance | Browser `localStorage` (survives navigation) |
| Flask session state (temp file paths, flags) | Server-side session cookie (stable key in `prefs.json`) |
| Last-used browse directories | `prefs.json` |
| Nickname patterns | `nicknames.json` |
| Review page match results | Temp `.json` file in system temp directory (deleted after confirm) |

---

## Backups

Each time the ODS file is saved, a timestamped backup is created in the same directory:

```
reconcile_backup_20260326_143022.ods
```

The backup is made before any changes are written.
