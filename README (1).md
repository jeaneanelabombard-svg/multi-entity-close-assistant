# Multi-Entity Close Assistant

A month-end close assistant for a multi-entity family office, built in Google Sheets with Apps Script, the Claude API and the QuickBooks Online API.

It codes bank activity across entities, flags what it isn't sure of, reconciles each bank account, and posts approved entries to QuickBooks Online. Nothing posts until a person approves it.

Built as a technical exercise. All entities and transactions are fictional demo data, and the QuickBooks connection was tested against an Intuit Developer sandbox company only.

## What it does

**Close Tools → ▶ Run full close** (`Code.gs`)

1. **Match.** Matches each bank line to entries already in the books.
2. **Categorize.** Keyword rules first, then Claude for what the rules can't place. Claude must pick an account from the chart of accounts or answer `UNSURE`, with a confidence score and a reason.
3. **Check.** Duplicates, transpositions (differences divisible by 9), intercompany transfers with no mirror in the other entity's bank, large items, and low-confidence AI picks.
4. **Reconcile.** Builds a bank reconciliation per entity that must tie to the penny.
5. **Export.** Writes a QuickBooks journal entry import file for clean lines only.
6. **Summarize.** A summary, a review queue, and a run log for every step.

**Close Tools → QuickBooks** (`QuickBooks.gs`)

1. Save QuickBooks app keys
2. Connect an entity (one QuickBooks company per entity, accounting scope only)
3. Test connection
4. Create missing accounts
5. Load demo books (sandbox only)
6. Pull books from QuickBooks
7. Post approved lines

## Controls

- **A person approves every entry.** Only lines that are "Ready to import" and ticked Approve are posted, after a confirmation showing the count, the dollars and the QuickBooks company.
- **Safety stop.** It won't read or post unless the QuickBooks company name matches the entity's legal name.
- **No double posting.** Journal numbers (`BF-<bank line>`) already in QuickBooks are skipped.
- **Ties are checked.** Each bank rec must tie, and opening cash in QuickBooks must tie to the Entities tab.
- **Audit trail.** `Run_Log` records every step; `QBO_Posted` records each entry's QuickBooks ID, amount, account and approver.
- **No secrets in the sheet or the code.** The Claude key and QuickBooks credentials are stored in the script's private user properties.
- **Consistent reruns.** Claude's answers are cached, so a rerun gives the same result and makes no new AI calls.

## Setup

1. Create a Google Sheet, open **Extensions → Apps Script**, and add `Code.gs` and `QuickBooks.gs`.
2. Add the OAuth2 for Apps Script library (identifier `OAuth2`), script ID `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`.
3. Reload the sheet, then run **Close Tools → 1. Set up / reset demo data** and **2. Save Claude API key**.
4. For QuickBooks: create an app at developer.intuit.com, add the redirect URI `https://script.google.com/macros/d/<SCRIPT_ID>/usercallback`, rename the sandbox company to match the entity's legal name, and turn on account numbers.

## Known limitations

- **Rules can be confidently wrong.** Trust transfers and trust tax payments should always be held for review.
- **"Ties" overstates it.** Unbooked bank activity counts as a reconciling item, so the rec should report open items and use the actual statement balance.
- **Notes aren't holds.** A note asking a person to confirm something should hold the line.
- **Pull books reads journal entries only.** In production it would read the general ledger for the cash account, which also includes checks, deposits and bank feed activity.
- **Bank feeds.** If QuickBooks bank feeds are on, feed lines must be matched to posted entries, or cash is counted twice.
