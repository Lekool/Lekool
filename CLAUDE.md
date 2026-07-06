# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

A Google Apps Script tool, bound to a Google Sheet, that ingests credit card CSV exports (Chase, Amex), normalizes them, and writes each upload to its own sheet within the active spreadsheet.

## Runtime & "Build"

There is no local build, test runner, package manager, or CI. The code runs inside Google's Apps Script V8 runtime, not Node. To "deploy":

1. Open the target Google Sheet → `Extensions` → `Apps Script`.
2. Replace the contents of `Code.gs` and the `FileUpload.html` file in the Apps Script editor with the versions from this repo.
3. Save, refresh the sheet, and use the **Finance Tool** menu.

Because there is no test harness, validate changes by uploading representative Chase and Amex CSVs through the dialog and inspecting the resulting sheet. The "Test Server Connection" button in `FileUpload.html` calls `testConnection()` and is the quickest smoke test that the client↔server bridge works.

## Architecture

Two files, two sides of a `google.script.run` boundary:

- **`Code.gs`** — server side (Apps Script). Entry points are `onOpen` (installs the menu) and `loadNewFile` (opens the modal). The real work is `processUploadedCsv(base64Content, cardName)`, invoked from the client. `getCardNames()` is also client-invoked — it returns the existing tab names to populate the dialog's card dropdown.
- **`FileUpload.html`** — client side. Renders the drag-and-drop dialog, populates a **card dropdown** from `getCardNames()` (plus a "+ New card…" free-text option), base64-encodes the file via `FileReader.readAsDataURL`, strips the `data:...;base64,` prefix, and posts the payload + chosen card name to `processUploadedCsv`. It surfaces the `{added, skipped}` summary the server returns.

### Server-side flow (`processUploadedCsv`)

1. **One tab per card, with append.** If no tab named `cardName` exists, it's created (the old behavior). If it *does* exist, the upload is **appended** to it — not rejected. The card name comes from a dropdown built by `getCardNames()`, so it's normally an existing tab.
2. Base64-decode → `Utilities.parseCsv` → 2D array.
3. **Format detection by header sniffing**, not filename:
   - Chase: presence of both `Transaction Date` AND `Post Date`.
   - Amex: presence of `Card Member`.
   - Anything else throws.
4. `validateHeaders` enforces the required column set for the detected format.
5. `processChaseData` / `processAmexData` map each row to the unified 8-column schema:
   `[Transaction Date, Post Date, Description, Category, Type, Amount, Card, Notes]`
6. **Sign convention is normalized to Chase's**: charges negative, credits positive. Amex amounts are negated (`-amount`) because Amex exports use the opposite sign.
7. **Amex has no Post Date** — column B is written as `''` (empty string, not null). The sort comparator falls back to Transaction Date when Post Date is missing (`a[1] instanceof Date ? a[1] : a[0]`). Don't "fix" this by writing a Date to column B for Amex; downstream sorting depends on the empty-string sentinel.
8. **Write path depends on whether the tab exists:**
   - **New tab:** sort ascending via `compareByDate`, `insertSheet(cardName)` → `setupNewSheet` (headers, `A:B` as `MM/dd/yyyy`, freeze row 1) → `setValues` → `applyAmountFormatting`.
   - **Existing tab:** read current rows, drop incoming rows whose `dedupKey` already appears (also collapses dupes within the same upload), append the survivors, re-sort the **whole tab** with `compareByDate`, clear the old data region, rewrite, and re-apply `applyAmountFormatting`.
   - Returns `{card, format, added, skipped, created}` so the client can report "Added N, skipped M".

### Append, dedup, and the mismatch guard

- **`dedupKey(row)`** = `TransactionDate.getTime() | Amount | Description | Type`. Two rows with the same key are the same transaction. Deliberate tradeoff: two genuinely identical charges (same day/amount/merchant) collapse to one. Reading an existing tab returns dates as `Date` objects and amounts as numbers, so keys match those built from freshly-processed rows.
- **`compareByDate`** is the single shared sort comparator (used for both new and appended tabs). Same rule as before: Post Date if it's a `Date`, else fall back to Transaction Date — which is why the `''` Post Date sentinel matters (see below).
- **`detectSheetFormat(sheet)`** is a best-effort Chase-vs-Amex guess used only for the cross-issuer mismatch guard: appending a Chase file into an Amex tab (or vice versa) throws. Heuristic: Amex populates the Notes column (Card Member), Chase leaves it empty, so a tab whose rows mostly have Notes is Amex. It **cannot** tell two Chase cards apart — that's inherent to the CSV format, so the dropdown is the sole source of Chase-card identity.
- **`applyAmountFormatting(sheet)`** sets two conditional-format rules on column **F** (`F2:F`): `< 0` → light-blue fill (`#cfe2f3`), `> 0` → light-green fill (`#d9ead3`). It's idempotent — `setConditionalFormatRules` replaces all rules — so it's safe to call on every write.

### Date parsing

`parseDateUniversal` tries `new Date(...)` first (handles `MM/dd/yyyy` and most native-parseable forms), then falls back to `parseYyyyMmDd` for `yyyy/mm/dd`. Both return `null` on failure. If you add a new format, add it to `parseDateUniversal` rather than calling a new parser from each processor.

**Row-drop strictness is asymmetric and deliberate.** A row is dropped (`.filter(row => row !== null)`) only when its *primary* date fails to parse — Transaction Date for Chase, Date for Amex. A Chase row with a valid Transaction Date but an **unparseable Post Date is kept**, writing `''` into slot 1 (same sentinel as Amex). Don't tighten `processChaseData` back to requiring both dates; that silently discards otherwise-valid transactions.

### Adding a new card format

1. Add a header-sniff branch in `processUploadedCsv` next to the existing Chase/Amex checks.
2. Add a `processXxxData(rows, headers, cardName)` that returns rows in the 8-column unified schema.
3. Decide the sign convention vs. the Chase baseline (negate if needed).
4. If Post Date isn't available, write `''` in slot 1 — do NOT invent one.
5. Revisit `detectSheetFormat` — its Notes-column heuristic only separates Chase from Amex. A third format needs a new signal (or the mismatch guard will mislabel it), and `dedupKey` assumes the standard 8-column schema.

## Constraints worth remembering

- `@OnlyCurrentDoc` at the top of `Code.gs` scopes OAuth to the bound spreadsheet only; broadening it changes the auth prompt users see.
- The 10MB client-side cap in `FileUpload.html` is a transport guard, not the Apps Script hard limit — `google.script.run` payloads get fragile well before Apps Script's documented limits.
- No external libraries. Stick to the Apps Script standard library (`Utilities`, `SpreadsheetApp`, `HtmlService`).
