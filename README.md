# Personal Finance Google Sheet Tool

## Overview

This project is a Google Apps Script-based tool designed to create a personal finance management system within a single Google Sheet. It automates the process of consolidating credit card transaction data from various sources into a tidy, per-card, sorted, and easy-to-review set of tabs.

The script is built to handle CSV files from different credit card providers (specifically Chase and American Express), normalize their disparate data formats, and present them in a standardized way.

## Features

- **Per-Card Tabs:** Each card gets its own sheet tab. Uploads for a card **append** into that card's existing tab, and the whole tab is re-sorted after every upload.
- **Card Dropdown:** Pick the card from a dropdown in the upload dialog instead of typing a name each time. The list is a fixed set of four built-in cards (`Chase Sapphire`, `Chase Amazon`, `Chase Unlimited`, `Amex`) plus any cards you add via **"+ New card…"**. Added cards can be removed under **"Manage cards"** (removal only takes the card off the dropdown — the tab and its data stay). The four built-ins are permanent.
- **Duplicate Skipping:** When you upload an overlapping statement period, rows that exactly match an existing row (Transaction Date + Amount + Description + Type) are skipped, and the dialog reports how many were added vs. skipped.
- **Smart Data Normalization:** Automatically processes data from different providers (Chase, Amex) and standardizes them into a consistent format.
  - Inverts amount signs where necessary to ensure charges are negative and credits/payments are positive.
  - Parses multiple common date formats (`MM/dd/yyyy`, `yyyy/mm/dd`).
- **Amount Highlighting:** The Amount column is conditionally formatted — light-blue fill for negative amounts (charges), light-green for positive (credits/payments).
- **Drag-and-Drop File Upload:** A modern, easy-to-use dialog for uploading your CSV files.
- **Automatic Sorting:** After every file upload, the tab is re-sorted in ascending order, prioritizing the `Post Date` and using the `Transaction Date` as a fallback.
- **Simple Interface:** All functionality is accessible through a custom "Finance Tool" menu directly within your Google Sheet.

## Setup Instructions

To get this tool working in your own Google Sheet, follow these steps carefully.

1.  **Create a New Google Sheet:** Start with a fresh, blank Google Sheet.

2.  **Open the Script Editor:** In the top menu of your sheet, navigate to `Extensions` > `Apps Script`. This will open the script editor in a new browser tab.

3.  **Add the Main Script (`Code.gs`):**
    *   You will see a file named `Code.gs`. Delete any default code inside it.
    *   Copy the entire contents of the `Code.gs` file from this repository and paste it into the `Code.gs` file in the editor.

4.  **Add the HTML User Interface (`FileUpload.html`):**
    *   In the script editor, click the **`+`** icon next to "Files" in the left-hand sidebar and select **HTML**.
    *   In the dialog box that appears, name the file `FileUpload` (the `.html` extension will be added automatically).
    *   A new `FileUpload.html` file will be created. Delete any default code inside it.
    *   Copy the entire contents of the `FileUpload.html` file from this repository and paste it into your new `FileUpload.html` file.

5.  **Save the Project:** Click the "Save project" icon (which looks like a floppy disk) at the top of the script editor.

6.  **Refresh:**
    *   Go back to your Google Sheet browser tab and **refresh the page**.
    *   A new menu named **"Finance Tool"** should now appear in the menu bar.

7.  **Authorize the Script (One-Time Step):**
    *   The first time you run `Finance Tool` > `Load New CSV File`, Google will ask for your permission. This is a standard security step.
    *   A dialog will say "Authorization required." Click **Continue**.
    *   Choose your Google account from the list.
    *   You will likely see a warning screen that says "Google hasn't verified this app." This is expected because you are running your own script. Click on the **"Advanced"** link.
    *   Click on **"Go to [Your Project Name] (unsafe)"**.
    *   Finally, review the permissions and click **"Allow"**.

Your sheet is now ready to use!

## How to Use

1.  Click on the **"Finance Tool"** menu.
2.  Select **"Load New CSV File"**.
3.  In the dialog box, **pick the card** from the dropdown. To add a card that isn't listed, choose **"+ New card…"**, type its name, and click **Add**. To tidy the list, open **"Manage cards"** and click ✕ next to an added card (this only removes it from the dropdown; its tab and data are kept).
4.  **Drag and drop** your CSV file into the box, or click the file input to select it from your computer.
5.  The script shows a "processing" message, then appends the new transactions to that card's tab, skips any exact duplicates, re-sorts the tab, and reports how many rows were added vs. skipped.

> **Note:** Because Chase CSVs are identical in format across cards, the tool can't tell your Chase cards apart — the dropdown choice decides which tab the rows land in. Make sure you pick the right card. (It *will* warn you if you drop a Chase file onto an Amex tab, or vice versa.)

## Supported CSV Formats

The script is designed to automatically detect the file type based on the column headers in the CSV.

### Chase Format

The script identifies a Chase file by the presence of both `Transaction Date` and `Post Date` columns. It expects the following columns:
- `Transaction Date`
- `Post Date`
- `Description`
- `Category`
- `Type`
- `Amount` (Charges are negative, credits are positive)

### American Express (Amex) Format

The script identifies an Amex file by the presence of the `Card Member` column. It expects the following columns:
- `Date` (This is used as the Transaction Date)
- `Description`
- `Card Member`
- `Amount` (Charges are positive, credits are negative)
- `Account #` (This column is ignored)
