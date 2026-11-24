# Order to Delivery: one-click PI split (Google Apps Script)

Turn one Proforma Invoice (PI) into **two documents with a single click**:

| Document | Who sees it | What it contains |
|---|---|---|
| **Full PI (PDF)** | Sales, CRM, accounts | Everything: customer name, address, GST, rates, payment terms |
| **Production order (PDF)** | Plant / production team | Items and quantities only. Customer name is replaced by a customer ID. No address, GST, rates or payment terms |

It also writes both orders to the right tracking sheets and keeps a **view-only mirror** of your order tracking (FMS) sheet for the plant team. Optionally, it sends every item line of the PI to the right product tab of a **planning sheet** (see [Planning sheet split](#planning-sheet-split-optional)).

Making the plant copy by hand took 10 to 15 minutes per order, and one missed cell could leak commercial details. This script does it in about 30 seconds and can test itself for leaks.

Built with Google Sheets and Apps Script only. No paid tools, no external servers.

---

## How it works

```
PI tab (full details)
        |
   [ Order to Delivery > Confirm this PI ]
        |
        +--> Full PI PDF ------------> office Drive folder
        |
        +--> Plant copy of the tab
        |      name -> customer ID
        |      address, GST, rates, amounts, terms, side notes removed
        |      title -> "PRODUCTION ORDER"
        |         |
        |         +--> Production order PDF --> plant Drive folder
        |
        +--> Main order-to-delivery sheet   (commercial row)
        +--> Plant sheet                    (no commercial details)
        +--> PI Register tab                (who confirmed, when, revision)
        |
        +--> Planning sheet (optional)      one row per item, in its product tab, status "New"

Every 10 minutes (time trigger)
   FMS tab (main sheet)  --->  read-only mirror in the plant sheet
                               party name -> customer ID
                               PI link -> production order link
```

## Features

- **One-click confirmation** from a custom menu, with a small dialog for sales person, CRM name, sample kits, test certificate and NBD flags.
- **Label-driven PI reading.** The script finds fields by their labels ("Customer Name", "PI No", "Basic Price" and so on), not by fixed cell addresses, so PI layouts can differ per product.
- **Automatic customer IDs.** The first time a customer appears, the next ID (`CUST1`, `CUST2`, ...) is created and saved. The same customer always gets the same ID.
- **Revision tracking.** Confirming the same PI again saves it as `Rev1`, `Rev2` and so on, and updates the existing rows instead of adding duplicates.
- **PI Register tab** with PI number, customer, customer ID, status, revision, time, who confirmed and both PDF links.
- **Plant mirror** of your FMS tab: same layout and colours, values only, no connection back to the main sheet.
- **Leak tests** that scan the plant copies and the mirror for customer names, foreign GST numbers, amounts, payment terms and full-PI links.
- **Backfill** for old orders that have no production order yet.
- **Locking**, so two people confirming at the same time cannot overwrite each other.
- **Optional planning split.** One row per PI item goes to its product tab of a planning sheet with status `New`. Confirming a PI again replaces rows the plant has not started on, never touches rows already in work, and flags possible duplicate orders.

## What gets removed from the plant copy

- Customer name (replaced by the customer ID)
- Address, GST number and contact lines under the name
- Other reference / PO number
- Every number from the price column rightwards (rates, amounts, taxes, grand total)
- Payment terms
- Working notes typed outside the PI block, to the right of the amount column or below the signature / bank details

Your own GSTIN (`OWN_GST`) is allowed to stay.

---

## Setup

### 1. What you need

- A **PI workbook**: one tab per PI. Tabs that hold PIs have `_PI_` in their name (for example `ABC_PI_26-27_0001`). A tab ending in `(B)` is treated as a hand-made plant copy and skipped.
- A **main order-to-delivery spreadsheet** with:
  - an `order form` tab (one row per confirmed PI, column map below)
  - an `FMS` tab (your order tracking sheet)
  - a customer code tab (column A: customer name, column B: customer ID, header in row 1)
- A **plant spreadsheet** (a blank one is fine, the script builds the rest).
- Two **Drive folders**: one for full PI PDFs (office only), one for production order PDFs (shared with the plant).

### 2. Install the script

1. Open the PI workbook, then **Extensions > Apps Script**.
2. Paste `Code.gs` into the editor.
3. In **Project Settings**, tick **Show "appsscript.json" manifest file in editor**, then paste in `appsscript.json` from this repo.
4. Fill in the `CFG` block at the top of `Code.gs` (table below).

### 3. Configuration

| Setting | What to put |
|---|---|
| `MASTER_ID` | ID of the main spreadsheet (the part of its URL between `/d/` and `/edit`) |
| `MASTER_TAB` | Tab in the main sheet that gets one row per confirmed PI |
| `CODES_TAB` | Customer name to customer ID table |
| `PLANT_ID` | ID of the plant spreadsheet |
| `FULL_PDF_FOLDER` | Drive folder ID for full PI PDFs |
| `PLANT_PDF_FOLDER` | Drive folder ID for production order PDFs |
| `REGISTER_TAB` | Name of the register tab (created automatically) |
| `CODE_PREFIX` | Prefix for new customer IDs |
| `OWN_GST` | Your own GSTIN, so it is not flagged as a leak. Leave `''` if you do not need it |
| `PLAN.ENABLED` | `false` (default) = no planning split. `true` = send every confirmed PI to the planning sheet |
| `PLAN.ID` | ID of the planning spreadsheet. Only needed when `PLAN.ENABLED` is `true` |

The planning split is off by default and needs more adapting, see [Planning sheet split](#planning-sheet-split-optional). Also check the `MIRROR` block further down: `SRC_TAB`, `HEAD_ROWS`, `PI_COL`, `NAME_COL` and `LINK_COL` must match your FMS tab. And set `BACKFILL.FY` and `BACKFILL.FROM` to match your PI numbering if you use the backfill.

### 4. First run

1. Run `setupPlantMirror()` once and approve the permissions. It builds the read-only FMS mirror in the plant sheet and hides the `PI list` helper tab.
2. Go to **Triggers** (clock icon) and add a **time-driven** trigger for `syncPlantMirror`, every 10 minutes. Create it with an account that can edit both spreadsheets and both Drive folders.
3. Reload the PI workbook. The **Order to Delivery** menu appears.
4. Open a PI tab and use **Check this PI (no changes)** to confirm it reads correctly.

### 5. Test before go-live

Use **copies** of your main and plant sheets first (and a copy of your planning sheet if you turn the planning split on).

| Function | What it does | Writes anything? |
|---|---|---|
| `auditTabs()` | Reads every PI tab and lists the ones it cannot parse | No |
| `leakTest()` | Builds plant copies of a sample of PI tabs and scans them for leaks | Temporary tabs only, deleted afterwards |
| `mirrorLeakTest()` | Scans the plant mirror for party names and full-PI links | No |
| `mirrorCompare()` | Compares the FMS tab with the mirror, column by column | No |
| `testActiveTab()` | Runs a full confirmation on the active tab | **Yes**: rows and PDFs |

Results go to **Executions / Logs** in the Apps Script editor.

---

## Using it

1. Open the PI tab after the party confirms.
2. **Order to Delivery > Confirm this PI (party confirmed)**.
3. Enter sales person and CRM name, set the three YES/NO flags, click the button.
4. After about 30 seconds you get links to both PDFs. The plant sheet shows the order within 10 minutes.

If the PI was already confirmed, the dialog warns you and saves it as a revision. With the planning split on, the dialog also warns when the PI number is not in a form the planning sheet accepts.

## Column maps

**Main sheet, `order form` tab** (written by the script)

| Column | Value |
|---|---|
| A | Timestamp |
| B | PI no (also used to find an existing row) |
| C | Customer |
| D | Category (item descriptions) |
| E | Quantity (second unit if present, else first) |
| F | Basic value |
| G | Order date |
| J | PI no |
| M | Production order PDF link |
| P | Sample kits / brochure / pads / pens (YES/NO) |
| Q | Sales person |
| R | Test certificate required (YES/NO) |
| S | CRM name |
| T | NBD client (YES/NO) |
| U | Full PI PDF link |

**Plant sheet, `PI list` tab:** PI no, Rev, Customer ID, Order date, Category, Quantity, Dispatch date, Kit, Test certificate, Production order PDF, Confirmed at, Note.

If your sheets use different columns, change the `setValues` / `getRange` calls in `confirmAndSplit`.

## Planning sheet split (optional)

Off by default. Set `PLAN.ENABLED = true` and `PLAN.ID` to turn it on.

**What it does.** After the PDFs and order rows are written, the script reads every item line of the PI and adds one row per item to the matching product tab of a planning spreadsheet, with status `New`. Only the customer ID is written, never the name, address, GST, rates or terms. If this step fails, the confirmation still stands and the dialog tells you to retry from the menu.

**This module is an example.** It comes from a plastic-sheet plant, so the tabs, routing rules, column layouts and weight formulas describe one specific planning sheet. Expect to adapt it:

| What | Where | Change it to |
|---|---|---|
| Product tabs and where their columns sit | `PLAN.TABS` | Your tab names. `first` = first data row, `orderCol` = column with the PI number, `statusCol` = status column, `kind` = row layout |
| Which item goes to which tab | `routeItem_` | Your own match rules on the item description |
| What each row contains | `planRow_` | Your own column order and formulas, one block per `kind` (`roll`, `mw`, `prof`, `base`) |
| Weight per item | `itemKg_` | Your own formula |
| Item columns read from the PI | `readItems_` | Your own column labels (`Colour`, `Width`, `Length`, ...) |
| PI number format that is accepted | `planPiOk_` | Your own numbering pattern |

**Rules it follows**

- **One PI, one set of rows.** Confirming the same PI again replaces the rows this script wrote earlier for it, as long as the plant has not started on them (status still starts with `New` and the customer column holds a customer ID).
- **Work in progress is never touched.** Rows the plant typed by hand, or that are already Pending / Running / Complete, stay as they are. Items the plant already has (same size and quantity) are not added again. If the revised PI has other items, only those are added, with status `New - REVISED, check old rows`.
- **Possible duplicates are flagged.** On tabs that have a customer column, the same customer ID with the same items under a different PI number gets status `New - POSSIBLE DUPLICATE` and a cell note, so the plant can hold the order.
- **Nothing is dropped.** Items that match no product tab go to a tab called `Unsorted PI items`.
- **Quotations are skipped.** Only a proper PI number goes to the planning sheet (see `planPiOk_`).

**Menu items** (shown only when the split is on)

| Item | What it does |
|---|---|
| Check planning split (no changes) | Shows which tab each item of the open PI would go to |
| Send this PI to planning sheet | Retries the step for an already confirmed PI |
| Remove a PI from the planning sheet | Deletes the not-started rows of a wrong, cancelled or duplicate PI, and lists any rows already in work so you can tell the plant |

## Limitations

- **This is a redaction helper, not a security boundary.** The plant team must only have access to the plant spreadsheet and the plant PDF folder, never to the PI workbook, the main sheet or the full-PI folder.
- Redaction is based on labels and position. A PI layout the script has not seen may leak something. Run `auditTabs()` and `leakTest()` when you add a new product layout, and look at the first few production orders yourself.
- A PI tab must contain these labels: `Customer Name`, `PI No`, `SL.`, `Basic Price`, `Order Quantity`, `Total Qty` (or a `Total Amount` row), and `Total Amount`. `Date :`, `Address`, `Other Ref` and `Payment` are used if present.
- PDF export uses an A4 landscape, fit-to-width layout. Change the URL parameters in `exportPdf_` for other page setups.
- The dispatch date is left blank on purpose: in this workflow the plant team fills it in.
- The planning split writes formulas into fixed columns of your planning tabs. If your layout differs and you do not adapt `PLAN.TABS` and `planRow_`, it writes into the wrong columns. Test on a copy of the planning sheet first.
- Planning rows are matched to a PI by its number, so keep PI numbers unique and in one format.
- Apps Script limits apply (execution time, daily URL fetch and Drive quotas). The backfill stops after about 4.5 minutes, so run it again until nothing is left.

## Customising

- **Different labels:** edit the regular expressions in `readPI_`.
- **Different things to hide:** edit `makePlantCopy_`.
- **Different ID format:** change `CODE_PREFIX` or `customerCode_`.
- **Different sheet columns:** see the column maps above.
- **Planning split:** see the table in [Planning sheet split](#planning-sheet-split-optional).

## Contributing

Issues and pull requests are welcome, especially for other PI layouts and extra leak checks. Please use sample data only. Never post real customer names, GST numbers, rates or sheet IDs in issues or screenshots.

## License

MIT. See [LICENSE](LICENSE).
