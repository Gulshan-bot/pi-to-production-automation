# PI to Production Automation: one PI in, every sheet updated (Google Apps Script)

**Confirm one Proforma Invoice (PI) and the script does the rest: the customer PDF, the plant's production order, and one row per item in the right production sheet (PC to the PC tab, PP to the PP tab, and so on).**

Nobody retypes the order, nobody builds a plant copy by hand, and nobody has to distribute items to the production sheets. Machine operators work straight from their own sheet.

Built with Google Sheets and Apps Script only. No paid tools, no external servers.

---

## What one click creates

| Output | Who uses it | What it contains |
|---|---|---|
| **Full PI (PDF)** | Sales, CRM, accounts | Everything: customer name, address, GST, rates, payment terms |
| **Production order (PDF)** | Plant | Items and quantities only. Customer name is replaced by a customer ID. No address, GST, rates or payment terms |
| **One row per item in its production sheet tab** | **Machine operators** | Customer ID, PI number, material, colour, size, thickness, quantity, the weight formulas, status `New` |
| Rows in the main and plant order-to-delivery sheets | Office and plant | The commercial row in the main sheet, and the same order without commercial details in the plant sheet |
| PI Register entry | CRM, MIS | Who confirmed it, when, which revision, links to both PDFs |

It also keeps a **view-only mirror** of your order tracking (FMS) sheet for the plant team, with customer IDs instead of names.

## Who saves time

| Role | Before | After |
|---|---|---|
| **CRM** | Built the plant copy by hand, 10 to 15 minutes per order, and one missed cell could show a rate or party name to the plant | One click, about 30 seconds |
| **Whoever distributed orders to the production sheets** | Read every PI and copied each item into the right sheet (PC, PP, and so on) | That step is gone. The script routes every item |
| **Machine operators** | Waited for the sheet to be updated | Open their own sheet, see new rows with status `New`, and take it from there |
| **Plant** | Got a hand-made copy | Gets only the generated production order, with no commercial details |

---

## How it works

```
PI tab (full details)
        |
   [ Order to Delivery > Confirm this PI ]
        |
        +--> Full PI PDF -----------------> office Drive folder
        |
        +--> Plant copy of the tab
        |      name -> customer ID
        |      address, GST, rates, amounts, terms, side notes removed
        |      title -> "PRODUCTION ORDER"
        |         +--> Production order PDF --> plant Drive folder
        |
        +--> Main order-to-delivery sheet   (commercial row)
        +--> Plant sheet                    (no commercial details)
        +--> PI Register tab                (who confirmed, when, revision)
        |
        +--> Production sheets              (each item to its own tab, status "New")
               PC item -> PC tab, PP item -> PP tab, ...

Every 10 minutes (time trigger)
   FMS tab (main sheet)  --->  read-only mirror in the plant sheet
                               party name -> customer ID
                               PI link -> production order link
```

### Where each item goes

The script reads the item description of every line on the PI. The first rule that matches wins.

| Item description contains | Goes to tab |
|---|---|
| `base plate`, `louver` | PC Base Plate |
| `profil...` | Profiling, plus a row in Compact (if it says compact / solid) or Embossed for the sheet that has to be made first |
| `multiwall`, `honeycomb` | PC Multiwall |
| `diamond` | Diamond |
| `emboss...` | Embossed |
| `acrylic`, `pmma` | PMMA Acrylic |
| `compact`, `solid` | Compact |
| `ps`, `gpps`, `polystyrene` | PS Sheet |
| `pp`, `polypropylene` | PP |
| anything else | Unsorted PI items (nothing is dropped) |

These tabs, rules and column layouts are the ones from one plastic-sheet plant. Change them to match your products, see [Adapting it to your plant](#adapting-it-to-your-plant).

## Features

- **One-click confirmation** from a custom menu, with a small dialog for sales person, CRM name, sample kits, test certificate and NBD flags.
- **Automatic production sheet split.** Every item goes to the right tab, with the weight formulas already filled in. Only the customer ID is written, never the name, address, GST, rates or terms.
- **Label-driven PI reading.** The script finds fields by their labels ("Customer Name", "PI No", "Basic Price" and so on), not by fixed cell addresses, so PI layouts can differ per product.
- **Automatic customer IDs.** The first time a customer appears, the next ID (`CUST1`, `CUST2`, ...) is created and saved. The same customer always gets the same ID.
- **Safe revisions.** Confirming the same PI again saves it as `Rev1`, `Rev2` and so on, and replaces only the production rows the operators have not started on.
- **Duplicate check.** Same customer and same items under a different PI number is flagged for the plant to hold.
- **PI Register tab** with PI number, customer, customer ID, status, revision, time, who confirmed and both PDF links.
- **Plant mirror** of your FMS tab: same layout and colours, values only, no connection back to the main sheet.
- **Leak tests** that scan the plant copies and the mirror for customer names, foreign GST numbers, amounts, payment terms and full-PI links.
- **Backfill** for old orders that have no production order yet.
- **Locking**, so two people confirming at the same time cannot overwrite each other.

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
- A **planning spreadsheet**: the production sheets the machine operators work from, one tab per product (see `PLAN.TABS`).
- Two **Drive folders**: one for full PI PDFs (office only), one for production order PDFs (shared with the plant).

### 2. Install the script

1. Open the PI workbook, then **Extensions > Apps Script**.
2. Paste `Code.gs` into the editor.
3. In **Project Settings**, tick **Show "appsscript.json" manifest file in editor**, then paste in `appsscript.json` from this repo.
4. Fill in the `CFG` and `PLAN` blocks at the top of `Code.gs` (table below).

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
| `PLAN.ENABLED` | Set to `true` to turn on the production sheet split. It ships as `false` so nothing runs against a sheet you have not set up |
| `PLAN.ID` | ID of the planning spreadsheet |
| `PLAN.TABS` | Your product tabs and where their columns sit, see [Adapting it to your plant](#adapting-it-to-your-plant) |

Also check the `MIRROR` block further down: `SRC_TAB`, `HEAD_ROWS`, `PI_COL`, `NAME_COL` and `LINK_COL` must match your FMS tab. And set `BACKFILL.FY` and `BACKFILL.FROM` to match your PI numbering if you use the backfill.

### 4. First run

1. Run `setupPlantMirror()` once and approve the permissions. It builds the read-only FMS mirror in the plant sheet and hides the `PI list` helper tab.
2. Go to **Triggers** (clock icon) and add a **time-driven** trigger for `syncPlantMirror`, every 10 minutes. Create it with an account that can edit both spreadsheets and both Drive folders.
3. Reload the PI workbook. The **Order to Delivery** menu appears.
4. Open a PI tab and use **Check this PI (no changes)** to confirm it reads correctly.
5. With the split on, use **Check planning split (no changes)** to see which tab each item would go to.

### 5. Test before go-live

Use **copies** of your main, plant and planning sheets first.

| Function | What it does | Writes anything? |
|---|---|---|
| `auditTabs()` | Reads every PI tab and lists the ones it cannot parse | No |
| `leakTest()` | Builds plant copies of a sample of PI tabs and scans them for leaks | Temporary tabs only, deleted afterwards |
| `mirrorLeakTest()` | Scans the plant mirror for party names and full-PI links | No |
| `mirrorCompare()` | Compares the FMS tab with the mirror, column by column | No |
| `testActiveTab()` | Runs a full confirmation on the active tab | **Yes**: rows, PDFs and production sheet rows |

Results go to **Executions / Logs** in the Apps Script editor.

---

## Using it

1. Open the PI tab after the party confirms.
2. **Order to Delivery > Confirm this PI (party confirmed)**.
3. Enter sales person and CRM name, set the three YES/NO flags, click the button.
4. After about 30 seconds you get links to both PDFs and a line saying how many rows went to which production tab. The plant sheet shows the order within 10 minutes.

If the PI was already confirmed, the dialog warns you and saves it as a revision. It also warns when the PI number is not in a form the production sheet split accepts.

**Menu items for the production sheet split**

| Item | What it does |
|---|---|
| Check planning split (no changes) | Shows which tab each item of the open PI would go to |
| Send this PI to planning sheet | Retries the split for an already confirmed PI, if that step failed during confirmation |
| Remove a PI from the planning sheet | Deletes the not-started rows of a wrong, cancelled or duplicate PI, and lists any rows already in work so you can tell the operators |

## Rules the production sheet split follows

- **One PI, one set of rows.** Confirming the same PI again replaces the rows this script wrote earlier for it, as long as the operators have not started on them (status still starts with `New` and the customer column holds a customer ID).
- **Work in progress is never touched.** Rows typed by hand, or already Pending / Running / Complete, stay as they are. Items the plant already has (same size and quantity) are not added again. If the revised PI has other items, only those are added, with status `New - REVISED, check old rows`.
- **Possible duplicates are flagged.** On tabs that have a customer column, the same customer ID with the same items under a different PI number gets status `New - POSSIBLE DUPLICATE` and a cell note.
- **Nothing is dropped.** Items that match no product tab go to `Unsorted PI items`.
- **Quotations are skipped.** Only a proper PI number goes to the planning sheet (see `planPiOk_`).
- **A failure does not undo the confirmation.** If the split fails, the PDFs and order rows stand and the dialog tells you to retry from the menu.

## Adapting it to your plant

The production sheet split is written for one specific planning sheet. Expect to change these:

| What | Where | Change it to |
|---|---|---|
| Product tabs and where their columns sit | `PLAN.TABS` | Your tab names. `first` = first data row, `orderCol` = column with the PI number, `statusCol` = status column, `kind` = row layout |
| Which item goes to which tab | `routeItem_` | Your own match rules on the item description |
| What each row contains | `planRow_` | Your own column order and formulas, one block per `kind` (`roll`, `mw`, `prof`, `base`) |
| Weight per item | `itemKg_` | Your own formula |
| Item columns read from the PI | `readItems_` | Your own column labels (`Colour`, `Width`, `Length`, ...) |
| PI number format that is accepted | `planPiOk_` | Your own numbering pattern |

Other things you may want to change:

- **Different PI labels:** edit the regular expressions in `readPI_`.
- **Different things to hide from the plant:** edit `makePlantCopy_`.
- **Different customer ID format:** change `CODE_PREFIX` or `customerCode_`.
- **Different main or plant sheet columns:** see the column maps below and the `getRange` / `setValues` calls in `confirmAndSplit`.

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

## Limitations

- **This is a redaction helper, not a security boundary.** The plant team and machine operators must only have access to the plant spreadsheet, the planning spreadsheet and the plant PDF folder, never to the PI workbook, the main sheet or the full-PI folder.
- Redaction is based on labels and position. A PI layout the script has not seen may leak something. Run `auditTabs()` and `leakTest()` when you add a new product layout, and look at the first few production orders yourself.
- A PI tab must contain these labels: `Customer Name`, `PI No`, `SL.`, `Basic Price`, `Order Quantity`, `Total Qty` (or a `Total Amount` row), and `Total Amount`. `Date :`, `Address`, `Other Ref` and `Payment` are used if present.
- The production sheet split writes formulas into fixed columns of your planning tabs. If your layout differs and you do not adapt `PLAN.TABS` and `planRow_`, it writes into the wrong columns. Test on a copy of the planning sheet first.
- Routing depends on the words in the item description. An item described in an unexpected way lands in `Unsorted PI items`, so someone should check that tab.
- Production rows are matched to a PI by its number, so keep PI numbers unique and in one format.
- PDF export uses an A4 landscape, fit-to-width layout. Change the URL parameters in `exportPdf_` for other page setups.
- The dispatch date is left blank on purpose: in this workflow the plant team fills it in.
- Apps Script limits apply (execution time, daily URL fetch and Drive quotas). The backfill stops after about 4.5 minutes, so run it again until nothing is left.

## Contributing

Issues and pull requests are welcome, especially for other PI layouts, other product routing rules and extra leak checks. Please use sample data only. Never post real customer names, GST numbers, rates or sheet IDs in issues or screenshots.

## License

MIT. See [LICENSE](LICENSE).
