/**
 * Order to Delivery - single PI entry, automatic split (Google Apps Script)
 *
 * Lives in the Proforma Invoice (PI) workbook. The CRM opens a PI tab and uses
 * menu "Order to Delivery" > "Confirm this PI (party confirmed)".
 * One confirmation creates:
 *   - full PI PDF + a row in the main order-to-delivery sheet (commercial details)
 *   - production order PDF + a row in the plant sheet
 *     (no customer name / address / GST / rates / payment terms)
 *   - one row per PI item in its product tab of the planning (production) sheet, so every item
 *     reaches the right machine operators with no manual distribution
 *     (switch on with PLAN.ENABLED, see PLAN further down)
 *
 * Setup: fill in CFG below, run setupPlantMirror() once, then add a time-driven
 * trigger for syncPlantMirror() (every 10 minutes). See README.md.
 */
var CFG = {
  MASTER_ID: 'YOUR_MAIN_SHEET_ID',               // main order-to-delivery spreadsheet
  MASTER_TAB: 'order form',                      // tab in the main sheet where each confirmed PI gets a row
  CODES_TAB: 'Customer Codes',                   // tab in the main sheet: column A customer name, column B customer ID
  PLANT_ID: 'YOUR_PLANT_SHEET_ID',               // plant spreadsheet (no commercial details)
  FULL_PDF_FOLDER: 'YOUR_FULL_PDF_FOLDER_ID',    // Drive folder for full PI PDFs (office only)
  PLANT_PDF_FOLDER: 'YOUR_PLANT_PDF_FOLDER_ID',  // Drive folder for production order PDFs (shared with plant)
  REGISTER_TAB: 'PI Register',                   // created automatically in the PI workbook
  CODE_PREFIX: 'CUST',                           // new customer IDs look like CUST1, CUST2, ...
  OWN_GST: ''                                    // your own GSTIN: allowed to stay on the plant copy. Leave '' if none.
};

function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu('Order to Delivery')
    .addItem('Confirm this PI (party confirmed)', 'showConfirmDialog')
    .addItem('Check this PI (no changes)', 'checkThisPI');
  if (PLAN.ENABLED) {
    menu.addSeparator()
      .addItem('Check planning split (no changes)', 'checkPlanSplit')
      .addItem('Send this PI to planning sheet (only if that step failed)', 'sendThisPIToPlan')
      .addItem('Remove a PI from the planning sheet (wrong / cancelled / duplicate)', 'removePIFromPlan');
  }
  menu.addToUi();
}

/* ---------- reading the PI tab (label driven, layouts differ per product) ---------- */

function norm_(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
function key_(s) { return norm_(s).toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function isLabel_(s) { return /:-?\s*$/.test(norm_(s)); }

/** True if the text contains a GSTIN that is not our own. */
function foreignGst_(t) {
  var g = norm_(t).match(/\d{2}[A-Z]{5}\d{4}[A-Z][A-Z\d]{3}/g);
  return !!(g && g.join('') !== CFG.OWN_GST);
}

function find_(disp, re, fromRow) {
  for (var r = fromRow || 0; r < disp.length; r++)
    for (var c = 0; c < disp[r].length; c++)
      if (re.test(norm_(disp[r][c]))) return { r: r, c: c };
  return null;
}
function findLast_(disp, re) {
  var hit = null;
  for (var r = 0; r < disp.length; r++)
    for (var c = 0; c < disp[r].length; c++)
      if (re.test(norm_(disp[r][c]))) hit = { r: r, c: c };
  return hit;
}
function rightOf_(disp, pos) {
  if (!pos) return '';
  for (var c = pos.c + 1; c < disp[pos.r].length; c++) {
    var v = norm_(disp[pos.r][c]);
    if (v) return isLabel_(v) ? '' : v;
  }
  return '';
}

function readPI_(sheet) {
  var rng = sheet.getDataRange();
  var disp = rng.getDisplayValues(), vals = rng.getValues();
  var L = {
    name: find_(disp, /^customer name/i),
    addr: find_(disp, /^address/i),
    pi: find_(disp, /^pi no/i),
    date: find_(disp, /^date\s*:/i),
    ref: find_(disp, /^other ref/i),
    sl: find_(disp, /^sl\.?$/i),
    price: find_(disp, /^basic price/i),
    qty: find_(disp, /^order quantity/i),
    tot: find_(disp, /^total qty/i),
    grand: findLast_(disp, /total amount/i),
    pay: find_(disp, /^payment/i)
  };
  if (!L.tot && L.sl) { var ft = find_(disp, /total amount/i, L.sl.r + 1); if (ft) L.tot = { r: ft.r, c: ft.c }; }   // layouts without a TOTAL QTY label
  var missing = [];
  ['name', 'pi', 'sl', 'price', 'qty', 'tot', 'grand'].forEach(function (k) { if (!L[k]) missing.push(k); });
  if (missing.length) throw new Error('This tab does not look like a PI. Could not find: ' + missing.join(', '));

  var descCol = L.sl.c + 1, items = [], seen = {}, cats = [];
  for (var r = L.sl.r + 2; r < L.tot.r; r++) {
    var d = norm_(disp[r][descCol]);
    if (!d) continue;
    items.push(d);
    if (!seen[key_(d)]) { seen[key_(d)] = 1; cats.push(d); }
  }
  var num = function (v) { return typeof v === 'number' ? v : 0; };
  var two = L.qty.c + 1 < L.price.c;                      // some PIs have a single quantity column
  var q1 = num(vals[L.tot.r][L.qty.c]), q2 = two ? num(vals[L.tot.r][L.qty.c + 1]) : 0;
  var u1 = norm_(disp[L.sl.r + 1][L.qty.c]), u2 = two ? norm_(disp[L.sl.r + 1][L.qty.c + 1]) : '';
  var value = 0, ta = find_([disp[L.tot.r]], /total amount/i);
  for (var c = (ta ? ta.c + 1 : L.price.c); c < vals[L.tot.r].length; c++)
    if (typeof vals[L.tot.r][c] === 'number') { value = vals[L.tot.r][c]; break; }

  var dateVal = '';
  if (L.date) {
    for (var c2 = L.date.c + 1; c2 < disp[L.date.r].length; c2++)
      if (norm_(disp[L.date.r][c2])) { dateVal = vals[L.date.r][c2] instanceof Date ? vals[L.date.r][c2] : norm_(disp[L.date.r][c2]); break; }
  }
  return {
    L: L, disp: disp, vals: vals,
    piNo: rightOf_(disp, L.pi), customer: rightOf_(disp, L.name), date: dateVal,
    category: cats.join(', '), itemCount: items.length,
    qty1: Math.round(q1 * 1000) / 1000, unit1: u1, qty2: Math.round(q2 * 1000) / 1000, unit2: u2, value: Math.round(value * 100) / 100
  };
}

/* ---------- customer code ---------- */

function customerCode_(master, name) {
  var sh = master.getSheetByName(CFG.CODES_TAB);
  if (!sh) throw new Error('Customer code tab "' + CFG.CODES_TAB + '" not found in the main sheet.');
  var data = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 2).getDisplayValues();
  var k = key_(name), maxN = 0;
  for (var i = 1; i < data.length; i++) {
    if (key_(data[i][0]) === k && norm_(data[i][1])) return { code: norm_(data[i][1]), isNew: false };
    var m = /(\d+)\s*$/.exec(data[i][1]);
    if (m) maxN = Math.max(maxN, parseInt(m[1], 10));
  }
  var code = CFG.CODE_PREFIX + (maxN + 1);
  sh.appendRow([norm_(name), code]);
  return { code: code, isNew: true };
}

/* ---------- PDFs ---------- */

function exportPdf_(ss, sheet, fileName, folderId) {
  SpreadsheetApp.flush();
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf&gid=' + sheet.getSheetId() +
    '&size=A4&portrait=false&fitw=true&gridlines=false&sheetnames=false&printtitle=false&pagenumbers=false&fzr=false';
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('PDF export failed (' + resp.getResponseCode() + ') for ' + fileName);
  var file = DriveApp.getFolderById(folderId).createFile(resp.getBlob().setName(fileName));
  return file.getUrl();
}

/** Copy of the PI tab with every commercial detail removed. Returns the temp sheet. */
function makePlantCopy_(ss, sheet, pi, code) {
  var tmp = sheet.copyTo(ss).setName('PI_' + new Date().getTime());
  var L = pi.L, disp = pi.disp, vals = pi.vals, lastCol = disp[0].length;
  var stopCol = L.pi.c;                                   // right-hand block (PI no / date) starts here

  // customer name -> code; wipe address, GST no., contact lines below it
  for (var c = L.name.c + 1; c < stopCol; c++) if (norm_(disp[L.name.r][c])) { tmp.getRange(L.name.r + 1, c + 1).setValue(code); break; }
  if (L.sl.r - L.name.r - 1 > 0 && stopCol - L.name.c - 1 > 0)
    tmp.getRange(L.name.r + 2, L.name.c + 2, L.sl.r - L.name.r - 1, stopCol - L.name.c - 1).clearContent();
  // anything else left of the labels' value columns that is not a label (e.g. address typed beside PI no)
  for (var r = L.name.r; r < L.sl.r; r++)
    for (var c3 = stopCol; c3 < lastCol; c3++) {
      var t = norm_(disp[r][c3]);
      if (!t || isLabel_(t)) continue;
      var keep = (r === L.pi.r) || (L.date && r === L.date.r);
      if (!keep) tmp.getRange(r + 1, c3 + 1).clearContent();   // also clears Other Ref / PO no.
    }
  // rates and amounts: every number from the price column rightwards, items through grand total
  for (var r2 = L.sl.r + 2; r2 <= L.grand.r; r2++)
    for (var c4 = L.price.c; c4 < lastCol; c4++)
      if (typeof vals[r2][c4] === 'number' || /[₹]|^\d[\d,\.]*%?$/.test(norm_(disp[r2][c4]))) tmp.getRange(r2 + 1, c4 + 1).clearContent();
  // payment terms
  if (L.pay) {
    var lab = norm_(disp[L.pay.r][L.pay.c]);
    if (/:-?\s*\S/.test(lab)) tmp.getRange(L.pay.r + 1, L.pay.c + 1).setValue('Payment :-');
    for (var c5 = L.pay.c + 1; c5 < lastCol; c5++)
      if (norm_(disp[L.pay.r][c5]) && !isLabel_(disp[L.pay.r][c5])) tmp.getRange(L.pay.r + 1, c5 + 1).clearContent();
  }
  // title
  var title = find_(disp, /^proforma invoice/i);
  if (title) tmp.getRange(title.r + 1, title.c + 1).setValue('PRODUCTION ORDER');
  // working notes typed outside the PI block (right of the amount column, below the signature) must not reach the plant
  var amtCol = L.price.c + 1;
  for (var c6 = L.grand.c + 1; c6 < lastCol; c6++) if (typeof vals[L.grand.r][c6] === 'number') { amtCol = Math.max(amtCol, c6); break; }
  var end = findLast_(disp, /authori[sz]ed signatory|ifsc|bank detail/i);
  var endRow = Math.max(L.grand.r, end ? end.r + 1 : (L.pay ? L.pay.r + 8 : L.grand.r + 10));
  var nR = tmp.getMaxRows(), nC = tmp.getMaxColumns();
  if (nC > amtCol + 1) tmp.getRange(1, amtCol + 2, nR, nC - amtCol - 1).clearContent();
  if (nR > endRow + 1) tmp.getRange(endRow + 2, 1, nR - endRow - 1, nC).clearContent();
  return tmp;
}

/* ---------- register ---------- */

var REG_HEAD = ['PI No', 'Tab', 'Customer', 'Customer ID', 'Status', 'Rev', 'Confirmed on', 'Confirmed by', 'Full PI PDF', 'Production order PDF'];
function register_(ss) {
  var sh = ss.getSheetByName(CFG.REGISTER_TAB);
  if (!sh) { sh = ss.insertSheet(CFG.REGISTER_TAB, 0); sh.appendRow(REG_HEAD); sh.setFrozenRows(1); }
  return sh;
}
function findRow_(sh, col, piNo) {
  var n = sh.getLastRow(); if (n < 2) return 0;
  var v = sh.getRange(2, col, n - 1, 1).getDisplayValues(), k = key_(piNo);
  for (var i = v.length - 1; i >= 0; i--) if (key_(v[i][0]) === k) return i + 2;
  return 0;
}
/** Last row of the real block of data. Stray entries far below it (after 200 empty rows) are ignored, so one stray cell at the bottom of a sheet cannot push new orders thousands of rows down. */
function blockEnd_(vals, start, test) {
  var last = start, gap = 0;
  for (var i = start; i < vals.length; i++) {
    if (test(vals[i])) { last = i + 1; gap = 0; }
    else if (++gap >= 200 && last > start) break;
  }
  return last;
}
function lastDataRow_(sh, col) {
  var v = sh.getRange(1, col, sh.getMaxRows(), 1).getDisplayValues();
  for (var i = v.length - 1; i >= 0; i--) if (norm_(v[i][0])) return i + 1;
  return 1;
}

/* ---------- menu actions ---------- */

function checkThisPI() {
  var pi = readPI_(SpreadsheetApp.getActiveSheet());
  SpreadsheetApp.getUi().alert('PI read OK\n\nPI No: ' + pi.piNo + '\nCustomer: ' + pi.customer + '\nCategory: ' + pi.category +
    '\nQty: ' + pi.qty1 + ' ' + pi.unit1 + ' / ' + pi.qty2 + ' ' + pi.unit2 + '\nBasic value: ' + pi.value + '\nItems: ' + pi.itemCount);
}

function showConfirmDialog() {
  var ss = SpreadsheetApp.getActive(), sheet0 = ss.getActiveSheet(), pi = readPI_(sheet0);
  if (!pi.piNo || !pi.customer) throw new Error('PI number or customer name is blank on this tab.');
  if (new RegExp('^' + CFG.CODE_PREFIX + '\\d+$', 'i').test(pi.customer)) throw new Error('This looks like a (B) plant copy. Run this on the main PI tab that has the customer name - the plant copy is now made automatically.');
  var prev = findRow_(register_(ss), 1, pi.piNo);
  if (ss.getActiveSheet().getSheetId() !== sheet0.getSheetId()) ss.setActiveSheet(sheet0);   // creating the register tab must not change which PI is confirmed
  var esc = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;'); };
  var html = '<style>body{font:13px Arial;margin:12px}label{display:block;margin-top:8px;color:#444}input,select{width:100%;padding:5px;box-sizing:border-box}' +
    '.box{background:#f3f6fb;padding:8px;border-radius:6px;line-height:1.5}button{margin-top:14px;padding:8px 14px}#msg{margin-top:10px;color:#b00}</style>' +
    '<div class="box"><b>' + esc(pi.piNo) + '</b><br>' + esc(pi.customer) + '<br>' + esc(pi.category) + '<br>Qty: ' + pi.qty1 + ' ' + esc(pi.unit1) + ' / ' + pi.qty2 + ' ' + esc(pi.unit2) +
    '<br>Basic value: ' + pi.value + (pi.value ? '' : '<br><b style="color:#b00">Basic value is 0 - rates are missing on this tab. Check before confirming.</b>') + (prev ? '<br><b style="color:#b00">Already confirmed once - this will be saved as a revision.</b>' : '') + ((!PLAN.ENABLED || planPiOk_(pi.piNo)) ? '' : '<br><b style="color:#b00">PI number "' + esc(pi.piNo) + '" is not in the form ABC/PI/26-27/0846 - it will NOT be sent to the planning sheet. Correct it first.</b>') + '</div>' +
    '<label>Sales person</label><input id="sales">' +
    '<label>CRM name</label><input id="crm">' +
    '<label>Sample kits / brochure / pads / pens</label><select id="kit"><option>NO</option><option>YES</option></select>' +
    '<label>Test certificate required</label><select id="tc"><option>NO</option><option>YES</option></select>' +
    '<label>NBD client</label><select id="nbd"><option>NO</option><option>YES</option></select>' +
    '<button id="go" onclick="go()">Party has confirmed - split this PI</button><div id="msg"></div>' +
    '<script>function go(){var g=function(i){return document.getElementById(i).value};var b=document.getElementById("go");' +
    'if(!g("sales")||!g("crm")){document.getElementById("msg").textContent="Sales person and CRM name are required.";return}' +
    'b.disabled=true;document.getElementById("msg").textContent="Working... about 30 seconds";' +
    'google.script.run.withSuccessHandler(function(m){document.body.innerHTML="<div class=box>"+m+"</div>"})' +
    '.withFailureHandler(function(e){b.disabled=false;document.getElementById("msg").textContent=e.message})' +
    '.confirmAndSplit({sheetId:' + sheet0.getSheetId() + ',sales:g("sales"),crm:g("crm"),kit:g("kit"),tc:g("tc"),nbd:g("nbd")})}</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(470), 'Confirm PI and split');
}

function confirmAndSplit(f) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(60000);
  var ss = SpreadsheetApp.getActive(), tmp = null;
  try {
    var sheet = ss.getSheets().filter(function (s) { return s.getSheetId() === f.sheetId; })[0];
    if (!sheet) throw new Error('PI tab not found.');
    var pi = readPI_(sheet);
    var master = SpreadsheetApp.openById(CFG.MASTER_ID), plant = SpreadsheetApp.openById(CFG.PLANT_ID);
    var reg = register_(ss), regRow = findRow_(reg, 1, pi.piNo);
    var rev = regRow ? (parseInt(reg.getRange(regRow, 6).getValue(), 10) || 0) + 1 : 0;
    if (!regRow) {                                                   // confirmed or backfilled earlier from another PI workbook
      var pl0 = plant.getSheetByName(MIRROR.LIST_TAB), plr = pl0 ? findRow_(pl0, 1, pi.piNo) : 0;
      if (plr) rev = (parseInt(pl0.getRange(plr, 2).getValue(), 10) || 0) + 1;
    }
    var cust = customerCode_(master, pi.customer);
    var now = new Date(), tz = ss.getSpreadsheetTimeZone();
    var safe = pi.piNo.replace(/[\/\\]/g, '_').replace(/\s+/g, '');
    var revTag = rev ? ' Rev' + rev : '';
    var dispDate = f.disp ? Utilities.formatDate(new Date(f.disp + 'T00:00:00'), tz, 'dd/MM/yyyy') : '';
    var orderDate = pi.date instanceof Date ? Utilities.formatDate(pi.date, tz, 'dd/MM/yyyy') : pi.date;

    // 1. two PDFs from the same tab
    var fullUrl = exportPdf_(ss, sheet, safe + ' ' + pi.customer + revTag + '.pdf', CFG.FULL_PDF_FOLDER);
    tmp = makePlantCopy_(ss, sheet, pi, cust.code);
    var plantUrl = exportPdf_(ss, tmp, 'PO_' + safe + '_' + cust.code + revTag + '.pdf', CFG.PLANT_PDF_FOLDER);
    ss.deleteSheet(tmp); tmp = null;

    try {
    // 2. main order-to-delivery sheet ("order form" tab layout, see README for the column map)
    var ms = master.getSheetByName(CFG.MASTER_TAB);
    var mRow = findRow_(ms, 2, pi.piNo) || blockEnd_(ms.getRange(1, 2, ms.getMaxRows(), 1).getDisplayValues(), 1, function (r) { return norm_(r[0]); }) + 1;
    if (mRow > ms.getMaxRows()) ms.insertRowsAfter(ms.getMaxRows(), 50);
    ms.getRange(mRow, 1, 1, 7).setValues([[now, pi.piNo, pi.customer, pi.category, pi.qty2 || pi.qty1, pi.value, orderDate]]);   // dispatch date is given by the plant team, not the CRM
    ms.getRange(mRow, 10).setValue(pi.piNo);
    ms.getRange(mRow, 13).setValue(plantUrl);
    ms.getRange(mRow, 16, 1, 6).setValues([[f.kit, f.sales, f.tc, f.crm, f.nbd, fullUrl]]);

    // 3. plant order-to-delivery sheet (no commercial details)
    var ps = plant.getSheetByName(MIRROR.LIST_TAB) || plant.getSheets()[0];
    var pRow = findRow_(ps, 1, pi.piNo) || ps.getLastRow() + 1;
    var qtyTxt = [pi.qty1 ? pi.qty1 + ' ' + pi.unit1.replace(/^in\s+/i, '') : '', pi.qty2 ? pi.qty2 + ' ' + pi.unit2.replace(/^in\s+/i, '') : ''].filter(String).join(' / ');
    ps.getRange(pRow, 1, 1, 12).setValues([[pi.piNo, rev, cust.code, orderDate, pi.category, qtyTxt, dispDate, f.kit, f.tc, plantUrl, now, rev ? 'Revised ' + Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm') : '']]);

    SpreadsheetApp.flush();
    } catch (e) { throw new Error('The two PDFs were created, but the order rows could not be written (' + e.message + '). This PI is NOT marked confirmed. Ask the sheet admin to check edit rights on the main and plant sheets.'); }

    // 4. register
    var line = [pi.piNo, sheet.getName(), pi.customer, cust.code, 'CONFIRMED', rev, now, Session.getActiveUser().getEmail(), fullUrl, plantUrl];
    if (regRow) reg.getRange(regRow, 1, 1, line.length).setValues([line]); else reg.appendRow(line);

    // 5. the plant team's mirror of the FMS tab is refreshed by the 10-minute timer (doing it here made each confirmation take about 3 minutes)

    // 6. optional planning sheet (PLAN.ENABLED): one row per item in its product tab. A failure here does not undo the confirmation.
    var planMsg = '';
    if (PLAN.ENABLED) {
      try { planMsg = sendToPlan_(pi, cust.code, plantUrl); }
      catch (e) { planMsg = '<b style="color:#b00">Planning sheet NOT updated (' + e.message + '). The PI is confirmed. Ask the sheet admin to check edit rights on the planning sheet, then use the menu item "Send this PI to planning sheet".</b>'; }
    }

    return '<b>Done: ' + pi.piNo + (rev ? ' (Rev ' + rev + ')' : '') + '</b><br>Customer ID: ' + cust.code + (cust.isNew ? ' (new code created)' : '') +
      '<br>Main sheet row ' + mRow + ', plant sheet row ' + pRow + '.<br>The plant sheet shows it within 10 minutes.<br>' + planMsg + '<br><a target="_blank" href="' + fullUrl + '">Full PI</a> | <a target="_blank" href="' + plantUrl + '">Production order</a>';
  } finally {
    if (tmp) { try { ss.deleteSheet(tmp); } catch (e) {} }
    lock.releaseLock();
  }
}


/* ---------- production sheet split (switch on with PLAN.ENABLED = true) ----------
 * Runs as the last step of a confirmation. Each item line of the PI goes to its product tab
 * of the planning (production) spreadsheet with status "New", so the machine operators can start
 * from their own sheet and nobody has to distribute the items by hand. Only the customer ID is written:
 * no name, address, GST, rates or terms. If this step fails the confirmation itself still stands.
 *
 * It ships with the tabs of a plastic-sheet plant as the example. The product tabs (PLAN.TABS), the routing
 * rules (routeItem_), the column layouts (planRow_) and the weight formulas (itemKg_) describe
 * one specific planning sheet. Change them to match yours, or leave PLAN.ENABLED = false.
 */
var PLAN = {
  ENABLED: false,                                   // true = send every confirmed PI to the planning sheet
  ID: 'YOUR_PLANNING_SHEET_ID',                     // planning spreadsheet
  UNSORTED: 'Unsorted PI items',
  NEW: 'New', REVISED: 'New - REVISED, check old rows', DUP: 'New - POSSIBLE DUPLICATE',
  TABS: {
    Compact:   { tab: 'Compact',       first: 3, orderCol: 5, statusCol: 25, kind: 'roll' },
    Embossed:  { tab: 'Embossed',      first: 3, orderCol: 5, statusCol: 25, kind: 'roll' },
    Diamond:   { tab: 'Diamond',       first: 3, orderCol: 5, statusCol: 23, kind: 'roll' },
    PMMA:      { tab: 'PMMA Acrylic',  first: 2, orderCol: 5, statusCol: 25, kind: 'roll' },
    PS:        { tab: 'PS Sheet',      first: 2, orderCol: 5, statusCol: 23, kind: 'roll' },
    PP:        { tab: 'PP',            first: 2, orderCol: 5, statusCol: 23, kind: 'roll' },
    Multiwall: { tab: 'PC Multiwall',  first: 2, orderCol: 5, statusCol: 24, kind: 'mw' },
    Profiling: { tab: 'Profiling',     first: 2, orderCol: 4, statusCol: 17, kind: 'prof' },
    BasePlate: { tab: 'PC Base Plate', first: 6, orderCol: 4, statusCol: 15, kind: 'base' }
  }
};

function colIn_(row, re, from) { for (var c = from || 0; c < row.length; c++) if (re.test(norm_(row[c]))) return c; return -1; }

/** Item lines of the PI, read by the column labels (layouts differ per product). */
function readItems_(pi) {
  var L = pi.L, d = pi.disp, v = pi.vals, h = d[L.sl.r], s = d[L.sl.r + 1], from = L.sl.c + 1;
  var C = {
    color: colIn_(h, /^colou?r$/i, from), mask: colIn_(h, /^masking/i, from),
    width: colIn_(s, /^width$/i, from), prof: colIn_(s, /^profile/i, from), len: colIn_(s, /^length/i, from),
    thk: colIn_(s, /^thick/i, from), den: colIn_(s, /^density/i, from),
    stick: colIn_(h, /packing sticker/i, from), uv: colIn_(h, /^uv/i, from), per: colIn_(h, /^per (roll|sheet)/i, from),
    nos: L.qty.c, q2: L.qty.c + 1 < L.price.c ? L.qty.c + 1 : -1
  };
  var get = function (r, c) { if (c < 0) return ''; var x = v[r][c]; return typeof x === 'number' ? Math.round(x * 10000) / 10000 : norm_(d[r][c]); };
  var items = [];
  for (var r = L.sl.r + 2; r < L.tot.r; r++) {
    var m = norm_(d[r][from]); if (!m) continue;
    items.push({ mat: m, color: get(r, C.color), mask: get(r, C.mask), width: get(r, C.width), prof: get(r, C.prof), len: get(r, C.len),
      thk: get(r, C.thk), den: get(r, C.den), stick: get(r, C.stick), uv: get(r, C.uv), per: get(r, C.per), nos: get(r, C.nos), q2: get(r, C.q2) });
  }
  return items;
}

/** Which product tab an item belongs to. feed = the sheet that has to be extruded first for a profiled item. */
function routeItem_(mat) {
  var m = norm_(mat).toLowerCase();
  if (/base\s*plate|louver|louvre/.test(m)) return { key: 'BasePlate' };
  if (/profil/.test(m)) return { key: 'Profiling', feed: /compact|solid/.test(m) ? 'Compact' : 'Embossed' };
  if (/multi\s*wall|honey\s*comb/.test(m)) return { key: 'Multiwall' };
  if (/diamond/.test(m)) return { key: 'Diamond' };
  if (/emboss/.test(m)) return { key: 'Embossed' };
  if (/acrylic|pmma/.test(m)) return { key: 'PMMA' };
  if (/compact|solid/.test(m)) return { key: 'Compact' };
  if (/(^|[^a-z])(gpps|ps)([^a-z]|$)|polystyrene/.test(m)) return { key: 'PS' };
  if (/(^|[^a-z])pp([^a-z]|$)|polyprop/.test(m)) return { key: 'PP' };
  return { key: '' };
}

function n_(x) { return typeof x === 'number' ? x : 0; }
function itemKg_(it, kind) {
  if (kind === 'base') return n_(it.per) * n_(it.nos);
  if (kind === 'mw') return n_(it.width) / 1000 * n_(it.len) / 1000 * n_(it.den) * n_(it.nos);
  return n_(it.width) / 1000 * n_(it.len) / 1000 * n_(it.thk) * n_(it.den) * n_(it.nos);
}

/** One row for a product tab, laid out exactly as the plant fills it by hand (same weight / kg formulas). */
function planRow_(kind, r, first, it, piNo, code, orderDate, matText) {
  var sr = r - 1 < first ? 1 : '=A' + (r - 1) + '+1', R = String(r);
  if (kind === 'roll')
    return [sr, '=D' + R, '=E' + R, orderDate, piNo, code, matText, it.color, it.mask, it.width, it.len, it.thk, it.den, it.stick,
      '=SUM(J' + R + '/1000*K' + R + '/1000)*L' + R + '*M' + R, it.nos, '', '=P' + R + '+Q' + R, '=R' + R + '*O' + R, '', '=R' + R + '-T' + R, '=U' + R + '*O' + R];
  if (kind === 'mw')
    return [sr, '=D' + R, '=E' + R, orderDate, piNo, code, matText, it.color, it.mask, it.uv, it.width, it.len, it.thk, it.den, it.stick,
      '=SUM(K' + R + '/1000*L' + R + '/1000)*N' + R, it.nos, '', '=Q' + R + '+R' + R, '=S' + R + '*P' + R, '', '=S' + R + '-U' + R, '=V' + R + '*P' + R];
  if (kind === 'prof') {
    var pw = it.prof !== '' ? it.prof : ((/\(([^)]+)\)/.exec(it.mat) || [])[1] || '');
    return [sr, orderDate, code, piNo, 'Polycarbonate Sheet', /compact|solid/i.test(it.mat) ? 'Compact' : 'Embossed', it.color, it.mask, it.width, pw, it.len, it.thk, it.den,
      '=SUM(I' + R + '/1000*K' + R + '/1000)*L' + R + '*M' + R, it.nos, '=O' + R + '*N' + R];
  }
  // base plate: Sr, Date, Order No (=Order Name), Order Name, Mould Design, Size, Thickness, Weight, In Nos, In KGS, In Stock, ready, Balance
  var size = [it.width, it.len, it.den].filter(function (x) { return typeof x === 'string' && x; })[0] || '';
  return [sr, orderDate, '=D' + R, piNo, it.color, size, it.thk, it.per, it.nos, '=I' + R + '*H' + R, '', '', '=I' + R + '-K' + R];
}

function planDate_(d, tz) {
  if (d instanceof Date) return Utilities.formatDate(d, tz, 'dd/MM/yyyy');
  var m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/.exec(norm_(d));
  return m ? ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3] : norm_(d);
}

/** Works out where every item goes. No changes are made here. */
function planSplit_(pi) {
  var items = readItems_(pi), jobs = [], unsorted = [], kg = 0;
  items.forEach(function (it) {
    var rt = routeItem_(it.mat);
    if (!rt.key) { unsorted.push(it); return; }
    var def = PLAN.TABS[rt.key];
    jobs.push({ key: rt.key, it: it, mat: it.mat });
    kg += itemKg_(it, def.kind);
    if (rt.feed) jobs.push({ key: rt.feed, it: it, mat: (rt.feed === 'Compact' ? 'PC Compact Sheet' : 'PC Embossed') + ' (for profiling)' });
  });
  return { items: items, jobs: jobs, unsorted: unsorted, kg: Math.round(kg * 10) / 10 };
}

/** A real PI number (not a quotation "QI", not a short / mistyped serial). Only these go to the planning sheet. */
function planPiOk_(piNo) { var s = norm_(piNo).replace(/\s+/g, ''); return /\/PI\//i.test(s) && /\d{2}-\d{2}\/\d{4,}$/.test(s); }
function samePi_(cell, k, q) {
  var kc = key_(cell); if (!kc) return false; if (kc === k) return true;
  var qc = /PI/i.test(cell) ? piSeq_(cell) : null; return !!(qc && q && qc.fy === q.fy && qc.n === q.n);
}
function planSig_(a) { return a.map(function (x) { return typeof x === 'number' ? Math.round(x * 100) / 100 : (isNaN(parseFloat(x)) ? key_(x) : Math.round(parseFloat(String(x).replace(/,/g, '')) * 100) / 100); }).join('|'); }
var PLAN_SIG = { roll: [10, 11, 12, 16], mw: [11, 12, 13, 17], prof: [9, 11, 12, 15], base: [5, 7, 9] };   // size + quantity columns per tab layout
var PLAN_CODE = { roll: 6, mw: 6, prof: 3, base: 0 };                                                      // customer code column (0 = none)
function jobSig_(kind, it) { return planSig_(kind === 'base' ? [it.color, it.thk, it.nos] : [it.width, it.len, it.thk, it.nos]); }

/** Rows of one PI on one tab. mine = rows this script wrote that the plant has not started on (status still "New..."): only these may be removed. */
function planRowsOf_(sh, def, piNo) {
  var oc = sh.getRange(1, def.orderCol, sh.getMaxRows(), 1).getDisplayValues();
  var last = blockEnd_(oc, def.first - 1, function (x) { return norm_(x[0]); });
  var k = key_(piNo), q = /PI/i.test(piNo) ? piSeq_(piNo) : null, hits = [];
  for (var i = def.first - 1; i < last; i++) if (samePi_(oc[i][0], k, q)) hits.push(i + 1);
  var out = { last: last, mine: [], kept: {}, nKept: 0 };
  if (!hits.length) return out;
  var W = Math.max(def.statusCol, 17), top = hits[0], ex = sh.getRange(top, 1, hits[hits.length - 1] - top + 1, W).getValues(), cc = PLAN_CODE[def.kind];
  var isCode = new RegExp('^' + CFG.CODE_PREFIX + '\\d+$', 'i');
  hits.forEach(function (rr) {
    var row = ex[rr - top], st = norm_(row[def.statusCol - 1]);
    if (/^new/i.test(st) && (!cc || isCode.test(norm_(row[cc - 1])))) out.mine.push(rr);
    else { out.nKept++; var g = planSig_(PLAN_SIG[def.kind].map(function (c) { return row[c - 1]; })); out.kept[g] = (out.kept[g] || 0) + 1; }
  });
  return out;
}
/** Deletes rows (bottom up) and repairs the Sr No formula of the row that moved up into the gap. */
function planDelete_(sh, def, rows, last) {
  if (!rows.length) return last;
  rows.slice().sort(function (a, b) { return b - a; }).forEach(function (rr) { sh.deleteRow(rr); });
  last -= rows.length;
  var from = Math.min.apply(null, rows);
  if (last >= from) {
    var f = sh.getRange(from, 1, last - from + 1, 1).getFormulas();
    for (var i = 0; i < f.length; i++) if (/#REF!/.test(f[i][0])) { var r = from + i; if (r - 1 < def.first) sh.getRange(r, 1).setValue(1); else sh.getRange(r, 1).setFormula('=A' + (r - 1) + '+1'); }
  }
  return last;
}

function sendToPlan_(pi, code, plantUrl) {
  if (!planPiOk_(pi.piNo)) return '<b style="color:#b00">Planning sheet NOT updated: "' + pi.piNo + '" is not a proper PI number (it must look like ABC/PI/26-27/0846). Correct the PI number on the tab and confirm again.</b>';
  var plan = SpreadsheetApp.openById(PLAN.ID), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var orderDate = planDate_(pi.date, tz), sp = planSplit_(pi), counts = {}, notes = [], removed = 0, dupAll = {};
  if (!sp.items.length) return 'Planning sheet: no item lines found on this PI.';

  Object.keys(PLAN.TABS).forEach(function (key) {
    var def = PLAN.TABS[key], mine = sp.jobs.filter(function (j) { return j.key === key; }), sh = plan.getSheetByName(def.tab);
    if (!sh) { if (mine.length) throw new Error('tab "' + def.tab + '" not found in the planning sheet'); return; }
    // 1. one PI = one set of rows: rows this script wrote earlier for this PI, and the plant has not started on, are replaced
    var cur = planRowsOf_(sh, def, pi.piNo), last = cur.last;
    if (cur.mine.length) { last = planDelete_(sh, def, cur.mine, last); removed += cur.mine.length; }
    if (!mine.length) return;
    // 2. items the plant already has for this PI (typed by hand, or already Pending / Running / Complete) are not added again
    var todo = mine.filter(function (j) { var g = jobSig_(def.kind, j.it); if (cur.kept[g] > 0) { cur.kept[g]--; return false; } return true; });
    if (!todo.length) { notes.push(def.tab + ': the plant already has all ' + mine.length + ' item(s) of this PI - nothing added'); return; }
    // 3. same customer, same items, under another PI number = possible duplicate order
    var dupOf = '', cc = PLAN_CODE[def.kind];
    if (cc && !cur.nKept && last >= def.first) {
      var from = Math.max(def.first, last - 399), blk = sh.getRange(from, 1, last - from + 1, Math.max(def.statusCol, 17)).getValues(), groups = {}, names = {};
      var k = key_(pi.piNo), q = piSeq_(pi.piNo);
      blk.forEach(function (row) {
        var o = norm_(row[def.orderCol - 1]); if (!o || samePi_(o, k, q) || key_(row[cc - 1]) !== key_(code)) return;
        if (/^(canc|cancle|cancel)/i.test(norm_(row[def.statusCol - 1]))) return;
        (groups[key_(o)] = groups[key_(o)] || []).push(planSig_(PLAN_SIG[def.kind].map(function (c) { return row[c - 1]; }))); names[key_(o)] = o;
      });
      var want = todo.map(function (j) { return jobSig_(def.kind, j.it); }).sort().join('~');
      Object.keys(groups).forEach(function (g) { if (!dupOf && groups[g].sort().join('~') === want) dupOf = names[g]; });
    }
    if (dupOf) dupAll[dupOf] = 1;
    var status = cur.nKept ? PLAN.REVISED : (dupOf ? PLAN.DUP : PLAN.NEW);
    if (last + todo.length > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), todo.length + 50);
    var rows = todo.map(function (j, n) {                         // status goes in the same write, so a rejected status can never leave half-written rows
      var row = planRow_(def.kind, last + 1 + n, def.first, j.it, pi.piNo, code, orderDate, j.mat);
      while (row.length < def.statusCol - 1) row.push('');
      row.push(status);
      return row;
    });
    try { sh.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows); }
    catch (e) { throw new Error(def.tab + ' rows not written: ' + e.message); }
    if (dupOf) sh.getRange(last + 1, def.statusCol, rows.length, 1).setNote('Same customer and same items as ' + dupOf + '. Check with the CRM before planning. If ' + dupOf + ' was the wrong one, remove it with the menu "Remove a PI from the planning sheet".');
    counts[def.tab] = rows.length;
    if (cur.nKept) notes.push(def.tab + ': the plant already has other rows for this PI - the ' + rows.length + ' changed item(s) were added and marked "' + PLAN.REVISED + '"');
  });

  // items that match no product tab are kept, never dropped (earlier unsorted lines of this PI are replaced)
  var us = plan.getSheetByName(PLAN.UNSORTED);
  if (us && us.getLastRow() > 1) {
    var uo = us.getRange(2, 2, us.getLastRow() - 1, 1).getDisplayValues(), uk = key_(pi.piNo);
    for (var u = uo.length - 1; u >= 0; u--) if (key_(uo[u][0]) === uk) us.deleteRow(u + 2);
  }
  if (sp.unsorted.length) {
    if (!us) { us = plan.insertSheet(PLAN.UNSORTED); us.appendRow(['Added on', 'Order No', 'Customer Code', 'Order Date', 'Material', 'Color', 'Masking', 'Width', 'Length', 'Thickness', 'Density', 'Sticker thickness', 'Qty (nos)', 'Qty 2', 'Moved to tab']); us.setFrozenRows(1); }
    var now = new Date(), ur = sp.unsorted.map(function (it) { return [now, pi.piNo, code, orderDate, it.mat, it.color, it.mask, it.width, it.len, it.thk, it.den, it.stick, it.nos, it.q2, '']; });
    us.getRange(us.getLastRow() + 1, 1, ur.length, ur[0].length).setValues(ur);
    counts[PLAN.UNSORTED] = ur.length;
  }

  // Summary is NOT written: its order columns are a live import from the FMS tab, and any typed value there breaks the import.
  SpreadsheetApp.flush();
  var parts = Object.keys(counts).map(function (t) { return counts[t] + ' in ' + t; }), dups = Object.keys(dupAll);
  return 'Planning sheet: ' + (parts.length ? parts.join(', ') : 'no new rows') + (removed ? ' (replaced ' + removed + ' earlier row(s) of this PI)' : '') + '.' +
    (dups.length ? '<br><b style="color:#b00">POSSIBLE DUPLICATE: the same customer already has the same items under ' + dups.join(', ') + '. The rows are marked for the plant to hold. If that PI was wrong, remove it with the menu "Remove a PI from the planning sheet".</b>' : '') +
    (notes.length ? '<br>' + notes.join('<br>') : '') +
    (counts[PLAN.UNSORTED] ? '<br><b style="color:#b00">' + counts[PLAN.UNSORTED] + ' item(s) matched no product tab - see "' + PLAN.UNSORTED + '".</b>' : '');
}

/** Takes a PI out of the planning sheet: rows the plant has not started on are deleted, rows already in work are only counted. */
function planRemove_(piNo) {
  var plan = SpreadsheetApp.openById(PLAN.ID), gone = [], kept = [];
  Object.keys(PLAN.TABS).forEach(function (key) {
    var def = PLAN.TABS[key], sh = plan.getSheetByName(def.tab); if (!sh) return;
    var cur = planRowsOf_(sh, def, piNo);
    if (cur.mine.length) { planDelete_(sh, def, cur.mine, cur.last); gone.push(cur.mine.length + ' from ' + def.tab); }
    if (cur.nKept) kept.push(cur.nKept + ' in ' + def.tab);
  });
  var us = plan.getSheetByName(PLAN.UNSORTED);
  if (us && us.getLastRow() > 1) {
    var uo = us.getRange(2, 2, us.getLastRow() - 1, 1).getDisplayValues(), uk = key_(piNo), n = 0;
    for (var u = uo.length - 1; u >= 0; u--) if (key_(uo[u][0]) === uk) { us.deleteRow(u + 2); n++; }
    if (n) gone.push(n + ' from ' + PLAN.UNSORTED);
  }
  SpreadsheetApp.flush();
  return 'Removed for ' + piNo + ': ' + (gone.length ? gone.join(', ') : 'nothing') + '.' +
    (kept.length ? '\n\nNOT removed (typed by the plant or already Pending / Running / Complete): ' + kept.join(', ') + '.\nTell the plant team directly so they stop or cancel these.' : '');
}

/** Menu: remove a wrong / cancelled / duplicate PI from the planning sheet. */
function removePIFromPlan() {
  var ui = SpreadsheetApp.getUi(), def = '';
  try { def = readPI_(SpreadsheetApp.getActiveSheet()).piNo; } catch (e) {}
  var a = ui.prompt('Remove a PI from the planning sheet', 'Type the PI number exactly as it was confirmed' + (def ? ' (this tab is ' + def + ')' : '') + ':', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var piNo = norm_(a.getResponseText()) || def; if (!piNo) return;
  if (ui.alert('Remove ' + piNo + ' from the planning sheet?', 'Rows the plant has not started on are deleted. This does not change the O2D sheets.', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var lock = LockService.getDocumentLock(); lock.waitLock(60000);
  try { ui.alert(planRemove_(piNo)); } finally { lock.releaseLock(); }
}

/** Menu: shows where each item of this PI would go. Changes nothing. */
function checkPlanSplit() {
  var pi = readPI_(SpreadsheetApp.getActiveSheet()), sp = planSplit_(pi), lines = [];
  sp.jobs.forEach(function (j) { lines.push(PLAN.TABS[j.key].tab + '  <-  ' + j.mat + ' | ' + j.it.width + ' x ' + j.it.len + ' x ' + j.it.thk + ' | ' + j.it.nos + ' nos'); });
  sp.unsorted.forEach(function (it) { lines.push(PLAN.UNSORTED + '  <-  ' + it.mat + ' | ' + it.nos + ' nos'); });
  SpreadsheetApp.getUi().alert('Planning split for ' + pi.piNo + ' (nothing changed)\n\n' + (lines.join('\n') || 'No item lines found.') + '\n\nApprox. total: ' + sp.kg + ' kg');
}

/** Menu: sends an already confirmed PI to the planning sheet, for use when that step failed during confirmation. */
function sendThisPIToPlan() {
  var ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive(), pi = readPI_(ss.getActiveSheet());
  if (new RegExp('^' + CFG.CODE_PREFIX + '\\d+$', 'i').test(pi.customer)) throw new Error('This looks like a (B) plant copy. Run this on the main PI tab.');
  var reg = register_(ss), row = findRow_(reg, 1, pi.piNo);
  if (!row) throw new Error('This PI is not confirmed yet. Use "Confirm this PI" first.');
  if (ui.alert('Send ' + pi.piNo + ' to the planning sheet?', 'Use this only if the planning step failed when the PI was confirmed. Rows of this PI that the plant has not started on are replaced, not added twice.', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var lock = LockService.getDocumentLock(); lock.waitLock(60000);
  try {
    var r = reg.getRange(row, 1, 1, 10).getDisplayValues()[0];
    var code = norm_(r[3]) || customerCode_(SpreadsheetApp.openById(CFG.MASTER_ID), pi.customer).code;
    ui.alert(sendToPlan_(pi, code, norm_(r[9])).replace(/<br>/g, '\n').replace(/<[^>]+>/g, ''));
  } finally { lock.releaseLock(); }
}


/* ---------- test helpers (safe to delete) ----------
 * Run these against TEST copies of the main and plant sheets first.
 */

/** Read-only: checks that every PI tab in this workbook can be read. PI tabs are recognised by "_PI_" in the tab name; "(B)" tabs are skipped. */
function auditTabs() {
  var ss = SpreadsheetApp.getActive(), ok = 0, bad = [], skipped = 0, sample = [];
  ss.getSheets().forEach(function (sh) {
    var n = sh.getName();
    if (!/_PI_/i.test(n) || /\(B\)/i.test(n)) { skipped++; return; }
    try {
      var p = readPI_(sh);
      if (!p.piNo || !p.customer || !p.itemCount || !(p.qty1 || p.qty2) || !p.value) bad.push(n.trim() + ' => blank: ' + JSON.stringify({pi: p.piNo, c: p.customer, items: p.itemCount, q1: p.qty1, q2: p.qty2, v: p.value}));
      else { ok++; if (ok % 25 === 1) sample.push(p.piNo + ' | ' + p.customer + ' | ' + p.category.slice(0, 40) + ' | ' + p.qty1 + ' ' + p.unit1 + ' / ' + p.qty2 + ' ' + p.unit2 + ' | ' + p.value + ' | ' + (p.date instanceof Date ? 'DATE' : p.date)); }
    } catch (e) { bad.push(n.trim() + ' => ' + e.message); }
  });
  Logger.log('OK=' + ok + ' BAD=' + bad.length + ' SKIPPED=' + skipped);
  Logger.log('BAD:\n' + bad.join('\n'));
  Logger.log('SAMPLE:\n' + sample.join('\n'));
}
/** Read-only: builds plant copies of a sample of PI tabs (about 30) and scans them for leftover names, GSTINs, amounts, terms. Temp tabs are deleted. */
function leakTest() {
  var ss = SpreadsheetApp.getActive(), tabs = ss.getSheets().filter(function (s) { return /_PI_/i.test(s.getName()) && !/\(B\)/i.test(s.getName()); });
  var step = Math.max(1, Math.floor(tabs.length / 30)), res = [], tested = 0;
  for (var i = 0; i < tabs.length; i += step) {
    var sh = tabs[i], tmp = null;
    try {
      var p = readPI_(sh); tmp = makePlantCopy_(ss, sh, p, CFG.CODE_PREFIX + '999'); SpreadsheetApp.flush();
      var d = tmp.getDataRange().getDisplayValues(), v = tmp.getDataRange().getValues(), leaks = [], ck = key_(p.customer);
      for (var r = 0; r < d.length; r++) for (var c = 0; c < d[r].length; c++) {
        var t = norm_(d[r][c]); if (!t) continue;
        if (ck.length > 3 && key_(t).indexOf(ck) >= 0) leaks.push('NAME@' + (r + 1) + ',' + (c + 1));
        if (foreignGst_(t)) leaks.push('GST@' + (r + 1) + ',' + (c + 1));
        if (r > p.L.sl.r + 1 && r <= p.L.grand.r && c >= p.L.price.c && typeof v[r][c] === 'number' && v[r][c] !== 0) leaks.push('NUM@' + (r + 1) + ',' + (c + 1) + '=' + t);
        if (r > p.L.name.r && r < p.L.sl.r && c > p.L.name.c && !isLabel_(t) && !(r === p.L.pi.r || (p.L.date && r === p.L.date.r))) leaks.push('HDRTEXT@' + (r + 1) + ',' + (c + 1) + '=' + t.slice(0, 25));
        if (p.L.pay && r === p.L.pay.r && c > p.L.pay.c && !isLabel_(t)) leaks.push('PAY@' + (r + 1) + ',' + (c + 1) + '=' + t.slice(0, 25));
        if (/days|advance/i.test(t) && /term/i.test(t)) leaks.push('TERMS@' + (r + 1) + ',' + (c + 1) + '=' + t.slice(0, 30));
      }
      tested++; if (leaks.length) res.push(sh.getName().trim() + ' => ' + leaks.slice(0, 8).join(' ; '));
    } catch (e) { res.push(sh.getName().trim() + ' => ERR ' + e.message); }
    finally { if (tmp) ss.deleteSheet(tmp); }
  }
  Logger.log('TESTED=' + tested + ' WITH_LEAKS_OR_ERR=' + res.length);
  Logger.log(res.join('\n'));
}
/** WRITES REAL ROWS AND PDFs: runs the full confirmation on the active tab. Use only with test copies of the main and plant sheets. */
function testActiveTab() {
  var sh = SpreadsheetApp.getActiveSheet();
  Logger.log(confirmAndSplit({ sheetId: sh.getSheetId(), sales: 'TEST', crm: 'TEST', kit: 'NO', tc: 'NO', nbd: 'NO' }));
}

/* ---------- plant mirror of the FMS tab ----------
 * The plant team gets a view-only copy of the FMS tab, same layout and colours,
 * with the party name replaced by the customer ID and the PI link replaced by
 * the production order link. Values only: the plant sheet has no connection to
 * the main sheet. setupPlantMirror() is run once; syncPlantMirror() refreshes it.
 *
 * Column settings (1-based): PI_COL = column holding the PI number, NAME_COL = party name,
 * LINK_COL = PI link. HEAD_ROWS = number of header rows at the top of the FMS tab.
 */
var MIRROR = { SRC_TAB: 'FMS', TAB: 'FMS', LIST_TAB: 'PI list', HEAD_ROWS: 8, PI_COL: 3, NAME_COL: 4, LINK_COL: 11, SPARE_ROWS: 300 };

function mirrorLastRow_(src) {
  var scan = src.getRange(1, 2, src.getMaxRows(), 2).getDisplayValues();
  return blockEnd_(scan, MIRROR.HEAD_ROWS, function (r) { return norm_(r[0]) && norm_(r[1]); });
}

function setupPlantMirror() {
  var master = SpreadsheetApp.openById(CFG.MASTER_ID), plant = SpreadsheetApp.openById(CFG.PLANT_ID);
  var src = master.getSheetByName(MIRROR.SRC_TAB);
  if (!src) throw new Error('Tab "' + MIRROR.SRC_TAB + '" not found in the main sheet.');
  if (!plant.getSheetByName(MIRROR.LIST_TAB)) {
    var first = plant.getSheets()[0];
    if (first.getName() === MIRROR.TAB) throw new Error('First tab of the plant sheet is already called ' + MIRROR.TAB + '.');
    first.setName(MIRROR.LIST_TAB);
  }
  var old = plant.getSheetByName(MIRROR.TAB);
  if (old) plant.deleteSheet(old);
  plant.setSpreadsheetTimeZone(master.getSpreadsheetTimeZone());
  plant.setSpreadsheetLocale(master.getSpreadsheetLocale());
  var m = src.copyTo(plant);
  m.clearContents();                                       // formats, widths, merges and colours stay; no data
  m.clearNotes();
  m.getRange(1, 1, m.getMaxRows(), m.getMaxColumns()).clearDataValidations();
  if (m.getFilter()) m.getFilter().remove();
  m.showRows(1, m.getMaxRows());
  var keep = mirrorLastRow_(src) + MIRROR.SPARE_ROWS;
  if (m.getMaxRows() > keep) m.deleteRows(keep + 1, m.getMaxRows() - keep);
  m.setName(MIRROR.TAB);
  m.protect().setDescription('Auto-generated mirror of the FMS tab - do not edit, it is overwritten').setWarningOnly(true);
  plant.setActiveSheet(m); plant.moveActiveSheet(1);
  plant.getSheetByName(MIRROR.LIST_TAB).hideSheet();
  SpreadsheetApp.flush();
  return syncPlantMirror();
}

function syncPlantMirror() {
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(30000)) return 'busy';
  try { var msg = syncPlantMirror_(); Logger.log(msg); return msg; } finally { lock.releaseLock(); }
}

/** Text that Sheets would re-read as a formula, number or date is written as plain text, so it shows exactly as in the FMS. */
function keepText_(x) {
  if (typeof x !== 'string' || !x) return x;
  return (/^[=+]/.test(x) || (/^[\d.,\/:\- ]+$/.test(x) && /\d/.test(x))) ? "'" + x : x;
}

function syncPlantMirror_() {
  var master = SpreadsheetApp.openById(CFG.MASTER_ID), plant = SpreadsheetApp.openById(CFG.PLANT_ID);
  var src = master.getSheetByName(MIRROR.SRC_TAB), dst = plant.getSheetByName(MIRROR.TAB);
  if (!src) throw new Error('Tab "' + MIRROR.SRC_TAB + '" not found in the main sheet.');
  if (!dst) throw new Error('Plant mirror tab not found - run setupPlantMirror once.');
  var H = MIRROR.HEAD_ROWS, last = mirrorLastRow_(src), nC = Math.min(src.getLastColumn(), dst.getMaxColumns());
  var rng = src.getRange(1, 1, last, nC), out = rng.getValues(), forms = rng.getFormulas(), disp = rng.getDisplayValues();

  // customer name -> customer ID (same table the confirmation uses; missing names get the next code)
  var cs = master.getSheetByName(CFG.CODES_TAB);
  if (!cs) throw new Error('Customer code tab "' + CFG.CODES_TAB + '" not found in the main sheet.');
  var cd = cs.getRange(1, 1, Math.max(cs.getLastRow(), 1), 2).getDisplayValues(), codes = {}, maxN = 0, fresh = [];
  for (var i = 1; i < cd.length; i++) {
    var ck = key_(cd[i][0]), cc = norm_(cd[i][1]);
    if (ck && cc && !codes[ck]) codes[ck] = cc;
    var mn = /(\d+)\s*$/.exec(cd[i][1]); if (mn) maxN = Math.max(maxN, parseInt(mn[1], 10));
  }
  // PI no -> production order PDF (only PIs confirmed through this workbook have one)
  // (read from the plant sheet's hidden "PI list" tab, so it works from whichever PI workbook runs this)
  var po = {}, pl = plant.getSheetByName(MIRROR.LIST_TAB);
  if (pl && pl.getLastRow() > 1) pl.getRange(2, 1, pl.getLastRow() - 1, 10).getDisplayValues().forEach(function (r) { if (norm_(r[9])) po[key_(r[0])] = norm_(r[9]); });

  var withPo = 0, rows = 0;
  for (var r = H; r < last; r++) {
    for (var c = 0; c < nC; c++) {
      var f = forms[r][c];
      if (f && /HYPERLINK\(/i.test(f) && disp[r][c]) {                 // keep working form links
        var hm = /HYPERLINK\(\s*"([^"]+)"\s*,\s*"([^"]*)"\s*\)/i.exec(f);
        if (hm && hm[2] === disp[r][c]) { out[r][c] = '=HYPERLINK("' + hm[1] + '","' + hm[2] + '")'; continue; }
      }
      out[r][c] = keepText_(out[r][c]);
    }
    var name = norm_(disp[r][MIRROR.NAME_COL - 1]), nk = key_(name), code = '';
    if (nk) { code = codes[nk]; if (!code) { code = CFG.CODE_PREFIX + (++maxN); codes[nk] = code; fresh.push([name, code]); } }
    out[r][MIRROR.NAME_COL - 1] = code;
    var pk = key_(disp[r][MIRROR.PI_COL - 1]);
    if (pk) rows++;
    if (pk && po[pk]) { out[r][MIRROR.LINK_COL - 1] = '=HYPERLINK("' + po[pk] + '","Production order")'; withPo++; }
    else out[r][MIRROR.LINK_COL - 1] = '';
  }
  for (var hr = 0; hr < H; hr++) for (var hc = 0; hc < nC; hc++) out[hr][hc] = keepText_(out[hr][hc]);
  out[H - 1][MIRROR.NAME_COL - 1] = 'Customer ID';
  out[H - 1][MIRROR.LINK_COL - 1] = 'Production order';

  // save new customer IDs; if the code tab cannot be written, never fall back to the name
  var pending = 0;
  if (fresh.length) {
    try { cs.getRange(cs.getLastRow() + 1, 1, fresh.length, 2).setValues(fresh); }
    catch (e) {
      var fk = {}; fresh.forEach(function (x) { fk[x[1]] = 1; });
      for (var r3 = H; r3 < last; r3++) if (fk[out[r3][MIRROR.NAME_COL - 1]]) { out[r3][MIRROR.NAME_COL - 1] = 'ID pending'; pending++; }
      fresh = [];
    }
  }

  if (dst.getMaxRows() < last + 50) dst.insertRowsAfter(dst.getMaxRows(), last + MIRROR.SPARE_ROWS - dst.getMaxRows());
  dst.getRange(1, 1, last, nC).setValues(out);
  if (dst.getMaxRows() > last) dst.getRange(last + 1, 1, dst.getMaxRows() - last, dst.getMaxColumns()).clearContent();
  if (dst.getMaxRows() > last + MIRROR.SPARE_ROWS + 500) dst.deleteRows(last + MIRROR.SPARE_ROWS + 1, dst.getMaxRows() - last - MIRROR.SPARE_ROWS);
  // header cells that carry a link (form links for the doers)
  var rt = src.getRange(1, 1, H, nC).getRichTextValues();
  for (var a = 0; a < H; a++) for (var b = 0; b < nC; b++) {
    if (!rt[a][b]) continue;
    if (rt[a][b].getRuns().some(function (x) { return x.getLinkUrl(); })) dst.getRange(a + 1, b + 1).setRichTextValue(rt[a][b]);
  }
  var stamp = Utilities.formatDate(new Date(), plant.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm');
  dst.getRange(1, 1).setNote('Plant view of the FMS sheet. Last refreshed ' + stamp + '. Read only - changes made here are overwritten.');
  SpreadsheetApp.flush();
  return 'Plant mirror refreshed ' + stamp + ': ' + rows + ' orders, ' + withPo + ' with a production order link, ' + fresh.length + ' new customer IDs' + (pending ? ', ' + pending + ' rows shown as ID pending because the ' + CFG.CODES_TAB + ' tab could not be written.' : '.');
}

/** Read-only check: looks for any party name or full-PI link inside the plant mirror. */
function mirrorLeakTest() {
  var master = SpreadsheetApp.openById(CFG.MASTER_ID), plant = SpreadsheetApp.openById(CFG.PLANT_ID);
  var src = master.getSheetByName(MIRROR.SRC_TAB), dst = plant.getSheetByName(MIRROR.TAB), last = mirrorLastRow_(src);
  var names = {}, links = {}, sd = src.getRange(1, 1, last, MIRROR.LINK_COL).getDisplayValues();
  var srt = src.getRange(1, MIRROR.LINK_COL, last, 1).getRichTextValues();
  for (var r = MIRROR.HEAD_ROWS; r < last; r++) {
    var k = key_(sd[r][MIRROR.NAME_COL - 1]); if (k.length > 3) names[k] = sd[r][MIRROR.NAME_COL - 1];
    var id = /id=([\w-]{20,})|\/d\/([\w-]{20,})/.exec(sd[r][MIRROR.LINK_COL - 1] + ' ' + ((srt[r][0] && srt[r][0].getLinkUrl()) || '')); if (id) links[id[1] || id[2]] = 1;
  }
  var dr = dst.getDataRange(), d = dr.getDisplayValues(), f = dr.getFormulas(), nk = Object.keys(names), hits = [], cells = 0;
  var col = function (i) { var s = ''; i++; while (i > 0) { var q = (i - 1) % 26; s = String.fromCharCode(65 + q) + s; i = Math.floor((i - 1) / 26); } return s; };
  for (var r2 = 0; r2 < d.length; r2++) for (var c = 0; c < d[r2].length; c++) {
    var t = d[r2][c], ff = f[r2][c]; if (!t && !ff) continue; cells++;
    var tk = key_(t);
    if (tk.length > 3 && /[A-Z]/.test(tk)) for (var n = 0; n < nk.length; n++) if (tk.indexOf(nk[n]) >= 0) { hits.push('NAME ' + col(c) + (r2 + 1) + ' = ' + String(t).slice(0, 40)); break; }
    var lid = /id=([\w-]{20,})|\/d\/([\w-]{20,})/.exec(t + ' ' + ff); if (lid && links[lid[1] || lid[2]]) hits.push('PI LINK ' + col(c) + (r2 + 1));
  }
  var byCol = {}; hits.forEach(function (h) { var cc = /^(?:NAME|PI LINK) ([A-Z]+)/.exec(h)[1]; byCol[cc] = (byCol[cc] || 0) + 1; });
  Logger.log('MIRROR rows=' + d.length + ' cols=' + (d[0] || []).length + ' cells checked=' + cells + ' party names known=' + nk.length + ' PI links known=' + Object.keys(links).length + ' HITS=' + hits.length + ' byCol=' + JSON.stringify(byCol));
  Logger.log(hits.slice(0, 40).join('\n'));
  Logger.log('SAMPLE row ' + (MIRROR.HEAD_ROWS + 3) + ': ' + JSON.stringify(d[MIRROR.HEAD_ROWS + 2].slice(0, 14)) + '\nLAST row ' + d.length + ': ' + JSON.stringify(d[last - 1].slice(0, 14)) + ' link cell formula: ' + f[last - 1][MIRROR.LINK_COL - 1]);
}

/* ---------- one-time backfill: production orders for PIs already in the FMS ----------
 * For every order in the FMS tab from BACKFILL.FROM onwards that has no production order yet:
 * finds its PI tab in this workbook, makes the plant copy (no name / address / GST / rates / terms),
 * saves the PDF and records it in the PI Register. Does not touch the order form.
 * One run works for about 4.5 minutes; run it again until it reports nothing left.
 *
 * FY   = financial-year part of your PI numbers, e.g. PI no "...26-27/0703" -> '26-27'
 * FROM = first PI sequence number to backfill (703 in "26-27/0703")
 */
var BACKFILL = { FY: '26-27', FROM: 1 };
function piSeq_(s) { var m = /(\d{2}-\d{2})[\/_\s-]+0*(\d+)/.exec(norm_(s)); return m ? { fy: m[1], n: parseInt(m[2], 10) } : null; }
function plantCopyLeaks_(tmp, p) {
  SpreadsheetApp.flush();
  var d = tmp.getDataRange().getDisplayValues(), ck = new RegExp('^' + CFG.CODE_PREFIX + '\\d+$', 'i').test(p.customer) ? '' : key_(p.customer), n = 0;
  for (var r = 0; r < d.length; r++) for (var c = 0; c < d[r].length; c++) {
    var t = norm_(d[r][c]); if (!t) continue;
    if (ck.length > 3 && key_(t).indexOf(ck) >= 0) n++;
    if (foreignGst_(t)) n++;
  }
  return n;
}
function backfillPlantOrders() {
  var t0 = Date.now(), lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  var ss = SpreadsheetApp.getActive(), tmp = null, done = [], failed = [], already = 0, left = 0, stop = '';
  try {
    ss.getSheets().forEach(function (s) { if (/^(PO|PI)_[0-9]{13}$/.test(s.getName())) ss.deleteSheet(s); });   // temp copies left by an interrupted run
    var master = SpreadsheetApp.openById(CFG.MASTER_ID), plant = SpreadsheetApp.openById(CFG.PLANT_ID);
    var src = master.getSheetByName(MIRROR.SRC_TAB), last = mirrorLastRow_(src), fd = src.getRange(1, 1, last, 4).getDisplayValues();
    var reg = register_(ss), have = {};
    if (reg.getLastRow() > 1) reg.getRange(2, 1, reg.getLastRow() - 1, 10).getDisplayValues().forEach(function (r) { if (norm_(r[9])) have[key_(r[0])] = 1; });
    var hl = plant.getSheetByName(MIRROR.LIST_TAB);                 // production orders already made from another PI workbook
    if (hl && hl.getLastRow() > 1) hl.getRange(2, 1, hl.getLastRow() - 1, 10).getDisplayValues().forEach(function (r) { if (norm_(r[9])) have[key_(r[0])] = 1; });
    var cs = master.getSheetByName(CFG.CODES_TAB), cd = cs.getRange(1, 1, Math.max(cs.getLastRow(), 1), 2).getDisplayValues(), codes = {};
    for (var i = 1; i < cd.length; i++) { var ck = key_(cd[i][0]); if (ck && norm_(cd[i][1]) && !codes[ck]) codes[ck] = norm_(cd[i][1]); }
    var tabs = {};
    ss.getSheets().forEach(function (s) {
      var n = s.getName(); if (!/_PI_/i.test(n) || /\(B\)/i.test(n)) return;
      var q = piSeq_(n.replace(/^.*_PI_/i, '')); if (!q || q.fy !== BACKFILL.FY || q.n < BACKFILL.FROM) return;
      (tabs[q.n] = tabs[q.n] || []).push(s);
    });
    var list = plant.getSheetByName(MIRROR.LIST_TAB), tz = ss.getSpreadsheetTimeZone(), me = Session.getActiveUser().getEmail();
    for (var r = MIRROR.HEAD_ROWS; r < last; r++) {
      var piNo = norm_(fd[r][MIRROR.PI_COL - 1]), q = piSeq_(piNo.replace(/^.*PI/i, ''));
      if (!q || q.fy !== BACKFILL.FY || q.n < BACKFILL.FROM) continue;
      if (have[key_(piNo)]) { already++; continue; }
      if (stop || Date.now() - t0 > 265000) { left++; continue; }
      try {
        var cand = tabs[q.n] || [], tab = null, pi = null;
        var isCode = new RegExp('^' + CFG.CODE_PREFIX + '\\d+$', 'i');
        for (var k = 0; k < cand.length; k++) {                     // prefer the tab that carries the customer name over a hand-made plant copy
          try { var p = readPI_(cand[k]); if (key_(p.piNo) !== key_(piNo)) continue; if (!tab || (isCode.test(pi.customer) && !isCode.test(p.customer))) { tab = cand[k]; pi = p; } } catch (e1) {}
        }
        if (!tab) { failed.push(piNo + ' - no PI tab in this workbook'); continue; }
        var name = norm_(fd[r][MIRROR.NAME_COL - 1]), code = codes[key_(name)] || codes[key_(pi.customer)];
        if (!code) { try { code = customerCode_(master, name || pi.customer).code; codes[key_(name || pi.customer)] = code; } catch (e2) { failed.push(piNo + ' - no customer ID (' + CFG.CODES_TAB + ' could not be written)'); continue; } }
        tmp = makePlantCopy_(ss, tab, pi, code);
        var leaks = plantCopyLeaks_(tmp, pi);
        if (leaks) { ss.deleteSheet(tmp); tmp = null; failed.push(piNo + ' - plant copy still showed customer details (' + leaks + '), not exported'); continue; }
        var url = exportPdf_(ss, tmp, 'PO_' + piNo.replace(/[\/\\]/g, '_').replace(/\s+/g, '') + '_' + code + '.pdf', CFG.PLANT_PDF_FOLDER);
        ss.deleteSheet(tmp); tmp = null;
        var now = new Date(), orderDate = pi.date instanceof Date ? Utilities.formatDate(pi.date, tz, 'dd/MM/yyyy') : pi.date;
        reg.appendRow([piNo, tab.getName(), pi.customer, code, 'BACKFILLED', 0, now, me, '', url]);
        if (list) {
          var qtyTxt = [pi.qty1 ? pi.qty1 + ' ' + pi.unit1.replace(/^in\s+/i, '') : '', pi.qty2 ? pi.qty2 + ' ' + pi.unit2.replace(/^in\s+/i, '') : ''].filter(String).join(' / ');
          list.getRange(list.getLastRow() + 1, 1, 1, 12).setValues([[piNo, 0, code, orderDate, pi.category, qtyTxt, '', '', '', url, now, 'Backfilled']]);
        }
        have[key_(piNo)] = 1; done.push(piNo);
        Utilities.sleep(700);
      } catch (e) {
        if (tmp) { try { ss.deleteSheet(tmp); } catch (e3) {} tmp = null; }
        failed.push(piNo + ' - ' + String(e.message).slice(0, 90));
        if (/PDF export failed/.test(e.message)) { stop = 'PDF export refused, stopping this run'; }
      }
    }
    var sync = ''; try { SpreadsheetApp.flush(); sync = syncPlantMirror_(); } catch (e4) { sync = 'mirror not refreshed: ' + e4.message; }
    Logger.log('BACKFILL made=' + done.length + ' alreadyHad=' + already + ' failed=' + failed.length + ' LEFT=' + left + (stop ? ' (' + stop + ')' : '') + ' secs=' + Math.round((Date.now() - t0) / 1000));
    Logger.log('MADE: ' + done.map(function (x) { return x.replace(/^.*\//, ''); }).join(','));
    Logger.log('FAILED:\n' + failed.join('\n'));
    Logger.log(sync);
  } finally {
    if (tmp) { try { ss.deleteSheet(tmp); } catch (e5) {} }
    lock.releaseLock();
  }
}

/** Read-only check: compares what the FMS tab shows with what the plant mirror shows, column by column. */
function mirrorCompare() {
  var master = SpreadsheetApp.openById(CFG.MASTER_ID), plant = SpreadsheetApp.openById(CFG.PLANT_ID);
  var src = master.getSheetByName(MIRROR.SRC_TAB), dst = plant.getSheetByName(MIRROR.TAB), last = mirrorLastRow_(src), nC = src.getLastColumn();
  var a = src.getRange(1, 1, last, nC).getDisplayValues(), b = dst.getRange(1, 1, last, nC).getDisplayValues(), diff = {}, ex = {}, filled = {};
  var col = function (i) { var s = ''; i++; while (i > 0) { var q = (i - 1) % 26; s = String.fromCharCode(65 + q) + s; i = Math.floor((i - 1) / 26); } return s; };
  for (var r = 0; r < last; r++) for (var c = 0; c < nC; c++) {
    if (c === MIRROR.NAME_COL - 1 || c === MIRROR.LINK_COL - 1) continue;
    if (norm_(a[r][c])) filled[col(c)] = (filled[col(c)] || 0) + 1;
    if (norm_(a[r][c]) !== norm_(b[r][c])) { var k = col(c); diff[k] = (diff[k] || 0) + 1; if (!ex[k]) ex[k] = 'r' + (r + 1) + ' FMS[' + String(a[r][c]).slice(0, 22) + '] mirror[' + String(b[r][c]).slice(0, 22) + ']'; }
  }
  Logger.log('COMPARE rows=' + last + ' cols=' + nC + ' columns that differ: ' + JSON.stringify(diff));
  Logger.log('EXAMPLES: ' + JSON.stringify(ex));
  Logger.log('FMS filled cells per column: ' + JSON.stringify(filled));
}