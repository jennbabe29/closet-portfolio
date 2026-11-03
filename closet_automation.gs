// Closet Portfolio
// A closet manager built like a financial portfolio, for fashion resellers.
// Google Apps Script for a Google Sheet with Assets, Asset for Sale and Sale Log tabs.
// Paste into Extensions > Apps Script, save, then reload the sheet.
// License: MIT (see LICENSE)


// ---------- automatic formulas ----------
// Type a Description and the row's formulas fill in. Filled cells are never overwritten.

// The formulas you use, per tab. {r} becomes the row number.
const ITEM_FORMULAS = {
  K: '=I{r}-J{r}',          // Discount = Retail - Cost
  L: '=TODAY()-A{r}',       // Hold (days since acquired)
  M: '=DATEDIF(0,L{r},"y") & " Years " & DATEDIF(0,L{r},"ym") & " Months " & DATEDIF(0,L{r},"md") & " Days "',
};
const TABS = {
  'Assets': ITEM_FORMULAS,
  'Asset for Sale': ITEM_FORMULAS,
  'Recycle/Donate/Gift': ITEM_FORMULAS,
  'Sale Log': {
    L: '=J{r}-K{r}',        // Discount = Retail - Cost
    Q: '=P{r}-K{r}',        // Net Loss/Gain = Sold Profit - Cost
    S: '=DATEDIF(0,T{r},"y")&" years " &DATEDIF(0,T{r},"ym")&" months "&DATEDIF(0,T{r},"md")&" days"',
    T: '=A{r}-B{r}',        // Hold = Date Sold - Date Acquired
  },
};
// Poshmark pays out 80% of the sale price, or the price minus $2.95 under $15.
const POSHMARK_PROFIT = '=IF(M{r}<15, M{r}-2.95, M{r}*0.8)';
const DESCRIPTION_COL = { 'Sale Log': 'F' };   // every other tab uses E

// Runs by itself every time you edit a cell.
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (!TABS[sheet.getName()]) return;
  const first = Math.max(e.range.getRow(), findHeaderRow(sheet) + 1);
  for (let row = first; row <= e.range.getLastRow(); row++) {
    fillRow(sheet, row);
  }
}

// Fill one row's blank formula cells.
function fillRow(sheet, row) {
  const name = sheet.getName();
  const formulas = TABS[name];
  const descCol = DESCRIPTION_COL[name] || 'E';
  if (sheet.getRange(descCol + row).getValue() === '') return;   // empty row: leave it alone

  for (const col in formulas) {
    const cell = sheet.getRange(col + row);
    if (cell.getFormula() === '' && cell.getValue() === '') {
      cell.setFormula(formulas[col].split('{r}').join(row));
    }
  }

  if (name === 'Sale Log') {
    const broker = sheet.getRange('R' + row).getValue();
    const profit = sheet.getRange('P' + row);
    if (broker === 'Poshmark' && profit.getFormula() === '' && profit.getValue() === '') {
      profit.setFormula(POSHMARK_PROFIT.split('{r}').join(row));
    }
  }
}

// The header is the first row whose first cell isn't blank (Asset for Sale starts on row 2).
function findHeaderRow(sheet) {
  let row = 1;
  while (sheet.getRange(row, 1).getValue() === '' && row < 10) row++;
  return row;
}


// ---------- Closet menu and button: Add asset ----------

// Runs by itself when the sheet opens and adds the Closet menu.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Closet')
    .addItem('Add asset', 'addAsset')
    .addToUi();
}

function addAsset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assets = ss.getSheetByName('Assets');
  const row = insertAtTop(assets);
  assets.getRange(row, 1).setValue(today());
  ss.setActiveSheet(assets);
  assets.setActiveRange(assets.getRange(row, 2));   // cursor on Brand
  // the formulas fill in by themselves once you type the Description
}

// Insert a blank row under the header, with the formatting and dropdowns of the row below it. Returns its number.
function insertAtTop(sheet) {
  const row = findHeaderRow(sheet) + 1;
  sheet.insertRowBefore(row);
  const width = sheet.getLastColumn();
  sheet.getRange(row + 1, 1, 1, width)
       .copyTo(sheet.getRange(row, 1, 1, width), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sheet.getRange(row + 1, 1, 1, width)
       .copyTo(sheet.getRange(row, 1, 1, width), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  return row;
}

// Today's date in the sheet's own time zone.
function today() {
  return dateOnly(new Date());
}

// A date as text like 2026-11-03, which Sheets reads as a plain date with no time.
// A Date object can carry the script's time zone, making Hold slightly negative.
function dateOnly(date) {
  const zone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, zone, 'yyyy-MM-dd');
}
