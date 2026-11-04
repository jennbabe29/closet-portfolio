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
     if (sheet.getName() === 'Assets') prepareAssetRow(sheet, row, e.range.getNumColumns());
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
    .addItem('List selected asset for sale', 'listForSale')
    .addItem('Mark selected item as sold', 'markAsSold')
    .addItem('Move selected asset to Recycle/Donate/Gift', 'moveToRecycleDonateGift')
    .addToUi();
      setUpAssets();
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
// ---------- easier data entry on Assets ----------

// Asset Class, Segment and Size. The Sale Log has the same columns one to the right.
const DROPDOWN_COLS = ['C', 'D', 'G'];
const SALE_LOG_DROPDOWN_COLS = ['D', 'E', 'H'];
const MONEY_FORMAT = '"$"#,##0.00';

// Runs when the sheet opens: rebuilds the dropdowns and keeps Retail, Cost and Discount in dollars.
function setUpAssets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assets = ss.getSheetByName('Assets');
  const first = findHeaderRow(assets) + 1;
  const last = assets.getMaxRows();
  DROPDOWN_COLS.forEach((col, i) => {
    const options = dropdownOptions(ss, i);
    if (options.length === 0) return;
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .setAllowInvalid(true)                // you can still type something new
      .build();
    assets.getRange(col + first + ':' + col + last).setDataValidation(rule);
  });
  assets.getRange('I' + first + ':K' + last).setNumberFormat(MONEY_FORMAT);
}

// Every value already used in this column on any tab, without blanks or repeats, A to Z.
function dropdownOptions(ss, i) {
  const options = new Set();
  for (const name in TABS) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    const col = name === 'Sale Log' ? SALE_LOG_DROPDOWN_COLS[i] : DROPDOWN_COLS[i];
    const first = findHeaderRow(sheet) + 1;
    const last = sheet.getLastRow();
    if (last < first) continue;
    sheet.getRange(col + first + ':' + col + last).getValues().forEach(cell => {
      const value = cell[0];
      if (value !== '' && !(value instanceof Date)) options.add(String(value).trim());
    });
  }
  return [...options].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

// A row you just started by hand gets today's date. Rows with other data are left alone,
// so you can change the date, and old rows without a date stay as they are.
function prepareAssetRow(sheet, row, cellsEdited) {
  const values = sheet.getRange(row, 1, 1, 10).getValues()[0];   // A:J
  const filled = values.filter(v => v !== '').length;
  if (values[0] === '' && filled >= 1 && filled <= cellsEdited) {
    sheet.getRange(row, 1).setValue(today());
    sheet.getRange(row + 1, 3, 1, 5)                             // C:G dropdowns from the row below
         .copyTo(sheet.getRange(row, 3, 1, 5), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
  sheet.getRange(row, 9, 1, 3).setNumberFormat(MONEY_FORMAT);    // I:K stay dollars, even after pasting
}
// ---------- button: List for sale (Assets -> Asset for Sale) ----------

function listForSale() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getActiveSheet();
  const row = selectedItemRow(source, 'Assets');
  if (!row) return;

  const item = source.getRange(row, 1, 1, 10).getValues()[0];      // A:J  Date Acquired ... Cost
  const notes = source.getRange('N' + row).getValue();
  const answer = ui.alert('List for sale?', item[1] + ' ' + item[4], ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const forSale = ss.getSheetByName('Asset for Sale');
  const newRow = insertAtTop(forSale);
  forSale.getRange(newRow, 1, 1, 10).setValues([item]);
  if (notes !== '') forSale.getRange('N' + newRow).setValue(notes);
  fillRow(forSale, newRow);

  source.deleteRow(row);
  ss.setActiveSheet(forSale);
  forSale.setActiveRange(forSale.getRange(newRow, 1));
  ss.toast('Moved to Asset for Sale.', 'Closet', 4);
}

// The row the cursor is on, if it's an item row on the expected tab. Otherwise explains and returns 0.
function selectedItemRow(sheet, expectedTab) {
  const ui = SpreadsheetApp.getUi();
  if (sheet.getName() !== expectedTab) {
    ui.alert('Go to the "' + expectedTab + '" tab and click on the item\'s row first.');
    return 0;
  }
  const row = sheet.getActiveRange().getRow();
  if (row <= findHeaderRow(sheet)) {
    ui.alert('Click on the item\'s row first, then try again.');
    return 0;
  }
  return row;
}
// ---------- button: Mark as sold (Asset for Sale -> Sale Log) ----------

function markAsSold() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getActiveSheet();
  const row = selectedItemRow(source, 'Asset for Sale');
  if (!row) return;

  const item = source.getRange(row, 1, 1, 10).getValues()[0];      // A:J
  const notes = source.getRange('N' + row).getValue();

  const price = ui.prompt('Sold: ' + item[1] + ' ' + item[4], 'Sold price (just the number, e.g. 85):', ui.ButtonSet.OK_CANCEL);
  if (price.getSelectedButton() !== ui.Button.OK) return;
  const broker = ui.prompt('Where did it sell?', 'Poshmark, eBay, The Real Real, GOAT, Depop...', ui.ButtonSet.OK_CANCEL);
  if (broker.getSelectedButton() !== ui.Button.OK) return;

  const log = ss.getSheetByName('Sale Log');
  const newRow = insertAtTop(log);
  log.getRange(newRow, 1).setValue(today());                        // A  Date Sold
  log.getRange(newRow, 2, 1, 10).setValues([item]);                 // B:K <- A:J
  log.getRange('M' + newRow).setValue(Number(price.getResponseText().replace(/[$,]/g, '')));
  log.getRange('R' + newRow).setValue(broker.getResponseText().trim());
  if (notes !== '') log.getRange('O' + newRow).setValue(notes);
  fillRow(log, newRow);

  source.deleteRow(row);
  ss.setActiveSheet(log);
  log.setActiveRange(log.getRange('P' + newRow));
  ss.toast('Moved to Sale Log. Check Sold Profit (type it in for eBay).', 'Closet', 6);
}
// ---------- button: Recycle, donate or gift (Assets -> Recycle/Donate/Gift) ----------

function moveToRecycleDonateGift() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getActiveSheet();
  const row = selectedItemRow(source, 'Assets');
  if (!row) return;

  const item = source.getRange(row, 1, 1, 10).getValues()[0];      // A:J
  const notes = source.getRange('N' + row).getValue();

  const reason = ui.prompt('Where did it go? ' + item[1] + ' ' + item[4], 'For example: Gave to Mom, Donated, Threw out due to wear', ui.ButtonSet.OK_CANCEL);
  if (reason.getSelectedButton() !== ui.Button.OK) return;

  const tab = ss.getSheetByName('Recycle/Donate/Gift');
  const newRow = insertAtTop(tab);
  tab.getRange(newRow, 1, 1, 10).setValues([item]);
  const newNotes = [reason.getResponseText().trim(), notes].filter(n => n !== '').join('; ');
  if (newNotes !== '') tab.getRange('N' + newRow).setValue(newNotes);
  fillRow(tab, newRow);

  source.deleteRow(row);
  ss.setActiveSheet(tab);
  tab.setActiveRange(tab.getRange(newRow, 1));
  ss.toast('Moved to Recycle/Donate/Gift.', 'Closet', 4);
}
