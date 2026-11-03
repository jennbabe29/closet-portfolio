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
