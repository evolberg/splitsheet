/**
 * SplitSheet 2.2 — 100/0 Split Options
 * Mobile web app + Google Sheet formula backend. The app writes raw data only; the sheet calculates split shares and balances.
 *
 * Install:
 * 1. Open your SplitSheet Google Sheet.
 * 2. Extensions → Apps Script.
 * 3. Replace Code.gs with this file.
 * 4. Add an HTML file named Index and paste Index.html.
 * 5. Run setupSplitSheet() once, authorize.
 * 6. Deploy → New deployment → Web app.
 */

const SS_CONFIG = {
  people: ['Ethan', 'Kaya'],
  expensesSheet: 'Expenses',
  paymentsSheet: 'Payments',
  activityLogSheet: 'Activity Log',
  listsSheet: 'Lists',
  categories: ['Groceries', 'Restaurants', 'Gas', 'Utilities', 'Rent', 'Vacation', 'Shopping', 'Entertainment', 'Other'],
  splitPresets: ['100/0', '90/10', '80/20', '70/30', '60/40', '50/50', '40/60', '30/70', '20/80', '10/90', '0/100']
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SplitSheet')
    .addItem('Open mobile app', 'showAppSidebar')
    .addItem('Add Expense', 'showExpenseForm')
    .addItem('Record Payment', 'showPaymentForm')
    .addItem('Settle Up', 'showSettleUpForm')
    .addItem('Undo Last Entry', 'undoLastEntryFromMenu')
    .addSeparator()
    .addItem('Setup / repair sheets', 'setupSplitSheet')
    .addToUi();
}

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('SplitSheet')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function showAppSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Index').setTitle('SplitSheet');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Desktop menu compatibility. These all use the same mobile-first app shell.
function showExpenseForm() { showAppSidebar(); }
function showPaymentForm() { showAppSidebar(); }
function showSettleUpForm() { showAppSidebar(); }

function setupSplitSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expenses = getOrCreateSheet_(ss, SS_CONFIG.expensesSheet);
  const payments = getOrCreateSheet_(ss, SS_CONFIG.paymentsSheet);
  const lists = getOrCreateSheet_(ss, SS_CONFIG.listsSheet);
  const activity = getOrCreateSheet_(ss, SS_CONFIG.activityLogSheet);

  ensureHeaders_(expenses, ['ID', 'Date', 'Description', 'Category', 'Paid By', 'Amount', 'Split Preset (Ethan/Kaya)', 'Ethan Share', 'Kaya Share', 'Notes', 'Receipt Link']);
  ensureHeaders_(payments, ['Date', 'From', 'To', 'Amount', 'Note']);
  ensureHeaders_(activity, ['Timestamp', 'Type', 'Sheet', 'Row', 'Summary', 'Amount', 'Status']);

  lists.clear();
  lists.getRange(1, 1).setValue('Categories');
  lists.getRange(2, 1, SS_CONFIG.categories.length, 1).setValues(SS_CONFIG.categories.map(v => [v]));
  lists.getRange(1, 3).setValue('Split Presets');
  lists.getRange(2, 3, SS_CONFIG.splitPresets.length, 1).setValues(SS_CONFIG.splitPresets.map(v => [v]));
  lists.getRange(1, 5).setValue('People');
  lists.getRange(2, 5, SS_CONFIG.people.length, 1).setValues(SS_CONFIG.people.map(v => [v]));

  expenses.setFrozenRows(1);
  payments.setFrozenRows(1);
  activity.setFrozenRows(1);
  expenses.autoResizeColumns(1, 9);
  payments.autoResizeColumns(1, 5);
  activity.autoResizeColumns(1, 7);

  return { ok: true, message: 'SplitSheet setup complete.' };
}

function getAppData() {
  setupIfNeeded_();
  const balance = calculateBalance_();
  const recentExpenses = getRecentExpenses_(8);
  const monthStats = getMonthStats_();
  const undoInfo = getUndoInfo_();

  return {
    ok: true,
    people: SS_CONFIG.people,
    categories: SS_CONFIG.categories,
    splitPresets: SS_CONFIG.splitPresets,
    balance,
    recentExpenses,
    monthStats,
    undoInfo,
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

function saveExpense(expense) {
  setupIfNeeded_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.expensesSheet);
  const headers = getHeaderMap_(sheet);

  const date = parseDate_(expense.date);
  const amount = parseMoney_(expense.amount);
  const description = String(expense.description || '').trim();
  const paidBy = String(expense.paidBy || '').trim();
  const splitPreset = String(expense.splitPreset || '50/50').trim();
  const category = String(expense.category || 'Other').trim();
  const notes = String(expense.notes || '').trim();
  const receiptLink = String(expense.receiptLink || '').trim();

  if (!description) throw new Error('Description is required.');
  if (!amount || amount <= 0) throw new Error('Amount must be greater than 0.');
  if (SS_CONFIG.people.indexOf(paidBy) === -1) throw new Error('Paid By must be Ethan or Kaya.');
  validateSplitPreset_(splitPreset); // validates only; the sheet does the actual split math.

  const rowIndex = findFirstEmptyRow_(sheet, 2, headers['date'] || 2);

  // Important architecture choice:
  // The web app writes only raw input fields. It does NOT calculate Ethan/Kaya shares.
  // Formula columns like ID, Ethan Share, and Kaya Share are left to the sheet.
  setCellByHeader_(sheet, rowIndex, headers, 'date', date);
  setCellByHeader_(sheet, rowIndex, headers, 'description', description);
  setCellByHeader_(sheet, rowIndex, headers, 'category', category);
  setCellByHeader_(sheet, rowIndex, headers, 'paid by', paidBy);
  setCellByHeader_(sheet, rowIndex, headers, 'amount', amount);
  setCellByHeader_(sheet, rowIndex, headers, 'split preset (ethan/kaya)', splitPreset);
  setCellByHeader_(sheet, rowIndex, headers, 'notes', notes);
  setCellByHeader_(sheet, rowIndex, headers, 'receipt link', receiptLink);

  ensureExpenseRowFormulas_(sheet, rowIndex, headers);

  if (headers['date']) sheet.getRange(rowIndex, headers['date']).setNumberFormat('yyyy-mm-dd');
  if (headers['amount']) sheet.getRange(rowIndex, headers['amount']).setNumberFormat('$#,##0.00');

  SpreadsheetApp.flush();
  logActivity_('expense', SS_CONFIG.expensesSheet, rowIndex, description, amount);
  return { ok: true, message: 'Expense added.', balance: calculateBalance_(), undoInfo: getUndoInfo_() };
}

function recordPayment(payment) {
  setupIfNeeded_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.paymentsSheet);
  const headers = getHeaderMap_(sheet);

  const date = parseDate_(payment.date);
  const from = String(payment.from || '').trim();
  const to = String(payment.to || '').trim();
  const amount = parseMoney_(payment.amount);
  const note = String(payment.note || '').trim();

  if (SS_CONFIG.people.indexOf(from) === -1 || SS_CONFIG.people.indexOf(to) === -1) throw new Error('Payment must be between Ethan and Kaya.');
  if (from === to) throw new Error('From and To cannot be the same person.');
  if (!amount || amount <= 0) throw new Error('Amount must be greater than 0.');

  const rowIndex = findFirstEmptyRow_(sheet, 2, headers['date'] || 1);
  const row = new Array(sheet.getLastColumn()).fill('');

  setByHeader_(row, headers, 'date', date);
  setByHeader_(row, headers, 'from', from);
  setByHeader_(row, headers, 'to', to);
  setByHeader_(row, headers, 'amount', amount);
  setByHeader_(row, headers, 'note', note);

  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  sheet.getRange(rowIndex, headers['date'] || 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(rowIndex, headers['amount'] || 4).setNumberFormat('$#,##0.00');

  logActivity_('payment', SS_CONFIG.paymentsSheet, rowIndex, from + ' → ' + to, amount);
  return { ok: true, message: 'Payment recorded.', balance: calculateBalance_(), undoInfo: getUndoInfo_() };
}


function undoLastEntry() {
  setupIfNeeded_();
  const info = getUndoInfo_();
  if (!info.available) throw new Error('Nothing to undo yet.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = ss.getSheetByName(info.sheet);
  if (!target) throw new Error('Could not find sheet for last entry: ' + info.sheet);

  const row = Number(info.row);
  if (!row || row < 2) throw new Error('Invalid undo row.');

  if (info.type === 'expense') {
    const h = getHeaderMap_(target);
    const lastCol = Math.max(target.getLastColumn(), 11);
    target.getRange(row, 1, 1, lastCol).clearContent();
    // Keep the sheet tidy: formulas will be re-created the next time this row is used.
  } else if (info.type === 'payment') {
    const lastCol = Math.max(target.getLastColumn(), 5);
    target.getRange(row, 1, 1, lastCol).clearContent();
  } else {
    throw new Error('Unknown activity type: ' + info.type);
  }

  markActivityUndone_(info.logRow);
  SpreadsheetApp.flush();
  return { ok: true, message: 'Undid last ' + info.type + '.', balance: calculateBalance_(), undoInfo: getUndoInfo_() };
}

function undoLastEntryFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const info = getUndoInfo_();
  if (!info.available) {
    ui.alert('Nothing to undo yet.');
    return;
  }
  const response = ui.alert(
    'Undo last entry?',
    info.type.toUpperCase() + ': ' + info.summary + ' — $' + Number(info.amount || 0).toFixed(2),
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) {
    const result = undoLastEntry();
    ui.alert(result.message);
  }
}

function getUndoInfo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(SS_CONFIG.activityLogSheet);
  if (!log || log.getLastRow() < 2) return { available: false, label: 'Nothing to undo' };

  const values = log.getRange(2, 1, log.getLastRow() - 1, 7).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const status = String(row[6] || '').trim().toUpperCase();
    if (status === 'UNDONE') continue;
    const type = String(row[1] || '').trim();
    const sheet = String(row[2] || '').trim();
    const entryRow = Number(row[3]);
    const summary = String(row[4] || '').trim();
    const amount = Number(row[5]) || 0;
    if (!type || !sheet || !entryRow) continue;
    return {
      available: true,
      logRow: i + 2,
      type,
      sheet,
      row: entryRow,
      summary,
      amount,
      label: 'Undo ' + type + ': ' + summary + ' — $' + amount.toFixed(2)
    };
  }
  return { available: false, label: 'Nothing to undo' };
}

function logActivity_(type, sheetName, rowIndex, summary, amount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = getOrCreateSheet_(ss, SS_CONFIG.activityLogSheet);
  ensureHeaders_(log, ['Timestamp', 'Type', 'Sheet', 'Row', 'Summary', 'Amount', 'Status']);
  const logRow = findFirstEmptyRow_(log, 2, 1);
  log.getRange(logRow, 1, 1, 7).setValues([[new Date(), type, sheetName, rowIndex, summary, amount, 'ACTIVE']]);
  log.getRange(logRow, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  log.getRange(logRow, 6).setNumberFormat('$#,##0.00');
}

function markActivityUndone_(logRow) {
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.activityLogSheet);
  if (log && logRow) log.getRange(logRow, 7).setValue('UNDONE');
}

function getSettleUpSuggestion() {
  setupIfNeeded_();
  const balance = calculateBalance_();
  if (Math.abs(balance.ethanNet) < 0.005) {
    return { ok: true, settled: true, from: '', to: '', amount: 0, meaning: 'You are settled up.' };
  }

  const ethanIsOwed = balance.ethanNet > 0;
  return {
    ok: true,
    settled: false,
    from: ethanIsOwed ? 'Kaya' : 'Ethan',
    to: ethanIsOwed ? 'Ethan' : 'Kaya',
    amount: Math.abs(balance.ethanNet),
    meaning: balance.meaning
  };
}

function calculateBalance_() {
  // Source of truth: the Google Sheet.
  // The app reads the Balances tab, which is formula-driven from Expenses + Payments.
  SpreadsheetApp.flush();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const balances = ss.getSheetByName('Balances');
  if (!balances || balances.getLastRow() < 3) {
    throw new Error('Balances sheet is missing or incomplete. Run setupSplitSheet(), then check the workbook formulas.');
  }

  const h = getHeaderMap_(balances);
  const personCol = h['person'] || 1;
  const netCol = h['net position'] || 5;
  const meaningCol = h['meaning'] || 6;
  const values = balances.getRange(2, 1, balances.getLastRow() - 1, balances.getLastColumn()).getValues();

  let ethanNet = 0;
  let kayaNet = 0;
  let ethanMeaning = '';
  let kayaMeaning = '';

  values.forEach(row => {
    const person = String(row[personCol - 1] || '').trim();
    const net = Number(row[netCol - 1]) || 0;
    const meaning = String(row[meaningCol - 1] || '').trim();
    if (person === 'Ethan') {
      ethanNet = net;
      ethanMeaning = meaning;
    }
    if (person === 'Kaya') {
      kayaNet = net;
      kayaMeaning = meaning;
    }
  });

  ethanNet = Math.round(ethanNet * 100) / 100;
  kayaNet = Math.round(kayaNet * 100) / 100;
  const amount = Math.abs(ethanNet);
  let meaning = 'You are settled up.';
  if (amount >= 0.01) meaning = ethanNet > 0 ? 'Kaya owes Ethan' : 'Ethan owes Kaya';

  return {
    ethanNet,
    kayaNet,
    amount: Math.round(amount * 100) / 100,
    meaning,
    settled: amount < 0.01,
    ethanMeaning,
    kayaMeaning,
    source: 'Balances sheet'
  };
}

function getRecentExpenses_(limit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.expensesSheet);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const h = getHeaderMap_(sheet);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()
    .filter(row => row[(h['description'] || 3) - 1] || row[(h['amount'] || 6) - 1])
    .slice(-limit)
    .reverse();

  return rows.map(row => ({
    date: formatDateForApp_(row[(h['date'] || 2) - 1]),
    description: String(row[(h['description'] || 3) - 1] || ''),
    category: String(row[(h['category'] || 4) - 1] || ''),
    paidBy: String(row[(h['paid by'] || 5) - 1] || ''),
    amount: Number(row[(h['amount'] || 6) - 1]) || 0,
    splitPreset: String(row[(h['split preset (ethan/kaya)'] || h['split preset'] || 7) - 1] || '')
  }));
}

function getMonthStats_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.expensesSheet);
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  const label = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM yyyy');
  let total = 0;
  let count = 0;
  const byCategory = {};

  if (sheet && sheet.getLastRow() >= 2) {
    const h = getHeaderMap_(sheet);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    rows.forEach(row => {
      const date = row[(h['date'] || 2) - 1];
      const d = date instanceof Date ? date : new Date(date);
      if (!d || isNaN(d.getTime()) || d.getFullYear() !== year || d.getMonth() !== month) return;
      const amount = Number(row[(h['amount'] || 6) - 1]) || 0;
      if (!amount) return;
      const category = String(row[(h['category'] || 4) - 1] || 'Other');
      total += amount;
      count += 1;
      byCategory[category] = (byCategory[category] || 0) + amount;
    });
  }

  return { label, total: Math.round(total * 100) / 100, count, byCategory };
}

function setupIfNeeded_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SS_CONFIG.expensesSheet) || !ss.getSheetByName(SS_CONFIG.paymentsSheet)) setupSplitSheet();
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders_(sheet, headers) {
  const existing = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0] : [];
  headers.forEach((header, i) => {
    if (!existing[i]) sheet.getRange(1, i + 1).setValue(header);
  });
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim().toLowerCase();
    if (key) map[key] = i + 1;
  });
  return map;
}

function setByHeader_(row, headers, headerName, value) {
  const col = headers[headerName.toLowerCase()];
  if (col) row[col - 1] = value;
}

function findFirstEmptyRow_(sheet, startRow, checkColumn) {
  const maxRows = sheet.getMaxRows();
  const values = sheet.getRange(startRow, checkColumn, maxRows - startRow + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) return startRow + i;
  }
  sheet.insertRowsAfter(maxRows, 50);
  return maxRows + 1;
}

function makeExpenseId_(sheet, headers) {
  const idCol = headers['id'] || 1;
  if (sheet.getLastRow() < 2) return 1;
  const values = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues().flat();
  const nums = values.map(v => Number(v)).filter(v => !isNaN(v));
  return nums.length ? Math.max.apply(null, nums) + 1 : 1;
}

function validateSplitPreset_(preset) {
  const clean = String(preset || '50/50').replace(/\s/g, '');
  const match = clean.match(/^(\d{1,3})\/(\d{1,3})$/);
  if (!match) throw new Error('Split must look like 50/50 or 70/30.');
  const ethan = Number(match[1]);
  const kaya = Number(match[2]);
  if (ethan + kaya !== 100) throw new Error('Split must add up to 100.');
  return true;
}

function setCellByHeader_(sheet, rowIndex, headers, headerName, value) {
  const col = headers[headerName.toLowerCase()];
  if (col) sheet.getRange(rowIndex, col).setValue(value);
}

function ensureExpenseRowFormulas_(sheet, rowIndex, headers) {
  const idCol = headers['id'];
  const ethanShareCol = headers['ethan share'];
  const kayaShareCol = headers['kaya share'];

  if (idCol && !sheet.getRange(rowIndex, idCol).getFormula()) {
    sheet.getRange(rowIndex, idCol).setFormula(`=IF(B${rowIndex}="","",ROW()-1)`);
  }

  if (ethanShareCol && !sheet.getRange(rowIndex, ethanShareCol).getFormula()) {
    sheet.getRange(rowIndex, ethanShareCol).setFormula(`=IF($F${rowIndex}="","",IFERROR(VALUE(LEFT($G${rowIndex},FIND("/",$G${rowIndex})-1))/100*$F${rowIndex},$F${rowIndex}/2))`);
  }

  if (kayaShareCol && !sheet.getRange(rowIndex, kayaShareCol).getFormula()) {
    sheet.getRange(rowIndex, kayaShareCol).setFormula(`=IF($F${rowIndex}="","",IFERROR(VALUE(MID($G${rowIndex},FIND("/",$G${rowIndex})+1,99))/100*$F${rowIndex},$F${rowIndex}/2))`);
  }
}

function parseMoney_(value) {
  if (typeof value === 'number') return value;
  return Number(String(value || '').replace(/[$,]/g, '').trim());
}

function parseDate_(value) {
  if (!value) return new Date();
  if (value instanceof Date) return value;
  const parts = String(value).split('-');
  if (parts.length === 3) return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  return new Date(value);
}

function formatDateForApp_(value) {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
