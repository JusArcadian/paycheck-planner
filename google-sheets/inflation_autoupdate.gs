/**
 * Paycheck Planner — Google Sheets add-on script
 * ------------------------------------------------
 * Adds a "Budget Tools" menu with:
 *   - Update inflation now      -> pulls the latest US inflation rate from the BLS API
 *   - Turn on monthly auto-update -> re-checks on the 15th of every month by itself
 *   - Import my data from an old copy -> moves everything a buyer typed into a NEW version
 *   - Overspending alerts: a pop-up when you log spending that puts you over, plus optional daily emails
 *   - Sort Spending by date / Sort Money in by date (both also run by themselves every time the sheet opens)
 *
 * Where it writes: the named cells InflationAuto and InflationAsOf
 * (fallback: Engine!B16 and Engine!B17, on the hidden Engine tab).
 *
 * The BLS public API (v1) needs no key. Docs: https://www.bls.gov/developers/
 */

var BLS_URL = 'https://api.bls.gov/publicAPI/v1/timeseries/data/CUUR0000SA0'; // CPI-U, all items

function onOpen() {
  try { sortSpending_(true); } catch (e) { /* never block opening the sheet */ }
  SpreadsheetApp.getUi()
    .createMenu('Budget Tools')
    .addItem('Sort Spending by date', 'sortSpending')
    .addItem('Sort Money in by date', 'sortMoneyIn')
    .addItem('Sort categories A to Z', 'sortCategoriesAZ')
    .addItem('Sort categories by group', 'sortCategoriesByGroup')
    .addSeparator()
    .addItem('Update inflation now', 'updateInflation')
    .addItem('Turn on monthly auto-update', 'installMonthlyTrigger')
    .addSeparator()
    .addItem('Import my data from an old copy', 'importFromOldCopy')
    .addSeparator()
    .addItem('Check my alerts now', 'showAlerts')
    .addItem('Email me alerts every morning', 'installAlertEmails')
    .addItem('Stop alert emails', 'removeAlertEmails')
    .addSeparator()
    .addItem("Save this month's net worth", 'saveNetWorth')
    .addItem('Send my data to the app', 'sendToApp')
    .addToUi();
}

function getCell_(name, fallbackA1) {
  var ss = SpreadsheetApp.getActive();
  return ss.getRangeByName(name) || ss.getRange(fallbackA1);
}

function updateInflation() {
  var res = UrlFetchApp.fetch(BLS_URL, { muteHttpExceptions: true });
  var json = JSON.parse(res.getContentText());
  if (json.status !== 'REQUEST_SUCCEEDED') {
    throw new Error('BLS API said: ' + (json.message || []).join(' '));
  }
  var data = json.Results.series[0].data;           // newest month first
  var latest = data[0];
  var yearAgo = data.filter(function (d) {
    return d.period === latest.period && Number(d.year) === Number(latest.year) - 1;
  })[0];
  if (!yearAgo) throw new Error('Could not find the same month last year.');

  // Year-over-year inflation = this month's price index / same month last year - 1
  var rate = Number(latest.value) / Number(yearAgo.value) - 1;
  rate = Math.round(rate * 1000) / 1000;             // round to 0.1%

  getCell_('InflationAuto', 'Engine!B16').setValue(rate);
  getCell_('InflationAsOf', 'Engine!B17')
    .setValue(latest.periodName + ' ' + latest.year + ' (US CPI-U, BLS, auto-updated)');

  try {
    SpreadsheetApp.getActive().toast('Inflation set to ' + (rate * 100).toFixed(1) + '%', 'Budget Tools');
  } catch (e) { /* no UI when run from a timer — that's fine */ }
}

function installMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'updateInflation') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('updateInflation').timeBased().onMonthDay(15).atHour(9).create();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'saveNetWorth') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('saveNetWorth').timeBased().onMonthDay(28).atHour(9).create();
  SpreadsheetApp.getUi().alert('Done! Inflation refreshes on the 15th, and your net worth is written down on the 28th of each month.');
}


/* =====================================================================
 * IMPORT MY DATA FROM AN OLD COPY
 * ---------------------------------------------------------------------
 * When a new version of the Paycheck Planner comes out, a buyer makes a
 * fresh copy, then runs this. It copies ONLY the yellow boxes (what they
 * typed) from their old copy into this one. Formulas, layout and new
 * features stay from the new version.
 *
 * ELI5: the new version is a new notebook with better pages. This moves
 * your handwriting from the old notebook into it.
 * ===================================================================== */

// Every place a buyer types. Sections are FOUND BY THEIR HEADINGS (not fixed cell addresses), so an
// old copy still imports correctly after rows move in a new version (v2.7 moved Setup down 3 rows).
//   label:  find this question text in column B, copy the answer in column C
//   table:  find this header text in column B, copy `rows` rows under it, columns `cols`
//   range:  fixed cells (these never moved)
var IMPORT_JOBS = [
  // Setup questions (old wording on the right still matches)
  {sheet: 'Setup', label: ['How often do you get paid?']},
  {sheet: 'Setup', label: ['First paycheck, if it was smaller (before taxes)']},
  {sheet: 'Setup', label: ['Full paycheck, before taxes (ONE paycheck)', 'How much is each paycheck, before taxes?']},
  {sheet: 'Setup', label: ['Taken out before taxes, per paycheck']},
  {sheet: 'Setup', label: ['When was your first payday this year?']},
  {sheet: 'Setup', label: ['Paid twice a month? Which days?']},     // new in v2.11: skipped for older copies
  {sheet: 'Setup', label: ['Payday on a weekend or bank holiday?', 'Payday lands on a weekend?']},
  {sheet: 'Setup', label: ['Does your bank pay you early?']},
  {sheet: 'Setup', label: ['How should we figure your taxes?']},
  {sheet: 'Setup', label: ['Filing status']},
  {sheet: 'Setup', label: ['State + local income tax rate']},
  {sheet: 'Setup', label: ['Take out Social Security + Medicare?']},
  {sheet: 'Setup', label: ['Your own total tax rate']},
  {sheet: 'Setup', label: ['Which inflation rate should we use?']},
  {sheet: 'Setup', label: ['Your own inflation guess']},
  {sheet: 'Setup', label: ['Pick a budget plan']},
  {sheet: 'Setup', table: 'Group name', rows: 6, cols: ['B', 'C', 'D', 'G']},
  {sheet: 'Setup', table: 'Category', rows: 30, cols: ['B', 'C', 'D', 'E']},
  // Spending + Money in: columns found by their header names, so new columns (Budget check, OK'd?...) don't break it
  {sheet: 'Spending', headers: ['Date', 'What was it?', 'Category', 'Amount', "OK'd?", 'Cover with', 'Notes (optional)'], rows: 1000},
  {sheet: 'Spending', headers: ['Date received', 'From (job, gig, gift...)', 'Amount', 'For (optional)', 'Notes (optional)'], rows: 300},
  {sheet: 'Goals', range: 'G8:G9'},           // count leftover as saved? + start tracking from
  {sheet: 'Goals', range: 'B12:G19'},         // savings goals
  {sheet: 'Accounts', range: 'G8'},           // show future amounts in today's dollars?
  {sheet: 'Accounts', range: 'B12:I19'},      // savings + investing accounts
  {sheet: 'Accounts', label: ['Other things you own (car value, cash at home...)'], col: 'F'},
  {sheet: 'Accounts', label: ['Other money you owe (not on the Debts tab)'], col: 'F'},
  {sheet: 'Accounts', table: 'Month', after: 'NET WORTH', rows: 36, byMonth: true},   // net worth log, matched month by month
  {sheet: 'Bills', table: 'What', rows: 8, cols: ['B', 'C', 'D', 'E']},            // yearly costs
  {sheet: 'Bills', table: 'Who / what', rows: 6, cols: ['B', 'C', 'D', 'E', 'F', 'G']},   // payment arrangements
  {sheet: 'Debts', range: 'F8:F9'},           // extra per month + payoff method
  {sheet: 'Debts', range: 'B12:F19'}          // debts
];

// Row number (1-based) of the first cell in column B that equals one of `texts`, starting after `afterText`.
function findRow_(sheet, texts, afterText) {
  var colB = sheet.getRange(1, 2, sheet.getLastRow(), 1).getValues();
  var start = 0;
  if (afterText) {
    for (var i = 0; i < colB.length; i++) { if (String(colB[i][0]) === afterText) { start = i + 1; break; } }
  }
  for (var j = start; j < colB.length; j++) {
    if (texts.indexOf(String(colB[j][0]).trim()) >= 0) return j + 1;
  }
  return -1;
}

// Finds the header row that starts with headers[0], then the column of each header to its right.
// Returns {row, cols: {header: column number}} or null.
function findHeaders_(sheet, headers) {
  var top = sheet.getRange(1, 1, 10, sheet.getLastColumn()).getValues();
  for (var r = 0; r < top.length; r++) {
    for (var c = 0; c < top[r].length; c++) {
      if (String(top[r][c]).trim() !== headers[0]) continue;
      var cols = {};
      cols[headers[0]] = c + 1;
      headers.slice(1).forEach(function (h) {
        for (var k = c + 1; k < top[r].length; k++) {
          if (String(top[r][k]).trim() === h) { cols[h] = k + 1; return; }
        }
      });
      return { row: r + 1, cols: cols };
    }
  }
  return null;
}

// Copy one block: keep a formula if the old cell had one (like =TODAY()), otherwise the typed value.
// Dropdowns here REJECT anything not on their list, and Google checks that when it saves, not when we write.
// So a try/catch can't catch it. Instead: lift the dropdown rules, write, flush (save now), put the rules back.
// Cells whose value isn't on the list keep it, get a small red corner, and are listed at the end.
var IMPORT_PROBLEMS = [];
function copyBlock_(fromRange, toRange) {
  var values = fromRange.getValues(), formulas = fromRange.getFormulas();
  var out = values.map(function (row, r) {
    return row.map(function (v, c) { return formulas[r][c] ? formulas[r][c] : v; });
  });
  var rules = toRange.getDataValidations();
  var hasRules = rules.some(function (row) { return row.some(function (x) { return x; }); });
  if (hasRules) toRange.clearDataValidations();
  toRange.setValues(out);
  SpreadsheetApp.flush();                          // save the values while no dropdown can refuse them
  if (hasRules) {
    toRange.setDataValidations(rules);
    SpreadsheetApp.flush();
    findOffList_(toRange, out, rules);
  }
}

// Adds "Sheet!A1 (value)" to IMPORT_PROBLEMS for every cell whose value isn't on its dropdown list.
function findOffList_(range, values, rules) {
  var T = SpreadsheetApp.DataValidationCriteria, cache = {};
  values.forEach(function (row, r) {
    row.forEach(function (v, c) {
      var rule = rules[r][c];
      if (!rule || v === '' || v === null) return;
      var type = rule.getCriteriaType(), args = rule.getCriteriaValues(), list;
      if (type === T.VALUE_IN_LIST) list = args[0];
      else if (type === T.VALUE_IN_RANGE) {
        var key = args[0].getA1Notation();
        list = cache[key] || (cache[key] = [].concat.apply([], args[0].getValues()));
      } else return;
      var ok = list.some(function (x) { return String(x) === String(v); });
      if (!ok) IMPORT_PROBLEMS.push(range.getSheet().getName() + '!' + range.getCell(r + 1, c + 1).getA1Notation() + ' (' + v + ')');
    });
  });
}

function importFromOldCopy() {
  var ui = SpreadsheetApp.getUi();
  var ask = ui.prompt('Import my data',
    'Paste the link to your OLD Paycheck Planner copy (open it, copy the address bar).',
    ui.ButtonSet.OK_CANCEL);
  if (ask.getSelectedButton() !== ui.Button.OK) return;
  var url = ask.getResponseText().trim();

  var sure = ui.alert('Replace what is in THIS copy?',
    'Everything you typed in this copy will be replaced by your old copy\'s info. ' +
    'Your old copy is not changed. Continue?', ui.ButtonSet.YES_NO);
  if (sure !== ui.Button.YES) return;

  var oldSs;
  try {
    oldSs = SpreadsheetApp.openByUrl(url);
  } catch (e) {
    ui.alert('Could not open that link. Make sure it is YOUR old copy and the whole link was pasted.');
    return;
  }
  if (oldSs.getId() === SpreadsheetApp.getActive().getId()) {
    ui.alert('That is this same sheet. Paste the link to your OLD copy instead.');
    return;
  }

  var newSs = SpreadsheetApp.getActive();
  var copied = 0, skipped = [];
  IMPORT_PROBLEMS = [];
  IMPORT_JOBS.forEach(function (job) {
    var from = oldSs.getSheetByName(job.sheet), to = newSs.getSheetByName(job.sheet);
    var name = job.sheet + ': ' + (job.range || (job.label ? job.label[0] : (job.table || job.headers[0])));
    if (!from || !to) { skipped.push(name); return; }            // e.g. old copy has no Bills tab
    if (job.headers) {
      var a = findHeaders_(from, job.headers), b = findHeaders_(to, job.headers);
      if (!a || !b) { skipped.push(name); return; }
      job.headers.forEach(function (h) {
        if (!a.cols[h] || !b.cols[h]) return;                    // that column isn't in the old version: skip it
        copyBlock_(from.getRange(a.row + 1, a.cols[h], job.rows, 1), to.getRange(b.row + 1, b.cols[h], job.rows, 1));
      });
    } else if (job.range) {
      copyBlock_(from.getRange(job.range), to.getRange(job.range));
    } else if (job.label) {
      var rOld = findRow_(from, job.label), rNew = findRow_(to, job.label), col = job.col || 'C';
      if (rOld < 0 || rNew < 0) { skipped.push(name); return; }
      copyBlock_(from.getRange(col + rOld), to.getRange(col + rNew));
    } else if (job.byMonth) {
      // Month lists can start on different months in old and new copies (old: your first payday,
      // new: Jan 2026). So copy each month's number to the row with the SAME month, not the same position.
      var hO = findRow_(from, [job.table], job.after), hN = findRow_(to, [job.table], job.after);
      if (hO < 0 || hN < 0) { skipped.push(name); return; }
      var nO = Math.min(job.rows, from.getMaxRows() - hO), nN = Math.min(job.rows, to.getMaxRows() - hN);
      if (nO < 1 || nN < 1) { skipped.push(name); return; }
      var oldRows = from.getRange(hO + 1, 2, nO, 2).getValues();         // B month, C number
      var newMonths = to.getRange(hN + 1, 2, nN, 1).getValues();
      oldRows.forEach(function (o) {
        if (!(o[0] instanceof Date) || o[1] === '' || o[1] === null) return;
        for (var k = 0; k < newMonths.length; k++) {
          var d = newMonths[k][0];
          if (d instanceof Date && d.getFullYear() === o[0].getFullYear() && d.getMonth() === o[0].getMonth()) {
            to.getRange(hN + 1 + k, 3).setValue(o[1]);
            return;
          }
        }
      });
    } else {
      var hOld = findRow_(from, [job.table], job.after), hNew = findRow_(to, [job.table], job.after);
      if (hOld < 0 || hNew < 0) { skipped.push(name); return; }
      job.cols.forEach(function (c) {
        copyBlock_(from.getRange(c + (hOld + 1) + ':' + c + (hOld + job.rows)),
                   to.getRange(c + (hNew + 1) + ':' + c + (hNew + job.rows)));
      });
    }
    copied++;
  });

  ui.alert('Done!', 'Moved ' + copied + ' sections from your old copy.' +
    (skipped.length ? '\nNot in your old version (that\'s fine): ' + skipped.join(', ') : '') +
    (IMPORT_PROBLEMS.length ? '\n\nNot on this copy\'s dropdown list (kept, with a red corner, pick again): ' +
      IMPORT_PROBLEMS.slice(0, 12).join(', ') + (IMPORT_PROBLEMS.length > 12 ? ' and ' + (IMPORT_PROBLEMS.length - 12) + ' more' : '') : '') +
    '\nCheck your Dashboard: your numbers should look the same as before.', ui.ButtonSet.OK);
}


/* =====================================================================
 * OVERSPENDING ALERTS
 * ---------------------------------------------------------------------
 * The Dashboard already works out every alert (hidden columns X:Y).
 * X = the message ("" if it isn't happening), Y = 1 red / 2 yellow.
 * These functions only READ that list and show it to you.
 *
 * ELI5: the sheet is the smoke detector; this script is the beeper.
 *  - onEdit: runs by itself every time you type. If you just logged spending
 *    and something is now over budget, a pop-up appears at the bottom right.
 *  - Email me alerts every morning: Google runs checkAlertsAndEmail at ~8 AM
 *    and emails YOU (the sheet owner) only if there is something to know.
 * ===================================================================== */

function getAlerts_() {
  var ss = SpreadsheetApp.getActive();
  var dash = ss.getSheetByName('Dashboard');
  if (!dash) return { month: '', red: [], yellow: [] };
  var rows = dash.getRange('X2:Y' + dash.getLastRow()).getValues();
  var out = { month: dash.getRange('C2').getDisplayValue(), red: [], yellow: [] };
  rows.forEach(function (r) {
    if (!r[0]) return;
    (Number(r[1]) === 1 ? out.red : out.yellow).push(String(r[0]));
  });
  return out;
}

// Runs automatically when you edit the sheet (a "simple trigger", no setup needed).
function onEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() !== 'Spending') return;
    var col = e.range.getColumn();
    if (col < 2 || col > 5 || e.range.getRow() < 5) return;   // only the purchase columns B:E
    SpreadsheetApp.flush();                                   // let the formulas catch up first
    var a = getAlerts_();
    if (a.red.length) {
      e.source.toast(a.red.slice(0, 2).join('\n') + (a.red.length > 2 ? '\n+ ' + (a.red.length - 2) + ' more on the Dashboard' : ''),
                     '⚠ Over budget (' + a.month + ')', 10);
    } else if (a.yellow.length && /^Heads up: .* used, only/.test(a.yellow[0])) {
      e.source.toast(a.yellow[0], 'Heads up', 6);            // "almost out" warning
    }
  } catch (err) { /* never block typing because of an alert */ }
}

function showAlerts() {
  var a = getAlerts_();
  var all = a.red.concat(a.yellow);
  SpreadsheetApp.getUi().alert('Alerts for ' + a.month,
    all.length ? all.map(function (t) { return '• ' + t; }).join('\n\n') : 'No alerts. Your spending is on track. ✔',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function checkAlertsAndEmail() {
  var a = getAlerts_();
  if (!a.red.length && !a.yellow.length) return;              // nothing to say = no email
  var ss = SpreadsheetApp.getActive();
  var lines = [];
  if (a.red.length) lines.push('OVER BUDGET', a.red.map(function (t) { return '• ' + t; }).join('\n'), '');
  if (a.yellow.length) lines.push('HEADS UP', a.yellow.map(function (t) { return '• ' + t; }).join('\n'), '');
  lines.push('Open your planner: ' + ss.getUrl());
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
    'Paycheck Planner: ' + (a.red.length ? a.red.length + ' over budget' : a.yellow.length + ' heads up') + ' (' + a.month + ')',
    lines.join('\n'));
}

function installAlertEmails() {
  removeAlertEmails_(true);
  ScriptApp.newTrigger('checkAlertsAndEmail').timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('Done! Each morning around 8 AM you will get an email, only on days with an alert. ' +
    'Keep the Dashboard month on the current month.');
}

function removeAlertEmails() { removeAlertEmails_(false); }
function removeAlertEmails_(quiet) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkAlertsAndEmail') ScriptApp.deleteTrigger(t);
  });
  if (!quiet) SpreadsheetApp.getUi().alert('Alert emails are off.');
}


/* =====================================================================
 * SAVE THIS MONTH'S NET WORTH
 * ---------------------------------------------------------------------
 * A formula can't remember last month's number, so this copies today's
 * NET WORTH (Accounts tab) into this month's row of the log below it.
 * ELI5: taking a photo of your piggy bank once a month.
 * ===================================================================== */
function saveNetWorth() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Accounts');
  if (!sh) return;
  var colB = sh.getRange(1, 2, sh.getLastRow(), 1).getValues();
  var nwRow = -1;
  for (var i = 0; i < colB.length; i++) { if (colB[i][0] === 'NET WORTH') { nwRow = i + 1; break; } }
  if (nwRow < 0) return;
  var value = sh.getRange(nwRow, 6).getValue();               // column F
  var now = new Date();
  for (var r = nwRow + 1; r <= colB.length; r++) {
    var d = colB[r - 1][0];
    if (d instanceof Date && d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth()) {
      sh.getRange(r, 3).setValue(Math.round(value * 100) / 100);   // column C
      try { SpreadsheetApp.getActive().toast('Saved: ' + Utilities.formatString('$%.0f', value), 'Net worth'); } catch (e) {}
      return;
    }
  }
  try { SpreadsheetApp.getUi().alert("This month isn't in the net worth log (it runs Jan 2026 to Dec 2028). Type it in by hand."); } catch (e) {}
}


/* =====================================================================
 * SORT SPENDING BY DATE
 * ---------------------------------------------------------------------
 * Logged something late? This puts every purchase (and every Money in
 * entry) back in date order, oldest first, blanks at the bottom.
 * It sorts only the columns you type in. The hidden helper columns use
 * "this row" formulas, so they follow along by themselves.
 * Runs from Budget Tools → Sort Spending by date, and every time the sheet opens.
 * ===================================================================== */
function sortSpending() { sortTables_(['spend'], false); }     // Budget Tools → Sort Spending by date
function sortMoneyIn() { sortTables_(['in'], false); }        // Budget Tools → Sort Money in by date
function sortSpending_(quiet) { sortTables_(['spend', 'in'], quiet); }   // when the sheet opens: both

// Each table: [first header, last header, how many rows, toast message].
// The table runs from its first header's column to its last header's column, so every column moves together.
var SORT_TABLES = {
  spend: ['Date', 'Notes (optional)', 1000, 'Spending is in date order.'],
  'in': ['Date received', 'Notes (optional)', 300, 'Money in is in date order.']
};
function sortTables_(which, quiet) {
  var sh = SpreadsheetApp.getActive().getSheetByName('Spending');
  if (!sh) return;
  var done = [];
  which.forEach(function (k) {
    var t = SORT_TABLES[k];
    var h = findHeaders_(sh, [t[0], t[1]]);
    if (!h || !h.cols[t[1]]) return;
    var first = h.cols[t[0]], width = h.cols[t[1]] - first + 1;
    sh.getRange(h.row + 1, first, t[2], width).sort({ column: first, ascending: true });   // oldest first, blanks last
    done.push(t[3]);
  });
  if (!quiet && done.length) { try { SpreadsheetApp.getActive().toast(done.join(' '), 'Sorted'); } catch (e) {} }
}


/* =====================================================================
 * SEND MY DATA TO THE APP
 * ---------------------------------------------------------------------
 * Builds one file with everything you typed in this sheet (pay, groups,
 * categories, spending, money in, goals, accounts, bills, debts...) in the
 * Paycheck Planner app's format. Download it, then in the app:
 * Setup → Your data → Load backup.
 * ELI5: packing your notebook into a box the app knows how to unpack.
 * Nothing is uploaded anywhere: the file goes to YOUR computer/phone only.
 * ===================================================================== */
var APP_EARLY = ['No, on payday', '1 day early', '2 days early', '3 days early', '4 days early', '5 days early'];

function sendToApp() {
  var data = buildAppData_(SpreadsheetApp.getActive());
  var t = HtmlService.createTemplate(
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">' +
    '<p><b>Your data is ready for the app.</b> ' +
    '<?= counts ?></p>' +
    '<p><a id="dl" style="display:inline-block;background:#2F5D50;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">Download file</a></p>' +
    '<p>Then in the app: <b>Setup → Your data → Load backup</b> and pick the file ' +
    '(on a phone, it is in Downloads / Files).</p>' +
    '<p style="color:#6B7280">This replaces what is in the app now. The file stays on your device.</p></div>' +
    '<script>var d=<?!= json ?>;var b=new Blob([JSON.stringify(d,null,1)],{type:"application/json"});' +
    'var a=document.getElementById("dl");a.href=URL.createObjectURL(b);a.download="paycheck-planner-for-app.json";</script>');
  t.json = JSON.stringify(data).replace(/</g, '\\u003c');      // safe inside <script>
  t.counts = data.spending.length + ' purchases, ' + data.moneyIn.length + ' money in, ' +
    data.setup.categories.length + ' categories, ' + data.goals.length + ' goals, ' + data.debts.length + ' debts.';
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(420).setHeight(260), 'Send my data to the app');
}

// --- small helpers -----------------------------------------------------------------------
function iso_(d) {                                       // Date → "2026-10-22" ("" if not a date)
  if (!(d instanceof Date) || isNaN(d)) return '';
  var m = d.getMonth() + 1, day = d.getDate();
  return d.getFullYear() + '-' + (m < 10 ? '0' : '') + m + '-' + (day < 10 ? '0' : '') + day;
}
function num_(v) { var n = Number(v); return isFinite(n) && v !== '' && v !== null ? n : 0; }
function cell_(sh, r, c) { return sh.getRange(r, c, 1, 1).getValues()[0][0]; }
function byLabel_(sh, labels, col) {                     // the answer next to a question in column B
  var r = findRow_(sh, labels);
  return r < 0 ? '' : cell_(sh, r, col || 3);
}
function table_(sh, header, rows, after) {               // the rows under a header in column B (B onward)
  var h = findRow_(sh, [header], after);
  if (h < 0) return [];
  rows = Math.min(rows, sh.getMaxRows() - h);             // never read past the bottom of the tab
  if (rows < 1) return [];
  return sh.getRange(h + 1, 2, rows, 10).getValues();    // columns B..K
}
function columns_(sh, headers, rows) {                   // rows under headers found by name (any column)
  var f = findHeaders_(sh, headers);
  if (!f) return [];
  var out = [];
  var cols = headers.map(function (h) { return f.cols[h] ? sh.getRange(f.row + 1, f.cols[h], rows, 1).getValues() : null; });
  for (var i = 0; i < rows; i++) out.push(cols.map(function (c) { return c ? c[i][0] : ''; }));
  return out;
}

function buildAppData_(ss) {
  var setup = ss.getSheetByName('Setup'), sp = ss.getSheetByName('Spending'), goals = ss.getSheetByName('Goals');
  var acc = ss.getSheetByName('Accounts'), bills = ss.getSheetByName('Bills'), debts = ss.getSheetByName('Debts');
  var id = 0, nextId = function () { id++; return 'sheet' + id; };

  var s = {
    freq: String(byLabel_(setup, ['How often do you get paid?']) || 'Biweekly'),
    firstPay: num_(byLabel_(setup, ['First paycheck, if it was smaller (before taxes)'])),
    pay: num_(byLabel_(setup, ['Full paycheck, before taxes (ONE paycheck)', 'How much is each paycheck, before taxes?'])),
    pretax: num_(byLabel_(setup, ['Taken out before taxes, per paycheck'])),
    first: iso_(byLabel_(setup, ['When was your first payday this year?'])),
    payDays: String(byLabel_(setup, ['Paid twice a month? Which days?']) || 'From first payday'),
    weekend: String(byLabel_(setup, ['Payday on a weekend or bank holiday?', 'Payday lands on a weekend?']) || 'Workday before')
      .replace('Friday before', 'Workday before').replace('Monday after', 'Workday after'),
    early: Math.max(0, APP_EARLY.indexOf(String(byLabel_(setup, ['Does your bank pay you early?'])))),
    taxMode: byLabel_(setup, ['How should we figure your taxes?']) === 'Use my own rate' ? 'own' : 'estimate',
    status: String(byLabel_(setup, ['Filing status']) || 'Single'),
    state: num_(byLabel_(setup, ['State + local income tax rate'])),
    fica: byLabel_(setup, ['Take out Social Security + Medicare?']) !== 'No',
    ownRate: num_(byLabel_(setup, ['Your own total tax rate'])),
    inflMode: byLabel_(setup, ['Which inflation rate should we use?']) === 'My own rate' ? 'own' : 'auto',
    inflOwn: num_(byLabel_(setup, ['Your own inflation guess'])),
    plan: String(byLabel_(setup, ['Pick a budget plan']) || ''),
    groups: [], categories: []
  };
  table_(setup, 'Group name', 6).forEach(function (r) {   // B name, C kind, D own %, G grows with inflation
    if (r[0]) s.groups.push({ name: String(r[0]), kind: String(r[1]), pct: num_(r[2]), infl: r[5] === 'Yes' });
  });
  table_(setup, 'Category', 30).forEach(function (r) {    // B name, C group, D goal, E due day
    if (r[0]) s.categories.push({ name: String(r[0]), group: String(r[1]), goal: num_(r[2]), due: num_(r[3]) || '' });
  });

  var spending = [];
  columns_(sp, ['Date', 'What was it?', 'Category', 'Amount', "OK'd?", 'Cover with', 'Notes (optional)'], 1000).forEach(function (r) {
    if (!iso_(r[0]) || !num_(r[3])) return;
    var row = { id: nextId(), date: iso_(r[0]), what: String(r[1] || ''), cat: String(r[2] || ''), amt: num_(r[3]), notes: String(r[6] || '') };
    if (r[4]) row.ok = String(r[4]);
    if (r[5]) row.cover = String(r[5]).indexOf('leftover') >= 0 ? 'leftover' : 'extra';
    spending.push(row);
  });
  var moneyIn = [];
  columns_(sp, ['Date received', 'From (job, gig, gift...)', 'Amount', 'For (optional)', 'Notes (optional)'], 300).forEach(function (r) {
    if (!iso_(r[0]) || !num_(r[2])) return;
    var row = { id: nextId(), date: iso_(r[0]), from: String(r[1] || ''), amt: num_(r[2]), notes: String(r[4] || '') };
    if (r[3]) row.for = String(r[3]);
    moneyIn.push(row);
  });

  var data = { v: 1, setupDone: true, setup: s, spending: spending, moneyIn: moneyIn,
    goalsLeftover: false, trackFrom: '', goals: [], accountsReal: false, accounts: [], yearly: [], arrangements: [],
    debts: [], debtExtra: 0, debtMethod: 'snowball', nwOwn: 0, nwOwe: 0, nwLog: {} };

  if (goals) {
    data.goalsLeftover = cell_(goals, 8, 7) === 'Yes';
    var tf = cell_(goals, 9, 7);
    data.trackFrom = iso_(tf) ? iso_(tf).slice(0, 7) + '-01' : '';
    goals.getRange(12, 2, 8, 6).getValues().forEach(function (r) {   // B name, C target, D start, E end, F already, G counts
      if (r[0]) data.goals.push({ id: nextId(), name: String(r[0]), target: num_(r[1]), start: iso_(r[2]), end: iso_(r[3]),
        already: num_(r[4]), counts: String(r[5] || '') });
    });
  }
  if (acc) {
    data.accountsReal = cell_(acc, 8, 7) === 'Yes';
    acc.getRange(12, 2, 8, 8).getValues().forEach(function (r) {     // B..I
      if (r[0]) data.accounts.push({ id: nextId(), name: String(r[0]), type: String(r[1] || ''), where: String(r[2] || ''),
        balance: num_(r[3]), asOf: iso_(r[4]), rate: num_(r[5]), deposit: num_(r[6]), fedBy: String(r[7] || '') });
    });
    data.nwOwn = num_(byLabel_(acc, ['Other things you own (car value, cash at home...)'], 6));
    data.nwOwe = num_(byLabel_(acc, ['Other money you owe (not on the Debts tab)'], 6));
    table_(acc, 'Month', 36, 'NET WORTH').forEach(function (r) {      // B month, C net worth
      if (iso_(r[0]) && r[1] !== '' && r[1] !== null) data.nwLog[iso_(r[0]).slice(0, 7)] = Math.round(num_(r[1]));
    });
  }
  if (bills) {
    table_(bills, 'What', 8).forEach(function (r) {                  // B name, C next due, D cost, E saved
      if (r[0]) data.yearly.push({ id: nextId(), name: String(r[0]), due: iso_(r[1]), cost: num_(r[2]), saved: num_(r[3]) });
    });
    table_(bills, 'Who / what', 6).forEach(function (r) {            // B..G
      if (r[0]) data.arrangements.push({ id: nextId(), name: String(r[0]), total: num_(r[1]), payment: num_(r[2]),
        freq: String(r[3] || 'Every 2 weeks'), first: iso_(r[4]), paid: num_(r[5]) });
    });
  }
  if (debts) {
    data.debtExtra = num_(cell_(debts, 8, 6));
    data.debtMethod = String(cell_(debts, 9, 6)).indexOf('Avalanche') === 0 ? 'avalanche' : 'snowball';
    debts.getRange(12, 2, 8, 5).getValues().forEach(function (r) {   // B name, C type, D balance, E APR, F minimum
      if (r[0]) data.debts.push({ id: nextId(), name: String(r[0]), type: String(r[1] || ''), balance: num_(r[2]),
        apr: num_(r[3]), min: num_(r[4]) });
    });
  }
  return data;
}


/* =====================================================================
 * SORT CATEGORIES (Setup tab)
 * ---------------------------------------------------------------------
 * Sorts the yellow categories table (Category, Group, Monthly goal, Due day)
 * A to Z, or by group then A to Z. Empty rows go to the bottom.
 * Safe: everything else finds categories by NAME, not by row.
 * ===================================================================== */
function sortCategoriesAZ() { sortCategories_(false); }
function sortCategoriesByGroup() { sortCategories_(true); }
function sortCategories_(byGroup) {
  var sh = SpreadsheetApp.getActive().getSheetByName('Setup');
  if (!sh) return;
  var h = findRow_(sh, ['Category']);
  if (h < 0) return;
  var range = sh.getRange(h + 1, 2, 30, 4);                 // B..E, 30 category rows
  range.sort(byGroup ? [{ column: 3, ascending: true }, { column: 2, ascending: true }] : [{ column: 2, ascending: true }]);
  try { SpreadsheetApp.getActive().toast(byGroup ? 'Categories sorted by group, then A to Z.' : 'Categories sorted A to Z.', 'Sorted'); } catch (e) {}
}