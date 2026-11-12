/**
 * Benny's Time Clock: Google Apps Script backend for the site's Time Clock and Checkout pages.
 * Punches and sales are saved in the Google Sheet this script is attached to.
 *
 * ONE-TIME SETUP
 * 1. Make a new Google Sheet (go to sheets.new) and name it "Benny's Time Clock".
 * 2. In the sheet: Extensions > Apps Script. Delete the sample code, paste this whole file, click Save.
 * 3. In the function dropdown at the top pick "setup", click Run, and allow access when Google asks.
 *    This creates the Employees, Shifts, Hours, Sales and Sales Totals tabs.
 * 4. In the Employees tab, type each employee's name in column A and their PIN in column B.
 * 5. Deploy > New deployment > gear icon > Web app.
 *    Execute as: Me.  Who has access: Anyone.  Click Deploy.
 * 6. Copy the "Web app" URL (it ends in /exec) and put it in CLOCK_URL in index.html.
 *
 * If you change this script later: paste it in, Save, run "setup" again (it keeps your data),
 * then Deploy > Manage deployments > pencil icon > Version: New version > Deploy. That keeps the same URL.
 *
 * TABS
 * Employees:    Name | PIN. Add or delete rows to manage staff.
 * Shifts:       one row per shift (Name | Clock In | Clock Out | Hours). You can fix times by hand; Hours updates itself.
 * Hours:        per employee: hours this week (since Monday) and hours all time.
 * Sales:        one row per completed sale (Time | Employee | Items | Total).
 * Sales Totals: per employee: sales this week (since Monday) and sales all time.
 */

const MAX_WRONG_PINS = 5;   // wrong PINs allowed per name...
const LOCKOUT_MINUTES = 10; // ...before that name is locked for this long
const HOURS_FORMULA = '=IF(RC[-1]="","",ROUND((RC[-1]-RC[-2])*24,2))';
const MONDAY = '(TODAY()-WEEKDAY(TODAY(),3))';
const DATE_FORMAT = 'ddd m/d/yyyy h:mm am/pm';

function setup() {
  const ss = SpreadsheetApp.getActive();
  const tab = (name, header) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (!sh.getLastRow()) sh.appendRow(header).setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    return sh;
  };
  // per-employee totals tab, rebuilt each time (formulas only, no data of its own)
  const totals = (name, headers, src, nameCol, valueCol, timeCol, format) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear();
    sh.getRange('A1:C1').setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    const n = `${src}!${nameCol}2:${nameCol}`, v = `${src}!${valueCol}2:${valueCol}`, t = `${src}!${timeCol}2:${timeCol}`;
    sh.getRange('A2').setFormula(`=IFERROR(SORT(UNIQUE(FILTER(${n},${n}<>""))),)`);
    sh.getRange('B2').setFormula(`=ARRAYFORMULA(IF(A2:A="",,SUMIFS(${v},${n},A2:A,${t},">="&${MONDAY})))`);
    sh.getRange('C2').setFormula(`=ARRAYFORMULA(IF(A2:A="",,SUMIFS(${v},${n},A2:A)))`);
    if (format) sh.getRange('B:C').setNumberFormat(format);
  };

  tab('Employees', ['Name', 'PIN']).getRange('B:B').setNumberFormat('@'); // text, so PINs like 0420 keep their 0
  tab('Shifts', ['Name', 'Clock In', 'Clock Out', 'Hours']).getRange('B:C').setNumberFormat(DATE_FORMAT);
  const sales = tab('Sales', ['Time', 'Employee', 'Items', 'Total']);
  sales.getRange('A:A').setNumberFormat(DATE_FORMAT);
  sales.getRange('D:D').setNumberFormat('$#,##0');
  sales.setColumnWidth(3, 420);

  totals('Hours', ['Employee', 'Hours this week (since Monday)', 'Hours all time'], 'Shifts', 'A', 'D', 'B');
  totals('Sales Totals', ['Employee', 'Sales this week (since Monday)', 'Sales all time'], 'Sales', 'B', 'D', 'A', '$#,##0');

  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

// GET: employee names + who is on shift (no PINs ever leave the sheet)
function doGet() {
  return json(status());
}

// POST {name, pin, action: 'in' | 'out'}  or  {name, pin, action: 'sale', items: [{name, qty, unit}]}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (!['in', 'out', 'sale'].includes(body.action)) return json({ ok: false, error: 'Unknown action.' });
    const who = checkPin(String(body.name || '').trim(), String(body.pin || '').trim());
    if (who.error) return json({ ok: false, error: who.error });
    return json(body.action === 'sale' ? sale(who.name, body.items) : punch(who.name, body.action));
  } catch (err) {
    return json({ ok: false, error: 'Something went wrong. Try again.' });
  } finally {
    lock.releaseLock();
  }
}

// returns {name} (as written in the Employees tab) or {error}; locks a name after too many wrong PINs
function checkPin(name, pin) {
  const cache = CacheService.getScriptCache();
  const failKey = 'wrong-pin:' + name.toLowerCase();
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_WRONG_PINS) return { error: `Too many wrong PINs. Try again in ${LOCKOUT_MINUTES} minutes.` };

  const emp = rows('Employees').find(r => String(r[0]).trim().toLowerCase() === name.toLowerCase());
  if (!emp || String(emp[1]).trim() !== pin) {
    if (emp) cache.put(failKey, String(fails + 1), LOCKOUT_MINUTES * 60);
    return { error: 'Wrong name or PIN.' };
  }
  cache.remove(failKey);
  return { name: String(emp[0]).trim() };
}

function openShiftRow(who) {
  return rows('Shifts').findIndex(r => String(r[0]).trim() === who && r[1] && !r[2]);
}

function punch(who, action) {
  const shifts = SpreadsheetApp.getActive().getSheetByName('Shifts');
  const open = openShiftRow(who);
  if (action === 'in') {
    if (open >= 0) return Object.assign(status(), { ok: false, error: `${who} is already punched in.` });
    shifts.appendRow([who, new Date(), '', '']);
    shifts.getRange(shifts.getLastRow(), 4).setFormulaR1C1(HOURS_FORMULA);
  } else {
    if (open < 0) return Object.assign(status(), { ok: false, error: `${who} isn't punched in.` });
    shifts.getRange(open + 2, 3).setValue(new Date());
  }
  return Object.assign(status(), { ok: true, message: `${who} punched ${action}.` });
}

// ponytail: prices come from the site (the menu lives there), so totals are only as honest as the employee.
// Move the price list into the sheet if that ever matters.
function sale(who, items) {
  if (openShiftRow(who) < 0) return { ok: false, error: `${who} must be punched in to complete a sale.` };
  if (!Array.isArray(items) || !items.length) return { ok: false, error: 'This order is empty.' };
  if (items.length > 100) return { ok: false, error: 'Too many items in one sale.' };

  const parts = [];
  let total = 0;
  for (const it of items) {
    const name = String((it && it.name) || '').trim().slice(0, 60);
    const qty = Math.floor(Number(it && it.qty));
    const unit = Number(it && it.unit);
    if (!name || !(qty >= 1 && qty <= 100000) || !(unit >= 0 && unit <= 10000000)) {
      return { ok: false, error: 'Something in this order looks wrong. Refresh and try again.' };
    }
    total += qty * unit;
    parts.push(qty === 1 ? name : `${qty}x ${name}`);
  }
  SpreadsheetApp.getActive().getSheetByName('Sales').appendRow([new Date(), who, parts.join(', '), total]);
  return { ok: true, total, message: `Sale saved: $${total.toLocaleString('en-US')} by ${who}.` };
}

function status() {
  const employees = rows('Employees').map(r => String(r[0]).trim()).filter(Boolean);
  const onShift = rows('Shifts')
    .filter(r => r[0] && r[1] && !r[2])
    .map(r => ({ name: String(r[0]).trim(), since: new Date(r[1]).getTime() }));
  return { ok: true, employees, onShift };
}

// data rows of a tab (skips the header row)
function rows(tabName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, sh.getLastColumn()).getValues() : [];
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
