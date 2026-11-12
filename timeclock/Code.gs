/**
 * Benny's Time Clock: Google Apps Script backend for the site's Time Clock page.
 * Punches are saved in the Google Sheet this script is attached to.
 *
 * ONE-TIME SETUP
 * 1. Make a new Google Sheet (go to sheets.new) and name it "Benny's Time Clock".
 * 2. In the sheet: Extensions > Apps Script. Delete the sample code, paste this whole file, click Save.
 * 3. In the function dropdown at the top pick "setup", click Run, and allow access when Google asks.
 *    This creates the Employees, Shifts and Hours tabs.
 * 4. In the Employees tab, type each employee's name in column A and their PIN in column B.
 * 5. Deploy > New deployment > gear icon > Web app.
 *    Execute as: Me.  Who has access: Anyone.  Click Deploy.
 * 6. Copy the "Web app" URL (it ends in /exec) and put it in CLOCK_URL in index.html.
 *
 * If you change this script later: Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 * That keeps the same URL.
 *
 * TABS
 * Employees: Name | PIN. Add or delete rows to manage staff.
 * Shifts:    one row per shift (Name | Clock In | Clock Out | Hours). You can fix times by hand; Hours updates itself.
 * Hours:     per employee: hours this week (since Monday) and hours all time.
 */

const MAX_WRONG_PINS = 5;   // wrong PINs allowed per name...
const LOCKOUT_MINUTES = 10; // ...before that name is locked for this long
const HOURS_FORMULA = '=IF(RC[-1]="","",ROUND((RC[-1]-RC[-2])*24,2))';

function setup() {
  const ss = SpreadsheetApp.getActive();
  const tab = (name, header) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (!sh.getLastRow()) sh.appendRow(header).setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    return sh;
  };

  tab('Employees', ['Name', 'PIN']).getRange('B:B').setNumberFormat('@'); // text, so PINs like 0420 keep their 0
  tab('Shifts', ['Name', 'Clock In', 'Clock Out', 'Hours']).getRange('B:C').setNumberFormat('ddd m/d/yyyy h:mm am/pm');

  // one row per employee who has shifts; SUMIFS skips open shifts (blank Hours), and works on an empty sheet
  const hours = ss.getSheetByName('Hours') || ss.insertSheet('Hours');
  hours.clear();
  hours.getRange('A1:C1').setValues([['Employee', 'Hours this week (since Monday)', 'Hours all time']]).setFontWeight('bold');
  hours.setFrozenRows(1);
  hours.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(Shifts!A2:A,Shifts!A2:A<>""))),)');
  hours.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="",,SUMIFS(Shifts!D2:D,Shifts!A2:A,A2:A,Shifts!B2:B,">="&(TODAY()-WEEKDAY(TODAY(),3)))))');
  hours.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="",,SUMIFS(Shifts!D2:D,Shifts!A2:A,A2:A)))');

  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

// GET: employee names + who is on shift (no PINs ever leave the sheet)
function doGet() {
  return json(status());
}

// POST {name, pin, action: 'in' | 'out'}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    return json(punch(String(body.name || '').trim(), String(body.pin || '').trim(), body.action));
  } catch (err) {
    return json({ ok: false, error: 'Something went wrong. Try again.' });
  } finally {
    lock.releaseLock();
  }
}

function punch(name, pin, action) {
  if (action !== 'in' && action !== 'out') return { ok: false, error: 'Unknown action.' };

  const cache = CacheService.getScriptCache();
  const failKey = 'wrong-pin:' + name.toLowerCase();
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_WRONG_PINS) {
    return { ok: false, error: `Too many wrong PINs. Try again in ${LOCKOUT_MINUTES} minutes.` };
  }

  const emp = rows('Employees').find(r => String(r[0]).trim().toLowerCase() === name.toLowerCase());
  if (!emp || String(emp[1]).trim() !== pin) {
    if (emp) cache.put(failKey, String(fails + 1), LOCKOUT_MINUTES * 60);
    return { ok: false, error: 'Wrong name or PIN.' };
  }
  cache.remove(failKey);

  const who = String(emp[0]).trim();
  const shifts = SpreadsheetApp.getActive().getSheetByName('Shifts');
  const open = rows('Shifts').findIndex(r => String(r[0]).trim() === who && r[1] && !r[2]);

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
