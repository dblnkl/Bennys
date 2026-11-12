// Checks the time clock logic with a fake spreadsheet. Run: node timeclock/test.js
const assert = require('assert'), fs = require('fs'), vm = require('vm');
function sheet(data){ return {
  data, getLastRow(){ return data.length; }, getLastColumn(){ return Math.max(...data.map(r=>r.length)); },
  getRange(r,c,nr=1,nc=1){ return {
    getValues:()=>data.slice(r-1,r-1+nr).map(row=>{ const a=row.slice(c-1,c-1+nc); while(a.length<nc)a.push(''); return a; }),
    setValue:v=>{ data[r-1][c-1]=v; }, setFormulaR1C1:f=>{ data[r-1][c-1]='(formula)'; } }; },
  appendRow(row){ data.push(row); return this; } }; }
const sheets = { Employees: sheet([['Name','PIN'],['Ana','0420'],['Bo','1111']]), Shifts: sheet([['Name','Clock In','Clock Out','Hours']]) };
const cache = {};
const ctx = { SpreadsheetApp:{ getActive:()=>({ getSheetByName:n=>sheets[n] }) },
  CacheService:{ getScriptCache:()=>({ get:k=>cache[k]??null, put:(k,v)=>{cache[k]=v;}, remove:k=>{delete cache[k];} }) },
  LockService:{ getScriptLock:()=>({ waitLock(){}, releaseLock(){} }) },
  ContentService:{ MimeType:{JSON:'json'}, createTextOutput:t=>({ setMimeType:()=>JSON.parse(t) }) } };
vm.createContext(ctx); vm.runInContext(fs.readFileSync(process.argv[2] || __dirname + '/Code.gs','utf8'), ctx);
const post = b => ctx.doPost({postData:{contents:JSON.stringify(b)}});

assert.deepStrictEqual(ctx.doGet().employees, ['Ana','Bo']);
assert.strictEqual(post({name:'ana', pin:'0420', action:'in'}).ok, true, 'case-insensitive name, PIN with leading 0');
assert.strictEqual(ctx.doGet().onShift[0].name, 'Ana');
assert.match(post({name:'Ana', pin:'0420', action:'in'}).error, /already punched in/);
assert.match(post({name:'Bo', pin:'1111', action:'out'}).error, /isn't punched in/);
assert.strictEqual(post({name:'Ana', pin:'0420', action:'out'}).ok, true);
assert.notStrictEqual(sheets.Shifts.data[1][2], '', 'clock-out time written');
assert.strictEqual(ctx.doGet().onShift.length, 0);
assert.match(post({name:'Nobody', pin:'1', action:'in'}).error, /Wrong name or PIN/);
assert.match(post({name:'Ana', pin:'0420', action:'dance'}).error, /Unknown action/);
for (let i=0;i<5;i++) assert.match(post({name:'Bo', pin:'0000', action:'in'}).error, /Wrong name or PIN/);
assert.match(post({name:'Bo', pin:'1111', action:'in'}).error, /Too many wrong PINs/, 'locked even with right PIN');
assert.strictEqual(post({name:'Ana', pin:'0420', action:'in'}).ok, true, 'other employees not locked');
assert.ok(!JSON.stringify(ctx.doGet()).includes('0420'), 'PINs never sent out');
console.log('all time clock checks passed');
