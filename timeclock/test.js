// Checks the time clock logic with a fake spreadsheet. Run: node timeclock/test.js
const assert = require('assert'), fs = require('fs'), vm = require('vm');
function sheet(data){ return {
  data, getLastRow(){ return data.length; }, getLastColumn(){ return Math.max(...data.map(r=>r.length)); },
  getRange(r,c,nr=1,nc=1){ return {
    getValues:()=>data.slice(r-1,r-1+nr).map(row=>{ const a=row.slice(c-1,c-1+nc); while(a.length<nc)a.push(''); return a; }),
    setValue:v=>{ data[r-1][c-1]=v; }, setFormulaR1C1:f=>{ data[r-1][c-1]='(formula)'; },
    setValues:vs=>vs.forEach((row,i)=>row.forEach((v,j)=>{ data[r-1+i][c-1+j]=v; })) }; },
  appendRow(row){ data.push(row); return this; } }; }
const sheets = { Employees: sheet([['Name','PIN'],['Ana','0420'],['Bo','1111'],['Cy','2222']]), Shifts: sheet([['Name','Clock In','Clock Out','Hours']]),
  Sales: sheet([['Time','Employee','Items','Total']]), Crafting: sheet([['Time','Employee','Items']]),
  Inventory: sheet([['Item','In Stock','Crafted (all time)','Sold (all time)','Last Change'],['Across',0,0,0,''],['Lockpick',0,0,0,'']]) };
const inv = n => sheets.Inventory.data.find(r=>r[0]===n);
const cache = {};
const ctx = { SpreadsheetApp:{ getActive:()=>({ getSheetByName:n=>sheets[n] }) },
  CacheService:{ getScriptCache:()=>({ get:k=>cache[k]??null, put:(k,v)=>{cache[k]=v;}, remove:k=>{delete cache[k];} }) },
  LockService:{ getScriptLock:()=>({ waitLock(){}, releaseLock(){} }) },
  ContentService:{ MimeType:{JSON:'json'}, createTextOutput:t=>({ setMimeType:()=>JSON.parse(t) }) } };
vm.createContext(ctx); vm.runInContext(fs.readFileSync(process.argv[2] || __dirname + '/Code.gs','utf8'), ctx);
const post = b => ctx.doPost({postData:{contents:JSON.stringify(b)}});

assert.deepStrictEqual(ctx.doGet().employees, ['Ana','Bo','Cy']);
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

// sales (Ana is punched in at this point, Cy is not)
const order = [{name:'S5 Brakes', qty:2, unit:50000}, {name:'Racing Chip', qty:1, unit:22000}];
const s = post({name:'Ana', pin:'0420', action:'sale', items:order});
assert.strictEqual(s.ok, true); assert.strictEqual(s.total, 122000);
const row = sheets.Sales.data[1];
assert.strictEqual(row[1], 'Ana'); assert.strictEqual(row[2], '2x S5 Brakes, Racing Chip'); assert.strictEqual(row[3], 122000);
assert.match(post({name:'Cy', pin:'2222', action:'sale', items:order}).error, /must be punched in/);
assert.match(post({name:'Ana', pin:'9999', action:'sale', items:order}).error, /Wrong name or PIN/);
assert.match(post({name:'Ana', pin:'0420', action:'sale', items:[]}).error, /empty/);
assert.match(post({name:'Ana', pin:'0420', action:'sale', items:[{name:'X', qty:-3, unit:5}]}).error, /looks wrong/);
assert.match(post({name:'Ana', pin:'0420', action:'sale', items:[{name:'X', qty:1, unit:'lots'}]}).error, /looks wrong/);
assert.strictEqual(sheets.Sales.data.length, 2, 'rejected sales are not saved');

// crafting log
const c = post({name:'Ana', pin:'0420', action:'craft', items:[{name:'Lockpick', qty:10, unit:2000}, {name:'Harness', qty:1, unit:4000}]});
assert.strictEqual(c.ok, true);
assert.strictEqual(sheets.Crafting.data[1].slice(1).join(' | '), 'Ana | 10x Lockpick, Harness');
assert.strictEqual(sheets.Sales.data.length, 2, 'crafting does not touch Sales');
assert.match(post({name:'Cy', pin:'2222', action:'craft', items:order}).error, /must be punched in to log crafting/);
assert.strictEqual(sheets.Crafting.data.length, 2);

// inventory: crafting above added 10 Lockpick and a new Harness row; earlier sale of untracked items changed nothing
assert.deepStrictEqual(inv('Lockpick').slice(1,4).join(), '10,10,0');
assert.deepStrictEqual(inv('Harness').slice(1,4).join(), '1,1,0', 'crafting a new item adds its row');
assert.ok(!inv('S5 Brakes'), 'selling an untracked item adds no row');
post({name:'Ana', pin:'0420', action:'craft', items:[{name:'Across', qty:5, unit:2000}]});
post({name:'Ana', pin:'0420', action:'sale', items:[{name:'Across (After Dark)', qty:2, unit:1700}, {name:'Lockpick', qty:3, unit:2000}]});
assert.strictEqual(inv('Across').slice(1,4).join(), '3,5,2', 'After Dark across comes out of Across stock');
assert.strictEqual(inv('Lockpick').slice(1,4).join(), '7,10,3');
assert.strictEqual(sheets.Sales.data.at(-1)[2], '2x Across (After Dark), 3x Lockpick');
assert.strictEqual(sheets.Sales.data.at(-1)[3], 9400);
post({name:'Cy', pin:'2222', action:'sale', items:[{name:'Across', qty:1, unit:2000}]});
assert.strictEqual(inv('Across')[1], 3, 'rejected sale leaves stock alone');
console.log('all time clock checks passed');
