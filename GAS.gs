// Spendly Pro -- Apps Script Backend
// 1. Open your Google Sheet → Extensions → Apps Script
// 2. Delete all code, paste this, Save
// 3. Deploy → New Deployment → Web App
//    Execute as: Me | Who has access: Anyone
// 4. Copy the /exec URL into Spendly

const SPREADSHEET_ID = '1IUfA7f4o8QoZ6LZqwaLpLfj16v39psGY3oQp04OKcVA';
const SHEET_NAME = 'Spendly';

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  if (!e) return json({ ok: false, error: 'Run via Web App URL only' });
  let action = null, params = {};
  try {
    if (e.parameter && e.parameter.action) { action = e.parameter.action; params = e.parameter; }
    else if (e.postData && e.postData.contents) { const p = JSON.parse(e.postData.contents); action = p.action; params = p; }
  } catch(err) { return json({ ok: false, error: 'Parse error: ' + err.message }); }
  try {
    if (action === 'ping') return json({ ok: true, sheet: getSheet().getName(), ts: new Date().toISOString() });
    if (action === 'push') { const del = params.deletedIds || []; if (del.length) deleteRows(del); const stats = writeAll(params.txns || []); if (params.budgets) writeBudgets(params.budgets); if (params.customCats) writeCustomCats(params.customCats); if (params.holdings) writeHoldings(params.holdings); if (params.dividends) writeDividends(params.dividends); if (params.recurringTxns) writeRecurring(params.recurringTxns); if (params.autoInvest) writeAutoInvest(params.autoInvest); if (params.pantryItems) writePantry(params.pantryItems); if (params.projects) writeProjects(params.projects); return json({ ok: true, ...stats }); }
    if (action === 'pull') { return json({ ok: true, txns: readAll(), budgets: readBudgets(), customCats: readCustomCats(), holdings: readHoldings(), dividends: readDividends(), recurringTxns: readRecurring(), autoInvest: readAutoInvest(), pantryItems: readPantry(), projects: readProjects() }); }
    return json({ ok: false, error: 'Unknown action: ' + action });
  } catch(err) { return json({ ok: false, error: err.message }); }
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id','date','type','amount','category','description','synced_at']);
    sh.getRange(1,1,1,7).setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('center');
  }
  return sh;
}

function deleteRows(ids) {
  const sh = getSheet(); const last = sh.getLastRow();
  if (last < 2) return;
  const idSet = new Set(ids.map(String));
  const vals = sh.getRange(2,1,last-1,1).getValues();
  const toDelete = [];
  vals.forEach((r,i) => { if (idSet.has(String(r[0]))) toDelete.push(i+2); });
  toDelete.reverse().forEach(n => sh.deleteRow(n));
}

function writeAll(txns) {
  if (!txns || !txns.length) return { written: 0, updated: 0 };
  const sh = getSheet(); const last = sh.getLastRow();
  const ts = new Date().toLocaleString();
  const existingIds = {};
  if (last > 1) sh.getRange(2,1,last-1,1).getValues().forEach((r,i) => { if (r[0]) existingIds[String(r[0])] = i+2; });
  const newRows = []; let updated = 0;
  txns.forEach(t => {
    if (!t.id || !t.amount) return;
    const row = [String(t.id), (t.date||'').slice(0,10), t.type||'expense', Number(t.amount)||0, t.cat||'other', t.desc||'', ts];
    if (existingIds[String(t.id)]) { sh.getRange(existingIds[String(t.id)],1,1,7).setValues([row]); updated++; }
    else newRows.push(row);
  });
  if (newRows.length) sh.getRange(sh.getLastRow()+1,1,newRows.length,7).setValues(newRows);
  const fl = sh.getLastRow();
  if (fl > 1) sh.getRange(2,1,fl-1,7).sort({ column: 2, ascending: false });
  return { written: newRows.length, updated };
}

function readAll() {
  const sh = getSheet(); const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2,1,last-1,7).getValues().filter(r => r[0]).map(r => {
    let d = r[1];
    if (d instanceof Date) { d = d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
    return { id: String(r[0]), date: String(d), type: r[2], amount: Number(r[3]), cat: r[4], desc: r[5] };
  });
}

function writeBudgets(budgets) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('Spendly_Budgets');
  if (!sh) {
    sh = ss.insertSheet('Spendly_Budgets');
    sh.appendRow(['cat', 'limit']);
    sh.getRange(1,1,1,2).setFontWeight('bold').setBackground('#f3f3f3');
  }
  const last = sh.getLastRow();
  sh.setFrozenRows(0); if (last > 1) sh.deleteRows(2, last - 1);
  if (budgets.length) {
    sh.getRange(2,1,budgets.length,2).setValues(budgets.map(b=>[b.cat, b.limit]));
  }
}

function readBudgets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Spendly_Budgets');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,2).getValues()
    .filter(r=>r[0])
    .map(r=>({ cat: String(r[0]), limit: Number(r[1]) }));
}

function writeCustomCats(cats) {
  if (!cats || !cats.length) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('Spendly_Categories');
  if (!sh) {
    sh = ss.insertSheet('Spendly_Categories');
    sh.appendRow(['id','emoji','label','color','type']);
    sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#f3f3f3');
  }
  const last = sh.getLastRow();
  const existingIds = {};
  if (last > 1) {
    sh.getRange(2,1,last-1,1).getValues().forEach(function(r,i){
      if (r[0]) existingIds[String(r[0])] = i + 2;
    });
  }
  const newRows = [];
  cats.forEach(function(c) {
    if (!c.id) return;
    const row = [String(c.id), c.emoji||'package', c.label||'', c.color||'#6b7280', c.type||'both'];
    if (existingIds[String(c.id)]) {
      sh.getRange(existingIds[String(c.id)], 1, 1, 5).setValues([row]);
    } else {
      newRows.push(row);
    }
  });
  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
  }
}

function readCustomCats() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Spendly_Categories');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,5).getValues()
    .filter(r => r[0])
    .map(r => ({ id:String(r[0]), emoji:r[1]||'package', label:String(r[2]), color:r[3]||'#6b7280', type:r[4]||'both' }));
}

function writeHoldings(holdings) {
  if (!holdings) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('Spendly_Holdings');
  if (!sh) {
    sh = ss.insertSheet('Spendly_Holdings');
    sh.appendRow(['ticker','name','units','avgPrice','currentPrice','color']);
    sh.getRange(1,1,1,6).setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('center');
  }
  const last = sh.getLastRow();
  sh.setFrozenRows(0); if (last > 1) sh.deleteRows(2, last - 1);
  if (holdings.length) {
    sh.getRange(2,1,holdings.length,6).setValues(
      holdings.map(function(h){ return [h.ticker||'', h.name||'', Number(h.units)||0, Number(h.avgPrice)||0, Number(h.currentPrice)||0, h.color||'#818cf8']; })
    );
  }
}

function readHoldings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Spendly_Holdings');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,6).getValues()
    .filter(function(r){ return r[0]; })
    .map(function(r){ return { ticker:String(r[0]), name:String(r[1]), units:Number(r[2]), avgPrice:Number(r[3]), currentPrice:Number(r[4]), color:String(r[5]) }; });
}

function writeDividends(divs) {
  if (!divs) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('Spendly_Dividends');
  if (!sh) {
    sh = ss.insertSheet('Spendly_Dividends');
    sh.appendRow(['ticker','amount','date','notes']);
    sh.getRange(1,1,1,4).setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('center');
  }
  const last = sh.getLastRow();
  sh.setFrozenRows(0); if (last > 1) sh.deleteRows(2, last - 1);
  if (divs.length) {
    sh.getRange(2,1,divs.length,4).setValues(
      divs.map(function(d){ return [d.ticker||'', Number(d.amount)||0, d.date||'', d.notes||'']; })
    );
  }
}

function readDividends() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Spendly_Dividends');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,4).getValues()
    .filter(function(r){ return r[0]; })
    .map(function(r){
      var d = r[2];
      if (d instanceof Date) { d = d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
      return { ticker:String(r[0]), amount:Number(r[1]), date:String(d), notes:String(r[3]) };
    });
}

function writeRecurring(items) {
  if (!items) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('Spendly_Recurring');
  if (!sh) {
    sh = ss.insertSheet('Spendly_Recurring');
    sh.appendRow(['name','type','amount','cat','freq','startDate','currency']);
    sh.getRange(1,1,1,7).setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('center');
  }
  const last = sh.getLastRow();
  sh.setFrozenRows(0); if (last > 1) sh.deleteRows(2, last - 1);
  if (items.length) {
    sh.getRange(2,1,items.length,7).setValues(
      items.map(function(r){ return [r.name||'', r.type||'expense', Number(r.amount)||0, r.cat||'other', r.freq||'monthly', r.startDate||'', r.currency||'AUD']; })
    );
  }
}

function readRecurring() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Spendly_Recurring');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,7).getValues()
    .filter(function(r){ return r[0]; })
    .map(function(r){
      var d = r[5];
      if (d instanceof Date) { d = d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
      return { name:String(r[0]), type:String(r[1]), amount:Number(r[2]), cat:String(r[3]), freq:String(r[4]), startDate:String(d), currency:String(r[6]||'AUD') };
    });
}

function writeAutoInvest(plan) {
  if (!plan || !plan.totalAmount) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('Spendly_AutoInvest');
  if (!sh) {
    sh = ss.insertSheet('Spendly_AutoInvest');
    sh.appendRow(['field', 'value']);
    sh.getRange(1,1,1,2).setFontWeight('bold').setBackground('#f3f3f3');
  }
  const last = sh.getLastRow();
  sh.setFrozenRows(0); if (last > 1) sh.deleteRows(2, last - 1);
  const splitsJson = JSON.stringify(plan.splits || []);
  sh.getRange(2,1,4,2).setValues([
    ['totalAmount', Number(plan.totalAmount)||0],
    ['startDate',   String(plan.startDate||'')],
    ['interval',    Number(plan.interval)||14],
    ['splits',      splitsJson]
  ]);
}

function readAutoInvest() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Spendly_AutoInvest');
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
  const map = {};
  rows.forEach(function(r){ if(r[0]) map[String(r[0])] = r[1]; });
  if (!map.totalAmount) return null;
  let splits = [];
  try { splits = JSON.parse(map.splits || '[]'); } catch(e) {}
  return { totalAmount: Number(map.totalAmount)||0, startDate: String(map.startDate||''), interval: Number(map.interval)||14, splits: splits };
}

function writePantry(items) {
  if (!items) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('Spendly_Pantry');
  if (!sh) {
    sh = ss.insertSheet('Spendly_Pantry');
    sh.appendRow(['name','emoji','qty','unit','min','max','category','expiry','notes']);
    sh.getRange(1,1,1,9).setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('center');
  }
  const last = sh.getLastRow();
  sh.setFrozenRows(0); if (last > 1) sh.deleteRows(2, last - 1);
  if (items.length) {
    sh.getRange(2,1,items.length,9).setValues(
      items.map(function(p){ return [p.name||'', p.emoji||'', Number(p.qty)||0, p.unit||'pcs', Number(p.min)||1, Number(p.max)||0, p.category||'pantry', p.expiry||'', p.notes||'']; })
    );
  }
}

function readPantry() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Spendly_Pantry');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,9).getValues()
    .filter(function(r){ return r[0]; })
    .map(function(r){
      return { name:String(r[0]), emoji:String(r[1]), qty:Number(r[2])||0, unit:String(r[3]||'pcs'), min:Number(r[4])||1, max:Number(r[5])||0, category:String(r[6]||'pantry'), expiry:String(r[7]||''), notes:String(r[8]||'') };
    });
}

function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function writeProjects(items) {
  if (!items || !items.length) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('Spendly_Projects');
  if (!sh) { sh = ss.insertSheet('Spendly_Projects'); }
  const headers = ['id','emoji','name','desc','start','end','budget','status','catBudgets'];
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#c8a96e');
  const last = sh.getLastRow();
  sh.setFrozenRows(0); if (last > 1) sh.deleteRows(2, last - 1);
  sh.getRange(2,1,items.length,headers.length).setValues(
    items.map(function(p){ return [
      String(p.id||''), String(p.emoji||''), String(p.name||''), String(p.desc||''),
      String(p.start||''), String(p.end||''), Number(p.budget)||0, String(p.status||'planned'),
      JSON.stringify(p.catBudgets||[])
    ]; })
  );
  sh.autoResizeColumns(1, headers.length);
}

function readProjects() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Spendly_Projects');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,9).getValues()
    .filter(function(r){ return r[0] && r[2]; })
    .map(function(r){
      var catBudgets = [];
      try { catBudgets = JSON.parse(r[8]||'[]'); } catch(e){}
      return { id:String(r[0]), emoji:String(r[1]||''), name:String(r[2]),
        desc:String(r[3]||''), start:String(r[4]||''), end:String(r[5]||''),
        budget:Number(r[6])||0, status:String(r[7]||'planned'), catBudgets:catBudgets };
    });
}
