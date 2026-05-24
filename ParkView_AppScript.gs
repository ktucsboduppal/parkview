// ============================================================
// ParkView — Google Apps Script Backend
// Keshava's Thirumala Residency, Boduppal, Hyderabad
// Spreadsheet: 1vhDXWZ_RAwsh-miBtl4oxC91Mr41_CAm14Fn-O2ezM0
// ============================================================
// HOW TO DEPLOY:
//   1. Open your Google Sheet
//   2. Extensions → Apps Script
//   3. Delete any existing code, paste this entire file
//   4. Click Save (Ctrl+S)
//   5. Click Deploy → New deployment
//   6. Type: Web app
//   7. Execute as: Me
//   8. Who has access: Anyone
//   9. Click Deploy → Copy the Web App URL
//  10. Paste that URL into ParkView app (SCRIPT_URL constant)
// ============================================================

const SS_ID = '1yUyNfIPxR8mPQAqbuxr35e3o1lXZOeh3hqWAFNeqI4M';

// ── Sheet names ──────────────────────────────────────────────
const SHEETS = {
  profiles:    'Profiles',
  vehicles:    'Vehicles',
  requests:    'Requests',
  visitors:    'Visitors',
  violations:  'Violations',
  concerns:    'Concerns',
  rollcall:    'RollCall',
  rcEntries:   'RollCallEntries',
  unknowns:    'Unknowns',
  broadcasts:  'Broadcasts',
  summary:     'Summary',
};

// ── Column headers per sheet ─────────────────────────────────
const HEADERS = {
  profiles: [
    'id','name','flatNo','type','phone','pin',
    'roles','isVacant','isActive','movedOutAt','movedOutNote','createdAt'
  ],
  vehicles: [
    'id','profileId','flatNo','ownerName','type','plate',
    'make','model','colour','status','photoUrl','addedAt','updatedAt'
  ],
  requests: [
    'id','profileId','flatNo','resName','by','reqType',
    'vtype','oldPlate','plate','make','model','colour',
    'reason','photo','status','decidedBy','decidedAt','rejectReason',
    'cancelledReason','submittedAt'
  ],
  visitors: [
    'id','plate','vtype','flatNo','exp','notes',
    'by','createdAt','status'
  ],
  violations: [
    'id','flatNo','plate','description','assignedTo','reportedBy',
    'reportedAt','status','closedBy','closedAt','resolution'
  ],
  concerns: [
    'id','profileId','flatNo','resName','subject','description',
    'assignedTo','raisedAt','status','closedBy','closedAt',
    'resolution','reopenReason','reopenedAt'
  ],
  rollcall: [
    'id','date','startedAt','closedAt','guardName','outOfWindow',
    'totalVehicles','parkedCount','notInSlotCount',
    'visitorCount','unknownCount','durationMins','closed'
  ],
  rcEntries: [
    'sessionId','flatNo','resName','plate','vtype','status','parked'
  ],
  unknowns: [
    'id','plate','location','note','reportedBy','reportedAt',
    'assignedTo','status','closedBy','closedAt','resolution'
  ],
  broadcasts: [
    'id','title','body','targets','priority','pinned',
    'sentBy','sentByName','sentByRole','sentAt','readBy',
    'status','approvedBy','approvedAt','rejectedBy','rejectedAt',
    'rejectReason','editedAt'
  ],
};

// ════════════════════════════════════════════════════════════
// ENTRY POINTS
// ════════════════════════════════════════════════════════════

function doGet(e){
  return handleRequest(e);
}

function doPost(e){
  return handleRequest(e);
}

// ── Change this to a secret string and also set it in the app's SCRIPT_KEY constant ──
const API_KEY = 'ParkView2024Thirumala'; // Change this!

function handleRequest(e){
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const params = e.parameter || {};
    let body = {};
    if(e.postData && e.postData.contents){
      try{ body = JSON.parse(e.postData.contents); }catch{ body = {}; }
    }
    const action = params.action || body.action;

    // ── API Key check ──
    const key = params.key || body.key;
    if(action !== 'ping' && key !== API_KEY){
      output.setContent(JSON.stringify({ok:false, error:'Unauthorized'}));
      return output;
    }

    if(!action){
      output.setContent(JSON.stringify({ok:false, error:'No action specified'}));
      return output;
    }

    ensureSheets(); // Make sure all tabs exist with headers

    let result;
    switch(action){
      // ── READ ──
      case 'getAll':       result = getAllData();         break;
      case 'getProfiles':  result = getSheet('profiles'); break;
      case 'getVehicles':  result = getSheet('vehicles'); break;
      case 'getRequests':  result = getSheet('requests'); break;
      case 'getVisitors':  result = getSheet('visitors'); break;
      case 'getViolations':result = getSheet('violations');break;
      case 'getConcerns':  result = getSheet('concerns'); break;
      case 'getRollCall':  result = getRollCallData();    break;
      case 'getUnknowns':  result = getSheet('unknowns'); break;
      case 'getBroadcasts':result = getSheet('broadcasts');break;

      // ── WRITE (upsert by id) ──
      case 'saveProfile':   result = upsert('profiles',  body.data); break;
      case 'saveVehicle':   result = upsert('vehicles',  body.data); break;
      case 'saveRequest':   result = upsert('requests',  body.data); break;
      case 'saveVisitor':   result = upsert('visitors',  body.data); break;
      case 'saveViolation': result = upsert('violations',body.data); break;
      case 'saveConcern':   result = upsert('concerns',  body.data); break;
      case 'saveRollCall':  result = saveRollCall(body.data);         break;
      case 'saveUnknown':   result = upsert('unknowns',  body.data); break;
      case 'saveBroadcast': result = upsert('broadcasts',body.data); break;

      // ── BULK SAVE (array) ──
      case 'bulkSave':      result = bulkSave(body.sheet, body.rows); break;

      // ── FULL SYNC (app pushes everything at once) ──
      case 'fullSync':      result = fullSync(body);                   break;

      // ── PING ──
      case 'ping': result = {ok:true, ts:Date.now(), msg:'ParkView backend alive'}; break;

      default:
        result = {ok:false, error:'Unknown action: ' + action};
    }

    output.setContent(JSON.stringify(result));
  } catch(err){
    output.setContent(JSON.stringify({ok:false, error:err.toString()}));
  }

  return output;
}

// ════════════════════════════════════════════════════════════
// SHEET HELPERS
// ════════════════════════════════════════════════════════════

function getSpreadsheet(){
  return SpreadsheetApp.openById(SS_ID);
}

function ensureSheets(){
  const ss = getSpreadsheet();
  Object.entries(SHEETS).forEach(([key, name])=>{
    if(key === 'summary') return; // handled separately
    let sheet = ss.getSheetByName(name);
    if(!sheet){
      sheet = ss.insertSheet(name);
      // Write headers
      const headers = HEADERS[key];
      if(headers){
        sheet.getRange(1,1,1,headers.length).setValues([headers]);
        sheet.getRange(1,1,1,headers.length)
          .setBackground('#1E40AF')
          .setFontColor('#FFFFFF')
          .setFontWeight('bold');
        sheet.setFrozenRows(1);
      }
    }
  });
  ensureSummarySheet();
}

function getSheetObj(key){
  const ss = getSpreadsheet();
  return ss.getSheetByName(SHEETS[key]);
}

function sheetToObjects(sheet){
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row=>{
    const obj = {};
    headers.forEach((h,i)=> obj[h] = row[i]);
    return obj;
  }).filter(obj => obj.id); // skip empty rows
}

function getSheet(key){
  const sheet = getSheetObj(key);
  return {ok:true, data: sheetToObjects(sheet)};
}

// ── Upsert: insert if new id, update row if existing ──
function upsert(key, data){
  if(!data || !data.id) return {ok:false, error:'Missing data or id'};
  const sheet  = getSheetObj(key);
  if(!sheet) return {ok:false, error:'Sheet not found: '+key};
  const headers = HEADERS[key];
  if(!headers) return {ok:false, error:'No headers for: '+key};

  const allData = sheet.getDataRange().getValues();
  const idCol   = headers.indexOf('id');
  let   foundRow = -1;

  for(let i=1; i<allData.length; i++){
    if(String(allData[i][idCol]) === String(data.id)){
      foundRow = i+1; // 1-indexed
      break;
    }
  }

  const row = cleanRowForSheets(headers, data);

  if(foundRow > 0){
    sheet.getRange(foundRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  updateSummary();
  return {ok:true, id:data.id};
}

// ── Bulk save: overwrite all rows for a sheet ──
function bulkSave(key, rows){
  if(!key || !rows) return {ok:false, error:'Missing key or rows'};
  const sheet   = getSheetObj(key);
  if(!sheet)    return {ok:false, error:'Sheet not found: '+key};
  const headers = HEADERS[key];
  if(!headers)  return {ok:false, error:'No headers for: '+key};

  // Clear existing data (keep header row)
  const lastRow = sheet.getLastRow();
  if(lastRow > 1) sheet.getRange(2,1,lastRow-1,headers.length).clearContent();

  if(!rows.length){
    updateSummary();
    return {ok:true, count:0};
  }

  const values = rows.map(data => cleanRowForSheets(headers, data));

  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  updateSummary();
  return {ok:true, count:rows.length};
}

// ── Force correct boolean strings before writing to Sheets ──
function cleanRowForSheets(headers, data){
  const boolFields=['isVacant','isActive','closed','outOfWindow','pinned'];
  return headers.map(h => {
    const v = data[h];
    if(v === undefined || v === null) return '';
    if(boolFields.includes(h)){
      // Always write as explicit TRUE or FALSE string
      if(v===true||v===1||String(v).toUpperCase()==='TRUE') return 'TRUE';
      return 'FALSE'; // default to FALSE for boolean fields
    }
    if(Array.isArray(v)) return JSON.stringify(v);
    if(typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return v;
  });
}

// ── Smart merge: only overwrite a record if incoming timestamp is newer ──
function smartMerge(key, rows){
  if(!key || !rows || !rows.length) return {ok:true, merged:0, skipped:0};
  const sheet   = getSheetObj(key);
  if(!sheet)    return {ok:false, error:'Sheet not found: '+key};
  const headers = HEADERS[key];
  if(!headers)  return {ok:false, error:'No headers: '+key};

  const idCol = headers.indexOf('id');
  const tsCol = headers.findIndex(h=>h==='t'||h==='startedAt'||h==='raisedAt'||h==='submittedAt'||h==='reportedAt'||h==='createdAt'||h==='sentAt');

  const existing = sheet.getDataRange().getValues();
  const existMap = {};
  for(let i=1; i<existing.length; i++){
    const id = String(existing[i][idCol]);
    if(id) existMap[id] = {row:i+1, ts: tsCol>-1 ? Number(existing[i][tsCol])||0 : 0};
  }

  let merged=0, skipped=0;
  const toAppend=[];

  rows.forEach(data=>{
    if(!data.id) return;
    const id = String(data.id);
    const inTs = tsCol>-1 ? Number(data[headers[tsCol]])||0 : Date.now();
    const row = headers.map(h=>{
      const v=data[h];
      if(v===undefined||v===null) return '';
      if(Array.isArray(v)) return JSON.stringify(v);
      if(typeof v==='boolean') return v?'TRUE':'FALSE';
      return v;
    });
    if(existMap[id]){
      if(inTs >= existMap[id].ts){
        sheet.getRange(existMap[id].row, 1, 1, row.length).setValues([row]);
        merged++;
      } else { skipped++; }
    } else {
      toAppend.push(row);
      merged++;
    }
  });

  if(toAppend.length){
    sheet.getRange(sheet.getLastRow()+1, 1, toAppend.length, headers.length).setValues(toAppend);
  }
  return {ok:true, merged, skipped};
}

// ════════════════════════════════════════════════════════════
// ROLL CALL (session + entries combined)
// ════════════════════════════════════════════════════════════

function getRollCallData(){
  const sessions = sheetToObjects(getSheetObj('rollcall'));
  const entries  = sheetToObjects(getSheetObj('rcEntries'));

  // Attach entries to sessions
  const result = sessions.map(s=>{
    s.entries = entries.filter(e=> String(e.sessionId) === String(s.id));
    s.entries.forEach(e=>{
      e.parked = e.parked === 'TRUE' ? true : e.parked === 'FALSE' ? false : null;
    });
    s.closed      = s.closed === 'TRUE';
    s.outOfWindow = s.outOfWindow === 'TRUE';
    return s;
  });

  return {ok:true, data:result};
}

function saveRollCall(session){
  if(!session || !session.id) return {ok:false, error:'Missing session or id'};

  // Save session row
  upsert('rollcall', {
    id:             session.id,
    date:           new Date(session.t).toLocaleDateString('en-IN'),
    startedAt:      session.t,
    closedAt:       session.closedAt || '',
    guardName:      session.guardName,
    outOfWindow:    session.outOfWindow || false,
    totalVehicles:  (session.entries||[]).length,
    parkedCount:    (session.entries||[]).filter(e=>e.parked===true).length,
    notInSlotCount: (session.entries||[]).filter(e=>e.parked===false).length,
    visitorCount:   (session.entries||[]).filter(e=>e.status==='visitor').length,
    unknownCount:   (session.unknowns||[]).length,
    durationMins:   session.summary?.duration || 0,
    closed:         session.closed || false,
  });

  // Save entries — remove old ones for this session first, then insert fresh
  const entrySheet = getSheetObj('rcEntries');
  const entryHeaders = HEADERS.rcEntries;
  const allEntries = entrySheet.getDataRange().getValues();
  const sidCol = entryHeaders.indexOf('sessionId');

  // Delete rows belonging to this session (bottom-up to preserve indices)
  for(let i=allEntries.length-1; i>=1; i--){
    if(String(allEntries[i][sidCol]) === String(session.id)){
      entrySheet.deleteRow(i+1);
    }
  }

  // Insert fresh entries
  if(session.entries && session.entries.length){
    const rows = session.entries.map(e=> entryHeaders.map(h=>{
      if(h==='sessionId') return session.id;
      const v = e[h];
      if(v===undefined||v===null) return '';
      if(typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      return v;
    }));
    entrySheet.getRange(entrySheet.getLastRow()+1, 1, rows.length, entryHeaders.length).setValues(rows);
  }

  updateSummary();
  return {ok:true, id:session.id};
}

// ════════════════════════════════════════════════════════════
// FULL SYNC — app sends all localStorage data at once
// ════════════════════════════════════════════════════════════

function fullSync(body){
  const results = {};
  const map = {
    profiles:   'profiles',
    vehicles:   'vehicles',
    requests:   'requests',
    visitors:   'visitors',
    violations: 'violations',
    concerns:   'concerns',
    unknowns:   'unknowns',
    broadcasts: 'broadcasts',
  };

  Object.entries(map).forEach(([appKey, sheetKey])=>{
    if(body[appKey]){
      results[appKey] = smartMerge(sheetKey, body[appKey]);
    }
  });

  // If vehicles sent separately — merge them
  // If NOT sent separately — extract from profiles as fallback
  if(body.vehicles && body.vehicles.length){
    results.vehicles = smartMerge('vehicles', body.vehicles);
  } else if(body.profiles && body.profiles.length){
    // Extract vehicles from nested profiles (legacy format)
    const vRows = [];
    body.profiles.forEach(p=>{
      if(!p.vehicles || !p.vehicles.length) return;
      p.vehicles.forEach(v=>{
        vRows.push({
          id: v.id, profileId: p.id, flatNo: p.flatNo, ownerName: p.name,
          type: v.type, plate: v.plate||'', make: v.make||'', model: v.model||'',
          colour: v.colour||'', status: v.status||'active',
          photoUrl: '', addedAt: v.addedAt||'', updatedAt: ''
        });
      });
    });
    if(vRows.length) results.vehicles = smartMerge('vehicles', vRows);
  }

  // Roll call sessions saved individually
  if(body.rollcall && Array.isArray(body.rollcall)){
    results.rollcall = {saved:0};
    body.rollcall.forEach(s=>{
      saveRollCall(s);
      results.rollcall.saved++;
    });
  }

  updateSummary();
  return {ok:true, results};
}

// ════════════════════════════════════════════════════════════
// GET ALL DATA — one call to load everything into the app
// ════════════════════════════════════════════════════════════

function getAllData(){
  const profiles   = sheetToObjects(getSheetObj('profiles'));
  const vehicles   = sheetToObjects(getSheetObj('vehicles'));
  const requests   = sheetToObjects(getSheetObj('requests'));
  const visitors   = sheetToObjects(getSheetObj('visitors'));
  const violations = sheetToObjects(getSheetObj('violations'));
  const concerns   = sheetToObjects(getSheetObj('concerns'));
  const unknowns   = sheetToObjects(getSheetObj('unknowns'));
  const broadcasts = sheetToObjects(getSheetObj('broadcasts'));
  const rc         = getRollCallData();

  // Attach vehicles to profiles
  profiles.forEach(p=>{
    // Don't attach vehicles here — returned separately
    p.vehicles = [];
    // Handle all possible boolean representations from Sheets
    p.isVacant = (p.isVacant===true||String(p.isVacant||'').toUpperCase()==='TRUE'||p.isVacant===1);
    p.isActive = (p.isActive===true||String(p.isActive||'').toUpperCase()==='TRUE'||p.isActive===1);
    // Default: no movedOutAt = active and not vacant
    if(!p.movedOutAt||p.movedOutAt===''){p.isActive=true;p.isVacant=false;}
    p.roles = parseJSON(p.roles, []);
    delete p.pin; // Never send PIN over the wire
  });
  // Parse vehicle fields
  vehicles.forEach(v=>{
    v.status = v.status || 'active';
  });

  // Parse booleans / arrays in other sheets
  broadcasts.forEach(b=>{
    b.targets = parseJSON(b.targets, ['residents']);
    b.readBy  = parseJSON(b.readBy,  []);
    b.pinned  = b.pinned === 'TRUE';
  });

  visitors.forEach(v=>{
    v.exp = parseInt(v.exp) || 0;
  });

  return {
    ok: true,
    ts: Date.now(),
    profiles,
    vehicles, // returned separately so app can reattach to profiles
    requests,
    visitors,
    violations,
    concerns,
    unknowns,
    broadcasts,
    rollcall: rc.data,
  };
}

// ════════════════════════════════════════════════════════════
// SUMMARY SHEET — auto-calculated, read-only reference
// ════════════════════════════════════════════════════════════

function ensureSummarySheet(){
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.summary);
  if(!sheet){
    sheet = ss.insertSheet(SHEETS.summary);
    sheet.getRange('A1').setValue('ParkView — Live Summary');
    sheet.getRange('A1').setFontSize(14).setFontWeight('bold').setFontColor('#1E40AF');
  }
}

function updateSummary(){
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.summary) || ss.insertSheet(SHEETS.summary);
  sheet.clearContents();

  const profiles   = sheetToObjects(getSheetObj('profiles'));
  const vehicles   = sheetToObjects(getSheetObj('vehicles'));
  const visitors   = sheetToObjects(getSheetObj('visitors'));
  const violations = sheetToObjects(getSheetObj('violations'));
  const concerns   = sheetToObjects(getSheetObj('concerns'));
  const unknowns   = sheetToObjects(getSheetObj('unknowns'));
  const now        = new Date();

  const occupied  = profiles.filter(p=>p.isActive==='TRUE'&&p.isVacant!=='TRUE'&&p.type!=='security').length;
  const vacant    = 35 - occupied;
  const activeV   = vehicles.filter(v=>v.status==='active').length;
  const cars      = vehicles.filter(v=>v.status==='active'&&v.type==='car').length;
  const bikes     = vehicles.filter(v=>v.status==='active'&&v.type==='bike').length;
  const cycles    = vehicles.filter(v=>v.status==='active'&&v.type==='cycle').length;
  const pendingReq= profiles.filter(p=>p.status==='pending').length;
  const openVio   = violations.filter(v=>v.status==='open').length;
  const openCon   = concerns.filter(c=>c.status==='open').length;
  const openUnk   = unknowns.filter(u=>u.status==='open').length;
  const activeVis = visitors.filter(v=>v.status==='active').length;

  const rows = [
    ['ParkView — Live Summary', ''],
    ['Last updated', now.toLocaleString('en-IN')],
    ['', ''],
    ['── FLATS ──', ''],
    ['Total flats', 35],
    ['Occupied', occupied],
    ['Vacant', vacant],
    ['', ''],
    ['── VEHICLES ──', ''],
    ['Total registered', activeV],
    ['Cars', cars],
    ['Two Wheelers', bikes],
    ['Cycles', cycles],
    ['', ''],
    ['── TODAY ──', ''],
    ['Active visitor passes', activeVis],
    ['Open violations', openVio],
    ['Open concerns', openCon],
    ['Open unknowns', openUnk],
    ['', ''],
    ['── PENDING ──', ''],
    ['Vehicle requests pending', pendingReq],
  ];

  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.getRange(1,1).setFontSize(14).setFontWeight('bold').setFontColor('#1E40AF');
  sheet.getRange(2,2).setFontColor('#64748B');
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 140);

  // Colour code the summary values
  [[5,'#DCFCE7'],[6,'#DCFCE7'],[7,'#FFEDD5'],
   [10,'#DBEAFE'],[11,'#DBEAFE'],[12,'#DCFCE7'],[13,'#FEF3C7'],
   [16,'#FFFBEB'],[17,'#FEE2E2'],[18,'#FEF3C7'],[19,'#FEE2E2'],
   [22,'#F3E8FF']].forEach(([r,col])=>{
    sheet.getRange(r,2).setBackground(col);
  });
}

// ════════════════════════════════════════════════════════════
// UTILITIES
// ════════════════════════════════════════════════════════════

function parseJSON(val, fallback){
  if(!val || val === '') return fallback;
  try{ return JSON.parse(val); }catch{ return fallback; }
}

// ── Test function — run this manually to verify setup ──
function testSetup(){
  ensureSheets();
  const result = getAllData();
  Logger.log('Profiles: ' + result.profiles.length);
  Logger.log('Setup OK');
  return result;
}
