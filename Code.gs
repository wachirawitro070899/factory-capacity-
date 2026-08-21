const SPREADSHEET_ID = '1eXby1xmCjhp_C8H_r7OC8JmnLu00WRYq';

// These tabs are not treated as machines. Add helper/master tab names here if needed.
const EXCLUDED_SHEETS = [
  'Department_Master', 'Capacity_Data', 'Settings', 'Config', 'Master', 'README', 'Template'
];

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'capacity').toLowerCase();
    if (action === 'capacity') return output_(getCapacityData_(), e);
    if (action === 'sheets') return output_(getMachineSheets_(), e);
    return output_({ok:false,error:'Unknown action: '+action}, e);
  } catch (err) {
    return output_({ok:false,error:String(err && err.message ? err.message : err)}, e);
  }
}

function getCapacityData_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = getMachineSheetObjects_(ss);
  const records = [];
  const warnings = [];

  sheets.forEach(function(sheet) {
    try {
      const parsed = parseMachineSheet_(sheet);
      parsed.records.forEach(function(r){ records.push(r); });
      if (parsed.warning) warnings.push(parsed.warning);
    } catch (err) {
      warnings.push(sheet.getName()+': '+String(err && err.message ? err.message : err));
    }
  });

  return {
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    machineSheets: sheets.map(function(s){return s.getName();}),
    records: records,
    warnings: warnings,
    generatedAt: new Date().toISOString()
  };
}

function getMachineSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = getMachineSheetObjects_(ss);
  return {ok:true,machineSheets:sheets.map(function(s){return s.getName();}),generatedAt:new Date().toISOString()};
}

function getMachineSheetObjects_(ss) {
  return ss.getSheets().filter(function(sheet) {
    const name = String(sheet.getName() || '').trim();
    if (!name) return false;
    if (name.charAt(0) === '_') return false;
    return !EXCLUDED_SHEETS.some(function(x){ return x.toLowerCase() === name.toLowerCase(); });
  });
}

function parseMachineSheet_(sheet) {
  const machineName = sheet.getName();
  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return {records:[],warning:null};

  const headerRow = findHeaderRow_(values);
  if (headerRow < 0) {
    return {records:[],warning:machineName+': header row not found'};
  }

  const headers = values[headerRow].map(normalize_);
  const index = {};
  headers.forEach(function(h,i){ if (h && index[h] === undefined) index[h]=i; });

  function pick(row,names,fallback) {
    for (var i=0;i<names.length;i++) {
      var pos=index[normalize_(names[i])];
      if (pos !== undefined && row[pos] !== '') return row[pos];
    }
    return fallback === undefined ? '' : fallback;
  }

  const out=[];
  values.slice(headerRow+1).forEach(function(row) {
    if (!row.some(function(v){return String(v).trim()!=='';})) return;
    const partNo=pick(row,['Part No.','Part No','PartNo','Part Number','P/N','PN']);
    const partName=pick(row,['Part Name','PartName','Description','Part Description']);
    const process=pick(row,['Process','Process Name','Operation','Operation Name']);
    const step=pick(row,['Step','Step No.','Step No','Process Step','Operation Step']);
    const ct=toNumber_(pick(row,['CT (sec/pc)','CT (sec)','CT','Cycle Time','Cycle Time (sec)','Time (sec)','Time']));
    const outputCycle=toNumber_(pick(row,['Output/Cycle','Output per Cycle','Output / Cycle','Cavity'],1)) || 1;
    const eff=toNumber_(pick(row,['Efficiency %','Eff %','Efficiency','Eff'],100)) || 100;
    const hours=toNumber_(pick(row,['Working Hours/Shift','Hours/Shift','Hours'],8)) || 8;
    const shifts=toNumber_(pick(row,['Shifts/Day','Shift/Day','Shifts'],2)) || 2;
    const department=pick(row,['Department','Section'],'');
    const status=pick(row,['Status'],'Active');
    const remark=pick(row,['Remark','Remarks','Note','Notes'],'');

    if (!partNo && !process && !step && !ct) return;
    out.push({
      partNo:String(partNo||'').trim(),
      partName:String(partName||'').trim(),
      department:String(department||'').trim(),
      process:String(process||'').trim(),
      step:String(step||'').trim(),
      machine:machineName,
      sourceSheet:machineName,
      ct:ct,
      outputCycle:outputCycle,
      eff:eff,
      hours:hours,
      shifts:shifts,
      status:String(status||'Active').trim(),
      remark:String(remark||'').trim()
    });
  });
  return {records:out,warning:null};
}

function findHeaderRow_(values) {
  const max=Math.min(values.length,10);
  for (var r=0;r<max;r++) {
    const hs=values[r].map(normalize_);
    const hasPart=hs.some(function(h){return ['part no.','part no','partno','part number','p/n','pn'].indexOf(h)>=0;});
    const hasProc=hs.some(function(h){return ['process','process name','operation','operation name'].indexOf(h)>=0;});
    const hasStep=hs.some(function(h){return ['step','step no.','step no','process step','operation step'].indexOf(h)>=0;});
    if (hasPart && (hasProc || hasStep)) return r;
  }
  return -1;
}

function setupMachineTemplate() {
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh=ss.getSheetByName('Machine Template');
  if (!sh) sh=ss.insertSheet('Machine Template');
  sh.clear();
  const headers=[['Part No.','Part Name','Process','Step','CT (sec/pc)','Output/Cycle','Efficiency %','Working Hours/Shift','Shifts/Day','Status','Remark']];
  sh.getRange(1,1,1,headers[0].length).setValues(headers);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers[0].length).setFontWeight('bold').setBackground('#dbeafe');
  sh.autoResizeColumns(1,headers[0].length);
}

function output_(payload,e) {
  const json=JSON.stringify(payload);
  const callback=e && e.parameter ? String(e.parameter.callback||'') : '';
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function normalize_(v) {
  return String(v||'').trim().toLowerCase().replace(/\s+/g,' ');
}
function toNumber_(v) {
  const n=Number(String(v==null?'':v).replace(/,/g,'').replace(/%/g,'').trim());
  return isFinite(n)?n:0;
}
