const SPREADSHEET_ID = '1eXby1xmCjhp_C8H_r7OC8JmnLu00WRYq';
const DATA_SHEET = 'Capacity_Data';
const DEPARTMENT_SHEET = 'Department_Master';

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'capacity').toLowerCase();
    let payload;
    if (action === 'capacity') payload = getCapacityData_();
    else payload = { ok: false, error: 'Unknown action: ' + action };
    return output_(payload, e);
  } catch (err) {
    return output_({ ok: false, error: String(err && err.message ? err.message : err) }, e);
  }
}

function getCapacityData_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) throw new Error('ไม่พบชีต ' + DATA_SHEET + ' — ให้รัน setupCapacitySheets() 1 ครั้ง');

  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return { ok: true, records: [], departments: getDepartments_(ss), generatedAt: new Date().toISOString() };

  const headers = values[0].map(normalize_);
  const index = {};
  headers.forEach((h, i) => index[h] = i);
  const pick = (row, names, fallback='') => {
    for (const n of names) {
      const i = index[normalize_(n)];
      if (i !== undefined && row[i] !== '') return row[i];
    }
    return fallback;
  };

  const records = values.slice(1).filter(row => row.some(v => String(v).trim() !== '')).map(row => ({
    partNo: pick(row, ['Part No.','Part No','PartNo']),
    partName: pick(row, ['Part Name','PartName']),
    department: pick(row, ['Department','Section']),
    process: pick(row, ['Process']),
    step: pick(row, ['Step']),
    machine: pick(row, ['Machine','Machine No.','Machine No']),
    ct: toNumber_(pick(row, ['CT (sec/pc)','CT (sec)','CT','Cycle Time'])),
    outputCycle: toNumber_(pick(row, ['Output/Cycle','Output per Cycle'], 1)) || 1,
    eff: toNumber_(pick(row, ['Efficiency %','Eff %','Efficiency'], 100)) || 100,
    hours: toNumber_(pick(row, ['Working Hours/Shift','Hours/Shift','Hours'], 8)) || 8,
    shifts: toNumber_(pick(row, ['Shifts/Day','Shift/Day','Shifts'], 2)) || 2,
    status: pick(row, ['Status'], 'Active'),
    remark: pick(row, ['Remark','Remarks'])
  })).filter(r => r.partNo || r.process || r.machine);

  return {
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    records,
    departments: getDepartments_(ss),
    generatedAt: new Date().toISOString()
  };
}

function getDepartments_(ss) {
  const sh = ss.getSheetByName(DEPARTMENT_SHEET);
  if (!sh || sh.getLastRow() < 2) {
    return ['Stamping','Welding','CNC','Tapping','Bending','Engineering Support','Machine Maintenance','Tooling Maintenance','Sorting 1','Sorting 2'];
  }
  return sh.getRange(2, 1, sh.getLastRow()-1, 1).getDisplayValues().flat().map(String).map(s=>s.trim()).filter(Boolean);
}

function output_(payload, e) {
  const json = JSON.stringify(payload);
  const callback = e && e.parameter ? String(e.parameter.callback || '') : '';
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function normalize_(v) {
  return String(v || '').trim().toLowerCase().replace(/\s+/g, ' ');
}
function toNumber_(v) {
  const n = Number(String(v == null ? '' : v).replace(/,/g, '').replace(/%/g, '').trim());
  return isFinite(n) ? n : 0;
}

// รันฟังก์ชันนี้เพียงครั้งแรก เพื่อสร้างหัวตารางที่เว็บต้องใช้
function setupCapacitySheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let data = ss.getSheetByName(DATA_SHEET);
  if (!data) data = ss.insertSheet(DATA_SHEET);
  if (data.getLastRow() === 0) {
    const headers = [['Part No.','Part Name','Department','Process','Step','Machine','CT (sec/pc)','Output/Cycle','Efficiency %','Working Hours/Shift','Shifts/Day','Status','Remark']];
    data.getRange(1,1,1,headers[0].length).setValues(headers);
    data.setFrozenRows(1);
    data.getRange(1,1,1,headers[0].length).setFontWeight('bold').setBackground('#dbeafe');
    data.autoResizeColumns(1, headers[0].length);
  }

  let dept = ss.getSheetByName(DEPARTMENT_SHEET);
  if (!dept) dept = ss.insertSheet(DEPARTMENT_SHEET);
  if (dept.getLastRow() === 0) {
    const names = ['Department','Stamping','Welding','CNC','Tapping','Bending','Engineering Support','Machine Maintenance','Tooling Maintenance','Sorting 1','Sorting 2'].map(x=>[x]);
    dept.getRange(1,1,names.length,1).setValues(names);
    dept.setFrozenRows(1);
    dept.getRange(1,1).setFontWeight('bold').setBackground('#dbeafe');
    dept.autoResizeColumn(1);
  }
}
