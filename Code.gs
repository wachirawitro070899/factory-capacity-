const SPREADSHEET_ID = '1PvNRR9uekifW7O-cmxh78bOWpYMuYhk3VtK76FGy6BE';

// 1 Google Sheet tab = 1 Machine
// ชื่อแท็บจะถูกใช้เป็นชื่อเครื่องจักรโดยอัตโนมัติ
function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'capacity').toLowerCase();
    let payload;

    if (action === 'capacity') {
      payload = getCapacityData_();
    } else if (action === 'ping') {
      payload = {
        ok: true,
        spreadsheetId: SPREADSHEET_ID,
        generatedAt: new Date().toISOString()
      };
    } else {
      payload = { ok: false, error: 'Unknown action: ' + action };
    }

    return output_(payload, e);
  } catch (err) {
    return output_({
      ok: false,
      error: String(err && err.message ? err.message : err)
    }, e);
  }
}

function getCapacityData_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const records = [];
  const machineSheets = [];
  const skippedSheets = [];

  sheets.forEach(sheet => {
    const sheetName = String(sheet.getName() || '').trim();
    if (!sheetName) return;

    const values = sheet.getDataRange().getDisplayValues();
    if (!values || values.length === 0 || values[0].length === 0) {
      skippedSheets.push({ sheet: sheetName, reason: 'empty' });
      return;
    }

    const headers = values[0].map(normalize_);
    const index = {};
    headers.forEach((h, i) => {
      if (h) index[h] = i;
    });

    // อ่านเฉพาะแท็บที่มีโครงสร้างข้อมูล Part/Process/Step
    const hasPart = hasAnyHeader_(index, ['Part No.', 'Part No', 'PartNo', 'Part Number']);
    const hasProcess = hasAnyHeader_(index, ['Process', 'Operation']);
    const hasStep = hasAnyHeader_(index, ['Step', 'Process Step', 'Operation Step']);

    if (!hasPart && !hasProcess && !hasStep) {
      skippedSheets.push({ sheet: sheetName, reason: 'header not recognized' });
      return;
    }

    const pick = (row, names, fallback = '') => {
      for (const n of names) {
        const i = index[normalize_(n)];
        if (i !== undefined && i < row.length && String(row[i]).trim() !== '') {
          return row[i];
        }
      }
      return fallback;
    };

    let sheetHasRecords = false;

    values.slice(1).forEach(row => {
      if (!row.some(v => String(v).trim() !== '')) return;

      const partNo = pick(row, ['Part No.', 'Part No', 'PartNo', 'Part Number']);
      const partName = pick(row, ['Part Name', 'PartName', 'Description', 'Part Description']);
      const process = pick(row, ['Process', 'Operation']);
      const step = pick(row, ['Step', 'Process Step', 'Operation Step']);
      const ct = toNumber_(pick(row, ['CT (sec/pc)', 'CT (sec)', 'CT', 'Cycle Time', 'Cycle Time (sec)', 'Time (sec)']));
      const outputCycle = toNumber_(pick(row, ['Output/Cycle', 'Output per Cycle', 'Output / Cycle', 'Qty/Cycle'], 1)) || 1;
      const eff = toNumber_(pick(row, ['Efficiency %', 'Eff %', 'Efficiency', 'Eff'], 100)) || 100;
      const department = pick(row, ['Department', 'Section', 'Dept']);
      const hours = toNumber_(pick(row, ['Working Hours/Shift', 'Hours/Shift', 'Hours'], 8)) || 8;
      const shifts = toNumber_(pick(row, ['Shifts/Day', 'Shift/Day', 'Shifts'], 2)) || 2;
      const status = pick(row, ['Status'], 'Active');
      const remark = pick(row, ['Remark', 'Remarks', 'Note', 'Notes']);

      // ข้ามแถวที่ไม่มีข้อมูลหลักเลย
      if (!partNo && !process && !step) return;

      records.push({
        machine: sheetName,
        sourceSheet: sheetName,
        partNo: String(partNo || '').trim(),
        partName: String(partName || '').trim(),
        department: String(department || '').trim(),
        process: String(process || '').trim(),
        step: String(step || '').trim(),
        ct: ct,
        outputCycle: outputCycle,
        eff: eff,
        hours: hours,
        shifts: shifts,
        status: String(status || 'Active').trim(),
        remark: String(remark || '').trim()
      });
      sheetHasRecords = true;
    });

    if (sheetHasRecords) machineSheets.push(sheetName);
  });

  return {
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    spreadsheetName: ss.getName(),
    machineSheets: machineSheets,
    machineCount: machineSheets.length,
    recordCount: records.length,
    records: records,
    skippedSheets: skippedSheets,
    generatedAt: new Date().toISOString()
  };
}

function hasAnyHeader_(index, names) {
  return names.some(n => index[normalize_(n)] !== undefined);
}

function output_(payload, e) {
  const json = JSON.stringify(payload);
  const callback = e && e.parameter ? String(e.parameter.callback || '') : '';

  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function normalize_(v) {
  return String(v || '')
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');
}

function toNumber_(v) {
  const raw = String(v == null ? '' : v)
    .replace(/,/g, '')
    .replace(/%/g, '')
    .trim();
  const n = Number(raw);
  return isFinite(n) ? n : 0;
}

// รันฟังก์ชันนี้เพื่อทดสอบจาก Apps Script Editor
function testCapacityData() {
  const data = getCapacityData_();
  Logger.log(JSON.stringify({
    ok: data.ok,
    spreadsheetName: data.spreadsheetName,
    machineCount: data.machineCount,
    recordCount: data.recordCount,
    machineSheets: data.machineSheets
  }, null, 2));
}
