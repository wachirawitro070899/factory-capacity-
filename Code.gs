const SPREADSHEET_ID = '1PvNRR9uekifW7O-cmxh78bOWpYMuYhk3VtK76FGy6BE';

// Factory Capacity backend
// โครงสร้างหลัก: 1 Google Sheet tab = 1 Machine
// ชื่อแท็บ = ชื่อเครื่องจักร
// ข้อมูลแต่ละแถว = Part No. + Process + Step + CT
function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'capacity').toLowerCase();

    if (action === 'capacity') {
      return output_(getCapacityData_(), e);
    }

    if (action === 'ping') {
      return output_({
        ok: true,
        spreadsheetId: SPREADSHEET_ID,
        structure: '1 sheet = 1 machine',
        generatedAt: new Date().toISOString()
      }, e);
    }

    return output_({ ok: false, error: 'Unknown action: ' + action }, e);
  } catch (err) {
    return output_({
      ok: false,
      error: String(err && err.message ? err.message : err)
    }, e);
  }
}

function getCapacityData_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const records = [];
  const machineSheets = [];
  const skippedSheets = [];

  ss.getSheets().forEach(sheet => {
    const machine = String(sheet.getName() || '').trim();
    if (!machine) return;

    const values = sheet.getDataRange().getDisplayValues();
    if (!values || values.length < 1 || values[0].length < 1) {
      skippedSheets.push({ sheet: machine, reason: 'empty' });
      return;
    }

    const headers = values[0].map(normalize_);
    const index = {};
    headers.forEach((header, i) => {
      if (header) index[header] = i;
    });

    const hasPart = hasAnyHeader_(index, ['Part No.', 'Part No', 'PartNo', 'Part Number']);
    const hasProcess = hasAnyHeader_(index, ['Process', 'Operation']);
    const hasStep = hasAnyHeader_(index, ['Step', 'Process Step', 'Operation Step']);

    // ไม่ใช่แท็บเครื่องจักรถ้าไม่มีหัวข้อมูลหลักเลย
    if (!hasPart && !hasProcess && !hasStep) {
      skippedSheets.push({ sheet: machine, reason: 'machine headers not recognized' });
      return;
    }

    const pick = (row, names, fallback) => {
      for (let i = 0; i < names.length; i++) {
        const col = index[normalize_(names[i])];
        if (col !== undefined && col < row.length && String(row[col]).trim() !== '') {
          return row[col];
        }
      }
      return fallback === undefined ? '' : fallback;
    };

    let rowCount = 0;

    values.slice(1).forEach(row => {
      if (!row.some(v => String(v).trim() !== '')) return;

      const partNo = String(pick(row, ['Part No.', 'Part No', 'PartNo', 'Part Number']) || '').trim();
      const partName = String(pick(row, ['Part Name', 'PartName', 'Description', 'Part Description']) || '').trim();
      const process = String(pick(row, ['Process', 'Operation']) || '').trim();
      const step = String(pick(row, ['Step', 'Process Step', 'Operation Step']) || '').trim();

      // ต้องมีอย่างน้อย Part / Process / Step อย่างใดอย่างหนึ่ง
      if (!partNo && !process && !step) return;

      const ct = toNumber_(pick(row, ['CT (sec/pc)', 'CT (sec)', 'CT', 'Cycle Time', 'Cycle Time (sec)', 'Time (sec)']));
      const outputCycle = toNumber_(pick(row, ['Output/Cycle', 'Output per Cycle', 'Output / Cycle', 'Qty/Cycle'], 1)) || 1;
      const eff = toNumber_(pick(row, ['Efficiency %', 'Eff %', 'Efficiency', 'Eff'], 100)) || 100;
      const hours = toNumber_(pick(row, ['Working Hours/Shift', 'Hours/Shift', 'Hours'], 8)) || 8;
      const shifts = toNumber_(pick(row, ['Shifts/Day', 'Shift/Day', 'Shifts'], 2)) || 2;
      const status = String(pick(row, ['Status'], 'Active') || 'Active').trim();
      const remark = String(pick(row, ['Remark', 'Remarks', 'Note', 'Notes']) || '').trim();

      records.push({
        machine: machine,
        sourceSheet: machine,
        partNo: partNo,
        partName: partName,
        process: process,
        step: step,
        ct: ct,
        outputCycle: outputCycle,
        eff: eff,
        hours: hours,
        shifts: shifts,
        status: status,
        remark: remark
      });

      rowCount++;
    });

    if (rowCount > 0) {
      machineSheets.push(machine);
    } else {
      skippedSheets.push({ sheet: machine, reason: 'no production rows' });
    }
  });

  return {
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    spreadsheetName: ss.getName(),
    structure: 'Machine -> Part No. -> Process -> Step -> CT',
    machineSheets: machineSheets,
    machineCount: machineSheets.length,
    recordCount: records.length,
    records: records,
    skippedSheets: skippedSheets,
    generatedAt: new Date().toISOString()
  };
}

function hasAnyHeader_(index, names) {
  return names.some(name => index[normalize_(name)] !== undefined);
}

function output_(payload, e) {
  const json = JSON.stringify(payload);
  const callback = e && e.parameter ? String(e.parameter.callback || '') : '';

  // JSONP สำหรับ GitHub Pages เพื่อเลี่ยงปัญหา CORS
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function normalize_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');
}

function toNumber_(value) {
  const raw = String(value == null ? '' : value)
    .replace(/,/g, '')
    .replace(/%/g, '')
    .trim();
  const number = Number(raw);
  return isFinite(number) ? number : 0;
}

// ใช้ทดสอบใน Apps Script Editor
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
