/**
 * ============================================================================
 *  Happinometer — Google Apps Script Backend (Google Sheets + Web App)
 * ============================================================================
 *  รองรับการใช้งานทั้ง 2 แบบ:
 *   1) เปิดเว็บแอพผ่านลิงก์ Apps Script โดยตรง (ไม่ต้องตั้งค่า URL เพิ่ม)
 *   2) โฮสต์ index.html แยก (เช่น GitHub Pages / Netlify / เปิดในเครื่อง)
 *      แล้วนำ Web App URL ไปวางในหน้า "ตั้งค่า" ของแอพ
 *
 *  วิธีติดตั้ง (สรุปย่อ — ดูรายละเอียดเพิ่มเติมใน README.md):
 *   1. สร้าง Google Sheet ใหม่ -> เมนู ส่วนขยาย (Extensions) -> Apps Script
 *   2. วางโค้ดนี้ลงในไฟล์ Code.gs
 *   3. สร้างไฟล์ HTML ชื่อ "index" แล้ววางโค้ดจากไฟล์ index.html
 *   4. กด ทำให้ใช้งานได้ (Deploy) -> การทำให้ใช้งานได้รายการใหม่ -> เว็บแอป
 *      - ดำเนินการในฐานะ: ฉัน (Me)
 *      - ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน (Anyone)
 *   5. คัดลอก Web App URL ไปเปิดบนมือถือพยาบาล หรือวางในหน้าตั้งค่าของแอพ
 * ============================================================================
 */

const CONFIG = {
  // ถ้าสร้าง Apps Script จากใน Google Sheet โดยตรง ให้เว้นว่าง '' ไว้ได้เลย
  // ถ้าสร้างเป็น Standalone Script ให้ใส่ Spreadsheet ID ที่นี่
  SPREADSHEET_ID: '',
  SHEET_RESPONSES: 'Responses',
  SHEET_SUMMARY: 'Summary',
  // แยกชีตตามหน่วยอัตโนมัติ (ตั้งเป็น false หากต้องการเก็บเฉพาะชีต Responses + Summary)
  SPLIT_BY_UNIT: true,
  TIMEZONE: 'Asia/Bangkok',
};

const HEADERS = [
  'ID',
  'Timestamp',
  'วันที่',
  'เวลา',
  'หน่วย',
  'ห้องตรวจ',
  'จุดบริการ',
  'เพศ',
  'ช่วงอายุ',
  'สิทธิ์การรักษา',
  'คะแนน',
  'ระดับความพึงพอใจ',
  'อุปกรณ์',
];

const SCORE_LABELS = {
  1: 'พอใจน้อยที่สุด',
  2: 'พอใจน้อย',
  3: 'ปานกลาง',
  4: 'พอใจมาก',
  5: 'พอใจมากที่สุด',
};

/* ============================================================================
 *  HTTP Entrypoints (Web App & JSON API)
 * ============================================================================ */

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = params.action || '';

  if (action === 'ping') {
    return jsonResponse(apiInfo());
  }
  if (action === 'list') {
    return jsonResponse(apiList(params.from || '', params.to || ''));
  }

  // Serve Happinometer Web App HTML directly
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Happinometer · ประเมินความพึงพอใจผู้ป่วย')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || '{}';
    const body = JSON.parse(raw);

    if (body.action === 'submit') {
      return jsonResponse(apiSubmit(body.records || []));
    }
    if (body.action === 'ping') {
      return jsonResponse(apiInfo());
    }
    return jsonResponse({ ok: false, error: 'Unknown action: ' + body.action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================================
 *  Public API (เรียกใช้ได้ทั้งจาก google.script.run และ doGet / doPost)
 * ============================================================================ */

function apiInfo() {
  try {
    const ss = getSpreadsheet();
    ensureSetup(ss);
    return {
      ok: true,
      app: 'Happinometer',
      version: '1.0',
      sheetName: ss.getName(),
      sheetUrl: ss.getUrl(),
      serverTime: new Date().toISOString(),
    };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function apiList(fromDate, toDate) {
  try {
    const ss = getSpreadsheet();
    const sheet = ensureResponsesSheet(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { ok: true, records: [], serverTime: new Date().toISOString() };
    }

    const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    const records = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const id = String(row[0] || '').trim();
      if (!id) continue;

      const dateStr = normalizeDateCell(row[2]);
      if (fromDate && dateStr && dateStr < fromDate) continue;
      if (toDate && dateStr && dateStr > toDate) continue;

      const score = Number(row[10]);
      if (!(score >= 1 && score <= 5)) continue;

      records.push({
        id: id,
        ts: normalizeTimestampCell(row[1]),
        date: dateStr,
        time: normalizeTimeCell(row[3]),
        unit: String(row[4] || ''),
        room: String(row[5] || ''),
        point: String(row[6] || ''),
        gender: String(row[7] || ''),
        age: String(row[8] || ''),
        right: String(row[9] || ''),
        score: score,
        device: String(row[12] || ''),
      });
    }

    return {
      ok: true,
      records: records,
      serverTime: new Date().toISOString(),
    };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err), records: [] };
  }
}

function apiSubmit(records) {
  if (!Array.isArray(records) || records.length === 0) {
    return { ok: true, saved: [] };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const ss = getSpreadsheet();
    const sheet = ensureResponsesSheet(ss);

    // ตรวจสอบ ID ซ้ำเพื่อรองรับ Offline Queue Retry
    const existingIds = new Set();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const idVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < idVals.length; i++) {
        const v = String(idVals[i][0] || '').trim();
        if (v) existingIds.add(v);
      }
    }

    const saved = [];
    const rowsToAppend = [];
    const byUnit = {};

    for (let i = 0; i < records.length; i++) {
      const r = records[i] || {};
      const id = String(r.id || '').trim();
      const score = Number(r.score);
      if (!id || !(score >= 1 && score <= 5)) continue;

      // ถ้ามี ID นี้อยู่แล้ว ถือว่าบันทึกสำเร็จแล้ว (กันข้อมูลซ้ำ)
      if (existingIds.has(id)) {
        saved.push(id);
        continue;
      }

      const now = new Date();
      const ts = r.ts ? String(r.ts) : now.toISOString();
      const dateStr = r.date ? String(r.date) : Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      const timeStr = r.time ? String(r.time) : Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm');
      const unit = String(r.unit || 'อื่นๆ').trim();
      const room = String(r.room || 'ไม่ระบุ').trim();

      const row = [
        id,
        ts,
        "'" + dateStr, // เก็บเป็น Text yyyy-MM-dd เพื่อไม่ให้ Sheet แปลงฟอร์แมตผิด
        "'" + timeStr,
        unit,
        room,
        String(r.point || ''),
        String(r.gender || 'ไม่ระบุ'),
        String(r.age || ''),
        String(r.right || ''),
        score,
        SCORE_LABELS[score] || '',
        String(r.device || ''),
      ];

      rowsToAppend.push(row);
      existingIds.add(id);
      saved.push(id);

      if (CONFIG.SPLIT_BY_UNIT) {
        if (!byUnit[unit]) byUnit[unit] = [];
        byUnit[unit].push(row);
      }
    }

    if (rowsToAppend.length > 0) {
      sheet
        .getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, HEADERS.length)
        .setValues(rowsToAppend);

      if (CONFIG.SPLIT_BY_UNIT) {
        const unitNames = Object.keys(byUnit);
        for (let u = 0; u < unitNames.length; u++) {
          const uName = unitNames[u];
          const uSheet = ensureUnitSheet(ss, uName);
          const uRows = byUnit[uName];
          uSheet
            .getRange(uSheet.getLastRow() + 1, 1, uRows.length, HEADERS.length)
            .setValues(uRows);
        }
      }

      updateSummarySheet(ss, sheet);
    }

    return { ok: true, saved: saved, added: rowsToAppend.length };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/* ============================================================================
 *  Sheet Setup & Summary Helpers
 * ============================================================================ */

function getSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('ไม่พบ Google Sheet — กรุณาผูก Script กับ Google Sheet หรือระบุ SPREADSHEET_ID ใน Code.gs');
}

function ensureSetup(ss) {
  const resp = ensureResponsesSheet(ss);
  ensureSummarySheet(ss, resp);
}

function ensureResponsesSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_RESPONSES);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_RESPONSES);
    formatDataSheet(sheet, '#7c6cff');
  } else if (sheet.getLastRow() === 0) {
    formatDataSheet(sheet, '#7c6cff');
  }
  return sheet;
}

function ensureUnitSheet(ss, unitName) {
  const safeName = sanitizeSheetName(unitName);
  let sheet = ss.getSheetByName(safeName);
  if (!sheet) {
    sheet = ss.insertSheet(safeName);
    formatDataSheet(sheet, '#ff6fa5');
  } else if (sheet.getLastRow() === 0) {
    formatDataSheet(sheet, '#ff6fa5');
  }
  return sheet;
}

function formatDataSheet(sheet, headerColor) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange
    .setBackground(headerColor)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);
  sheet.setFrozenRows(1);

  const widths = [130, 175, 100, 75, 130, 150, 175, 80, 105, 125, 70, 135, 110];
  for (let i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}

function ensureSummarySheet(ss, respSheet) {
  let sum = ss.getSheetByName(CONFIG.SHEET_SUMMARY);
  if (!sum) {
    sum = ss.insertSheet(CONFIG.SHEET_SUMMARY, 0);
    updateSummarySheet(ss, respSheet);
  }
  return sum;
}

function updateSummarySheet(ss, respSheet) {
  let sum = ss.getSheetByName(CONFIG.SHEET_SUMMARY);
  if (!sum) sum = ss.insertSheet(CONFIG.SHEET_SUMMARY, 0);

  const lastRow = respSheet.getLastRow();
  const values = lastRow >= 2 ? respSheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues() : [];

  const groups = {};
  let totalN = 0;
  let totalSum = 0;
  const totalDist = [0, 0, 0, 0, 0, 0];

  for (let i = 0; i < values.length; i++) {
    const unit = String(values[i][4] || 'อื่นๆ');
    const room = String(values[i][5] || 'ไม่ระบุ');
    const score = Number(values[i][10]);
    if (!(score >= 1 && score <= 5)) continue;

    totalN++;
    totalSum += score;
    totalDist[score]++;

    const key = unit + '||' + room;
    if (!groups[key]) {
      groups[key] = { unit: unit, room: room, n: 0, sum: 0, dist: [0, 0, 0, 0, 0, 0] };
    }
    groups[key].n++;
    groups[key].sum += score;
    groups[key].dist[score]++;
  }

  const rows = Object.keys(groups)
    .map(function (k) { return groups[k]; })
    .sort(function (a, b) { return b.n - a.n || (b.sum / b.n) - (a.sum / a.n); });

  sum.clearContents();
  sum.clearFormats();

  const updatedAt = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  sum.getRange('A1:K1').merge()
    .setValue('Happinometer — สรุปความพึงพอใจผู้ป่วยแยกตามหน่วยและห้องตรวจ (อัปเดต: ' + updatedAt + ')')
    .setBackground('#2b2350')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setVerticalAlignment('middle');
  sum.setRowHeight(1, 36);

  const sumHeaders = [
    'หน่วย', 'ห้องตรวจ', 'จำนวน (N)', 'คะแนนเฉลี่ย', 'ร้อยละ (%)',
    'พอใจมาก-มากที่สุด (%)', '5 คะแนน', '4 คะแนน', '3 คะแนน', '2 คะแนน', '1 คะแนน',
  ];
  sum.getRange(2, 1, 1, sumHeaders.length)
    .setValues([sumHeaders])
    .setBackground('#7c6cff')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sum.setRowHeight(2, 28);
  sum.setFrozenRows(2);

  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const g = rows[i];
    const avg = g.sum / g.n;
    const pct = (avg / 5) * 100;
    const top2 = ((g.dist[4] + g.dist[5]) / g.n) * 100;
    out.push([
      g.unit,
      g.room,
      g.n,
      Number(avg.toFixed(2)),
      Number(pct.toFixed(1)),
      Number(top2.toFixed(1)),
      g.dist[5],
      g.dist[4],
      g.dist[3],
      g.dist[2],
      g.dist[1],
    ]);
  }

  if (totalN > 0) {
    const tAvg = totalSum / totalN;
    const tPct = (tAvg / 5) * 100;
    const tTop2 = ((totalDist[4] + totalDist[5]) / totalN) * 100;
    out.push([
      'รวมทั้งหมด',
      rows.length + ' ห้องตรวจ',
      totalN,
      Number(tAvg.toFixed(2)),
      Number(tPct.toFixed(1)),
      Number(tTop2.toFixed(1)),
      totalDist[5],
      totalDist[4],
      totalDist[3],
      totalDist[2],
      totalDist[1],
    ]);
  }

  if (out.length > 0) {
    sum.getRange(3, 1, out.length, sumHeaders.length).setValues(out);
    sum.getRange(3, 3, out.length, 9).setHorizontalAlignment('center');
    sum.getRange(3, 4, out.length, 1).setNumberFormat('0.00');
    sum.getRange(3, 5, out.length, 2).setNumberFormat('0.0"%"');

    // ไฮไลต์แถวรวมทั้งหมด
    const totalRowIdx = 3 + out.length - 1;
    sum.getRange(totalRowIdx, 1, 1, sumHeaders.length)
      .setBackground('#f2edff')
      .setFontWeight('bold');
  }

  const widths = [140, 170, 95, 105, 100, 150, 80, 80, 80, 80, 80];
  for (let i = 0; i < widths.length; i++) {
    sum.setColumnWidth(i + 1, widths[i]);
  }
}

/* ============================================================================
 *  Cell Normalization Helpers
 * ============================================================================ */

function normalizeDateCell(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }
  return String(v).replace(/^'/, '').trim().slice(0, 10);
}

function normalizeTimeCell(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, CONFIG.TIMEZONE, 'HH:mm');
  }
  return String(v).replace(/^'/, '').trim().slice(0, 5);
}

function normalizeTimestampCell(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toISOString();
  return String(v).trim();
}

function sanitizeSheetName(name) {
  const cleaned = String(name || 'อื่นๆ')
    .replace(/[\\/?*[\]:]/g, '_')
    .trim()
    .slice(0, 31);
  return cleaned || 'อื่นๆ';
}

/**
 * ฟังก์ชันสำหรับทดสอบใน Apps Script Editor (กด Run เพื่อสร้างชีตเริ่มต้น)
 */
function setupSheets() {
  const ss = getSpreadsheet();
  ensureSetup(ss);
  Logger.log('Happinometer setup complete: ' + ss.getUrl());
}
