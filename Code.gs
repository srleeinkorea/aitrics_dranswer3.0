const SHEET_MASTER  = "WBS_MASTER";
const SHEET_STEPS   = "WBS_STEPS";
const SHEET_LOOKUPS = "LOOKUPS";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("AITRICS 2026 Dr.Answer 3.0 | Master Roadmap");
}

/** 날짜를 YYYY-MM-DD로 최대한 정규화 */
function toYMD_(v) {
  if (v === null || v === undefined || v === "") return "";

  // Date object
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  // Numeric serial date (in case a date cell is stored as number)
  if (typeof v === "number" && isFinite(v)) {
    // Google Sheets/Excel serial day -> JS Date (1899-12-30 base)
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
  }

  const s = String(v).trim();
  if (!s) return "";

  // Already yyyy-mm-dd
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // yyyy.mm.dd or yyyy/mm/dd
  const m1 = s.match(/^(\d{4})[./](\d{1,2})[./](\d{1,2})$/);
  if (m1) {
    const yyyy = m1[1];
    const mm = String(m1[2]).padStart(2, "0");
    const dd = String(m1[3]).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  // Fallback: return as-is (UI에서 그대로 표시)
  return s;
}

/** 시트(헤더 1행) → 객체 배열 */
function readTable_(sheetName, keyHeader) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => String(h || "").trim());
  const keyIdx = headers.indexOf(keyHeader);

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    // 완전 공백행 skip
    const any = row.some(v => v !== "" && v !== null && v !== undefined);
    if (!any) continue;

    // key가 있으면 key 비는 행 skip
    if (keyIdx >= 0 && (row[keyIdx] === "" || row[keyIdx] === null || row[keyIdx] === undefined)) continue;

    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c];
      if (!h) continue;

      if (h === "start_date" || h === "due_date") obj[h] = toYMD_(row[c]);
      else obj[h] = (row[c] === null || row[c] === undefined) ? "" : row[c];
    }
    out.push(obj);
  }
  return out;
}

/** LOOKUPS 시트(열별 리스트) → {workstream_4:[], dept:[], ...} */
function readLookups_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_LOOKUPS);
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return {};

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x || "").trim());
  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const out = {};
  for (let c = 0; c < lastCol; c++) {
    const key = header[c];
    if (!key) continue;

    const arr = [];
    for (let r = 0; r < body.length; r++) {
      const v = body[r][c];
      if (v === "" || v === null || v === undefined) continue;
      arr.push(String(v).trim());
    }
    out[key] = [...new Set(arr)];
  }
  return out;
}

/** WebApp → loadAll() */
function loadAll() {
  const master = readTable_(SHEET_MASTER, "task_id");
  const steps  = readTable_(SHEET_STEPS, "task_id");
  const lookups = readLookups_();

  return {
    ok: true,
    master,
    steps,
    lookups,
    meta: {
      loadedAt: new Date(),
      tz: Session.getScriptTimeZone()
    }
  };
}

/* =========================
   드롭다운(LOOKUPS) 자동 적용
   - 1회 실행 후, WBS_MASTER/WBS_STEPS의 주요 컬럼에 Data validation 부여
========================= */

function setupValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shL = ss.getSheetByName(SHEET_LOOKUPS);
  if (!shL) throw new Error("LOOKUPS 시트 없음");

  const lookups = readLookups_();

  const shM = ss.getSheetByName(SHEET_MASTER);
  const shS = ss.getSheetByName(SHEET_STEPS);
  if (!shM || !shS) throw new Error("WBS_MASTER 또는 WBS_STEPS 시트 없음");

  // Header -> col index
  const mHeaders = shM.getRange(1, 1, 1, shM.getLastColumn()).getValues()[0].map(x => String(x || "").trim());
  const sHeaders = shS.getRange(1, 1, 1, shS.getLastColumn()).getValues()[0].map(x => String(x || "").trim());

  const colOf = (headers, name) => headers.indexOf(name) + 1; // 1-based, 0이면 없음

  // LOOKUPS column range builder
  const lookupRangeOf_ = (key) => {
    if (!lookups[key] || lookups[key].length === 0) return null;

    const lHeaders = shL.getRange(1, 1, 1, shL.getLastColumn()).getValues()[0].map(x => String(x || "").trim());
    const c = lHeaders.indexOf(key);
    if (c < 0) return null;

    const lastRow = shL.getLastRow();
    const height = Math.max(lastRow - 1, 1); // from row2
    return shL.getRange(2, c + 1, height, 1);
  };

  // Apply validation for WBS_MASTER
  const masterMap = [
    { header: "workstream_4", lookup: "workstream_4" },
    { header: "dept",        lookup: "dept" },
    { header: "owner",       lookup: "owner" },
    { header: "status",      lookup: "status" },
    { header: "deliverable", lookup: "deliverable" },
    { header: "doc_type",    lookup: "doc_type" },
    { header: "priority",    lookup: "priority" },
  ];

  masterMap.forEach(m => {
    const col = colOf(mHeaders, m.header);
    const lr = lookupRangeOf_(m.lookup);
    if (col <= 0 || !lr) return;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(lr, true)
      .setAllowInvalid(false)
      .build();

    const target = shM.getRange(2, col, shM.getMaxRows() - 1, 1);
    target.setDataValidation(rule);
  });

  // Apply validation for WBS_STEPS
  const stepMap = [
    { header: "step_week", lookup: "step_week" },
  ];

  stepMap.forEach(m => {
    const col = colOf(sHeaders, m.header);
    const lr = lookupRangeOf_(m.lookup);
    if (col <= 0 || !lr) return;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(lr, true)
      .setAllowInvalid(false)
      .build();

    const target = shS.getRange(2, col, shS.getMaxRows() - 1, 1);
    target.setDataValidation(rule);
  });

  // UX: Freeze header rows
  shM.setFrozenRows(1);
  shS.setFrozenRows(1);
  shL.setFrozenRows(1);
}
