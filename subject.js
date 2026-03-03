// ============================================================
//  LICH GIANG DAY - Google Apps Script (rewrite theo yeu cau)
//
//  Yeu cau chinh:
//  - Doc du lieu tu Sheet1 va Sheet2, KHONG sua du lieu goc
//  - Tao sheet moi "Lich giang day" bang cach copy nguyen Sheet2
//  - Hien thi dialog nhap: ngay bat dau + (ngay ket thuc hoac so ngay cong them)
//  - Chen cac cot ngay ben phai sheet moi
//  - Ngay bat dau (khai giang): de trong, khong tinh tiet
//  - Ngay ket thuc (be giang): de trong, khong tinh tiet
//  - Cac ngay o giua: phan bo tiet, tong moi ngay huong toi 8 tiet
//  - Cac dong cung ma MH/MD nhan cung mot gia tri theo ngay
// ============================================================

const SHEET_TEACHERS = 'Sheet1';
const SHEET_SOURCE = 'Sheet2';
const SHEET_SCHEDULE = 'Lịch giảng dạy';

const TIETS_PER_DAY = 8;
const MIN_CELL_TIETS = 4;
const DEFAULT_PLUS_DAYS = 45;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Lich giang day')
    .addItem('Tao lich...', 'showScheduleDialog')
    .addItem('Xoa sheet lich', 'clearScheduleSheet')
    .addToUi();
}

function showScheduleDialog() {
  const html = HtmlService.createHtmlOutput(buildDialogHTML())
    .setWidth(520)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Thiet lap Lich giang day');
}

function buildDialogHTML() {
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <style>
    *{box-sizing:border-box}
    body{font-family:Arial,sans-serif;background:#f5f7fb;color:#1f2937;padding:18px}
    .card{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:16px}
    h2{margin:0 0 8px 0;font-size:18px}
    p{margin:0 0 14px 0;color:#4b5563;font-size:12px}
    label{display:block;font-weight:700;font-size:12px;margin:10px 0 6px 0}
    input{width:100%;padding:10px;border:1px solid #d1d5db;border-radius:8px;font-size:13px}
    .row{display:grid;grid-template-columns:1fr 1fr;gap:10px}
    .muted{font-size:11px;color:#6b7280;margin-top:5px}
    .btns{display:flex;gap:10px;margin-top:14px}
    button{flex:1;border:0;border-radius:8px;padding:10px;font-weight:700;cursor:pointer}
    .pri{background:#2563eb;color:#fff}
    .sec{background:#e5e7eb;color:#111827}
    #st{margin-top:12px;border-radius:8px;padding:10px;font-size:12px;display:none;white-space:pre-wrap}
    .ok{display:block;background:#ecfdf3;color:#065f46}
    .err{display:block;background:#fef2f2;color:#991b1b}
    .run{display:block;background:#eff6ff;color:#1d4ed8}
  </style>
</head>
<body>
  <div class="card">
    <h2>Tao Lich giang day</h2>
    <p>Nhap ngay bat dau, sau do chon 1 trong 2 cach: nhap ngay ket thuc hoac so ngay cong them.</p>

    <label>Ngay bat dau (khai giang)</label>
    <input type="date" id="start" value="2026-03-02" />

    <div class="row">
      <div>
        <label>Ngay ket thuc (be giang)</label>
        <input type="date" id="end" />
      </div>
      <div>
        <label>So ngay cong them</label>
        <input type="number" id="plus" min="1" value="${DEFAULT_PLUS_DAYS}" />
      </div>
    </div>

    <div class="muted">Neu co ca 2 gia tri, he thong uu tien Ngay ket thuc.</div>

    <div class="btns">
      <button class="pri" onclick="runCreate()">Tao lich</button>
      <button class="sec" onclick="runClear()">Xoa lich</button>
    </div>

    <div id="st"></div>
  </div>

  <script>
    function setSt(msg, cls){
      var el = document.getElementById('st');
      el.className = cls;
      el.textContent = msg;
    }

    function runCreate(){
      var start = document.getElementById('start').value;
      var end = document.getElementById('end').value;
      var plus = document.getElementById('plus').value;
      if(!start){ setSt('Vui long nhap ngay bat dau.', 'err'); return; }
      setSt('Dang tao lich...', 'run');
      google.script.run
        .withSuccessHandler(function(msg){ setSt(msg, 'ok'); })
        .withFailureHandler(function(err){ setSt(err.message || String(err), 'err'); })
        .generateScheduleFromDialog(start, end, plus);
    }

    function runClear(){
      setSt('Dang xoa...', 'run');
      google.script.run
        .withSuccessHandler(function(msg){ setSt(msg, 'ok'); })
        .withFailureHandler(function(err){ setSt(err.message || String(err), 'err'); })
        .clearScheduleSheet();
    }
  </script>
</body>
</html>`;
}

function clearScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(SHEET_SCHEDULE);
  if (old) ss.deleteSheet(old);
  return `Da xoa sheet "${SHEET_SCHEDULE}".`;
}

function generateScheduleFromDialog(startDateStr, endDateStr, plusDaysRaw) {
  const startDate = parseLocalDate(startDateStr);
  if (!isValidDate(startDate)) {
    throw new Error('Ngay bat dau khong hop le.');
  }

  let endDate = null;
  if (endDateStr) {
    endDate = parseLocalDate(endDateStr);
    if (!isValidDate(endDate)) throw new Error('Ngay ket thuc khong hop le.');
  } else {
    const plusDays = Number(plusDaysRaw);
    if (!Number.isFinite(plusDays) || plusDays < 1) {
      throw new Error('So ngay cong them phai >= 1.');
    }
    endDate = addDays(startDate, plusDays);
  }

  if (startDate >= endDate) {
    throw new Error('Ngay ket thuc phai sau ngay bat dau.');
  }

  return generateSchedule(startDate, endDate);
}

function generateSchedule(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const teacherSheet = ss.getSheetByName(SHEET_TEACHERS);
  if (!teacherSheet) {
    throw new Error(`Khong tim thay sheet "${SHEET_TEACHERS}".`);
  }
  const teachers = readTeachersBasic(teacherSheet);

  const src = ss.getSheetByName(SHEET_SOURCE);
  if (!src) {
    throw new Error(`Khong tim thay sheet "${SHEET_SOURCE}".`);
  }

  const scheduleSheet = recreateScheduleSheetFromSource(ss, src);

  const data = scheduleSheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error(`Sheet "${SHEET_SOURCE}" khong co du lieu de lap lich.`);
  }

  const header = data[0];
  const rows = data.slice(1);

  const idxCode = detectCodeColumn(header);
  const idxTotal = detectTotalTietColumn(header);

  if (idxCode < 0) {
    throw new Error('Khong tim thay cot ma MH/MD trong header Sheet2.');
  }

  const days = buildAllDays(startDate, endDate);
  if (days.length < 3) {
    throw new Error('Khoang ngay can it nhat 3 ngay (khai giang + hoc + be giang).');
  }

  const startCol = header.length + 1;
  scheduleSheet.insertColumnsAfter(header.length, days.length);

  writeDateHeaders(scheduleSheet, header.length, days, startDate, endDate);

  const groups = buildGroups(rows, idxCode, idxTotal);
  const planByGroup = buildDailyPlan(groups, days.length, days);

  writeScheduleValues(scheduleSheet, {
    rows,
    groups,
    planByGroup,
    idxCode,
    startCol,
    dayCount: days.length,
  });

  applyBasicStyles(scheduleSheet, {
    headerCols: header.length,
    startCol,
    days,
    startDate,
    endDate,
    rowCount: rows.length,
  });

  const stat = summarizePlan(groups, planByGroup, days.length);
  return [
    `Da tao sheet "${SHEET_SCHEDULE}" tu ban sao "${SHEET_SOURCE}".`,
    `Ngay: ${formatDateVN(startDate)} -> ${formatDateVN(endDate)} (${days.length} cot ngay).`,
    `Da doc ${teachers.length} giang vien tu "${SHEET_TEACHERS}".`,
    `Tong tiet nguon: ${stat.totalSource}. Tong tiet da xep: ${stat.totalPlaced}.`,
    stat.note,
  ].join('\n');
}

function readTeachersBasic(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values
    .slice(1)
    .map(r => String(r[0] || '').trim())
    .filter(Boolean);
}

function recreateScheduleSheetFromSource(ss, sourceSheet) {
  const old = ss.getSheetByName(SHEET_SCHEDULE);
  if (old) ss.deleteSheet(old);

  const cloned = sourceSheet.copyTo(ss);
  cloned.setName(SHEET_SCHEDULE);
  ss.setActiveSheet(cloned);
  ss.moveActiveSheet(ss.getNumSheets());
  return cloned;
}

function detectCodeColumn(header) {
  const lowers = header.map(h => String(h || '').trim().toLowerCase());
  let idx = lowers.findIndex(h => h.includes('ma') && (h.includes('mh') || h.includes('md')));
  if (idx >= 0) return idx;

  idx = lowers.findIndex(h => h.includes('ma'));
  if (idx >= 0) return idx;

  return header.length >= 3 ? 2 : -1;
}

function detectTotalTietColumn(header) {
  const lowers = header.map(h => String(h || '').trim().toLowerCase());

  let idx = lowers.findIndex(h => h.includes('tong') && h.includes('tiet'));
  if (idx >= 0) return idx;

  if (header.length > 7) {
    return -2; // special marker: sum col 6..8 (index 5,6,7)
  }

  return -1;
}

function buildAllDays(startDate, endDate) {
  const days = [];
  const d = new Date(startDate);
  while (d <= endDate) {
    days.push(new Date(d));
    d.setDate(d.getDate() + 1);
  }
  return days;
}

function writeDateHeaders(sheet, oldColCount, days, startDate, endDate) {
  const titles = days.map(d => formatDateVN(d));
  sheet.getRange(1, oldColCount + 1, 1, days.length).setValues([titles]);

  const row2 = days.map(d => {
    const dk = dayKey(d);
    if (dk === dayKey(startDate)) return 'Khai giang';
    if (dk === dayKey(endDate)) return 'Be giang';
    return thuVN(d.getDay());
  });

  sheet.insertRowAfter(1);
  sheet.getRange(2, oldColCount + 1, 1, days.length).setValues([row2]);
}

function buildGroups(rows, idxCode, idxTotal) {
  const map = {};

  rows.forEach((row, i) => {
    const code = String(row[idxCode] || '').trim();
    if (!code) return;

    let totalTiet = 0;
    if (idxTotal >= 0) {
      totalTiet = Number(row[idxTotal]) || 0;
    } else if (idxTotal === -2) {
      totalTiet = (Number(row[5]) || 0) + (Number(row[6]) || 0) + (Number(row[7]) || 0);
    }

    if (!map[code]) {
      map[code] = {
        code,
        rowIndexes: [],
        rowTotals: [],
      };
    }

    map[code].rowIndexes.push(i);
    map[code].rowTotals.push(totalTiet);
  });

  return Object.keys(map)
    .sort((a, b) => a.localeCompare(b))
    .map(k => {
      const g = map[k];
      g.rowCount = g.rowIndexes.length;
      g.sourceTiet = g.rowTotals.reduce((a, b) => a + b, 0);
      g.baseTiet = g.rowTotals.length ? g.rowTotals[0] : 0;
      g.perRowCap = g.rowTotals.length ? Math.min.apply(null, g.rowTotals) : 0;

      // Tong theo ma: de giu "cung ma cung gia tri cell", ta phai quy ve tong theo 1 dong
      const allEqual = g.rowTotals.every(v => Number(v) === Number(g.rowTotals[0]));
      // Dung gia tri chuan theo dong dau cua ma de giu pattern nhu mong muon (vd 10 -> 6,4)
      g.targetPerRow = Number(g.rowTotals[0]) || 0;
      g.sourceRemainder = g.sourceTiet - (g.targetPerRow * g.rowCount);
      if (!allEqual) g._nonUniform = true;
      return g;
    });
}

function buildDailyPlan(groups, dayCount, days) {
  // day index:
  // 0 = khai giang -> KG
  // dayCount-1 = be giang -> rong
  // Chu nhat: rong
  // Ngay hoc: xep block theo ma MH/MD
  const teachDayIdxList = [];
  for (let d = 1; d <= dayCount - 2; d++) {
    if (days[d].getDay() !== 0) teachDayIdxList.push(d);
  }
  const teachDayCount = teachDayIdxList.length;
  const planByGroup = {};

  if (teachDayCount <= 0) {
    groups.forEach(g => {
      planByGroup[g.code] = new Array(dayCount).fill(0);
    });
    return planByGroup;
  }

  groups.forEach(g => {
    planByGroup[g.code] = new Array(dayCount).fill(0);
  });

  const dayCaps = {};
  const dayGroupCount = {};
  teachDayIdxList.forEach(di => {
    dayCaps[di] = TIETS_PER_DAY;
    dayGroupCount[di] = 0;
  });

  const stats = {
    sourceRemainderTiet: 0,
    unscheduledPerRowByCode: {},
    notes: [],
  };

  groups.forEach(g => {
    stats.sourceRemainderTiet += Math.max(0, g.sourceRemainder || 0);
    if ((g.sourceRemainder || 0) > 0) {
      stats.notes.push(`Ma ${g.code}: tong tiet khong chia deu cho ${g.rowCount} dong.`);
    }
  });

  // Xep ma co tong tiet lon truoc de giam nguy co thieu ngay
  const orderedGroups = groups.slice().sort((a, b) => (b.targetPerRow || 0) - (a.targetPerRow || 0));
  orderedGroups.forEach(g => {
    const chunks = splitPerRowToValidChunks(g.targetPerRow || 0);
    if (!chunks.ok) {
      stats.unscheduledPerRowByCode[g.code] = chunks.unplanned || (g.targetPerRow || 0);
      stats.notes.push(`Ma ${g.code}: con duoi ${MIN_CELL_TIETS} tiet/row nen khong xep duoc.`);
      return;
    }

    const q = chunks.parts.slice().sort((a, b) => b - a);
    q.forEach(part => {
      const dayIdx = pickBestDayForPart(teachDayIdxList, dayCaps, dayGroupCount, planByGroup[g.code], part);
      let placed = false;
      if (dayIdx >= 0) {
        planByGroup[g.code][dayIdx] = part;
        dayCaps[dayIdx] -= part;
        dayGroupCount[dayIdx] += 1;
        placed = true;
      }

      if (!placed) {
        if (!stats.unscheduledPerRowByCode[g.code]) stats.unscheduledPerRowByCode[g.code] = 0;
        stats.unscheduledPerRowByCode[g.code] += part;
      }
    });
  });

  planByGroup._stats = stats;
  return planByGroup;
}

function pickBestDayForPart(teachDayIdxList, dayCaps, dayGroupCount, planOfCode, part) {
  let bestDay = -1;
  let bestScore = Number.POSITIVE_INFINITY;

  for (let i = 0; i < teachDayIdxList.length; i++) {
    const dayIdx = teachDayIdxList[i];
    if (planOfCode[dayIdx] > 0) continue;       // 1 ma chi 1 cell/ngay
    if (dayGroupCount[dayIdx] >= 2) continue;   // toi da 2 ma/ngay
    if (dayCaps[dayIdx] < part) continue;       // khong vuot 8/ngay

    // uu tien lap day cot ngay (remaining = 0), sau do remaining nho nhat
    const remaining = dayCaps[dayIdx] - part;
    const score = remaining === 0 ? -1 : remaining;
    if (score < bestScore) {
      bestScore = score;
      bestDay = dayIdx;
    }
  }

  return bestDay;
}

function splitPerRowToValidChunks(totalPerRow) {
  let total = Math.max(0, Number(totalPerRow) || 0);
  const res = { ok: true, parts: [], unplanned: 0 };
  if (total === 0) return res;
  if (total < MIN_CELL_TIETS) {
    res.ok = false;
    res.unplanned = total;
    return res;
  }

  const k = Math.floor(total / TIETS_PER_DAY);
  let rem = total % TIETS_PER_DAY;

  for (let i = 0; i < k; i++) res.parts.push(TIETS_PER_DAY);
  if (rem === 0) return res;
  if (rem >= MIN_CELL_TIETS) {
    res.parts.push(rem);
    return res;
  }

  // rem = 1..3 => muon giu moi cell >=4 thi can "muon" 1 chunk 8 de doi thanh cap hop le
  if (res.parts.length === 0) {
    res.ok = false;
    res.unplanned = total;
    return res;
  }
  res.parts.pop();
  if (rem === 1) res.parts.push(5, 4); // 8+1 -> 5+4
  if (rem === 2) res.parts.push(6, 4); // 8+2 -> 6+4
  if (rem === 3) res.parts.push(7, 4); // 8+3 -> 7+4
  return res;
}

function writeScheduleValues(sheet, ctx) {
  const { rows, groups, planByGroup, idxCode, startCol, dayCount } = ctx;
  if (!rows.length) return;

  // Du lieu dong bat dau tu row 3 vi da chen them 1 dong weekday
  const startRow = 3;
  const matrix = Array.from({ length: rows.length }, () => new Array(dayCount).fill(''));

  for (let i = 0; i < rows.length; i++) {
    const code = String(rows[i][idxCode] || '').trim();
    if (!code || !planByGroup[code]) continue;

    const plan = planByGroup[code];
    for (let d = 0; d < dayCount; d++) {
      if (d === 0) {
        matrix[i][d] = 'KG';
        continue;
      }
      matrix[i][d] = plan[d] > 0 ? plan[d] : '';
    }
  }

  sheet.getRange(startRow, startCol, rows.length, dayCount).setValues(matrix);
}

function applyBasicStyles(sheet, ctx) {
  const { headerCols, startCol, days, startDate, endDate, rowCount } = ctx;

  // Header styles
  sheet.getRange(1, startCol, 2, days.length)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground('#e8f0fe');

  // Highlight khai giang va be giang columns
  const startIdx = days.findIndex(d => dayKey(d) === dayKey(startDate));
  const endIdx = days.findIndex(d => dayKey(d) === dayKey(endDate));

  if (startIdx >= 0) {
    const col = startCol + startIdx;
    sheet.getRange(1, col, rowCount + 2, 1).setBackground('#fff3cd');
  }

  if (endIdx >= 0) {
    const col = startCol + endIdx;
    sheet.getRange(1, col, rowCount + 2, 1).setBackground('#fde2e2');
  }

  // Saturday/Sunday tint in day area
  days.forEach((d, i) => {
    const dow = d.getDay();
    const col = startCol + i;
    if (dow === 0) {
      sheet.getRange(1, col, rowCount + 2, 1).setBackground('#f1f5f9');
    } else if (dow === 6) {
      sheet.getRange(1, col, rowCount + 2, 1).setBackground('#f5f3ff');
    }
  });

  // Value cells center
  if (rowCount > 0) {
    sheet.getRange(3, startCol, rowCount, days.length)
      .setHorizontalAlignment('center')
      .setFontColor('#0f172a');
  }

  // Keep source part readable
  sheet.getRange(1, 1, 2, headerCols)
    .setBackground('#eef2ff')
    .setFontWeight('bold');

  sheet.setFrozenRows(2);
  sheet.autoResizeColumns(1, headerCols + days.length);
}

function summarizePlan(groups, planByGroup, dayCount) {
  let totalSource = 0;
  let totalPlaced = 0;

  groups.forEach(g => {
    totalSource += g.sourceTiet;
    const arr = planByGroup[g.code] || [];
    arr.forEach(v => { totalPlaced += (Number(v) || 0) * g.rowCount; });
  });

  const notPlaced = totalSource - totalPlaced;
  const teachDays = Math.max(0, dayCount - 2);
  const capacity = teachDays * TIETS_PER_DAY;

  const st = planByGroup._stats || {};
  const unscheduledByRules = Object.keys(st.unscheduledPerRowByCode || {}).reduce((acc, code) => {
    const g = groups.find(x => x.code === code);
    const perRow = st.unscheduledPerRowByCode[code] || 0;
    return acc + perRow * (g ? g.rowCount : 1);
  }, 0);
  const sourceRemainder = Number(st.sourceRemainderTiet || 0);

  let note = `Suc chua lich: ${capacity} tiet (${teachDays} ngay hoc x ${TIETS_PER_DAY}).`;
  note += ` Lech tong nguon-xep: ${notPlaced} tiet.`;
  if (sourceRemainder > 0) {
    note += ` Khong chia deu theo ma: ${sourceRemainder} tiet.`;
  }
  if (unscheduledByRules > 0) {
    note += ` Chua xep do rang buoc (>=${MIN_CELL_TIETS}, <=2 ma/ngay, <=8/ngay): ${unscheduledByRules} tiet.`;
  }
  if (st.notes && st.notes.length) {
    note += ` Ghi chu: ${st.notes.join(' | ')}`;
  }

  return { totalSource, totalPlaced, note };
}

function parseLocalDate(str) {
  const parts = String(str || '').split('-').map(Number);
  if (parts.length !== 3 || parts.some(n => !Number.isFinite(n))) return new Date('invalid');
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function isValidDate(d) {
  return d instanceof Date && !isNaN(d.getTime());
}

function addDays(d, n) {
  const x = new Date(d);
  x.setDate(x.getDate() + Number(n));
  return x;
}

function dayKey(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function formatDateVN(date) {
  return `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()}`;
}

function thuVN(n) {
  return ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'][n] || '';
}

function pad(n) {
  return String(n).padStart(2, '0');
}

function testRun() {
  generateScheduleFromDialog('2026-03-02', '2026-04-16', '');
}
