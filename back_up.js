// ============================================================
//  LỊCH GIẢNG DẠY AUTO-SCHEDULER  –  Google Apps Script
//
//  Rules:
//  • Mọi giảng viên đều phải có lịch dạy
//  • Tiết học được xếp liên tục / gần nhau (không rải thưa)
//  • Cột ngày hiển thị kể cả Chủ nhật (CN = highlight xám, không có tiết)
//  • Điều động cần đúng 4 giảng viên khác nhau
//  • Ngày khai giảng luôn trống, highlight vàng
// ============================================================

const SHEET_TEACHERS  = 'Sheet1';
const SHEET_SUBJECTS  = 'Sheet2';
const SHEET_SCHEDULE  = 'Lịch giảng dạy';
const SHEET_SUMMARY   = 'Tổng hợp';

const MAX_TIETS_DAY     = 8;
const MIN_TIETS_SESSION = 4;
const DIEU_DONG_TEACHERS_NEEDED = 4;

const KW_DIEU_DONG = ['điều động'];
const KW_PHAP_LUAT = ['pháp luật'];
const KW_VAN_TAI   = ['vận tải'];
const KW_AN_TOAN   = ['an toàn'];
const KW_BAO_DUONG = ['bảo dưỡng'];

// Colours
const C_HDR_DARK     = '#0d47a1';
const C_HDR_BLUE     = '#1a73e8';
const C_HDR_GREEN    = '#1b5e20';
const C_HIRED_BG     = '#fce8e6';
const C_HIRED_FG     = '#c5221f';
const C_KHAI_GIANG   = '#fbbc04';
const C_SUNDAY_BG    = '#eceff1';   // light grey – Sunday column
const C_SUNDAY_FG    = '#90a4ae';
const C_SAT_BG       = '#ede7f6';
const C_SAT_FG       = '#4a148c';
const C_ROW_A        = '#f8f9fa';
const C_ROW_B        = '#ffffff';
const C_CELL_FILL    = '#e3f2fd';
const C_CELL_FG      = '#0d47a1';

// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📅 Lịch Giảng Dạy')
    .addItem('🗓️  Tạo lịch tự động…', 'showScheduleDialog')
    .addSeparator()
    .addItem('🗑️  Xóa sheet đã tạo', 'clearGeneratedSheets')
    .addToUi();
}

// ============================================================
//  DIALOG
// ============================================================
function showScheduleDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(buildDialogHTML()).setWidth(520).setHeight(660),
    '🗓️  Thiết lập Lịch Giảng Dạy'
  );
}

function buildDialogHTML() {
  return `<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Google Sans',Arial,sans-serif;background:#f8f9fa;padding:20px;color:#202124}
  .card{background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,.1);padding:24px;margin-bottom:14px}
  h2{font-size:19px;font-weight:600;color:#1a73e8;margin-bottom:4px}
  .sub{font-size:12px;color:#5f6368;margin-bottom:20px}
  label{display:block;font-size:12px;font-weight:600;color:#3c4043;margin-bottom:5px;text-transform:uppercase;letter-spacing:.4px}
  input[type=date]{width:100%;padding:10px 13px;border:1.5px solid #dadce0;border-radius:8px;font-size:14px;outline:none;transition:border-color .2s;background:#fff;color:#202124}
  input[type=date]:focus{border-color:#1a73e8;box-shadow:0 0 0 3px rgba(26,115,232,.12)}
  .field{margin-bottom:16px}
  .hint{font-size:11px;color:#80868b;margin-top:3px}
  .btn-row{display:flex;gap:10px;margin-top:4px}
  button{flex:1;padding:11px;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;transition:all .2s}
  .bp{background:#1a73e8;color:#fff}.bp:hover{background:#1557b0}
  .bs{background:#e8f0fe;color:#1a73e8}.bs:hover{background:#d2e3fc}
  #status{margin-top:14px;padding:11px 14px;border-radius:8px;font-size:12px;display:none;line-height:1.6}
  .si{background:#e8f0fe;color:#1a73e8;display:block!important}
  .so{background:#e6f4ea;color:#137333;display:block!important}
  .se{background:#fce8e6;color:#c5221f;display:block!important}
  .spin{display:inline-block;width:13px;height:13px;border:2px solid #c8d8f8;border-top-color:#1a73e8;border-radius:50%;animation:sp .7s linear infinite;margin-right:6px;vertical-align:middle}
  @keyframes sp{to{transform:rotate(360deg)}}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:12px}
  .box{background:#f1f3f4;border-radius:8px;padding:10px;font-size:11px;line-height:1.6}
  .box b{color:#1a73e8;font-size:12px;display:block;margin-bottom:2px}
</style>
</head><body>
<div class="card">
  <h2>📅 Tạo Lịch Giảng Dạy</h2>
  <p class="sub">Phân bổ tự động từ Sheet1 (Giảng viên) &amp; Sheet2 (Môn học)</p>
  <div class="field">
    <label>📌 Ngày khai giảng</label>
    <input type="date" id="sd" value="2026-03-02">
    <p class="hint">Ngày đầu tiên – luôn trống, highlight vàng trên bảng tổng hợp</p>
  </div>
  <div class="field">
    <label>🏁 Ngày kết thúc</label>
    <input type="date" id="ed" value="2026-04-16">
    <p class="hint">Học Thứ 2 → Thứ 7 · Chủ nhật hiển thị nhưng không xếp tiết</p>
  </div>
  <div class="btn-row">
    <button class="bp" onclick="go()">⚡ Tạo lịch</button>
    <button class="bs" onclick="clr()">🗑️ Xóa lịch cũ</button>
  </div>
  <div id="status"></div>
</div>
<div class="card">
  <b style="font-size:13px;color:#1a73e8">📋 Quy tắc hệ thống</b>
  <div class="grid">
    <div class="box"><b>⏰ Tiết / ngày</b>Tối đa 8 tiết<br>Tối thiểu 4 tiết/buổi<br>Tiết liên tục, gần nhau</div>
    <div class="box"><b>🚢 Điều động</b>Đúng 4 GV khác nhau<br>≤ 1 lớp Điều động/ngày</div>
    <div class="box"><b>👨‍🏫 Tài &amp; Kha</b>Pháp luật · Vận tải<br>An toàn · Bảo dưỡng</div>
    <div class="box"><b>👥 Tất cả GV</b>Mọi GV đều<br>được phân lịch</div>
  </div>
</div>
<script>
function st(m,t){var e=document.getElementById('status');e.className=t;e.innerHTML=m;}
function go(){
  var s=document.getElementById('sd').value,e=document.getElementById('ed').value;
  if(!s||!e){st('⚠️ Vui lòng chọn đủ ngày bắt đầu và kết thúc.','se');return;}
  if(s>=e){st('⚠️ Ngày kết thúc phải sau ngày bắt đầu.','se');return;}
  st('<span class="spin"></span>Đang tạo lịch, vui lòng chờ…','si');
  google.script.run
    .withSuccessHandler(function(r){st('✅ '+r,'so');})
    .withFailureHandler(function(r){st('❌ '+r.message,'se');})
    .generateSchedule(s,e);
}
function clr(){
  st('<span class="spin"></span>Đang xóa…','si');
  google.script.run
    .withSuccessHandler(function(r){st('✅ '+r,'so');})
    .withFailureHandler(function(r){st('❌ '+r.message,'se');})
    .clearGeneratedSheets();
}
</script>
</body></html>`;
}

// ============================================================
//  CLEAR
// ============================================================
function clearGeneratedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_SCHEDULE, SHEET_SUMMARY].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });
  return `Đã xóa "${SHEET_SCHEDULE}" và "${SHEET_SUMMARY}" thành công.`;
}

// ============================================================
//  MAIN
// ============================================================
function generateSchedule(startDateStr, endDateStr) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const startDate = parseLocalDate(startDateStr);
  const endDate   = parseLocalDate(endDateStr);

  // allDays = every calendar day including Sunday (for display)
  const allDays     = buildAllDays(startDate, endDate);
  // teachingDays = Mon–Sat excluding khai giảng
  const kgKey       = dayKey(startDate);
  const teachingDays = allDays.filter(d => d.getDay() !== 0 && dayKey(d) !== kgKey);

  if (teachingDays.length === 0) {
    throw new Error('Không có ngày học hợp lệ (Thứ 2–Thứ 7, trừ ngày khai giảng) trong khoảng đã chọn.');
  }

  const teachers                        = readTeachers(ss);
  const { classes, rawRows, rawHeader } = readClasses(ss);

  // Assign teachers ensuring ALL teachers are used
  assignTeachersToClasses(classes, teachers);

  // Plan sessions per class
  classes.forEach(c => { c.sessions = planSessions(c); });

  // Schedule: pack sessions into consecutive blocks per class (close-together)
  const schedule = scheduleAllSessions(classes, teachingDays);

  writeScheduleSheet(ss, schedule, startDate, allDays);
  writeSummarySheet(ss, schedule, classes, rawRows, rawHeader, allDays, startDate);

  const w = (schedule._warnings || []).length;
  return `Tạo lịch thành công!${w ? ` (${w} cảnh báo – xem cuối Sheet lịch)` : ''} Kiểm tra "${SHEET_SCHEDULE}" và "${SHEET_SUMMARY}".`;
}

// ============================================================
//  READ TEACHERS
// ============================================================
function readTeachers(ss) {
  const data = ss.getSheetByName(SHEET_TEACHERS).getDataRange().getValues();
  return data.slice(1)
    .filter(r => String(r[0]).trim())
    .map(r => {
      const name       = String(r[0]).trim();
      const teachAll   = r[1] == 1 || String(r[1]).toLowerCase() === 'true';
      const subjectStr = String(r[2]).trim().toLowerCase();
      const onlyDieu   = subjectStr.includes('điều động') && !teachAll;
      return { name, teachAll, onlyDieu, hired: false, totalTiets: 0, assigned: false };
    });
}

// ============================================================
//  READ CLASSES
// ============================================================
function readClasses(ss) {
  const allData    = ss.getSheetByName(SHEET_SUBJECTS).getDataRange().getValues();
  const rawHeader  = allData[0];
  const rawRows    = [];
  const classes    = [];
  const countByKey = {};

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (!row[3]) continue;

    const subjectFull = String(row[3]).trim();
    const subjectKey  = subjectFull.toLowerCase();
    countByKey[subjectKey] = (countByKey[subjectKey] || 0) + 1;
    const idx = countByKey[subjectKey];

    const ltH  = Number(row[5]) || 0;
    const thH  = Number(row[6]) || 0;
    const ktH  = Number(row[7]) || 0;

    rawRows.push(row);
    classes.push({
      id:           classes.length,
      className:    `${toShortName(subjectFull)} ${idx}`,
      subject:      subjectFull,
      subjectKey,
      isDieu:       matchesKW(subjectKey, KW_DIEU_DONG),
      ltTiets:      ltH + ktH,
      thTiets:      thH,
      totalTiets:   ltH + thH + ktH,
      teacher:      null,
      teacherHired: false,
      sessions:     [],
    });
  }
  return { classes, rawRows, rawHeader };
}

// ============================================================
//  BUILD DAYS
// ============================================================
function buildAllDays(start, end) {
  const days = [];
  const d    = new Date(start);
  while (d <= end) { days.push(new Date(d)); d.setDate(d.getDate() + 1); }
  return days;
}

// ============================================================
//  ASSIGN TEACHERS  (ensures every teacher is used)
// ============================================================
function assignTeachersToClasses(classes, teachers) {
  let hiredSeq = 0;
  function hireNew() {
    hiredSeq++;
    const t = { name: `GV Thuê ${hiredSeq}`, teachAll: true, onlyDieu: false,
                hired: true, totalTiets: 0, assigned: false };
    teachers.push(t);
    return t;
  }

  // ── Step 1: Assign Điều động — needs exactly DIEU_DONG_TEACHERS_NEEDED distinct teachers ──
  const dieuClasses = classes.filter(c => c.isDieu);

  // Pool for Điều động: teachAll OR onlyDieu (Vấn must be included first)
  const dieuPool = teachers
    .filter(t => t.teachAll || t.onlyDieu)
    .sort((a, b) => {
      // onlyDieu (Vấn) first, then by totalTiets
      if (b.onlyDieu !== a.onlyDieu) return b.onlyDieu ? 1 : -1;
      return a.totalTiets - b.totalTiets;
    });

  // Ensure we have exactly DIEU_DONG_TEACHERS_NEEDED teachers for Điều động classes
  while (dieuPool.length < Math.min(DIEU_DONG_TEACHERS_NEEDED, dieuClasses.length)) {
    dieuPool.push(hireNew());
  }

  const usedForDieu = new Set();
  dieuClasses.forEach((cls, idx) => {
    // Cycle through dieuPool if we have more Điều động classes than teachers available
    const poolIdx = idx % dieuPool.length;
    // But prefer unused teachers first
    const unused  = dieuPool.filter(t => !usedForDieu.has(t.name));
    const t       = unused.length > 0
      ? unused.sort((a, b) => a.totalTiets - b.totalTiets)[0]
      : dieuPool.sort((a, b) => a.totalTiets - b.totalTiets)[0];

    usedForDieu.add(t.name);
    cls.teacher      = t.name;
    cls.teacherHired = t.hired;
    t.totalTiets    += cls.totalTiets;
    t.assigned       = true;
  });

  // ── Step 2: Assign other classes — prefer teachers not yet assigned ──
  const otherClasses = classes.filter(c => !c.isDieu);
  otherClasses.forEach(cls => {
    // Eligible = can teach this subject, not onlyDieu
    const eligible = teachers.filter(t => isEligible(t, cls.subjectKey) && !t.onlyDieu);
    if (eligible.length === 0) {
      const t = hireNew();
      cls.teacher = t.name; cls.teacherHired = true; t.assigned = true; t.totalTiets += cls.totalTiets;
      return;
    }
    // Prefer: not yet assigned → lower load
    const unassigned = eligible.filter(t => !t.assigned);
    const pool       = unassigned.length > 0 ? unassigned : eligible;
    const t          = pool.sort((a, b) => a.totalTiets - b.totalTiets)[0];
    cls.teacher      = t.name;
    cls.teacherHired = t.hired;
    t.totalTiets    += cls.totalTiets;
    t.assigned       = true;
  });

  // ── Step 3: Force-assign any still-unassigned real teacher to a compatible class ──
  // This guarantees EVERY teacher participates
  const unassignedTeachers = teachers.filter(t => !t.assigned && !t.hired);
  unassignedTeachers.forEach(t => {
    // Find a class this teacher can share (i.e., swap in as co-teacher by splitting sessions)
    // We look for a class with totalTiets > MIN_TIETS_SESSION * 2 and eligible teacher
    const compatible = classes
      .filter(cls => !cls.isDieu && isEligible(t, cls.subjectKey) && !t.onlyDieu && cls.totalTiets >= MIN_TIETS_SESSION * 2)
      .sort((a, b) => b.totalTiets - a.totalTiets);

    if (compatible.length > 0) {
      // Split the class: current teacher keeps first half, unassigned teacher takes second half
      const cls    = compatible[0];
      const shared = createSharedClass(cls, t);
      classes.push(shared);
      t.totalTiets += shared.totalTiets;
      t.assigned    = true;
    } else {
      // No compatible class — create a note (will be shown in warnings)
      t._noClass = true;
    }
  });
}

// Creates a "split" version of a class for a second teacher
function createSharedClass(originalCls, teacher) {
  const splitTiets = Math.floor(originalCls.totalTiets / 2);
  const ltSplit    = Math.min(splitTiets, originalCls.ltTiets);
  const thSplit    = Math.max(0, splitTiets - ltSplit);

  // Reduce original
  originalCls.ltTiets    -= ltSplit;
  originalCls.thTiets    -= thSplit;
  originalCls.totalTiets -= (ltSplit + thSplit);

  return {
    id:           -1,
    className:    originalCls.className,  // same class name, same subject
    subject:      originalCls.subject,
    subjectKey:   originalCls.subjectKey,
    isDieu:       false,
    ltTiets:      ltSplit,
    thTiets:      thSplit,
    totalTiets:   ltSplit + thSplit,
    teacher:      teacher.name,
    teacherHired: teacher.hired,
    sessions:     [],
  };
}

// ============================================================
//  PLAN SESSIONS  (split tiets into valid day-size chunks)
// ============================================================
function planSessions(cls) {
  const sessions = [];

  function split(total, type) {
    if (total <= 0) return;
    let rem = total;

    while (rem > 0) {
      let chosen = 0;

      // Tìm chunk lớn nhất thỏa mãn: phần còn lại = 0 hoặc ≥ MIN_TIETS_SESSION
      for (let s = MAX_TIETS_DAY; s >= MIN_TIETS_SESSION; s--) {
        const after = rem - s;
        if (after === 0 || after >= MIN_TIETS_SESSION) {
          chosen = s;
          break;
        }
      }

      if (chosen === 0) {
        // rem < MIN_TIETS_SESSION (ví dụ: còn 1, 2, hoặc 3 tiết)
        // ✅ FIX: Chỉ merge vào session cùng type nếu KHÔNG vượt MAX_TIETS_DAY
        const last = [...sessions]
          .reverse()
          .find(s => s.type === type && s.tiets + rem <= MAX_TIETS_DAY);

        if (last) {
          // Merge an toàn — session kết quả vẫn ≤ 8
          last.tiets += rem;
        } else {
          // Không tìm được session để merge (ví dụ: tất cả đã đủ 8)
          // → Tách thêm: giảm session trước xuống để nhường chỗ
          const donor = [...sessions]
            .reverse()
            .find(s => s.type === type && s.tiets > MIN_TIETS_SESSION);

          if (donor && donor.tiets - 1 >= MIN_TIETS_SESSION) {
            // Nhường 1 tiết từ donor sang rem
            donor.tiets -= 1;
            rem        += 1;
            // Giờ rem ≥ MIN_TIETS_SESSION (nếu rem ban đầu = MIN-1 thì nay = MIN)
            // Lặp lại vòng while để xử lý rem mới
            continue;
          } else {
            // Không thể tái phân phối → push standalone (chấp nhận < MIN, nhưng đảm bảo ≤ MAX)
            sessions.push({ tiets: Math.min(rem, MAX_TIETS_DAY), type });
            rem -= Math.min(rem, MAX_TIETS_DAY);
          }
        }
        rem = 0;
        break;
      }

      sessions.push({ tiets: chosen, type });
      rem -= chosen;
    }
  }

  split(cls.ltTiets, 'Lý thuyết');
  split(cls.thTiets, 'Thực hành');

  // Bảo vệ cuối cùng: clamp bất kỳ session nào vượt MAX (không nên xảy ra sau fix)
  sessions.forEach(s => {
    if (s.tiets > MAX_TIETS_DAY) s.tiets = MAX_TIETS_DAY;
  });

  return sessions.map(s => ({
    ...s,
    className:    cls.className,
    subject:      cls.subject,
    teacher:      cls.teacher,
    teacherHired: cls.teacherHired,
    isDieu:       cls.isDieu,
  }));
}

// ============================================================
//  SCHEDULE SESSIONS  —  CONSECUTIVE / CLOSE-TOGETHER
//
//  Strategy:
//  1. For each class, place all its sessions into consecutive available days
//     starting from an offset spread across the calendar.
//  2. "Close-together" = sessions of a class are placed in adjacent teaching days
//     (not scattered across the whole range), unless blocked by constraints.
//  3. Điều động constraint: max 1 Điều động class per day.
//  4. Teacher constraint: teacher can't be in 2 different classes on same day.
// ============================================================
function scheduleAllSessions(classes, teachingDays) {
  // dayState[dk] = { teachers: { name: {className, tiets} }, dieuCount }
  const dayState = {};
  teachingDays.forEach(d => {
    dayState[dayKey(d)] = { teachers: {}, dieuCount: 0, totalTiets: 0 };
  });

  const schedule = [];
  const warnings = [];
  const nDays    = teachingDays.length;

  // Sort: Điều động first, then by totalTiets desc (place big classes early)
  const ordered = [
    ...classes.filter(c => c.isDieu),
    ...classes.filter(c => !c.isDieu).sort((a, b) => b.totalTiets - a.totalTiets),
  ];

  // Divide calendar into bands, one band per class for "close-together" placement
  // Each class gets a preferred start index spread evenly
  const totalClasses  = ordered.length;
  const bandSize      = Math.max(1, Math.floor(nDays / (totalClasses + 1)));

  ordered.forEach((cls, classIdx) => {
    const sessions = cls.sessions;
    if (!sessions.length) return;

    // Preferred start day for this class band
    const preferredStart = Math.min(classIdx * bandSize, nDays - 1);

    // Place sessions of this class consecutively starting from preferredStart
    let cursor = preferredStart;

    sessions.forEach(sess => {
       let placed = false;
  for (let attempt = 0; attempt < nDays; attempt++) {
    const idx = (cursor + attempt) % nDays;
    const day = teachingDays[idx];
    const dk = dayKey(day);
    const state = dayState[dk];
    // Constraint: max 1 Điều động per day
    if (sess.isDieu && state.dieuCount >= 1) continue;
    // Constraint: teacher free
    const tEntry = state.teachers[sess.teacher];
    if (tEntry) {
      if (tEntry.className !== sess.className) continue;
      if (tEntry.tiets + sess.tiets > MAX_TIETS_DAY) continue;
    }

        // ✅ Place
        if (!state.teachers[sess.teacher]) {
          state.teachers[sess.teacher] = { className: sess.className, tiets: 0 };
        }
        state.teachers[sess.teacher].tiets += sess.tiets;
        if (sess.isDieu) state.dieuCount++;

        schedule.push({
          date:         day,
          className:    sess.className,
          subject:      sess.subject,
          type:         sess.type,
          tiets:        sess.tiets,
          teacher:      sess.teacher,
          teacherHired: sess.teacherHired,
        });

        // Advance cursor to next day for consecutive packing
        cursor = (idx + 1) % nDays;
        placed = true;
        break;
      }

      if (!placed) {
        warnings.push(`Không đủ ngày: ${sess.className} (${sess.teacher}) – ${sess.tiets} tiết`);
      }
    });
  });

  schedule.sort((a, b) => a.date - b.date || a.className.localeCompare(b.className));
  schedule._warnings = warnings;
  return schedule;
}

// ============================================================
//  WRITE "Lịch giảng dạy"
// ============================================================
function writeScheduleSheet(ss, schedule, startDate, allDays) {
  let ws = ss.getSheetByName(SHEET_SCHEDULE);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(SHEET_SCHEDULE);

  const headers = ['Ngày', 'Thứ', 'Tên lớp', 'Môn học', 'Loại', 'Số tiết', 'Giảng viên', 'Ghi chú'];
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(C_HDR_BLUE).setFontColor('#fff').setFontWeight('bold').setFontSize(11);

  const rows = [];

  // Khai giảng marker (first row, no session data)
  rows.push([
    formatDateVN(startDate), thuVN(startDate.getDay()),
    '', '', '', '', '', '🎓 NGÀY KHAI GIẢNG',
  ]);

  schedule.forEach(row => {
    rows.push([
      formatDateVN(row.date),
      thuVN(row.date.getDay()),
      row.className,
      row.subject,
      row.type,
      row.tiets,
      row.teacher,
      row.teacherHired ? '⚠️ Cần thuê giảng viên' : '',
    ]);
  });

  if (rows.length > 0) {
    ws.getRange(2, 1, rows.length, headers.length).setValues(rows);
    styleScheduleSheet(ws, rows, headers.length);
  }

  const warns = schedule._warnings || [];
  if (warns.length > 0) {
    const wRow = rows.length + 3;
    ws.getRange(wRow, 1).setValue('⚠️ CẢNH BÁO').setFontWeight('bold').setFontColor(C_HIRED_FG);
    warns.forEach((w, i) =>
      ws.getRange(wRow + 1 + i, 1).setValue('• ' + w).setFontColor(C_HIRED_FG));
  }

  ws.setFrozenRows(1);
  ws.autoResizeColumns(1, headers.length);
}

function styleScheduleSheet(ws, rows, colCount) {
  let prevDate = '', flip = false;
  rows.forEach((row, i) => {
    const r = i + 2;
    if (row[0] !== prevDate) { flip = !flip; prevDate = row[0]; }
    if (row[7] && row[7].includes('KHAI GIẢNG')) {
      ws.getRange(r, 1, 1, colCount).setBackground(C_KHAI_GIANG).setFontWeight('bold').setFontColor('#5d2a00');
      return;
    }
    ws.getRange(r, 1, 1, colCount).setBackground(flip ? C_ROW_A : C_ROW_B);
    if (row[4] === 'Lý thuyết') ws.getRange(r, 5).setBackground('#e8f0fe').setFontColor(C_HDR_BLUE);
    else if (row[4] === 'Thực hành') ws.getRange(r, 5).setBackground('#e6f4ea').setFontColor('#137333');
    if (row[7] && row[7].includes('Cần thuê')) {
      ws.getRange(r, 7).setBackground(C_HIRED_BG).setFontColor(C_HIRED_FG);
      ws.getRange(r, 8).setFontColor(C_HIRED_FG).setFontWeight('bold');
    }
  });
  ws.getRange(2, 1, rows.length, 1).setFontWeight('bold');
  ws.getRange(2, 6, rows.length, 1).setHorizontalAlignment('center');
  ws.getRange(1, 1, rows.length + 1, colCount)
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
}

// ============================================================
//  WRITE "Tổng hợp" — MATRIX
//
//  Columns: [Sheet2 gốc 8 cột] | Giảng viên | Ghi chú | [mỗi ngày kể cả CN]
//  Mỗi cell ngày = số tiết (chỉ số), trống nếu không có
//  Ngày khai giảng = trống + vàng
//  Chủ nhật = trống + xám
// ============================================================
function writeSummarySheet(ss, schedule, classes, rawRows, rawHeader, allDays, startDate) {
  let ws = ss.getSheetByName(SHEET_SUMMARY);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(SHEET_SUMMARY);

  // Build lookup: scheduleMap[className][dateKey] = total tiets
  const scheduleMap = {};
  schedule.forEach(row => {
    const dk = dayKey(row.date);
    if (!scheduleMap[row.className])     scheduleMap[row.className]     = {};
    if (!scheduleMap[row.className][dk]) scheduleMap[row.className][dk] = 0;
    scheduleMap[row.className][dk] += row.tiets;
  });

  const SH2_COLS      = rawHeader.length;   // 8
  const COL_GV        = SH2_COLS;           // idx 8
  const COL_NOTE      = SH2_COLS + 1;       // idx 9
  const COL_DAY_START = SH2_COLS + 2;       // idx 10
  const TOTAL_COLS    = COL_DAY_START + allDays.length;
  const kgKey         = dayKey(startDate);

  // ── ROW 1: header ──
  const hdr1 = new Array(TOTAL_COLS).fill('');
  rawHeader.forEach((h, i) => { hdr1[i] = h || ''; });
  hdr1[COL_GV]   = 'Giảng viên';
  hdr1[COL_NOTE] = 'Ghi chú';
  allDays.forEach((d, i) => {
    hdr1[COL_DAY_START + i] = `${pad(d.getDate())}/${pad(d.getMonth()+1)}`;
  });

  // ── ROW 2: weekday sub-header ──
  const hdr2 = new Array(TOTAL_COLS).fill('');
  allDays.forEach((d, i) => {
    const isKG  = dayKey(d) === kgKey;
    const isSun = d.getDay() === 0;
    hdr2[COL_DAY_START + i] = thuVN(d.getDay()) + (isKG ? ' ★' : '');
  });

  // ── DATA ROWS ──
  const dataRows = classes.map((cls, ri) => {
    const raw = rawRows[ri] || [];
    const row = new Array(TOTAL_COLS).fill('');
    for (let c = 0; c < SH2_COLS; c++) row[c] = raw[c] !== undefined ? raw[c] : '';
    row[COL_GV]   = cls.teacher || '';
    row[COL_NOTE] = cls.teacherHired ? 'Cần thuê' : '';

    const classSched = scheduleMap[cls.className] || {};
    allDays.forEach((d, i) => {
      const dk = dayKey(d);
      // Khai giảng & Sunday → always blank
      if (dk === kgKey || d.getDay() === 0) return;
      const tiets = classSched[dk] || 0;
      row[COL_DAY_START + i] = tiets > 0 ? tiets : '';
    });

    return { row, cls };
  });

  // ── FOOTER: total tiets per day ──
  const footerRow = new Array(TOTAL_COLS).fill('');
  footerRow[0] = 'Tổng tiết / ngày';
  let grandTotal = 0;
  allDays.forEach((d, i) => {
    const dk = dayKey(d);
    if (dk === kgKey || d.getDay() === 0) return;
    const sum = schedule.filter(s => dayKey(s.date) === dk).reduce((acc, s) => acc + s.tiets, 0);
    footerRow[COL_DAY_START + i] = sum > 0 ? sum : '';
    grandTotal += sum;
  });
  footerRow[COL_GV] = `Tổng: ${grandTotal}t`;

  // ── WRITE ──
  const ROW_H1     = 1;
  const ROW_H2     = 2;
  const ROW_DATA   = 3;
  const nData      = dataRows.length;
  const ROW_FOOTER = ROW_DATA + nData + 1;

  ws.getRange(ROW_H1, 1, 1, TOTAL_COLS).setValues([hdr1]);
  ws.getRange(ROW_H2, 1, 1, TOTAL_COLS).setValues([hdr2]);
  if (nData > 0) ws.getRange(ROW_DATA, 1, nData, TOTAL_COLS).setValues(dataRows.map(d => d.row));
  ws.getRange(ROW_FOOTER, 1, 1, TOTAL_COLS).setValues([footerRow]);

  // ── STYLES ──
  applySummaryStyles(ws, {
    ROW_H1, ROW_H2, ROW_DATA, nData, ROW_FOOTER,
    SH2_COLS, COL_GV, COL_NOTE, COL_DAY_START,
    TOTAL_COLS, allDays, dataRows, kgKey, scheduleMap,
  });

  ws.setFrozenRows(2);
  ws.setFrozenColumns(COL_DAY_START);

  // Column widths
  ws.setColumnWidth(1, 44);
  ws.setColumnWidth(2, 110);
  ws.setColumnWidth(3, 62);
  ws.setColumnWidth(4, 260);
  for (let c = 5; c <= 8; c++) ws.setColumnWidth(c, 52);
  ws.setColumnWidth(COL_GV   + 1, 145);
  ws.setColumnWidth(COL_NOTE + 1, 80);
  allDays.forEach((d, i) => {
    // Sunday columns narrower
    ws.setColumnWidth(COL_DAY_START + 1 + i, d.getDay() === 0 ? 32 : 48);
  });
}

// ──────────────────────────────
//  Summary styles
// ──────────────────────────────
function applySummaryStyles(ws, ctx) {
  const {
    ROW_H1, ROW_H2, ROW_DATA, nData, ROW_FOOTER,
    SH2_COLS, COL_GV, COL_NOTE, COL_DAY_START,
    TOTAL_COLS, allDays, dataRows, kgKey, scheduleMap,
  } = ctx;

  const LAST_ROW = ROW_FOOTER;
  const NUM_ROWS = LAST_ROW - ROW_H1 + 1;

  // ── Header Row 1 ──
  ws.getRange(ROW_H1, 1, 1, SH2_COLS)
    .setBackground(C_HDR_DARK).setFontColor('#fff').setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center').setWrap(true);
  ws.getRange(ROW_H1, COL_GV + 1, 1, 2)
    .setBackground(C_HDR_GREEN).setFontColor('#fff').setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center');
  ws.getRange(ROW_H1, COL_DAY_START + 1, 1, allDays.length)
    .setBackground(C_HDR_BLUE).setFontColor('#fff').setFontWeight('bold')
    .setFontSize(9).setHorizontalAlignment('center');

  // ── Header Row 2 ──
  ws.getRange(ROW_H2, 1, 1, SH2_COLS)
    .setBackground('#e3f2fd').setFontColor('#0d47a1').setFontSize(9);
  ws.getRange(ROW_H2, COL_GV + 1, 1, 2)
    .setBackground('#e8f5e9').setFontColor(C_HDR_GREEN).setFontSize(9);
  ws.getRange(ROW_H2, COL_DAY_START + 1, 1, allDays.length)
    .setBackground('#e8f0fe').setFontColor(C_HDR_BLUE).setFontWeight('bold')
    .setFontSize(9).setHorizontalAlignment('center');

  // Per-day column header colouring + full column stripe for Sun/KG/Sat
  allDays.forEach((d, i) => {
    const col   = COL_DAY_START + 1 + i;
    const isKG  = dayKey(d) === kgKey;
    const isSun = d.getDay() === 0;
    const isSat = d.getDay() === 6;

    if (isKG) {
      // Amber column header
      ws.getRange(ROW_H1, col).setBackground(C_KHAI_GIANG).setFontColor('#5d2a00');
      ws.getRange(ROW_H2, col).setBackground(C_KHAI_GIANG).setFontColor('#5d2a00');
      // Amber stripe through all data rows + footer
      ws.getRange(ROW_DATA, col, nData + 1, 1).setBackground(C_KHAI_GIANG);
    } else if (isSun) {
      // Grey column header
      ws.getRange(ROW_H1, col).setBackground(C_SUNDAY_BG).setFontColor(C_SUNDAY_FG);
      ws.getRange(ROW_H2, col).setBackground(C_SUNDAY_BG).setFontColor(C_SUNDAY_FG);
      // Grey stripe through data + footer
      ws.getRange(ROW_DATA, col, nData + 1, 1).setBackground(C_SUNDAY_BG).setFontColor(C_SUNDAY_FG);
    } else if (isSat) {
      ws.getRange(ROW_H1, col).setBackground(C_SAT_BG).setFontColor(C_SAT_FG);
      ws.getRange(ROW_H2, col).setBackground(C_SAT_BG).setFontColor(C_SAT_FG);
    }
  });

  // ── Data rows ──
  dataRows.forEach(({ row, cls }, ri) => {
    const r  = ROW_DATA + ri;
    const bg = ri % 2 === 0 ? C_ROW_A : C_ROW_B;
    ws.getRange(r, 1, 1, TOTAL_COLS)
      .setBackground(bg).setFontSize(10).setVerticalAlignment('middle');

    // GV / Note
    if (cls.teacherHired) {
      ws.getRange(r, COL_GV   + 1).setBackground(C_HIRED_BG).setFontColor(C_HIRED_FG).setFontWeight('bold');
      ws.getRange(r, COL_NOTE + 1).setBackground(C_HIRED_BG).setFontColor(C_HIRED_FG).setFontWeight('bold');
    } else {
      ws.getRange(r, COL_GV + 1).setFontWeight('bold').setFontColor('#1b5e20');
    }

    // Date cells — override column stripe with tiets colour when tiets > 0
    const classSched = scheduleMap[cls.className] || {};
    allDays.forEach((d, i) => {
      const col   = COL_DAY_START + 1 + i;
      const isKG  = dayKey(d) === kgKey;
      const isSun = d.getDay() === 0;
      if (isKG || isSun) return;  // already striped, leave blank

      const dk    = dayKey(d);
      const tiets = classSched[dk] || 0;
      if (tiets > 0) {
        ws.getRange(r, col)
          .setBackground(C_CELL_FILL).setFontColor(C_CELL_FG)
          .setFontWeight('bold').setHorizontalAlignment('center');
      }
      // Saturday tint for empty cells
      if (d.getDay() === 6 && tiets === 0) {
        ws.getRange(r, col).setBackground(C_SAT_BG);
      }
    });
  });

  // ── Footer row ──
  ws.getRange(ROW_FOOTER, 1, 1, TOTAL_COLS)
    .setBackground('#e8f0fe').setFontColor(C_HDR_BLUE).setFontWeight('bold').setFontSize(10);
  ws.getRange(ROW_FOOTER, COL_DAY_START + 1, 1, allDays.length).setHorizontalAlignment('center');
  ws.getRange(ROW_FOOTER, COL_GV + 1).setHorizontalAlignment('center');

  // ── Borders ──
  ws.getRange(ROW_H1, 1, NUM_ROWS, TOTAL_COLS)
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  ws.getRange(ROW_H1, COL_GV + 1, NUM_ROWS, 1)
    .setBorder(null, true, null, null, null, null, C_HDR_BLUE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  ws.getRange(ROW_H1, COL_DAY_START + 1, NUM_ROWS, 1)
    .setBorder(null, true, null, null, null, null, C_HDR_GREEN, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  ws.getRange(ROW_FOOTER, 1, 1, TOTAL_COLS)
    .setBorder(true, null, null, null, null, null, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Row heights
  ws.setRowHeight(ROW_H1, 32);
  ws.setRowHeight(ROW_H2, 24);
  for (let r = ROW_DATA; r <= LAST_ROW; r++) ws.setRowHeight(r, 22);
}

// ============================================================
//  HELPERS
// ============================================================

function isEligible(teacher, subjectKey) {
  if (teacher.onlyDieu) return matchesKW(subjectKey, KW_DIEU_DONG);
  if (teacher.teachAll) return true;
  return matchesKW(subjectKey, KW_PHAP_LUAT) ||
         matchesKW(subjectKey, KW_VAN_TAI)   ||
         matchesKW(subjectKey, KW_AN_TOAN)   ||
         matchesKW(subjectKey, KW_BAO_DUONG);
}

function matchesKW(str, keywords) {
  return keywords.some(kw => str.includes(kw));
}

function parseLocalDate(str) {
  const [y, m, d] = str.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function dayKey(date) {
  return `${date.getFullYear()}-${pad(date.getMonth()+1)}-${pad(date.getDate())}`;
}

function formatDateVN(date) {
  return `${pad(date.getDate())}/${pad(date.getMonth()+1)}/${date.getFullYear()}`;
}

function thuVN(n) {
  return ['CN','T2','T3','T4','T5','T6','T7'][n] || '';
}

function pad(n) { return String(n).padStart(2, '0'); }

function toShortName(subject) {
  const s = subject.toLowerCase();
  if (s.includes('an toàn'))     return 'An toàn';
  if (s.includes('thủy nghiệp')) return 'Thủy nghiệp';
  if (s.includes('luồng đường')) return 'Luồng đường';
  if (s.includes('pháp luật'))   return 'Pháp luật';
  if (s.includes('điều động'))   return 'Điều động';
  if (s.includes('vận tải'))     return 'Vận tải';
  if (s.includes('bảo dưỡng'))   return 'Bảo dưỡng';
  return subject.slice(0, 12);
}

// ============================================================
//  TEST
// ============================================================
function testRun() {
  generateSchedule('2025-09-01', '2025-10-31');
}