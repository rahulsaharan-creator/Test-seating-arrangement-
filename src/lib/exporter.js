// Builds the output workbook in the same layout as the 4-Nov example sheet.
import { naturalCompare, phaseOf } from './engine.js';

const C = {
  navy: 'FF1E3A5F', green: 'FF14532D', brown: 'FF713F12', text: 'FF333333',
  titleFill: 'FFDBEAFE', subFill: 'FFE0E7FF', headFill: 'FFDCFCE7', totFill: 'FFFEF9C3',
  bTotFill: 'FFE0F2FE', white: 'FFFFFFFF', grey: 'FFF3F4F6',
};
const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const thin = { style: 'thin', color: { argb: 'FF9CA3AF' } };
const med = { style: 'medium', color: { argb: 'FF1E3A5F' } };
const box = (s) => ({ top: s, left: s, bottom: s, right: s });
const center = { horizontal: 'center', vertical: 'middle', wrapText: true };

function style(cell, { bold = false, size = 10, color = C.text, bg = C.white, border = thin, align = center } = {}) {
  cell.font = { name: 'Arial', bold, size, color: { argb: color } };
  cell.fill = fill(bg);
  cell.alignment = align;
  if (border) cell.border = box(border);
}
const asNum = (v) => (/^\d+$/.test(String(v)) ? Number(v) : v);
const roomVal = asNum;

export const fmtTime = (hhmm) => {
  if (!hhmm) return '';
  const [h, m] = hhmm.split(':').map(Number);
  const ap = h >= 12 ? 'PM' : 'AM';
  return `${((h + 11) % 12) + 1}:${String(m).padStart(2, '0')} ${ap}`;
};
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const parseDate = (iso) => { const [y, m, d] = iso.split('-').map(Number); return { y, m, d }; };
export const fmtDateDMY = (iso) => { const { y, m, d } = parseDate(iso); return `${String(d).padStart(2,'0')}-${String(m).padStart(2,'0')}-${y}`; };
const fmtDateLong = (iso) => { const { m, d } = parseDate(iso); return `${d} ${MONTHS[m - 1]}`; };

function titleRow(ws, row, text, cols, opts) {
  ws.mergeCells(row, 1, row, cols);
  const c = ws.getCell(row, 1);
  c.value = text;
  style(c, opts);
  for (let i = 2; i <= cols; i++) style(ws.getCell(row, i), opts);
}

/**
 * meta: { title, dateISO, shifts: ['07:00', '10:00'?], batchesLabel, modeLabel }
 */
export function buildWorkbook(ExcelJS, result, meta) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Seating Arrangement Automation';
  const rows = result.assigned.slice().sort((a, b) => naturalCompare(a.Roll_Number, b.Roll_Number));
  const phases = [...new Set(rows.map((r) => r.CB))].sort();
  const rooms = result.usedRooms;
  const dateVal = (() => { const { y, m, d } = parseDate(meta.dateISO); return new Date(Date.UTC(y, m - 1, d)); })();
  const shiftLabels = meta.slots && meta.slots.length
    ? meta.slots.map((s) => (s.label ? `${s.label}: ` : '') + fmtTime(s.start) + (s.end ? ` to ${fmtTime(s.end)}` : ''))
    : meta.shifts.map(fmtTime);
  const shiftCols = shiftLabels.map((_, i) => `Shift ${i + 1}`);
  if (shiftCols.length < 2) shiftCols.push('Shift 2');

  // ---------- Raw_Data
  {
    const ws = wb.addWorksheet('Raw_Data');
    ws.columns = [{ width: 12 }, { width: 12 }, { width: 18 }, { width: 18 }];
    ['Building', 'Room No.', 'Seating Capacity', 'Students Allotted'].forEach((h, i) => {
      const c = ws.getCell(1, i + 1);
      c.value = h;
      style(c, { bold: true, size: 11, color: C.green, bg: C.headFill, border: med });
    });
    rooms.forEach((r, i) => {
      [r.building, roomVal(r.room), r.capacity, r.used].forEach((v, j) => { const c = ws.getCell(i + 2, j + 1); c.value = v; style(c); });
    });
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  // ---------- Student_list
  {
    const ws = wb.addWorksheet('Student_list');
    const heads = ['Roll_Number','Student_Name','Gender','Course','Batch','Medium','Date',...shiftCols,'Building','Room No.','Seat No.','CB','Mode'];
    ws.columns = heads.map((h) => ({ header: h, key: h, width: h === 'Student_Name' ? 30 : 12 }));
    ws.getRow(1).eachCell((c) => style(c, { bold: true, size: 11, color: C.green, bg: C.headFill }));
    for (const r of rows) {
      ws.addRow({
        ...r, Roll_Number: asNum(r.Roll_Number), Date: dateVal,
        ...Object.fromEntries(shiftCols.map((c, i) => [c, shiftLabels[i] || ''])), 'Room No.': roomVal(r['Room No.']),
      });
    }
    ws.getColumn('Date').numFmt = 'dd-mm-yyyy';
    ws.autoFilter = { from: 'A1', to: 'N1' };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  // ---------- Seating Arrangement (two side-by-side columns, sorted by roll number)
  {
    const ws = wb.addWorksheet('Seating Arrangement');
    ws.columns = [12.7, 10.8, 10.6, 9.1, 4, 12.7, 10.8, 10.6, 9.1].map((w) => ({ width: w }));
    titleRow(ws, 1, meta.title, 9, { bold: true, size: 22, color: C.navy, bg: C.titleFill, border: med });
    titleRow(ws, 2, `Mode: ${meta.modeLabel}`, 9, { bold: true, size: 13, color: C.navy, bg: C.subFill, border: med });
    titleRow(ws, 3, `Batch: ${meta.batchesLabel}`, 9, { bold: true, size: 13, color: C.navy, bg: C.subFill, border: med });
    titleRow(ws, 4, `Date:- ${fmtDateDMY(meta.dateISO)}          Time:- ${shiftLabels.join(' & ')}`, 9, { bold: true, size: 13, color: C.navy, bg: C.subFill, border: med });
    ws.getRow(1).height = 32;
    const H = ['Reg. No.', 'Building', 'Room No.', 'Seat No.'];
    H.forEach((h, i) => {
      [i + 1, i + 6].forEach((col) => { const c = ws.getCell(5, col); c.value = h; style(c, { bold: true, size: 11, color: C.green, bg: C.headFill, border: med }); });
    });
    const half = Math.ceil(rows.length / 2);
    for (let i = 0; i < half; i++) {
      [[rows[i], 1], [rows[i + half], 6]].forEach(([r, col]) => {
        if (!r) return;
        [asNum(r.Roll_Number), r.Building, roomVal(r['Room No.']), r['Seat No.']].forEach((v, j) => { const c = ws.getCell(6 + i, col + j); c.value = v; style(c); });
      });
    }
    ws.pageSetup = { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '5:5' };
  }

  // ---------- Attendance sheets, one per shift, one block per room with a page break
  shiftLabels.forEach((shift, si) => {
    const ws = wb.addWorksheet(`Attendance_${si + 1}Shift`);
    ws.columns = [{ width: 12.7 }, { width: 34.6 }, { width: 9.1 }, { width: 43.9 }];
    let r = 1;
    for (const room of rooms) {
      const occ = rows.filter((s) => s.Building === room.building && String(s['Room No.']) === String(room.room))
        .sort((a, b) => a['Seat No.'] - b['Seat No.']);
      titleRow(ws, r, `${room.building}  |  Room ${room.room}  |  ${fmtDateLong(meta.dateISO)}  |  ${shift}  (${occ.length} students)`, 4,
        { bold: true, size: 12, color: C.navy, bg: C.titleFill });
      r++;
      ['Reg. No.', 'Student Name', 'Seat No.', `Signature (${shift})`].forEach((h, i) => {
        const c = ws.getCell(r, i + 1);
        c.value = h;
        style(c, { bold: true, size: 11, color: C.green, bg: C.headFill });
      });
      r++;
      for (const s of occ) {
        [asNum(s.Roll_Number), s.Student_Name, s['Seat No.'], null].forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v; style(c); });
        ws.getRow(r).height = 20;
        r++;
      }
      const t = ws.getCell(r, 4);
      t.value = `Total: ${occ.length} students`;
      style(t, { bold: true, color: C.brown, bg: C.totFill, align: { horizontal: 'right' } });
      ws.getRow(r).addPageBreak();
      r += 2;
    }
    ws.pageSetup = { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  });

  // ---------- SMD: count per room by phase, block per building
  {
    const ws = wb.addWorksheet('SMD');
    const ncol = 3 + phases.length;
    ws.columns = [{ width: 12 }, { width: 11 }, ...phases.map(() => ({ width: 11 })), { width: 13 }];
    const grand = Object.fromEntries(phases.map((p) => [p, 0]));
    let r = 1;
    const buildings = [...new Set(rooms.map((x) => x.building))];
    for (const b of buildings) {
      titleRow(ws, r++, `🏢  ${b}`, ncol, { bold: true, size: 12, color: C.navy, bg: C.titleFill });
      ['Building', 'Room No.', ...phases, 'Grand Total'].forEach((h, i) => { const c = ws.getCell(r, i + 1); c.value = h; style(c, { bold: true, size: 11, color: C.green, bg: C.headFill, border: med }); });
      r++;
      const bt = Object.fromEntries(phases.map((p) => [p, 0]));
      for (const room of rooms.filter((x) => x.building === b)) {
        const occ = rows.filter((s) => s.Building === b && String(s['Room No.']) === String(room.room));
        const counts = phases.map((p) => occ.filter((s) => s.CB === p).length);
        counts.forEach((n, i) => { bt[phases[i]] += n; });
        [b, roomVal(room.room), ...counts.map((n) => n || null), occ.length].forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v; style(c); });
        r++;
      }
      const btTotal = phases.reduce((a, p) => a + bt[p], 0);
      [`${b} Total`, null, ...phases.map((p) => bt[p]), btTotal].forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v; style(c, { bold: true, color: C.navy, bg: C.bTotFill, border: med }); });
      phases.forEach((p) => { grand[p] += bt[p]; });
      r += 2;
    }
    ['Grand Total', null, ...phases.map((p) => grand[p]), rows.length].forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v; style(c, { bold: true, size: 11, color: C.navy, bg: C.subFill, border: med }); });
  }

  // ---------- Room_Layout: seat grid per room (seats run down each column, like the room list)
  {
    const ws = wb.addWorksheet('Room_Layout');
    let r = 1;
    const maxCols = Math.max(4, ...rooms.map((x) => x.cols || Math.min(10, Math.max(2, Math.ceil(x.capacity / 12)))));
    ws.columns = Array.from({ length: maxCols }, () => ({ width: 15 }));
    for (const room of rooms) {
      const cols = room.cols || Math.min(10, Math.max(2, Math.ceil(room.capacity / 12)));
      const nrows = Math.ceil(room.capacity / cols);
      const bySeat = {};
      rows.filter((s) => s.Building === room.building && String(s['Room No.']) === String(room.room)).forEach((s) => { bySeat[s['Seat No.']] = s; });
      titleRow(ws, r++, `${room.building} | Room No. ${room.room} | Strength ${room.capacity} | Allotted ${room.used}`, cols,
        { bold: true, size: 12, color: C.navy, bg: C.titleFill });
      for (let i = 0; i < cols; i++) { const c = ws.getCell(r, i + 1); c.value = 'Seat / Reg. No.'; style(c, { bold: true, color: C.green, bg: C.headFill }); }
      r++;
      for (let i = 0; i < nrows; i++) {
        for (let j = 0; j < cols; j++) {
          const seat = j * nrows + i + 1;
          if (seat > room.capacity) continue;
          const s = bySeat[seat];
          const c = ws.getCell(r, j + 1);
          c.value = s ? `Seat ${seat}\n${s.Roll_Number}` : `Seat ${seat}\n—`;
          style(c, { bg: s ? C.white : C.grey });
        }
        ws.getRow(r).height = 28;
        r++;
      }
      ws.getRow(r - 1).addPageBreak();
      r += 1;
    }
  }

  // ---------- Master_List: building entrance display
  {
    const ws = wb.addWorksheet('Master_List');
    ws.columns = [{ width: 12 }, { width: 11 }, { width: 11 }, { width: 12 }, { width: 15 }, { width: 15 }, { width: 30 }];
    titleRow(ws, 1, `${meta.title} — BUILDING MASTER LIST`, 7, { bold: true, size: 16, color: C.navy, bg: C.titleFill, border: med });
    titleRow(ws, 2, `Date: ${fmtDateDMY(meta.dateISO)}   Time: ${shiftLabels.join(' & ')}   —   Students check Room No. & Seat No. here`, 7,
      { bold: true, size: 11, color: C.navy, bg: C.subFill });
    let r = 4;
    for (const b of [...new Set(rooms.map((x) => x.building))]) {
      titleRow(ws, r++, b, 7, { bold: true, size: 12, color: C.navy, bg: C.titleFill });
      ['Building', 'Room No.', 'Strength', 'Allotted', 'Reg. No. From', 'Reg. No. To', 'Batches'].forEach((h, i) => { const c = ws.getCell(r, i + 1); c.value = h; style(c, { bold: true, color: C.green, bg: C.headFill }); });
      r++;
      for (const room of rooms.filter((x) => x.building === b)) {
        const occ = rows.filter((s) => s.Building === b && String(s['Room No.']) === String(room.room));
        const batches = [...new Set(occ.map((s) => s.Batch))].sort().join(', ');
        [b, roomVal(room.room), room.capacity, occ.length, asNum(occ[0]?.Roll_Number), asNum(occ[occ.length - 1]?.Roll_Number), batches]
          .forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v ?? null; style(c); });
        r++;
      }
      r++;
    }
  }

  // ---------- Not_Seated (only when someone could not get a seat)
  if (result.unseated.length) {
    const ws = wb.addWorksheet('Not_Seated');
    ws.columns = ['Roll_Number', 'Student_Name', 'Batch', 'Course', 'Reason'].map((h) => ({ header: h, key: h, width: h === 'Student_Name' ? 30 : 14 }));
    ws.getRow(1).eachCell((c) => style(c, { bold: true, color: C.brown, bg: C.totFill }));
    result.unseated.forEach((s) => ws.addRow({ ...s, Reason: 'No seat available / no building set' }));
  }

  return wb;
}

export { phaseOf };
