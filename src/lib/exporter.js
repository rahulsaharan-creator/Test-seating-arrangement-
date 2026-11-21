// Builds the output workbook in the same layout as the 4-Nov example sheet.
import { naturalCompare, phaseOf } from './engine.js';

const C = {
  navy: 'FF1E3A5F', green: 'FF14532D', brown: 'FF713F12', text: 'FF333333',
  titleFill: 'FFDBEAFE', subFill: 'FFE0E7FF', headFill: 'FFDCFCE7', totFill: 'FFFEF9C3',
  bTotFill: 'FFE0F2FE', white: 'FFFFFFFF', grey: 'FFF3F4F6',
};
const MARGINS = { left: 0.19685, right: 0.19685, top: 0.19685, bottom: 0.19685, header: 0.19685, footer: 0.19685 };
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

const labelsOf = (meta) => (meta.slots && meta.slots.length
  ? meta.slots.map((s) => (s.label ? `${s.label}: ` : '') + fmtTime(s.start))
  : (meta.shifts || []).map(fmtTime));
const inRoom = (rows, room) => rows.filter((s) => s.Building === room.building && String(s['Room No.']) === String(room.room));

/** One paper's workbook. meta: { title, dateISO, slots: [{ label, start, end }], batchesLabel, modeLabel } */
// Capacity printed in the sheets is the original one (room.base when the test changed it); the seat layout uses the
// seats really set for the test
const capOf = (room) => Number(room.base ?? room.capacity);

export function buildWorkbook(ExcelJS, result, meta) {
  return buildWorkbookAll(ExcelJS, [{ result, meta }], { dateISO: meta.dateISO, dayTitle: meta.title });
}

/**
 * One workbook for every paper of the day. items: [{ result, meta }] (meta.paper = paper name); day: { dateISO, dayTitle }.
 * With several papers each sheet has a Paper column or one block per paper.
 */
export function buildWorkbookAll(ExcelJS, items, day) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Seating Arrangement Automation';
  const many = items.length > 1;
  const P = items.map(({ result, meta }) => ({
    meta, result, name: meta.paper || meta.title,
    rows: result.assigned.slice().sort((a, b) => naturalCompare(a.Roll_Number, b.Roll_Number)),
    rooms: result.usedRooms, shiftLabels: labelsOf(meta), shiftStarts: (meta.slots && meta.slots.length ? meta.slots.map((x) => x.start) : meta.shifts || []).filter(Boolean),
  }));
  const dateISO = day.dateISO || items[0]?.meta.dateISO;
  const dateVal = (() => { const { y, m, d } = parseDate(dateISO); return new Date(Date.UTC(y, m - 1, d)); })();
  const shiftCols = Array.from({ length: Math.max(2, ...P.flatMap((p) => p.rows.map((x) => (x._shifts || p.shiftLabels).length))) }, (_, i) => `Shift ${i + 1}`);
  const paperCol = many || P.some((p) => p.rows.some((x) => x.Paper)) ? ['Paper'] : [];
  const head = (ws, row, labels, opts = { bold: true, size: 11, color: C.green, bg: C.headFill, border: med }) => labels.forEach((h, i) => { const c = ws.getCell(row, i + 1); c.value = h; style(c, opts); });

  // ---------- Raw_Data
  {
    const ws = wb.addWorksheet('Raw_Data');
    ws.columns = [...paperCol.map(() => ({ width: 24 })), { width: 12 }, { width: 12 }, { width: 18 }, { width: 18 }];
    head(ws, 1, [...paperCol, 'Building', 'Room No.', 'Seating Capacity', 'Students Allotted']);
    let r = 2;
    for (const p of P) for (const room of p.rooms) {
      [...(many ? [p.name] : []), room.building, roomVal(room.room), capOf(room), room.used].forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v; style(c); });
      r++;
    }
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  // ---------- Student_list (the main sheet: every student of every paper)
  {
    const ws = wb.addWorksheet('Student_list');
    const heads = ['Roll_Number', 'Student_Name', 'Gender', 'Course', 'Batch', 'Medium', 'Date', ...paperCol, ...shiftCols, 'Building', 'Room No.', 'Seat No.', 'CB', 'Mode', 'Flag'];
    ws.columns = heads.map((h) => ({ header: h, key: h, width: h === 'Student_Name' ? 30 : h === 'Paper' ? 24 : 12 }));
    ws.getRow(1).eachCell((c) => style(c, { bold: true, size: 11, color: C.green, bg: C.headFill }));
    const all = P.flatMap((p) => p.rows.map((r) => ({ r, p }))).sort((a, b) => naturalCompare(a.r.Roll_Number, b.r.Roll_Number));
    for (const { r, p } of all) {
      ws.addRow({
        ...r, Roll_Number: asNum(r.Roll_Number), Date: dateVal, Paper: r.Paper || p.name,
        ...Object.fromEntries(shiftCols.map((c, i) => [c, (r._shifts || p.shiftLabels)[i] || ''])), 'Room No.': roomVal(r['Room No.']),
      });
    }
    ws.getColumn('Date').numFmt = 'dd-mm-yyyy';
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: heads.length } };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  // ---------- Seating Arrangement (two side-by-side columns, sorted by roll number; one block per paper)
  // Two sheets, as in the 22 Nov Impulse sheet: "Seating Arrangement" leaves out the due fee students,
  // "Seating Arrangement (All)" has everyone (due fee, monitoring and all others).
  const seatingSheet = (sheetName, keep) => {
    const ws = wb.addWorksheet(sheetName);
    ws.columns = [9.66, 10.55, 10.66, 10.11, 10.78, 9.66, 10.55, 10.66, 10.11].map((w) => ({ width: w }));
    // One list for the whole day (operator, 21 Nov 2026): every paper, online and offline together, sorted by roll number
    const seen = new Set();
    const rows = P.flatMap((p) => p.rows).filter(keep).filter((st) => (seen.has(st.Roll_Number) ? false : seen.add(st.Roll_Number)))
      .sort((a, b) => naturalCompare(a.Roll_Number, b.Roll_Number));
    const on = rows.some((st) => st.Mode === 'Online'), off = rows.some((st) => st.Mode !== 'Online');
    const modeLabel = on && off ? 'Online+Offline Test' : on ? 'Online Test' : 'Offline Test';
    const batches = [...new Set(P.flatMap((p) => String(p.meta.batchesLabel || '').split(', ').filter(Boolean)))];
    const starts = [...new Set(P.flatMap((p) => p.shiftStarts))].sort((a, b) => a.localeCompare(b));
    let r = 1;
    titleRow(ws, r, day.dayTitle || P[0]?.meta.title || 'Seating Arrangement', 9, { bold: true, size: 22, color: C.navy, bg: C.titleFill, border: med });
    ws.getRow(r).height = 41.7; r++;
    ws.getRow(r).height = 23.85; titleRow(ws, r++, `Mode: ${modeLabel}`, 9, { bold: true, size: 13, color: C.navy, bg: C.subFill, border: med });
    ws.getRow(r).height = 23.85; titleRow(ws, r++, `Batch: ${batches.join(', ')}`, 9, { bold: true, size: 12, color: C.navy, bg: C.bTotFill, border: med });
    const left = { horizontal: 'left', vertical: 'middle', wrapText: true };
    ws.mergeCells(r, 1, r, 4); ws.mergeCells(r, 6, r, 9);
    ws.getCell(r, 1).value = `Date:- ${fmtDateDMY(dateISO)}`;
    ws.getCell(r, 6).value = starts.map((t, k) => `Shift${k + 1}:- ${fmtTime(t)}`).join(',    ');
    for (let c = 1; c <= 9; c++) style(ws.getCell(r, c), { bold: true, size: 12, color: C.navy, bg: C.bTotFill, border: c === 5 ? null : med, align: c < 5 ? left : center });
    ws.getRow(r).height = 23.85; r++;
    ws.getRow(r).height = 19.8;
    ['Reg. No.', 'Building', 'Room No.', 'Seat No.'].forEach((h, k) => {
      [k + 1, k + 6].forEach((col) => { const c = ws.getCell(r, col); c.value = h; style(c, { bold: true, size: 11, color: C.green, bg: C.headFill, border: med }); });
    });
    r++;
    const half = Math.ceil(rows.length / 2);
    for (let k = 0; k < half; k++) {
      [[rows[k], 1], [rows[k + half], 6]].forEach(([st, col]) => {
        if (!st) return;
        [asNum(st.Roll_Number), st.Building, roomVal(st['Room No.']), st['Seat No.']].forEach((v, jj) => { const c = ws.getCell(r + k, col + jj); c.value = v; style(c); });
      });
    }
    ws.pageSetup = { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: MARGINS, printTitlesRow: '5:5' };
  };
  seatingSheet('Seating Arrangement', (st) => st.Flag !== 'Due fee');
  seatingSheet('Seating Arrangement (All)', () => true);

  // ---------- Attendance: one sheet, a block per paper and room, each on its own page. A paper sat in two shifts has
  // one block per room with a signature column for each shift (the students keep the same seat).
  {
    const most = Math.max(1, ...P.map((p) => {
      const n = {}; for (const sl of p.meta.slots || []) n[sl.paper || ''] = (n[sl.paper || ''] || 0) + 1;
      return Math.max(1, ...Object.values(n));
    }));
    // named as in the 22 Nov Impulse sheet: Attendance_2Shift when a paper is sat in two shifts
    const ws = wb.addWorksheet(most > 1 ? `Attendance_${most}Shift` : 'Attendance');
    ws.columns = [{ width: 9.66 }, { width: 38.3 }, { width: 9.11 }, ...Array.from({ length: most }, (_, i) => ({ width: most > 1 ? (i ? 20.66 : 20.3) : 40.5 }))];
    let r = 1;
    for (const p of P) {
      // Shifts grouped by paper (slots carry the paper name when a test has several papers)
      const groups = [];
      p.shiftStarts.map(fmtTime).forEach((label, si) => {
        const paper = (p.meta.slots || [])[si]?.paper;
        const g = groups.find((x) => x.paper === paper);
        if (g) g.labels.push(label); else groups.push({ paper, labels: [label] });
      });
      for (const g of groups) for (const room of p.rooms) {
        const occ = inRoom(p.rows, room).filter((s) => !g.paper || s.Paper === g.paper).sort((a, b) => a['Seat No.'] - b['Seat No.']);
        if (g.paper && !occ.length) continue;
        const ncol = 3 + g.labels.length;
        titleRow(ws, r++, `${room.building}  |  Room ${room.room}  |  ${fmtDateLong(dateISO)}  |  ${g.labels.join(' & ')}  (${occ.length} students)`, ncol,
          { bold: true, size: 12, color: C.navy, bg: C.titleFill });
        ws.getRow(r - 1).height = 21.75;
        head(ws, r, ['Reg. No.', 'Student Name', 'Seat No.', ...g.labels.map((l) => `Signature  (${l})`)], { bold: true, size: 11, color: C.green, bg: C.headFill });
        ws.getRow(r++).height = 19.8;
        occ.forEach((s, k) => {
          const bg = k % 2 ? C.grey : C.white;
          [asNum(s.Roll_Number), s.Student_Name, s['Seat No.'], ...g.labels.map(() => null)].forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v; style(c, { color: 'FF000000', bg, align: j === 1 ? { horizontal: 'left', vertical: 'middle' } : center }); });
          ws.getRow(r).height = 16.8;
          r++;
        });
        titleRow(ws, r, `Total: ${occ.length} students`, ncol, { bold: true, color: C.brown, bg: C.totFill, align: { horizontal: 'right', vertical: 'middle' } });
        ws.getRow(r).addPageBreak();
        r += 3;
      }
    }
    ws.pageSetup = { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: MARGINS };
  }

  // ---------- SMD: count per room by phase, block per building (per paper)
  {
    const ws = wb.addWorksheet('SMD');
    const phases = [...new Set(P.flatMap((p) => p.rows.map((x) => x.CB)))].sort();
    const ncol = 3 + phases.length;
    ws.columns = [{ width: 12 }, { width: 11 }, ...phases.map(() => ({ width: 11 })), { width: 13 }];
    let r = 1;
    for (const p of P) {
      if (many) titleRow(ws, r++, `${p.name}  |  ${p.shiftLabels.join(' & ')}`, ncol, { bold: true, size: 13, color: C.white, bg: C.navy });
      const grand = Object.fromEntries(phases.map((x) => [x, 0]));
      for (const b of [...new Set(p.rooms.map((x) => x.building))]) {
        titleRow(ws, r++, `🏢  ${b}`, ncol, { bold: true, size: 12, color: C.navy, bg: C.titleFill });
        head(ws, r++, ['Building', 'Room No.', ...phases, 'Grand Total']);
        const bt = Object.fromEntries(phases.map((x) => [x, 0]));
        for (const room of p.rooms.filter((x) => x.building === b)) {
          const occ = inRoom(p.rows, room);
          const counts = phases.map((x) => occ.filter((s) => s.CB === x).length);
          counts.forEach((n, i) => { bt[phases[i]] += n; });
          [b, roomVal(room.room), ...counts.map((n) => n || null), occ.length].forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v; style(c); });
          r++;
        }
        const btTotal = phases.reduce((a, x) => a + bt[x], 0);
        [`${b} Total`, null, ...phases.map((x) => bt[x]), btTotal].forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v; style(c, { bold: true, color: C.navy, bg: C.bTotFill, border: med }); });
        phases.forEach((x) => { grand[x] += bt[x]; });
        r += 2;
      }
      [many ? `${p.name} Total` : 'Grand Total', null, ...phases.map((x) => grand[x]), p.rows.length].forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v; style(c, { bold: true, size: 11, color: C.navy, bg: C.subFill, border: med }); });
      r += 3;
    }
  }

  // ---------- Feedback List: per building, each room, and the batches in it with their student count
  {
    const ws = wb.addWorksheet('Feedback List');
    ws.columns = [{ width: 20 }, { width: 18.7 }];
    let r = 1;
    const put = (a, b, o) => { [a, b].forEach((v, jj) => { const c = ws.getCell(r, jj + 1); c.value = v; style(c, o); }); r++; };
    put('Building Rooms', 'Student Count', { bold: true, size: 11, color: C.navy, bg: 'FFDDEBF7' });
    let zebra = 0;
    const leftIn = { horizontal: 'left', vertical: 'middle', indent: 1 };
    for (const p of P) {
      if (many) put(p.name, p.rows.length, { bold: true, size: 11, color: C.white, bg: C.navy, align: leftIn });
      for (const b of [...new Set(p.rooms.map((x) => x.building))]) {
        const inB = p.rows.filter((st) => st.Building === b);
        put(b, inB.length, { bold: true, size: 11, color: C.navy, bg: 'FFE0F2FE', align: leftIn });
        for (const room of p.rooms.filter((x) => x.building === b)) {
          const occ = inRoom(p.rows, room);
          if (!occ.length) continue;
          put(roomVal(room.room), occ.length, { bold: true, bg: zebra++ % 2 ? C.grey : C.white, align: { horizontal: 'left', vertical: 'middle' } });
          const byBatch = {};
          occ.forEach((st) => { byBatch[st.Batch] = (byBatch[st.Batch] || 0) + 1; });
          for (const bt of Object.keys(byBatch).sort(naturalCompare)) put(bt, byBatch[bt], { bg: zebra++ % 2 ? C.grey : C.white, align: { horizontal: 'left', vertical: 'middle' } });
        }
      }
    }
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  // ---------- Room_Layout: seat grid per room (seats run down each column, like the room list)
  {
    const ws = wb.addWorksheet('Room_Layout');
    let r = 1;
    const allRooms = P.flatMap((p) => p.rooms);
    const colsOf = (x) => x.cols || Math.min(10, Math.max(2, Math.ceil(x.capacity / 12)));
    ws.columns = Array.from({ length: Math.max(4, ...allRooms.map(colsOf)) }, () => ({ width: 15 }));
    for (const p of P) for (const room of p.rooms) {
      const cols = colsOf(room);
      const nrows = Math.ceil(room.capacity / cols);
      const bySeat = {};
      inRoom(p.rows, room).forEach((s) => { bySeat[s['Seat No.']] = s; });
      titleRow(ws, r++, `${many ? `${p.name} | ` : ''}${room.building} | Room No. ${room.room} | Strength ${capOf(room)} | Allotted ${room.used}`, cols,
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

  // ---------- Master_List: building entrance display (notice board), every paper on one sheet
  noticeBoard(wb.addWorksheet('Master_List'), P.map((p) => ({ title: p.meta.title, rows: p.rows, rooms: p.rooms, shiftLabels: p.shiftLabels })), { dateISO, dayTitle: day.dayTitle });

  // ---------- Not_Seated (only when someone could not get a seat)
  const unseated = P.flatMap((p) => p.result.unseated.map((s) => ({ ...s, Paper: p.name })));
  if (unseated.length) {
    const ws = wb.addWorksheet('Not_Seated');
    ws.columns = ['Roll_Number', 'Student_Name', 'Batch', 'Course', ...paperCol, 'Reason'].map((h) => ({ header: h, key: h, width: h === 'Student_Name' || h === 'Paper' ? 30 : 14 }));
    ws.getRow(1).eachCell((c) => style(c, { bold: true, color: C.brown, bg: C.totFill }));
    unseated.forEach((s) => ws.addRow({ ...s, Reason: s.Reason || 'No seat available' }));
  }

  return wb;
}

export { phaseOf };

// Notice board: one sheet per building entrance, rooms with roll-number ranges. With several papers,
// each paper is a section on the same sheet (Operations prints one notice board for the day).
function noticeBoard(ws, papers, meta) {
  ws.columns = [{ width: 12 }, { width: 11 }, { width: 11 }, { width: 12 }, { width: 15 }, { width: 15 }, { width: 30 }];
  const many = papers.length > 1;
  titleRow(ws, 1, `${many ? meta.dayTitle || 'Test' : papers[0].title} — NOTICE BOARD`, 7, { bold: true, size: 16, color: C.navy, bg: C.titleFill, border: med });
  titleRow(ws, 2, `Date: ${fmtDateDMY(meta.dateISO)}${many ? '' : `   Time: ${papers[0].shiftLabels.join(' & ')}`}   —   Students check Room No. & Seat No. here`, 7,
    { bold: true, size: 11, color: C.navy, bg: C.subFill });
  let r = 4;
  for (const p of papers) {
    if (many) { titleRow(ws, r++, `${p.title}  |  ${p.shiftLabels.join(' & ')}`, 7, { bold: true, size: 13, color: C.white, bg: C.navy }); }
    for (const b of [...new Set(p.rooms.map((x) => x.building))]) {
      titleRow(ws, r++, b, 7, { bold: true, size: 12, color: C.navy, bg: C.titleFill });
      ['Building', 'Room No.', 'Strength', 'Allotted', 'Reg. No. From', 'Reg. No. To', 'Batches'].forEach((h, i) => { const c = ws.getCell(r, i + 1); c.value = h; style(c, { bold: true, color: C.green, bg: C.headFill }); });
      r++;
      for (const room of p.rooms.filter((x) => x.building === b)) {
        const occ = p.rows.filter((s) => s.Building === b && String(s['Room No.']) === String(room.room)).sort((x, y) => naturalCompare(x.Roll_Number, y.Roll_Number));
        const batches = [...new Set(occ.map((s) => s.Batch))].sort().join(', ');
        [b, roomVal(room.room), capOf(room), occ.length, asNum(occ[0]?.Roll_Number), asNum(occ[occ.length - 1]?.Roll_Number), batches]
          .forEach((v, j) => { const c = ws.getCell(r, j + 1); c.value = v ?? null; style(c); });
        r++;
      }
      r++;
    }
    r++;
  }
}

/** One workbook with a single notice board for every paper of the day. items: [{ result, meta }] */
export function buildNoticeBoard(ExcelJS, items, dayMeta) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Seating Arrangement Automation';
  const papers = items.filter((x) => x.result.assigned.length).map(({ result, meta }) => ({
    title: meta.title,
    rows: result.assigned,
    rooms: result.usedRooms,
    shiftLabels: (meta.slots || []).map((s) => (s.label ? `${s.label}: ` : '') + fmtTime(s.start)),
  }));
  noticeBoard(wb.addWorksheet('Notice_Board'), papers, dayMeta);
  return wb;
}
