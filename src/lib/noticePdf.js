// Notice board PDF: the "Seating Arrangement" sheet of the seating Excel file, as a PDF for the notice board.
// One PDF per building (or one building), optionally one floor. With every floor, each floor starts on a new page.
import { naturalCompare } from './engine.js';
import { fmtDateDMY, fmtTime } from './exporter.js';

// Floor of a room: the Floor column of the room list when it has one, else the first digit of a 3-digit room number
// (101 -> 1, 001 -> 0). '' when it is not known.
export function floorOf(room) {
  if (room && room.floor != null && String(room.floor).trim() !== '') return String(room.floor).trim();
  const m = String(room && room.room != null ? room.room : '').trim().match(/^(\d)\d\d(?!\d)/);
  return m ? m[1] : '';
}
export const floorText = (f) => (f === '' ? 'No floor set' : f === '0' ? 'Ground floor' : `Floor ${f}`);
const byFloor = (a, b) => naturalCompare(a === '' ? '~' : a, b === '' ? '~' : b);

const labelsOf = (meta) => (meta.slots && meta.slots.length
  ? meta.slots.map((s) => (s.label ? `${s.label}: ` : '') + fmtTime(s.start) + (s.end ? ` to ${fmtTime(s.end)}` : ''))
  : (meta.shifts || []).map(fmtTime));

// items: [{ result, meta }] as for the Excel file. Every seated student gets building and floor for the filters.
export function noticeData(items) {
  return items.filter((x) => x.result.assigned.length).map(({ result, meta }) => {
    const fl = new Map(result.usedRooms.map((r) => [`${r.building}|${r.room}`, floorOf(r)]));
    return {
      meta, shiftLabels: labelsOf(meta),
      rows: result.assigned.map((s) => ({ ...s, _floor: fl.get(`${s.Building}|${s['Room No.']}`) ?? floorOf({ room: s['Room No.'] }) }))
        .sort((a, b) => naturalCompare(a.Roll_Number, b.Roll_Number)),
    };
  });
}
export const noticeBuildings = (data) => [...new Set(data.flatMap((p) => p.rows.map((s) => s.Building)))].sort(naturalCompare);
export const noticeFloors = (data, building) => [...new Set(data.flatMap((p) => p.rows.filter((s) => s.Building === building).map((s) => s._floor)))].sort(byFloor);

const NAVY = [30, 58, 95], GREEN = [20, 83, 45], TITLE_BG = [219, 234, 254], SUB_BG = [224, 231, 255], HEAD_BG = [220, 252, 231], LINE = [156, 163, 175];

// One PDF for one building: floors = the floors to print (each starts on a new page)
export function noticePdf(jsPDF, autoTable, data, { dateISO, dayTitle, building, floors }) {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4', compress: true });
  const W = doc.internal.pageSize.getWidth(), M = 10;
  let first = true;
  const band = (y, text, { size, bg, bold = true, h }) => {
    doc.setFillColor(...bg); doc.setDrawColor(...NAVY); doc.setLineWidth(0.5); doc.rect(M, y, W - 2 * M, h, 'FD');
    doc.setFont('helvetica', bold ? 'bold' : 'normal'); doc.setFontSize(size); doc.setTextColor(...NAVY);
    doc.text(text, W / 2, y + h / 2, { align: 'center', baseline: 'middle', maxWidth: W - 2 * M - 4 });
    return y + h;
  };
  for (const f of floors) {
    const blocks = data.map((p) => ({ ...p, rows: p.rows.filter((s) => s.Building === building && s._floor === f) })).filter((p) => p.rows.length);
    if (!blocks.length) continue;
    if (!first) doc.addPage();
    first = false;
    let y = M;
    y = band(y, `${dayTitle || 'Test Seating Arrangement'} · ${building}${floors.length > 1 || f !== '' ? ` · ${floorText(f)}` : ''}`, { size: 15, bg: TITLE_BG, h: 11 }) + 4;
    blocks.forEach((p, bi) => {
      if (bi > 0) { doc.addPage(); y = M; }
      y = band(y, p.meta.title, { size: 13, bg: TITLE_BG, h: 9 });
      y = band(y, `Mode: ${p.meta.modeLabel}`, { size: 10, bg: SUB_BG, h: 7 });
      y = band(y, `Batch: ${p.meta.batchesLabel}`, { size: 10, bg: SUB_BG, h: 7 });
      y = band(y, `Date:- ${fmtDateDMY(dateISO)}          Time:- ${p.shiftLabels.join(' & ')}`, { size: 10, bg: SUB_BG, h: 7 });
      // two side-by-side columns, sorted by roll number: first half on the left, second half on the right
      const half = Math.ceil(p.rows.length / 2);
      const cell = (s) => (s ? [String(s.Roll_Number), s.Building, String(s['Room No.']), String(s['Seat No.'])] : ['', '', '', '']);
      const body = Array.from({ length: half }, (_, i) => [...cell(p.rows[i]), '', ...cell(p.rows[i + half])]);
      const head = ['Reg. No.', 'Building', 'Room No.', 'Seat No.'];
      autoTable(doc, {
        startY: y, margin: { left: M, right: M }, head: [[...head, '', ...head]], body, theme: 'grid',
        styles: { font: 'helvetica', fontSize: 9, halign: 'center', valign: 'middle', textColor: [51, 51, 51], lineColor: LINE, lineWidth: 0.2, cellPadding: 1.2 },
        headStyles: { fillColor: HEAD_BG, textColor: GREEN, fontStyle: 'bold', fontSize: 10, lineColor: NAVY, lineWidth: 0.4 },
        columnStyles: { 4: { cellWidth: 4, fillColor: [255, 255, 255], lineWidth: 0 } },
        didParseCell: (d) => { if (d.column.index === 4) { d.cell.styles.lineWidth = 0; d.cell.styles.fillColor = [255, 255, 255]; } },
      });
      y = doc.lastAutoTable.finalY + 6;
    });
  }
  const n = doc.getNumberOfPages();
  for (let i = 1; i <= n; i++) {
    doc.setPage(i); doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(120);
    doc.text(`${building} · ${fmtDateDMY(dateISO)} · page ${i} of ${n}`, W / 2, doc.internal.pageSize.getHeight() - 5, { align: 'center' });
  }
  return first ? null : doc.output('blob');
}
