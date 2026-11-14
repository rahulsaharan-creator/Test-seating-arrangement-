// Read uploaded Excel files (admission list, room list, batch-building list) into plain objects.

const BUILDING_ALIASES = {
  'GP': 'GP', 'M1': 'M1', 'M2': 'M2', 'M3': 'M3', 'M3 OLD': 'M3 Old', 'M3OLD': 'M3 Old', 'M4': 'M4',
  'MHS-1': 'MHS-1', 'MHS1': 'MHS-1', 'MHS-2': 'MHS-2', 'MHS2': 'MHS-2', 'MND-2': 'MND-2', 'MND2': 'MND-2', 'PC': 'PC',
};
export const canonBuilding = (b) => {
  const k = String(b ?? '').trim().toUpperCase().replace(/\s+/g, ' ');
  return BUILDING_ALIASES[k] || String(b ?? '').trim();
};

const cellText = (v) => {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if ('result' in v) return cellText(v.result);
    if (v.formula || v.sharedFormula || v.error) return ''; // a formula Excel never calculated, or an error cell
    if (v.text) return String(v.text);
    if (v instanceof Date) return v.toISOString().slice(0, 10);
  }
  return String(v).trim();
};

function sheetRows(ws) {
  const out = [];
  ws.eachRow({ includeEmpty: false }, (row) => {
    const vals = [];
    row.eachCell({ includeEmpty: true }, (c, col) => { vals[col - 1] = cellText(c.value); });
    out.push(vals);
  });
  return out;
}

const norm = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');

export async function readAdmissions(ExcelJS, buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  for (const ws of wb.worksheets) {
    const rows = sheetRows(ws);
    // common header variants are accepted (Roll_No., Reg No, Name, ...)
    const ALIAS = { rollnumber: ['rollnumber', 'rollno', 'roll', 'regno', 'registrationno', 'registrationnumber'], studentname: ['studentname', 'name', 'studentsname'], fathername: ['fathername', 'fathersname', 'father', 'fatherfullname'], contactnumber: ['contactnumber', 'contact', 'contactno', 'mobile', 'mobileno', 'mobilenumber', 'phone', 'phoneno', 'phonenumber'], dob: ['dob', 'dateofbirth', 'birthdate', 'birthday'] };
    const hi = rows.findIndex((r) => r.some((v) => ALIAS.rollnumber.includes(norm(v || ''))) && r.some((v) => norm(v || '') === 'batch'));
    if (hi < 0) continue;
    const head = rows[hi].map((h) => norm(h || ''));
    // Exact header first; then a looser match for headers like "Name of Student", "Father's Full Name", "Mobile No. (Parent)"
    const LOOSE = { studentname: (h) => h.includes('name') && !/father|mother|parent|guardian|course|batch|school|centre|center/.test(h), fathername: (h) => h.includes('father'),
      contactnumber: (h) => /mobile|contact|phone/.test(h), dob: (h) => h.includes('birth') };
    const col = (name) => { const k = norm(name); for (const a of ALIAS[k] || [k]) { const i = head.indexOf(a); if (i >= 0) return i; } return LOOSE[k] ? head.findIndex(LOOSE[k]) : -1; };
    const ix = {
      roll: col('Roll_Number'), name: col('Student_Name'), gender: col('Gender'), course: col('Course'),
      batch: col('Batch'), medium: col('Medium'), status: col('Status'),
      father: col('Father_Name'), contact: col('Contact_Number'), dob: col('DOB'),
    };
    const students = [];
    const seen = new Set();
    let duplicates = 0, inactive = 0;
    for (const r of rows.slice(hi + 1)) {
      const roll = r[ix.roll];
      const batch = r[ix.batch];
      if (!roll || !batch) continue;
      if (ix.status >= 0 && r[ix.status] && r[ix.status].toLowerCase() !== 'active') { inactive++; continue; }
      if (seen.has(roll)) { duplicates++; continue; }
      seen.add(roll);
      students.push({
        Roll_Number: roll, Student_Name: r[ix.name] || '', Gender: r[ix.gender] || '', Course: r[ix.course] || '',
        Batch: batch.trim(), Medium: r[ix.medium] || '',
        Father_Name: ix.father >= 0 ? r[ix.father] || '' : '', Contact_Number: ix.contact >= 0 ? r[ix.contact] || '' : '', DOB: ix.dob >= 0 ? r[ix.dob] || '' : '',
      });
    }
    return { students, duplicates, inactive, sheet: ws.name };
  }
  throw new Error('Could not find a sheet with "Roll_Number" and "Batch" columns.');
}

// Room list file: sheet "Room_Capacity" (Building, Room No., Seating Capacity) and optional "LAB_Capacity".
export async function readRooms(ExcelJS, buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const rooms = [], labs = [];
  const rc = wb.getWorksheet('Room_Capacity') || wb.worksheets[0];
  for (const r of sheetRows(rc).slice(1)) {
    if (!r[0] || !r[2] || isNaN(Number(r[2]))) continue;
    const building = canonBuilding(r[0]);
    const room = r[1];
    const note = (r[3] || '').trim();
    const entry = { building, room, capacity: Number(r[2]), enabled: !/^self/i.test(note), note, cols: 0 };
    if (building === 'PC' && room.toUpperCase() === 'PC-LAB') labs.push(entry);
    else rooms.push(entry);
  }
  const lc = wb.getWorksheet('LAB_Capacity');
  if (lc) {
    for (const r of sheetRows(lc).slice(1)) {
      if (!r[0] || isNaN(Number(r[2]))) continue;
      labs.push({ building: canonBuilding(r[0]), room: r[1], capacity: Number(r[2]), enabled: true, note: r[3] ? `Total systems ${r[3]}` : '', cols: 0 });
    }
  }
  // Seat-layout columns from "Room_List" if present
  const rl = wb.getWorksheet('Room_List');
  if (rl) {
    const rows = sheetRows(rl);
    rows.forEach((r, i) => {
      if (!r[0] || !r[0].includes('|')) return;
      const p = r[0].split('|').map((x) => x.trim());
      const b = canonBuilding(p[0]);
      const rn = p[1].replace(/Room No\.?\s*/i, '').trim();
      let cols = 0;
      for (let j = i + 2; j < rows.length && rows[j][0] && /^\d+$/.test(rows[j][0]); j++) cols = Math.max(cols, rows[j].filter((v) => v).length);
      const room = rooms.find((x) => x.building === b && String(x.room) === rn);
      if (room) room.cols = cols;
    });
  }
  if (!rooms.length) throw new Error('No rooms found. Expected columns: Building, Room No., Seating Capacity.');
  return { rooms, labs };
}

// Batch-building file: Building | Floor | Room | Morning | Evening (building cell may be merged)
export async function readBatchMap(ExcelJS, buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  const map = {};
  let building = null;
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const v = (i) => cellText(row.getCell(i).value);
    if (v(1)) building = canonBuilding(v(1));
    [[4, 'Morning'], [5, 'Evening']].forEach(([i, shift]) => {
      if (!v(i) || !building) return;
      v(i).split('/').forEach((b) => { map[b.trim()] = { building, room: v(3), shift }; });
    });
  });
  if (!Object.keys(map).length) throw new Error('No batches found. Expected columns: Building, Floor, Room, Morning, Evening.');
  return map;
}

// Roll-number list (monitoring, due fee): any sheet; the column headed Roll/Reg, else the first column.
export async function readRollList(ExcelJS, buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const out = new Set();
  for (const ws of wb.worksheets) {
    const rows = sheetRows(ws);
    const hi = rows.findIndex((r) => r.some((v) => /roll|reg/i.test(v || '')));
    const col = hi >= 0 ? rows[hi].findIndex((v) => /roll|reg/i.test(v || '')) : 0;
    for (const r of rows.slice(hi + 1)) { const v = String(r[col] || '').trim(); if (/^[A-Za-z0-9-]{3,}$/.test(v) && !/roll|reg/i.test(v)) out.add(v); }
    if (out.size) break;
  }
  if (!out.size) throw new Error('No roll numbers found. Put them in a column headed Roll_Number.');
  return [...out];
}

// Roll numbers pasted as text: commas, spaces or new lines between them.
export const parseRolls = (text) => [...new Set(String(text || '').split(/[\s,;]+/).map((v) => v.trim()).filter(Boolean))];
