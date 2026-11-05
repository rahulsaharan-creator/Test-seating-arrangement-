// End-to-end check: real admission list -> allocation -> workbook. Run: npm test
import fs from 'fs';
import ExcelJS from 'exceljs';
import defaults from '../src/data/defaults.json' with { type: 'json' };
import { allocate } from '../src/lib/engine.js';
import { buildWorkbook } from '../src/lib/exporter.js';
import { readAdmissions, readRooms, readBatchMap } from '../src/lib/readers.js';

const P = process.env.DATA || '/mnt/project';
const assert = (c, m) => { if (!c) { console.error('FAIL:', m); process.exit(1); } console.log('ok -', m); };

const adm = await readAdmissions(ExcelJS, fs.readFileSync(`${P}/admissions_list_29-Oct-2026.xlsx`));
assert(adm.students.length === 14485, `read ${adm.students.length} students`);
const rf = await readRooms(ExcelJS, fs.readFileSync(`${P}/Room_List_With_Seat_No__2.xlsx`));
assert(rf.rooms.length === 101 && rf.labs.length === 12, `room file: ${rf.rooms.length} rooms, ${rf.labs.length} labs`);
const bm = await readBatchMap(ExcelJS, fs.readFileSync(`${P}/Batch_and_Building.xlsx`));
assert(bm.XB3?.building === 'PC' && bm.ICB1?.building === 'M2', 'batch file parsed (XB3->PC, ICB1->M2)');

const batchBuilding = Object.fromEntries(Object.entries(defaults.batchMap).map(([k, v]) => [k, v.building]));
const base = { rooms: defaults.rooms, labs: defaults.labs, batchBuilding, buildingOrder: defaults.buildingOrder, farBuildings: defaults.farBuildings, seed: 42 };

// Scenario A: the 4-Nov example (PC batches, XC1 online)
const sel = ['XB1','XB2','XB3','XC1','XC2','XC3','YB1','YB2','ZB1'];
const students = adm.students.filter((s) => sel.includes(s.Batch));
const batchMode = Object.fromEntries(sel.map((b) => [b, b === 'XC1' ? 'Online' : 'Offline']));
const res = allocate(students, { ...base, batchMode });
const n = students.length;
assert(res.assigned.length + res.unseated.length === n, `everyone accounted for (${n})`);
assert(new Set(res.assigned.map((s) => s.Roll_Number)).size === res.assigned.length, 'no student seated twice');
const key = (s) => `${s.Building}|${s['Room No.']}|${s['Seat No.']}`;
assert(new Set(res.assigned.map(key)).size === res.assigned.length, 'no seat used twice');
for (const r of res.usedRooms) if (r.used > r.capacity) assert(false, `room over capacity ${r.building} ${r.room}`);
assert(true, 'no room over capacity');
console.log('  moved:', JSON.stringify(res.moved), '| warnings:', res.warnings);
// fill order: within a building, a later room is used only if every earlier enabled room is full
const offRooms = res.usedRooms.filter((r) => !defaults.labs.some((l) => l.building === r.building && l.room === r.room));
const pcAll = defaults.rooms.filter((r) => r.building === 'PC' && r.enabled);
const pcUsed = pcAll.map((r) => offRooms.find((u) => u.room === r.room)?.used || 0);
const firstNotFull = pcUsed.findIndex((u, i) => u < pcAll[i].capacity);
assert(firstNotFull === -1 || pcUsed.slice(firstNotFull + 1).every((u) => u === 0), 'rooms fill one after another');
const mixed = offRooms.filter((r) => new Set(r.occupants.map((s) => s.Batch)).size > 1).length;
assert(mixed > 0, `batches mixed within rooms (${mixed} rooms mixed)`);

const wb = buildWorkbook(ExcelJS, res, { title: 'Momentum Seating Arrangement', dateISO: '2026-11-04', shifts: ['07:00', '10:30'], batchesLabel: 'XB, XC, YB, ZB', modeLabel: 'Offline + Online Test' });
fs.mkdirSync('tests/out', { recursive: true });
await wb.xlsx.writeFile('tests/out/scenarioA.xlsx');
assert(true, 'workbook written: ' + wb.worksheets.map((w) => w.name).join(', '));

// Scenario B: big overflow test — all Momentum batches offline in one shift
const mom = adm.students.filter((s) => s.Course === 'Momentum');
const modeB = Object.fromEntries([...new Set(mom.map((s) => s.Batch))].map((b) => [b, 'Offline']));
const rB = allocate(mom, { ...base, batchMode: modeB });
assert(rB.assigned.length + rB.unseated.length === mom.length, `B: ${mom.length} Momentum students accounted for, ${rB.unseated.length} unseated`);
const pcGotOthers = rB.assigned.some((s) => s.Building === 'PC' && batchBuilding[s.Batch] !== 'PC');
const nearFree = rB.usedRooms.length;
console.log('  B moved (first 8):', JSON.stringify(rB.moved.slice(0, 8)), '| PC used for non-PC students:', pcGotOthers);
console.log('  B warnings:', rB.warnings);
