import React, { useEffect, useMemo, useRef, useState } from 'react';
import defaults from './data/defaults.json';
import { allocate, capacity, naturalCompare, isGirl, roomKey, FLAGS, SEAT_ORDERS } from './lib/engine.js';
import { buildWorkbookAll, buildNoticeBoard, fmtDateDMY, fmtTime } from './lib/exporter.js';
import { readAdmissions, readRooms, readRollList, parseRolls } from './lib/readers.js';
import { parsePhaseList, groupPhases, seatSummary, fillPreview } from './lib/plan.js';

const loadExcel = () => import('exceljs').then((m) => m.default || m);
const STORE = 'seating-settings-v1';
const LAST = 'seating-last-test-v5';
const GAP = 15; // minutes: papers closer than this count as overlapping (students leaving and arriving)

const allBuildingsOf = (rooms, labs, order = []) => [...new Set([...order, ...rooms.map((r) => r.building), ...labs.map((r) => r.building)])];
const freshSettings = () => {
  const rooms = JSON.parse(JSON.stringify(defaults.rooms));
  const labs = JSON.parse(JSON.stringify(defaults.labs)).map((l) => ({ ...l, girls: true }));
  const order = allBuildingsOf(rooms, labs, defaults.buildingOrder);
  return { rooms, labs, buildingOrder: order, buildingsOn: order.slice(), girlsOk: {}, cbGroups: [], pins: {} };
};
function loadSettings() {
  try {
    const s = JSON.parse(localStorage.getItem(STORE));
    if (s && s.rooms) {
      const f = freshSettings();
      const labs = (s.labs || f.labs).map((l) => ({ girls: true, ...l }));
      const order = allBuildingsOf(s.rooms, labs, s.buildingOrder || f.buildingOrder);
      return { ...f, ...s, labs, buildingOrder: order, buildingsOn: s.buildingsOn || order.slice() };
    }
  } catch { /* storage unavailable or empty */ }
  return freshSettings();
}
function loadLast() {
  try { return JSON.parse(localStorage.getItem(LAST)) || null; } catch { return null; }
}
const uid = () => Math.random().toString(36).slice(2, 8);
// A test: its courses and one or more papers (stored as slots).
// A paper (slot): name, 1 or 2 shifts and their times, mode (Online / Offline / Both), phases, batches
// (online ones listed in online when the mode is Both), labs, buildings and removed rooms.
const newSlot = (start = '09:00', end = '12:00') => {
  const s2 = hhmm(mins(end) + 120);
  return { id: uid(), name: '', mode: '', phases: [], start, end, two: false, s2: { start: s2, end: hhmm(mins(s2) + (mins(end) - mins(start))) }, batches: [], online: [], labs: null, buildings: null, roomsOff: [] };
};
// A time slot in 1 shift, or in 2 shifts with the same seating (a student keeps one seat; an attendance sheet per shift)
const slotLabel = (s) => (s.two ? `${range(s)} and ${range(s.s2)}` : range(s));
const slotOk = (s) => shiftOk(s) && (!s.two || (shiftOk(s.s2) && mins(s.s2.start) >= mins(s.end)));
const slotName = (p, s) => (s.name || '').trim() || `Paper ${p.slots.indexOf(s) + 1}`;
const newPaper = () => ({ id: uid(), name: '', courses: [], slots: [newSlot()], done: false });
function initialTest() {
  const last = loadLast();
  if (last && Array.isArray(last.papers) && last.papers.length) return { title: last.title || '', date: '', papers: last.papers.map((p) => ({ ...newPaper(), ...p, slots: (p.slots || [newSlot()]).map((s) => ({ ...newSlot(), ...s })), done: false })) };
  return { title: '', date: '', papers: [newPaper()] };
}

const pl = (n, w) => `${n.toLocaleString('en-IN')} ${w}${n === 1 ? '' : /(ch|sh|s|x)$/.test(w) ? 'es' : 's'}`;
const mins = (t) => { const [h, m] = String(t || '0:0').split(':').map(Number); return h * 60 + m; };
const hhmm = (m) => `${String(Math.floor(Math.min(m, 1439) / 60)).padStart(2, '0')}:${String(Math.min(m, 1439) % 60).padStart(2, '0')}`;
// A paper's shifts. Two shifts: the same students and the same seats, an attendance sheet for each shift.
const shiftsOf = (p) => (p.shifts ? p.shifts : p.two ? [{ label: 'Shift 1', ...p.s1 }, { label: 'Shift 2', ...p.s2 }] : [{ label: '', ...p.s1 }]);
// Online time of a paper: online students may sit at a different time from offline students (default: same as Shift 1)
const onlineOf = (p) => (p.mode === 'Both' && p.online && p.online.start ? p.online : p.s1 || shiftsOf(p)[0]);
const windowOf = (p) => {
  const sh = shiftsOf(p), on = onlineOf(p);
  return { start: hhmm(Math.min(...sh.map((x) => mins(x.start)), mins(on.start))), end: hhmm(Math.max(...sh.map((x) => mins(x.end)), mins(on.end))) };
};
export const overlaps = (a, b) => {
  if (a.id === b.id) return true;
  const x = windowOf(a), y = windowOf(b);
  return mins(x.start) < mins(y.end) + GAP && mins(y.start) < mins(x.end) + GAP;
};
const paperName = (p, i) => p.name || `Paper ${i + 1}`;
const range = (s) => `${fmtTime(s.start)}${s.end ? ` to ${fmtTime(s.end)}` : ''}`;
const shiftText = (p) => shiftsOf(p).map((s) => (s.label ? `${s.label} ${range(s)}` : range(s))).join(' · ');
const paperTime = (p) => (p.mode === 'Both' ? `Offline ${shiftText(p)} · Online ${range(onlineOf(p))}` : shiftText(p));
const shiftOk = (s) => s.start && s.end && mins(s.end) > mins(s.start);

const Ico = {
  upload: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M12 16V4M7 9l5-5 5 5M4 20h16" /></svg>,
  down: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M12 4v12M7 11l5 5 5-5M4 20h16" /></svg>,
  bin: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14" /></svg>,
  plus: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" aria-hidden="true"><path d="M12 5v14M5 12h14" /></svg>,
  spark: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m12 3 1.9 5.8L20 10l-4.9 3.6L16.5 20 12 16.5 7.5 20l1.4-6.4L4 10l6.1-1.2z" /></svg>,
  file: <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z" /><path d="M14 3v6h6M8 13h8M8 17h5" /></svg>,
};

// Blank Excel templates, made in the browser
const TEMPLATES = {
  admission: { file: 'Admission_list_template.xlsx', sheets: [['Admission list', ['Roll_Number', 'Student_Name', 'Gender', 'Course', 'Batch', 'Medium', 'Status', 'Father_Name', 'Contact_Number', 'DOB']]] },
  rooms: { file: 'Room_list_template.xlsx', sheets: [['Room_Capacity', ['Building', 'Room No.', 'Seating Capacity', 'Note']], ['LAB_Capacity', ['Building', 'Lab', 'Capacity', 'Total systems']]] },
  rolls: { file: 'Roll_number_list_template.xlsx', sheets: [['Students', ['Roll_Number']]] },
};
async function downloadTemplate(kind) {
  const ExcelJS = await loadExcel();
  const wb = new ExcelJS.Workbook();
  for (const [name, heads] of TEMPLATES[kind].sheets) {
    const ws = wb.addWorksheet(name);
    ws.columns = heads.map((h) => ({ header: h, key: h, width: Math.max(14, h.length + 4) }));
    ws.getRow(1).font = { bold: true };
  }
  saveBlob(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), TEMPLATES[kind].file);
}
function saveBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

const STEPS = [
  { id: 'students', label: 'Admission list' },
  { id: 'test', label: 'Test & papers' },
  { id: 'labs', label: 'Online seating' },
  { id: 'rooms', label: 'Offline seating' },
  { id: 'result', label: 'Seating' },
];

export default function App() {
  const [step, setStep] = useState('students');
  const [settings, setSettings] = useState(loadSettings);
  const [adm, setAdm] = useState(null);
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [test, setTest] = useState(initialTest);
  const [cur, setCur] = useState('');
  const [studentMode, setStudentMode] = useState({});
  const [top, setTop] = useState(() => new Set(loadLast()?.top || []));
  const [lists, setLists] = useState({ Monitoring: [], 'Due fee': [] });
  const [seed, setSeed] = useState(() => Math.floor(Math.random() * 1e9));
  const [results, setResults] = useState(null); // [{ paper (one time slot), index, result, batches, students, overlapsWith }]
  const [view, setView] = useState(0);

  useEffect(() => {
    try { localStorage.setItem(STORE, JSON.stringify(settings)); } catch { /* ignore */ }
  }, [settings]);
  useEffect(() => { window.scrollTo(0, 0); }, [step]);

  const batches = useMemo(() => {
    if (!adm) return [];
    const m = {};
    for (const s of adm.students) {
      m[s.Batch] ||= { batch: s.Batch, course: s.Course, count: 0, girls: 0 };
      m[s.Batch].count++;
      if (isGirl(s)) m[s.Batch].girls++;
    }
    return Object.values(m).sort((a, b) => naturalCompare(a.batch, b.batch));
  }, [adm]);
  const batchMap = useMemo(() => Object.fromEntries(batches.map((b) => [b.batch, b])), [batches]);
  const studentsByBatch = useMemo(() => {
    const m = {};
    for (const s of adm?.students || []) (m[s.Batch] ||= []).push(s);
    return m;
  }, [adm]);
  const courses = useMemo(() => [...new Set(batches.map((b) => b.course))].sort(naturalCompare), [batches]);
  const phaseList = useMemo(() => parsePhaseList(settings.phaseText || ''), [settings.phaseText]);
  const phases = useMemo(() => groupPhases(batches, phaseList.phases), [batches, phaseList]);

  const papers = test.papers;
  const curP = papers.find((p) => p.id === cur) || null;
  const curIdx = papers.findIndex((p) => p.id === cur);
  // Any change to a paper means its seating must be prepared again
  const setPaper = (id, patch) => setTest((t) => ({ ...t, papers: t.papers.map((p) => (p.id === id ? { ...p, ...patch, done: false } : p)) }));
  const setSlot = (p, sid, patch) => setPaper(p.id, { slots: p.slots.map((s) => (sid === '*' || s.id === sid ? { ...s, ...patch } : s)) });

  // Phase owner (a phase is in one paper only) and batch owner (a batch is in one time slot only)
  const ownerOf = {}, slotOf = {};
  for (const p of papers) for (const s of p.slots) {
    for (const k of s.phases) ownerOf[k] ||= s.id;
    for (const b of s.batches) slotOf[b] ||= { paper: p.id, slot: s.id };
  }
  const phaseBatches = (s) => phases.filter((g) => s.phases.includes(g.key)).flatMap((g) => g.batches);
  const modeIn = (p, s, b) => (s.mode === 'Online' ? 'Online' : s.mode === 'Offline' ? 'Offline' : s.online.includes(b) ? 'Online' : 'Offline');
  const labsOn = settings.labs.filter((l) => l.enabled && Number(l.capacity) > 0);
  const roomsOn = settings.rooms.filter((r) => r.enabled && Number(r.capacity) > 0);
  // Until the operator picks labs and buildings for a time slot, every lab and every building in use counts for the estimate
  // Rooms and labs kept for monitoring or due fee students are always part of the seating (and of every seat count),
  // even when their building is not ticked, so the counts match what the seating really uses.
  const pinKeys = new Set(['Monitoring', 'Due fee'].filter((f) => lists[f].length).flatMap((f) => settings.pins[f] || []));
  const slotLabs = (s) => (s.labs ? labsOn.filter((l) => s.labs.includes(roomKey(l)) || pinKeys.has(roomKey(l))) : labsOn);
  const slotBuildings = (s) => s.buildings || settings.buildingsOn;
  const slotRooms = (s) => roomsOn.filter((r) => (slotBuildings(s).includes(r.building) && !s.roomsOff.includes(roomKey(r))) || pinKeys.has(roomKey(r)));
  const sumSeats = (list) => list.reduce((a, r) => a + Number(r.capacity || 0), 0);
  const countSlot = (p, s) => {
    const c = { online: 0, offline: 0 };
    for (const b of s.batches) for (const st of studentsByBatch[b] || []) c[((s.mode === 'Both' && studentMode[st.Roll_Number]) || modeIn(p, s, b)) === 'Online' ? 'online' : 'offline']++;
    return c;
  };
  const sumSlot = (p, s) => seatSummary({ ...countSlot(p, s), labSeats: sumSeats(slotLabs(s)), roomSeats: sumSeats(slotRooms(s)) });
  const addUp = (list) => list.reduce((a, x) => { for (const k of Object.keys(x)) a[k] = (a[k] || 0) + x[k]; return a; }, {});
  const paperOk = (p) => !!(p.courses.length && p.slots.every((s) => slotOk(s) && s.mode && s.phases.length && s.batches.length));

  // Special students: Monitoring and Due fee from the lists on step 1, Top batch from the Top ticks
  const flags = useMemo(() => {
    const f = {};
    for (const b of top) for (const s of studentsByBatch[b] || []) f[s.Roll_Number] = 'Top batch';
    for (const r of lists['Due fee']) f[r] = 'Due fee';
    for (const r of lists.Monitoring) f[r] = 'Monitoring';
    return f;
  }, [top, lists, studentsByBatch]);
  const appearingRolls = new Set(papers.flatMap((p) => p.slots.flatMap((s) => s.batches.flatMap((b) => (studentsByBatch[b] || []).map((x) => x.Roll_Number)))));
  const appearing = appearingRolls.size;
  const flagCount = (f) => [...appearingRolls].filter((r) => flags[r] === f).length;
  const curChoices = [...new Map([...labsOn.map((r) => ({ ...r, kind: 'Lab' })), ...roomsOn.map((r) => ({ ...r, kind: 'Room' }))].map((r) => [roomKey(r), r])).values()];
  const curKeys = new Set(curChoices.map(roomKey));
  const special = ['Monitoring', 'Due fee'].map((f) => ({ flag: f, n: flagCount(f), kept: (settings.pins[f] || []).some((k) => curKeys.has(k)) })).filter((x) => x.n);
  // One unit per test: all its papers are seated together, in one seating, on the same labs and rooms.
  // Tests are grouped by time: tests that overlap share the labs and rooms.
  const allSlots = papers.map((p, i) => {
    const used = p.slots.filter((x) => x.batches.length);
    if (!used.length) return null;
    const w = { start: hhmm(Math.min(...used.map((x) => mins(x.start)))), end: hhmm(Math.max(...used.map((x) => mins(x.two ? x.s2.end : x.end)))) };
    const labs = new Map(), rooms = new Map();
    for (const x of used) { for (const l of slotLabs(x)) labs.set(roomKey(l), l); for (const r of slotRooms(x)) rooms.set(roomKey(r), r); }
    return {
      p, s: used[0], used, w, name: (p.name || '').trim() || `Test ${i + 1}`,
      c: addUp(used.map((x) => countSlot(p, x)).concat([{ online: 0, offline: 0 }])),
      labs: [...labs.values()], rooms: [...rooms.values()],
      labsSet: used.some((x) => x.labs && x.labs.length), roomsSet: used.some((x) => x.buildings && x.buildings.length),
    };
  }).filter(Boolean).sort((a, b) => mins(a.w.start) - mins(b.w.start));
  const timeGroups = [];
  for (const x of allSlots) {
    const g = timeGroups[timeGroups.length - 1];
    if (g && mins(x.w.start) < mins(g.end) + GAP) { g.items.push(x); if (mins(x.w.end) > mins(g.end)) g.end = x.w.end; } else timeGroups.push({ start: x.w.start, end: x.w.end, items: [x] });
  }
  for (const g of timeGroups) {
    const c = addUp(g.items.map((x) => x.c).concat([{ online: 0, offline: 0 }]));
    const labs = new Map(), rooms = new Map();
    for (const x of g.items) { for (const l of x.labs) labs.set(roomKey(l), l); for (const r of x.rooms) rooms.set(roomKey(r), r); }
    Object.assign(g, seatSummary({ ...c, labSeats: sumSeats([...labs.values()]), roomSeats: sumSeats([...rooms.values()]) }), { papers: g.items.flatMap((x) => x.used.map((y) => slotName(x.p, y))) });
  }
  // Seat plan in seating order (time order): a test that overlaps an earlier test sees the lab and classroom seats the
  // earlier one fills. Papers of the same test never block each other's seats (they are seated together).
  // plan[slotId] (the same for every paper of a test) = { taken: { labs, rooms }, labs, rooms, offNeed, c, roomList }
  const near = (a, b) => mins(a.start) < mins(b.end) + GAP && mins(b.start) < mins(a.end) + GAP;
  const plan = {};
  for (const x of allSlots) {
    const taken = { labs: {}, rooms: {} };
    for (const y of allSlots) {
      if (y === x) break;
      if (!near(x.w, y.w)) continue;
      for (const kind of ['labs', 'rooms']) for (const r of plan[y.s.id][kind]) {
        if (!r.filled) continue;
        const k = roomKey(r), o = taken[kind][k];
        taken[kind][k] = { filled: (o?.filled || 0) + r.filled, by: [...new Set([...(o?.by || []), y.name])] };
      }
    }
    const freeOf = (list, kind) => list.map((r) => ({ ...r, capacity: Math.max(0, Number(r.capacity) - (taken[kind][roomKey(r)]?.filled || 0)) }));
    const c = x.c;
    const lab = x.labsSet ? fillPreview(c.online, freeOf(x.labs, 'labs'), settings.buildingOrder) : { rooms: [], left: Math.max(0, c.online - sumSeats(x.labs)) };
    const offNeed = c.offline + lab.left;
    const room = x.roomsSet ? fillPreview(offNeed, freeOf(x.rooms, 'rooms'), settings.buildingOrder) : { rooms: [] };
    const one = { taken, labs: lab.rooms, rooms: room.rooms, offNeed, c, roomList: x.rooms };
    for (const y of x.used) plan[y.id] = one;
  }
  const planOf = (s) => plan[s.id] || { taken: { labs: {}, rooms: {} }, labs: [], rooms: [], offNeed: 0, c: { online: 0, offline: 0 }, roomList: [] };
  // Room capacity for one building: by a percentage or by a number of seats shared equally, always from the original capacity
  const adjustCap = (b, how, v) => setSettings((st) => {
    const idx = st.rooms.map((r, i) => [r, i]).filter(([r]) => r.building === b && r.enabled);
    const n = idx.length;
    if (!n) return st;
    const rooms = st.rooms.slice();
    idx.forEach(([r, i], j) => {
      const base = r.base ?? Number(r.capacity);
      const cap = how === 'pct' ? Math.round((base * v) / 100) : how === 'num' ? base + Math.trunc(v / n) + (j < Math.abs(v % n) ? Math.sign(v) : 0) : base;
      rooms[i] = how === 'reset' ? { ...r, capacity: base, base: undefined } : { ...r, base, capacity: Math.max(0, cap) };
    });
    return { ...st, rooms };
  });
  const setRoomCap = (k, v) => setSettings((st) => ({ ...st, rooms: st.rooms.map((r) => (roomKey(r) === k ? { ...r, base: r.base ?? Number(r.capacity), capacity: Math.max(0, Number(v) || 0) } : r)) }));

  // ---------- handlers
  async function onAdmission(file) {
    if (!file) return;
    setError(''); setNotice(''); setBusy('Reading admission list…');
    try {
      const ExcelJS = await loadExcel();
      const r = await readAdmissions(ExcelJS, await file.arrayBuffer());
      if (!r.students.length) throw new Error(`no Active students found${r.sheet ? ` on sheet “${r.sheet}”` : ''}${r.inactive ? ` (${r.inactive} rows are not marked Active in the Status column)` : ''}. Check that the Status column says Active, or use the Template link for the expected layout`);
      setAdm({ ...r, fileName: file.name });
      setResults(null);
      const names = new Set(r.students.map((s) => s.Batch));
      setStudentMode({});
      setTop((t) => new Set([...t].filter((b) => names.has(b))));
    } catch (e) { setError(`Could not read “${file.name}”: ${e.message}`); }
    setBusy('');
  }
  function removeAdmission() {
    if (!confirm('Remove the admission list from this page? Your rooms and settings stay.')) return;
    setAdm(null); setStudentMode({}); setResults(null); setNotice(''); setCur(''); setStep('students');
  }
  async function onRoomFile(file) {
    if (!file) return;
    setError(''); setBusy('Reading room list…');
    try {
      const ExcelJS = await loadExcel();
      const { rooms, labs } = await readRooms(ExcelJS, await file.arrayBuffer());
      setSettings((s) => {
        const newLabs = labs.length ? labs.map((l) => ({ ...l, girls: s.labs.find((x) => roomKey(x) === roomKey(l))?.girls ?? true })) : s.labs;
        const order = allBuildingsOf(rooms, newLabs, s.buildingOrder);
        return { ...s, rooms, labs: newLabs, buildingOrder: order, buildingsOn: [...new Set([...s.buildingsOn, ...order.filter((b) => !s.buildingOrder.includes(b))])] };
      });
      setNotice(`Room list loaded: ${pl(rooms.length, 'room')}${labs.length ? `, ${pl(labs.length, 'lab')}` : ''}.`);
    } catch (e) { setError(`Could not read the room list: ${e.message}`); }
    setBusy('');
  }
  async function onRollFile(flag, file) {
    if (!file) return;
    setError(''); setBusy(`Reading the ${flag.toLowerCase()} list…`);
    try {
      const ExcelJS = await loadExcel();
      const rolls = await readRollList(ExcelJS, await file.arrayBuffer());
      setLists((l) => ({ ...l, [flag]: rolls }));
      setNotice(`${flag} list: ${pl(rolls.length, 'roll number')} loaded.`);
    } catch (e) { setError(`Could not read the ${flag.toLowerCase()} list: ${e.message}`); }
    setBusy('');
  }

  // "Create seating for this test": labs and buildings are picked automatically (in fill order, just enough seats for
  // every online and offline student of all its papers, after seats used by an overlapping test), then the seating is
  // prepared at once. The operator can change labs or buildings afterwards on the Online / Offline seating steps.
  function autoPick(p) {
    const used = p.slots.filter((s) => s.batches.length);
    const c = addUp(used.map((s) => countSlot(p, s)).concat([{ online: 0, offline: 0 }]));
    const taken = { labs: {}, rooms: {} };
    for (const s of used) for (const kind of ['labs', 'rooms']) for (const [k, v] of Object.entries(planOf(s).taken[kind])) taken[kind][k] = Math.max(taken[kind][k] || 0, v.filled);
    const free = (r, kind) => Math.max(0, Number(r.capacity) - (taken[kind][roomKey(r)] || 0));
    const rank = (b) => { const i = settings.buildingOrder.indexOf(b); return i < 0 ? 1e9 : i; };
    // Rooms and labs kept for monitoring or due fee students are used anyway: count them first
    const labs = [];
    let labSeats = labsOn.filter((l) => pinKeys.has(roomKey(l))).reduce((a, l) => a + free(l, 'labs'), 0);
    for (const l of labsOn.map((r, i) => ({ r, i })).sort((a, b) => rank(a.r.building) - rank(b.r.building) || a.i - b.i).map((x) => x.r)) {
      if (labSeats >= c.online) break;
      if (free(l, 'labs') > 0 && !pinKeys.has(roomKey(l))) { labs.push(roomKey(l)); labSeats += free(l, 'labs'); }
    }
    const offNeed = c.offline + Math.max(0, c.online - labSeats);
    const buildings = [];
    let roomSeats = roomsOn.filter((r) => pinKeys.has(roomKey(r))).reduce((a, r) => a + free(r, 'rooms'), 0);
    for (const b of settings.buildingOrder.filter((x) => settings.buildingsOn.includes(x))) {
      if (roomSeats >= offNeed) break;
      const seats = roomsOn.filter((r) => r.building === b && !pinKeys.has(roomKey(r))).reduce((a, r) => a + free(r, 'rooms'), 0);
      if (seats > 0) { buildings.push(b); roomSeats += seats; }
    }
    return { labs, buildings };
  }
  function startPaper(p) {
    const unset = (s) => !(s.labs && s.labs.length) && !(s.buildings && s.buildings.length);
    const pick = p.slots.some((s) => s.batches.length && unset(s)) ? autoPick(p) : null;
    const np = { ...p, done: true, slots: p.slots.map((s) => (pick && unset(s) ? { ...s, labs: pick.labs, buildings: pick.buildings, roomsOff: [] } : { ...s, labs: s.labs || [], buildings: s.buildings || [] })) };
    const list = test.papers.map((x) => (x.id === p.id ? np : x));
    setCur(p.id);
    setTest((t) => ({ ...t, papers: list }));
    run(seed, list);
  }

  // One seating per test: the students of all its papers are seated together (mixed by the shuffle rules), in one
  // allocation, then shown and exported as an Online seating (labs) and an Offline seating (classrooms).
  // A test that overlaps an earlier test (or starts within GAP minutes of it) uses the seats that test left free.
  const shiftsOfTest = (p) => p.slots.filter((s) => s.batches.length).sort((a, b) => mins(a.start) - mins(b.start)).flatMap((s) => {
    const n = slotName(p, s);
    return s.two ? [{ label: `${n} shift 1`, paper: n, start: s.start, end: s.end }, { label: `${n} shift 2`, paper: n, start: s.s2.start, end: s.s2.end }] : [{ label: n, paper: n, start: s.start, end: s.end }];
  });
  function run(newSeed, list) {
    const cbOf = {};
    for (const g of settings.cbGroups) for (const b of parseRolls(g.batches)) if (g.name.trim()) cbOf[b] = g.name.trim();
    const units = list.filter((p) => p.done && p.slots.some((s) => s.batches.length))
      .map((p, i) => ({ p, i, unit: { id: p.id, name: (p.name || '').trim() || `Test ${list.indexOf(p) + 1}`, mode: '', two: false, shifts: shiftsOfTest(p) } }))
      .sort((a, b) => mins(windowOf(a.unit).start) - mins(windowOf(b.unit).start));
    const done = [], out = [];
    units.forEach(({ p, unit }, i) => {
      const used = p.slots.filter((s) => s.batches.length);
      const taken = {};
      const over = done.filter((d) => overlaps(d.unit, unit));
      for (const d of over) for (const room of d.result.usedRooms) {
        const k = roomKey(room);
        taken[k] = Math.max(taken[k] || 0, room.taken + room.used);
      }
      const batchMode = {}, paperOf = {}, bs = [];
      for (const s of used) for (const b of s.batches) {
        if (!batchMap[b]) continue;
        bs.push(batchMap[b]); batchMode[b] = modeIn(p, s, b); paperOf[b] = s;
      }
      const shiftLabel = (x) => `${fmtTime(x.start)}${x.end ? ` to ${fmtTime(x.end)}` : ''}`;
      const st = bs.flatMap((b) => (studentsByBatch[b.batch] || []).map((x) => {
        const s = paperOf[b.batch];
        return { ...x, Paper: slotName(p, s), _shifts: (s.two ? [s, s.s2] : [s]).map(shiftLabel) };
      }));
      // Labs and rooms: every one chosen for any paper of the test, plus the rooms kept for monitoring and due fee students
      const labKeys = new Set(used.flatMap((s) => slotLabs(s).map(roomKey)));
      const roomKeys = new Set(used.flatMap((s) => slotRooms(s).map(roomKey)));
      const studentModeAll = used.some((s) => s.mode === 'Both') ? studentMode : {};
      const cfg = {
        rooms: settings.rooms.map((r) => ({ ...r, enabled: !!r.enabled && roomKeys.has(roomKey(r)) })),
        labs: settings.labs.map((l) => ({ ...l, enabled: !!l.enabled && labKeys.has(roomKey(l)) })),
        buildingOrder: settings.buildingOrder, girlsOk: settings.girlsOk,
        batchMode, studentMode: studentModeAll, flags, pins: Object.fromEntries(Object.entries(settings.pins || {}).filter(([f]) => !(f in lists) || lists[f].length)), cbOf,
        seatOrder: settings.seatOrder || 'mixed', apart: !!settings.apart, seed: newSeed + i, taken,
      };
      const result = allocate(st, cfg);
      done.push({ unit, result, index: i });
      const overNames = over.map((d) => d.unit.name);
      for (const kind of ['Online', 'Offline']) {
        const part = {
          ...result,
          assigned: result.assigned.filter((x) => (x.Mode === 'Online') === (kind === 'Online')),
          unseated: result.unseated.filter((x) => (x.Mode === 'Online') === (kind === 'Online')),
          usedRooms: result.usedRooms.filter((r) => r.kind === kind),
          warnings: kind === 'Online' ? result.warnings : [],
        };
        if (!part.assigned.length && !part.unseated.length) continue;
        out.push({ paper: { ...unit, id: `${unit.id}-${kind}`, name: `${unit.name} · ${kind} seating` }, kind, index: out.length, result: part, batches: [...new Set(part.assigned.concat(part.unseated).map((x) => x.Batch))].map((b) => batchMap[b]).filter(Boolean), students: part.assigned.length + part.unseated.length, overlapsWith: overNames });
      }
    });
    setResults(out); setSeed(newSeed); setView(0); setStep('result');
    try { localStorage.setItem(LAST, JSON.stringify({ title: test.title, papers: list.map((p) => ({ ...p, done: false })), top: [...top] })); } catch { /* ignore */ }
  }
  function prepare() {
    const list = test.papers.map((p) => (p.id === cur ? { ...p, done: true } : p));
    setTest((t) => ({ ...t, papers: list }));
    run(seed, list);
  }

  const metaFor = (x) => {
    const cs = [...new Set(x.batches.map((b) => b.course))];
    const n = { online: x.result.assigned.filter((s) => s.Mode === 'Online').length, offline: x.result.assigned.filter((s) => s.Mode !== 'Online').length };
    const head = test.title || `${cs.join(' & ') || 'Test'} Seating Arrangement`;
    const sh = shiftsOf(x.paper);
    return {
      paper: x.paper.name,
      title: `${head} - ${x.paper.name}`,
      dateISO: test.date,
      shifts: sh.map((s) => s.start),
      slots: sh,
      batchesLabel: [...new Set(x.batches.map((b) => b.batch.slice(0, 2)))].sort().join(', '),
      modeLabel: x.kind === 'Online' && n.offline ? 'Online Test (moved offline listed separately)' : n.online && n.offline ? 'Offline + Online Test' : n.online ? 'Online Test' : 'Offline Test',
    };
  };

  const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const dayTitle = () => test.title || `${[...new Set((results || []).flatMap((x) => x.batches.map((b) => b.course)))].join(' & ') || 'Test'} Seating Arrangement`;
  // One Excel file for every paper of the day
  async function download() {
    const todo = results.filter((x) => x.batches.length);
    setBusy('Building the Excel file…');
    try {
      const ExcelJS = await loadExcel();
      const wb = buildWorkbookAll(ExcelJS, todo.map((x) => ({ result: x.result, meta: metaFor(x) })), { dateISO: test.date, dayTitle: dayTitle() });
      saveBlob(new Blob([await wb.xlsx.writeBuffer()], { type: XLSX }), `${dayTitle().replace(/[^\w]+/g, '_')}_${fmtDateDMY(test.date)}.xlsx`);
    } catch (e) { setError(`Could not build the file: ${e.message}`); }
    setBusy('');
  }
  async function downloadBoard() {
    setBusy('Building the notice board…');
    try {
      const ExcelJS = await loadExcel();
      const wb = buildNoticeBoard(ExcelJS, results.filter((x) => x.batches.length).map((x) => ({ result: x.result, meta: metaFor(x) })), { dateISO: test.date, dayTitle: dayTitle() });
      saveBlob(new Blob([await wb.xlsx.writeBuffer()], { type: XLSX }), `Notice_board_${fmtDateDMY(test.date)}.xlsx`);
    } catch (e) { setError(`Could not build the notice board: ${e.message}`); }
    setBusy('');
  }

  const setupOk = !!adm && !!test.date && !!curP && paperOk(curP);
  const canOpen = (id) => (id === 'students' ? true : id === 'test' ? !!adm : id === 'result' ? !!results : setupOk);
  const doneMark = (id) => (id === 'students' ? !!adm : id === 'test' ? papers.some((p) => p.done) : false);
  const labsNeeded = curP ? curP.slots.some((s) => countSlot(curP, s).online > 0) : false;
  const offlineNeeded = curP ? curP.slots.some((s) => s.batches.length && planOf(s).offNeed > 0) : false;

  return (
    <div className="app">
      <header className="top">
        <div className="top-in">
          <div className="logo">
            <img src="/brand/matrix-logo-green.png" alt="Matrix Academy" className="logo-img" />
            <span className="sep" aria-hidden="true" />
            <div className="brand"><b>Test <em>Seating</em></b><span className="chip">{test.date ? `${fmtDateDMY(test.date)} · ${pl(papers.reduce((a, p) => a + p.slots.length, 0), 'paper')}` : 'No test date yet'}</span></div>
          </div>
          <nav className="stepper" aria-label="Steps">
            {STEPS.filter((s) => s.id !== 'labs' || labsNeeded || step === 'labs').map((s, i) => {
              const done = doneMark(s.id) && step !== s.id;
              return (
                <button key={s.id} type="button" disabled={!canOpen(s.id)} aria-current={step === s.id ? 'step' : undefined}
                  className={`st ${step === s.id ? 'on' : ''} ${done ? 'done' : ''}`} onClick={() => setStep(s.id)}>
                  <i>{done ? '✓' : i + 1}</i><span>{s.label}</span>
                </button>
              );
            })}
          </nav>
        </div>
        {adm && step !== 'students' && (
          <SummaryBar adm={adm} papers={papers} appearing={appearing} groups={timeGroups} special={special}
            results={step === 'result' ? results : null} goRules={() => setStep('students')} />
        )}
      </header>

      <main className="main">
        {error && <div className="alert bad" role="alert"><span>{error}</span><button className="link" onClick={() => setError('')}>Dismiss</button></div>}
        {notice && !error && <div className="alert info"><span>{notice}</span><button className="link" onClick={() => setNotice('')}>OK</button></div>}
        {busy && <div className="busy" role="status"><span className="spin" aria-hidden="true" />{busy}</div>}

        {step === 'students' && (
          <section>
            <Head kicker="Step 1" title="Upload the admission list" lead="The Excel file with every student. Only students marked Active are used. Files are read inside this browser and are never uploaded anywhere." />
            {!adm && <Guide />}
            <FileCard title="Admission list" hint="Columns: Roll_Number, Student_Name, Gender, Course, Batch, Medium, Status"
              loaded={adm && <><b>{adm.fileName}</b><span>{pl(adm.students.length, 'active student')} · {pl(batches.length, 'batch')}{adm.inactive ? ` · ${adm.inactive} inactive left out` : ''}{adm.duplicates ? ` · ${adm.duplicates} duplicate rolls skipped` : ''}</span></>}
              onFile={onAdmission} onRemove={removeAdmission} template={() => downloadTemplate('admission')} big />
            {adm && !adm.students.some((x) => String(x.Gender || '').trim()) && <div className="alert warn"><span>The admission list has no Gender values, so every student counts as a boy: “Girls allowed” rules for buildings and labs cannot work. Add a Gender column (M / F) to use them.</span></div>}
            <details className="fold">
              <summary>Monitoring and due fee students <span className="muted">· optional · {lists.Monitoring.length} monitoring, {lists['Due fee'].length} due fee</span></summary>
              <p className="muted">Upload or paste their roll numbers, then pick the rooms where they sit. They are seated in these rooms first; other students fill the free seats only after every other room is full.</p>
              {['Monitoring', 'Due fee'].map((f) => (
                <div className="card special" key={f}>
                  <div className="special-head"><b>{f} students</b><span className="muted small">{pl(lists[f].length, 'roll number')}</span></div>
                  <div className="special-list">
                    <RollBox rolls={lists[f]} onChange={(rolls) => setLists((l) => ({ ...l, [f]: rolls }))} label={`${f} roll numbers`} />
                    <RollUpload onFile={(file) => onRollFile(f, file)} />
                  </div>
                  {lists[f].length > 0 && <PinPicker choices={curChoices} value={settings.pins[f] || []} onChange={(keys) => setSettings((x) => ({ ...x, pins: { ...x.pins, [f]: keys } }))} label={f} />}
                </div>
              ))}
            </details>
            <details className="fold">
              <summary>Seating rules <span className="muted">· {SEAT_ORDERS[settings.seatOrder || 'mixed']}{settings.apart ? ' · brothers and sisters apart' : ''}</span></summary>
              <SeatRules settings={settings} setSettings={setSettings} students={adm?.students || []} />
            </details>
            <details className="fold">
              <summary>Phase list <span className="muted">· optional · {settings.phaseText ? 'pasted list in use' : 'phases from batch codes (XA1, XA2 → XA)'}</span></summary>
              <p className="muted">Paste the course and phase list from the Counselling Planner, one phase per line: Course | Phase | Batches. Batches not in the list are grouped by their batch code. Saved in this browser.</p>
              <textarea rows="5" placeholder={'Momentum | XA | XA1, XA2, XA3\nMomentum | XB | XB1, XB2'} value={settings.phaseText || ''}
                onChange={(e) => setSettings((s) => ({ ...s, phaseText: e.target.value }))} aria-label="Phase list" />
              {phaseList.issues.map((x) => <p key={x} className="muted small bad">{x}</p>)}
            </details>
            {adm && <Nav next={() => setStep('test')} />}
          </section>
        )}

        {step === 'test' && (
          <PaperStep test={test} setTest={setTest} setPaper={setPaper} setSlot={setSlot} courses={courses} phases={phases} ownerOf={ownerOf} slotOf={slotOf}
            batchMap={batchMap} phaseBatches={phaseBatches} studentsByBatch={studentsByBatch} studentMode={studentMode} setStudentMode={setStudentMode}
            top={top} setTop={setTop} sumSlot={sumSlot} paperOk={paperOk} start={startPaper} back={() => setStep('students')} />
        )}

        {step === 'labs' && curP && (
          <LabStep p={curP} idx={curIdx} labsOn={labsOn} order={settings.buildingOrder} setSlot={setSlot} countSlot={countSlot} planOf={planOf}
            back={() => setStep('test')} next={() => (offlineNeeded ? setStep('rooms') : prepare())} label={offlineNeeded ? 'Continue' : 'Prepare seating'} />
        )}

        {step === 'rooms' && curP && (
          <RoomStep p={curP} idx={curIdx} settings={settings} setSettings={setSettings} setSlot={setSlot} planOf={planOf}
            roomsOn={roomsOn} slotRooms={slotRooms} onRoomFile={onRoomFile} adjustCap={adjustCap} setRoomCap={setRoomCap}
            back={() => setStep(labsNeeded ? 'labs' : 'test')} next={prepare} label="Prepare seating" />
        )}

        {step === 'result' && results && (
          <ResultStep results={results} view={view} setView={setView} metaFor={metaFor} multi={results.length > 1} order={settings.buildingOrder}
            reshuffle={() => { if (confirm('Reshuffle every seat? The current seating is replaced.')) run(Math.floor(Math.random() * 1e9), test.papers); }}
            download={download} downloadBoard={downloadBoard} back={() => setStep('rooms')} />
        )}
        {step === 'result' && results && papers.some((p) => !p.done) && (
          <div className="alert info"><span>{papers.map((p, i) => (p.done ? '' : `Test ${i + 1}`)).filter(Boolean).join(', ')}: seating not made yet.</span><button className="btn sm" onClick={() => setStep('test')}>Set up next paper</button></div>
        )}
      </main>
      <footer className="foot">Matrix Academy, Sikar · Operations · Student data stays in this browser</footer>
    </div>
  );
}

// Summary box at the top, always visible: each number with its own label.
function SummaryBar({ adm, papers, appearing, groups, special, results, goRules }) {
  const n = (v) => Number(v || 0).toLocaleString('en-IN');
  const Tile = ({ label, value, sub, tone = '', children }) => (
    <div className={`tile ${tone}`}><span className="tile-l">{label}</span><b className="tile-v">{value}</b>{sub && <span className="tile-s">{sub}</span>}{children}</div>
  );
  const specials = special.map((x) => (
    <Tile key={x.flag} label={`${x.flag} students`} value={n(x.n)} tone={x.kept ? 'good' : 'bad'}
      sub={x.kept ? 'Room kept for them' : 'Not assigned a seat yet'}>
      {!x.kept && goRules && <button className="link small" onClick={goRules}>Choose their rooms</button>}
    </Tile>
  ));
  if (results) {
    const all = results.flatMap((x) => x.result.assigned);
    const on = all.filter((s) => s.Mode === 'Online').length;
    const moved = results.reduce((a, x) => a + (x.result.movedOffline || 0), 0);
    const notSeated = results.reduce((a, x) => a + x.result.unseated.length, 0);
    return (
      <div className="sumbox" role="status">
        <Tile label="Students seated" value={n(all.length)} sub={`in ${pl(results.length, 'seating')}`} />
        <Tile label="Online (labs)" value={n(on)} tone="good" sub="students in computer labs" />
        <Tile label="Offline (classrooms)" value={n(all.length - on)} sub={moved ? `${n(moved)} moved from online, labs full` : 'students in classrooms'} />
        <Tile label="Without a seat" value={n(notSeated)} tone={notSeated ? 'bad' : 'good'} sub={notSeated ? 'see the Not_Seated sheet' : 'everyone has a seat'} />
        {specials}
      </div>
    );
  }
  return (
    <div className="sumbox" role="status">
      <Tile label="Active students" value={n(adm.students.length)} sub={`${n(appearing)} appearing · ${pl(papers.reduce((a, p) => a + p.slots.length, 0), 'paper')}`} />
      {!groups.length && <Tile label="Seats" value="–" sub="Tick batches to see online and offline seats" />}
      {groups.map((g) => (
        <Tile key={g.start} label={`${range(g)} · ${g.papers.join(', ')}`} value={`${n(g.online + g.offline - g.overflow)} students`} tone={g.offlineShort ? 'bad' : g.onlineShort ? 'warn' : 'good'}>
          {g.online > 0 && <span className="tile-s"><b>Online:</b> {n(g.online - g.overflow)} / {n(g.onlineSeats)} lab seats · {g.onlineShort ? `${n(g.onlineShort)} go offline (labs full)` : `${n(g.onlineLeft)} free`}</span>}
          {(g.offline > 0 || g.overflow > 0) && <span className="tile-s"><b>Offline:</b> {n(g.offline)} / {n(g.offlineSeats)} classroom seats · {g.offlineShort ? `${n(g.offlineShort)} without a seat` : `${n(g.offlineLeft)} free`}</span>}
          {g.items.length > 1 && <span className="tile-s muted">{g.items.length} papers at this time share these labs and rooms</span>}
        </Tile>
      ))}
      {specials}
    </div>
  );
}


function Head({ kicker, title, lead }) {
  return (
    <div className="head">
      {kicker && <span className="kicker">{kicker}</span>}
      <h1>{title}</h1>
      {lead && <p className="lead">{lead}</p>}
    </div>
  );
}

function Guide() {
  return (
    <div className="guide">
      <div><span className="kicker">First time here?</span><h2>Seating sheet in three steps</h2></div>
      <ol>
        <li><i>1</i><b>Admission list</b><span>Upload the admission list. Optional: monitoring and due fee students with their rooms, and seating rules.</span></li>
        <li><i>2</i><b>Test and papers</b><span>Courses, then each paper: name, shifts, online or offline, phases, batches. Create seating: labs and buildings are picked for you.</span></li>
        <li><i>3</i><b>Seating</b><span>Download the Excel file. Want other labs or buildings? Change them on Online seating or Offline seating.</span></li>
      </ol>
    </div>
  );
}

function FileCard({ title, hint, loaded, onFile, onRemove, template, big }) {
  const [over, setOver] = useState(false);
  const inp = useRef(null);
  return (
    <div className={`filecard ${big ? 'big' : ''} ${over ? 'over' : ''} ${loaded ? 'loaded' : ''}`}
      onDragOver={(e) => { e.preventDefault(); setOver(true); }} onDragLeave={() => setOver(false)}
      onDrop={(e) => { e.preventDefault(); setOver(false); onFile(e.dataTransfer.files[0]); }}>
      <span className="fc-ic">{Ico.file}</span>
      <div className="fc-body">
        <b className="fc-title">{title}</b>
        {loaded ? <div className="fc-stat">{loaded}</div> : <span className="muted">Drop the .xlsx file here, or click Upload. {hint}</span>}
      </div>
      <div className="fc-act">
        <button type="button" className={`btn ${loaded ? 'ghost' : 'primary'}`} onClick={() => inp.current.click()}>{Ico.upload}{loaded ? 'Replace' : 'Upload'}</button>
        <input ref={inp} hidden type="file" accept=".xlsx" onChange={(e) => { onFile(e.target.files[0]); e.target.value = ''; }} />
        {template ? <button type="button" className="tpl" onClick={template}>{Ico.down}Template</button> : <span />}
        {loaded && onRemove ? <button type="button" className="iconbtn danger" onClick={onRemove} aria-label={`Remove the ${title.toLowerCase()}`} title={`Remove the ${title.toLowerCase()}`}>{Ico.bin}</button> : <span />}
      </div>
    </div>
  );
}

function Nav({ back, next, nextOk = true, nextLabel = 'Continue', hint, children }) {
  return (
    <div className="nav">
      {back && <button className="btn ghost" onClick={back}>Back</button>}
      <span className="nav-hint muted">{hint}</span>
      {children}
      {next && <button className="btn primary" disabled={!nextOk} onClick={next}>{nextLabel} →</button>}
    </div>
  );
}

function PaperStep({ test, setTest, setPaper, setSlot, courses, phases, ownerOf, slotOf, phaseBatches, studentsByBatch, studentMode, setStudentMode, top, setTop, sumSlot, paperOk, start, back }) {
  const addPaper = () => setTest({ ...test, papers: [...test.papers, newPaper()] });
  const removePaper = (id) => setTest({ ...test, papers: test.papers.filter((p) => p.id !== id) });
  return (
    <section>
      <Head kicker="Step 2" title="Test and papers" lead="Pick the date and the courses. Then set up each paper: name, 1 or 2 shifts, online or offline, phases and batches. The seat summary at the top updates as you click." />
      <div className="card form">
        <label className="field">
          <span>Test date</span>
          <input type="date" value={test.date} onChange={(e) => setTest({ ...test, date: e.target.value })} autoFocus />
        </label>
        <label className="field grow">
          <span>Sheet title <em className="muted">(optional)</em></span>
          <input type="text" placeholder="e.g. Momentum PRT Seating Arrangement" value={test.title} onChange={(e) => setTest({ ...test, title: e.target.value })} />
        </label>
      </div>
      {!test.date && <p className="muted small">Pick the test date first.</p>}
      {test.date && (
        <>
          <div className="papers">
            {test.papers.map((p, i) => (
              <PaperCard key={p.id} p={p} i={i} n={test.papers.length} setPaper={setPaper} setSlot={setSlot} remove={() => removePaper(p.id)}
                courses={courses} phases={phases} ownerOf={ownerOf} slotOf={slotOf} phaseBatches={phaseBatches}
                studentsByBatch={studentsByBatch} studentMode={studentMode} setStudentMode={setStudentMode}
                top={top} setTop={setTop} sumSlot={sumSlot} ok={paperOk(p)} start={() => start(p)} />
            ))}
          </div>
          <button type="button" className="addslot" onClick={addPaper}>{Ico.plus}Add another test (other courses)</button>
        </>
      )}
      <Nav back={back} />
    </section>
  );
}

function TimeRow({ label, value, onChange, onRemove }) {
  const bad = !shiftOk(value);
  const set = (key, v) => {
    const o = { start: value.start, end: value.end };
    if (key === 'start') { const len = o.end && o.start ? mins(o.end) - mins(o.start) : 180; o.start = v; o.end = v ? hhmm(mins(v) + (len > 0 ? len : 180)) : o.end; } else o.end = v;
    onChange(o);
  };
  return (
    <div className={`slot ${bad ? 'bad' : ''}`}>
      <b className="shift-l">{label}</b>
      <label className="field"><span>Starts</span><input type="time" value={value.start} onChange={(e) => set('start', e.target.value)} /></label>
      <label className="field"><span>Ends</span><input type="time" value={value.end} onChange={(e) => set('end', e.target.value)} /></label>
      <span className="slot-time muted">{bad ? 'Set the end time' : `${range(value)} · ${Math.round((mins(value.end) - mins(value.start)) / 6) / 10} h`}</span>
      {onRemove && <button type="button" className="iconbtn danger" onClick={onRemove} aria-label={`Remove ${label}`} title="Remove this time">{Ico.bin}</button>}
    </div>
  );
}

// Big coloured seat summary of one paper
function SeatBoard({ s, mode }) {
  const n = (v) => Number(v || 0).toLocaleString('en-IN');
  return (
    <div className="seatboard">
      {mode !== 'Offline' && (
        <div className={`sb ${s.onlineShort ? 'warn' : 'good'}`}>
          <span>Online (computer labs)</span>
          <b>{n(s.online)} students</b>
          <small>{n(s.onlineSeats)} lab seats · {s.onlineShort ? `${n(s.onlineShort)} more than the labs: they sit offline` : `${n(s.onlineLeft)} empty online seats`}</small>
        </div>
      )}
      {(mode !== 'Online' || s.overflow > 0) && (
        <div className={`sb ${s.offlineShort ? 'bad' : 'good'}`}>
          <span>Offline (classrooms)</span>
          <b>{n(s.offline)} students</b>
          <small>{n(s.offlineSeats)} classroom seats · {s.offlineShort ? `${n(s.offlineShort)} without a seat` : `${n(s.offlineLeft)} empty offline seats`}{s.overflow ? ` · includes ${n(s.overflow)} from online` : ''}</small>
        </div>
      )}
    </div>
  );
}

const toggleIn = (list, x) => (list.includes(x) ? list.filter((y) => y !== x) : [...list, x]);
// Keep only the batches that belong to the given phases
const keepBatches = (s, phases, keys) => {
  const keep = new Set(phases.filter((g) => keys.includes(g.key)).flatMap((g) => g.batches.map((b) => b.batch)));
  return { phases: keys, batches: s.batches.filter((b) => keep.has(b)), online: s.online.filter((b) => keep.has(b)) };
};

function PaperCard({ p, i, n, setPaper, setSlot, remove, courses, phases, ownerOf, slotOf, phaseBatches, studentsByBatch, studentMode, setStudentMode, top, setTop, sumSlot, ok, start }) {
  const toggleCourse = (c) => {
    const cs = toggleIn(p.courses, c);
    setPaper(p.id, { courses: cs, slots: p.slots.map((s) => ({ ...s, ...keepBatches(s, phases, s.phases.filter((k) => cs.includes(k.split('|')[0]))) })) });
  };
  const addSlot = () => {
    const last = p.slots[p.slots.length - 1];
    const lastEnd = last.two ? last.s2.end : last.end;
    const st = hhmm(mins(lastEnd) + 60);
    setPaper(p.id, { slots: [...p.slots, newSlot(st, hhmm(mins(st) + (mins(last.end) - mins(last.start))))] });
  };
  return (
    <div className="card paper">
      <div className="paper-top">
        <span className="slot-n">{i + 1}</span>
        <b className="grow">{n > 1 ? `Test ${i + 1}` : 'Test'}{p.courses.length ? ` · ${p.courses.join(', ')}` : ''}</b>
        <button type="button" className="iconbtn danger" disabled={n === 1} onClick={remove} aria-label={`Remove test ${i + 1}`} title="Remove this test">{Ico.bin}</button>
      </div>

      <div className="q">
        <b className="q-l">1. Courses</b>
        <div className="chips">
          {courses.map((c) => <button key={c} type="button" className={`chip-btn ${p.courses.includes(c) ? 'on' : ''}`} onClick={() => toggleCourse(c)}>{c}</button>)}
        </div>
      </div>

      {p.courses.length > 0 && (
        <div className="q">
          <b className="q-l">2. Papers</b>
          <p className="muted small">For each paper: its name, 1 or 2 shifts, online or offline, its phases, then the batches of those phases.</p>
          {p.slots.map((s, k) => (
            <PaperBox key={s.id} p={p} s={s} k={k} setPaper={setPaper} setSlot={setSlot} phases={phases} ownerOf={ownerOf} slotOf={slotOf}
              pb={phaseBatches(s)} studentsByBatch={studentsByBatch} studentMode={studentMode} setStudentMode={setStudentMode}
              top={top} setTop={setTop} sum={sumSlot(p, s)} />
          ))}
          <button type="button" className="addslot sm" onClick={addSlot}>{Ico.plus}Add another paper</button>
        </div>
      )}

      {ok && (
        <div className="paper-cta">
          {p.done && <span className="pill">Seating ready ✓</span>}
          <button className="btn cta" onClick={start}>{Ico.spark}{p.done ? 'Change seating for this test' : 'Create seating for this test'} →</button>
        </div>
      )}
    </div>
  );
}

// One paper: name → shifts and times → online or offline → phases → batches of those phases
function PaperBox({ p, s, k, setPaper, setSlot, phases, ownerOf, slotOf, pb, studentsByBatch, studentMode, setStudentMode, top, setTop, sum }) {
  const [open, setOpen] = useState('');
  const set = (patch) => setSlot(p, s.id, patch);
  const label = `Paper ${k + 1}`;
  const toggleTop = (b) => setTop((t) => { const x = new Set(t); x.has(b) ? x.delete(b) : x.add(b); return x; });
  const free = pb.filter((b) => !slotOf[b.batch] || slotOf[b.batch].slot === s.id);
  const tickBatch = (b) => {
    const inThis = s.batches.includes(b);
    set({ batches: inThis ? s.batches.filter((x) => x !== b) : [...s.batches, b], online: inThis ? s.online.filter((x) => x !== b) : s.online });
  };
  return (
    <div className="timebox paperbox">
      <div className="timebox-h">
        <label className="field grow"><span>{label} name</span>
          <input type="text" placeholder={`${label} (e.g. PRT-3 Paper ${k + 1})`} value={s.name} onChange={(e) => set({ name: e.target.value })} /></label>
        {p.slots.length > 1 && <button type="button" className="iconbtn danger" onClick={() => setPaper(p.id, { slots: p.slots.filter((x) => x.id !== s.id) })} aria-label={`Remove ${label}`} title="Remove this paper">{Ico.bin}</button>}
      </div>

      <div className="q">
        <b className="q-l">a. Shifts</b>
        <div className="seg sm" role="group" aria-label="Shifts">
          <button className={!s.two ? 'on' : ''} onClick={() => set({ two: false })}>1 shift</button>
          <button className={s.two ? 'on' : ''} onClick={() => set({ two: true })}>2 shifts</button>
        </div>
        <div className="times">
          <TimeRow label={s.two ? 'Shift 1' : 'Time'} value={s} onChange={(v) => set(v)} />
          {s.two && <TimeRow label="Shift 2" value={s.s2} onChange={(v) => set({ s2: v })} />}
          {s.two && mins(s.s2.start) < mins(s.end) && <p className="muted small bad">Shift 2 must start after Shift 1 ends.</p>}
          {s.two && <p className="muted small">Both shifts use the same seating: a student keeps one seat. The Excel file has an attendance sheet for each shift.</p>}
        </div>
      </div>

      <div className="q">
        <b className="q-l">b. Online or offline?</b>
        <div className="seg" role="group" aria-label="Online or offline">
          {[['Online', 'Online'], ['Offline', 'Offline'], ['Both', 'Both online and offline']].map(([m, l]) => (
            <button key={m} className={s.mode === m ? 'on' : ''} onClick={() => set({ mode: m })}>{l}</button>
          ))}
        </div>
      </div>

      {s.mode && (
        <div className="q">
          <b className="q-l">c. Phases in {s.name.trim() || label}</b>
          <div className="chips">
            {phases.filter((g) => p.courses.includes(g.course)).map((g) => {
              const other = ownerOf[g.key] && ownerOf[g.key] !== s.id;
              return (
                <button key={g.key} type="button" className={`chip-btn ${s.phases.includes(g.key) ? 'on' : ''}`} disabled={other}
                  title={other ? 'Already in another paper' : g.batches.map((b) => b.batch).join(', ')} onClick={() => set(keepBatches(s, phases, toggleIn(s.phases, g.key)))}>
                  {p.courses.length > 1 ? `${g.course} · ` : ''}{g.phase} <small>{g.count}</small>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {s.mode && s.phases.length > 0 && (
        <div className="q">
          <b className="q-l">d. Batches</b>
          <p className="muted small">Batches of the phases picked above.{s.mode === 'Both' ? ' Mark each one Online or Offline.' : ''} <button className="link" onClick={() => set({ batches: free.map((b) => b.batch) })}>Tick all</button> · <button className="link" onClick={() => set({ batches: [], online: [] })}>Clear</button></p>
          <div className="batch-grid">
            {pb.map((b) => {
              const where = slotOf[b.batch];
              const here = s.batches.includes(b.batch);
              const elsewhere = where && where.slot !== s.id;
              const on = here && s.online.includes(b.batch);
              return (
                <div key={b.batch} className={`batch ${here ? 'on' : ''} ${elsewhere ? 'other' : ''}`}>
                  <label className="batch-top">
                    <input type="checkbox" checked={here} onChange={() => tickBatch(b.batch)} disabled={!!elsewhere} />
                    <span className="batch-name">{b.batch}</span>
                    {elsewhere ? <span className="pill muted-pill">Other paper</span> : <span className="batch-count">{b.count}</span>}
                  </label>
                  {here && (
                    <div className="batch-opts">
                      {s.mode === 'Both'
                        ? <div className="seg sm">{['Online', 'Offline'].map((m) => <button key={m} className={(on ? 'Online' : 'Offline') === m ? 'on' : ''} onClick={() => set({ online: m === 'Online' ? [...new Set([...s.online, b.batch])] : s.online.filter((x) => x !== b.batch) })}>{m}</button>)}</div>
                        : <span className="small muted">{s.mode}</span>}
                      {s.mode === 'Both' && <button className="link small" onClick={() => setOpen(open === b.batch ? '' : b.batch)}>Students</button>}
                      {s.mode !== 'Offline' && <label className="small top-tick" title="Top batch students take lab seats first"><input type="checkbox" checked={top.has(b.batch)} onChange={() => toggleTop(b.batch)} /> Top</label>}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
          {open && s.batches.includes(open) && (
            <StudentModes batch={open} students={studentsByBatch[open] || []} batchMode={s.online.includes(open) ? 'Online' : 'Offline'}
              studentMode={studentMode} setStudentMode={setStudentMode} close={() => setOpen('')} />
          )}
        </div>
      )}

      {s.batches.length > 0 && <SeatBoard s={sum} mode={s.mode} />}
    </div>
  );
}

function StudentModes({ batch, students, batchMode, studentMode, setStudentMode, close }) {
  const [q, setQ] = useState('');
  const list = students.filter((s) => !q || `${s.Roll_Number} ${s.Student_Name}`.toLowerCase().includes(q.toLowerCase()));
  const modeOf = (s) => studentMode[s.Roll_Number] || batchMode;
  const set = (s, m) => setStudentMode((o) => { const n = { ...o }; if (m === batchMode) delete n[s.Roll_Number]; else n[s.Roll_Number] = m; return n; });
  const online = students.filter((s) => modeOf(s) === 'Online').length;
  return (
    <div className="card smodes">
      <div className="smodes-head">
        <b>{batch}: {online} online · {students.length - online} offline</b>
        <input className="search" type="search" placeholder="Find a roll number or name" value={q} onChange={(e) => setQ(e.target.value)} />
        <button className="btn ghost sm" onClick={close}>Done</button>
      </div>
      <div className="smodes-list">
        {list.map((s) => (
          <div key={s.Roll_Number} className={`srow ${modeOf(s) !== batchMode ? 'changed' : ''}`}>
            <span className="srow-roll">{s.Roll_Number}</span>
            <span className="srow-name">{s.Student_Name}</span>
            <span className="muted small">{s.Gender}</span>
            <div className="seg sm">
              {['Offline', 'Online'].map((m) => <button key={m} className={modeOf(s) === m ? 'on' : ''} onClick={() => set(s, m)}>{m}</button>)}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// One line that says what is left to place, in big type
function BigLine({ left, placed, empty, what, more }) {
  const n = (v) => Number(v || 0).toLocaleString('en-IN');
  return left > 0
    ? <div className="bigline bad"><b>{n(left)}</b> {what} students still need a seat. {more}</div>
    : <div className="bigline good">All {n(placed)} {what} students have a seat · <b>{n(empty)}</b> seats empty</div>;
}

// One card for the whole test (all its papers share one seating); a change applies to every paper of the test
const testTitle = (p) => { const used = p.slots.filter((s) => s.batches.length); return `${used.map((s) => `${slotName(p, s)} (${slotLabel(s)})`).join(' + ')} · ${pl(used.reduce((a, s) => a + s.batches.length, 0), 'batch')}`; };
function LabStep({ p, idx, labsOn, order, setSlot: setOne, planOf, back, next, label }) {
  const setSlot = (pp, sid, patch) => setOne(pp, '*', patch);
  const slots = p.slots.filter((s) => s.batches.length && planOf(s).c.online > 0).slice(0, 1);
  const groups = {};
  for (const l of labsOn) (groups[l.building] ||= []).push(l);
  const blds = Object.keys(groups).sort((a, b) => order.indexOf(a) - order.indexOf(b));
  return (
    <section>
      <Head kicker={`Test ${idx + 1}: ${p.courses.join(', ')}`} title="Online seating: which labs?" lead="Tick the computer labs for the online batches. Labs fill in the order shown, one lab completely before the next. Top batch students are seated first. Online students left without a lab seat sit offline." />
      {slots.map((s) => {
        const need = planOf(s).c.online;
        const sel = new Set(p.slots.flatMap((x) => x.labs || []));
        const used = planOf(s).taken.labs;
        const free = (l) => Math.max(0, Number(l.capacity) - (used[roomKey(l)]?.filled || 0));
        const chosen = labsOn.filter((l) => sel.has(roomKey(l))).map((l) => ({ ...l, capacity: free(l) }));
        const prev = fillPreview(need, chosen, order);
        const fillOf = Object.fromEntries(prev.rooms.map((r) => [roomKey(r), r]));
        const empty = prev.rooms.reduce((a, r) => a + r.empty, 0);
        const toggle = (k) => setSlot(p, s.id, { labs: sel.has(k) ? [...sel].filter((x) => x !== k) : [...sel, k] });
        return (
          <div className="card slotcard" key={s.id}>
            <h2>{testTitle(p)}</h2>
            <BigLine left={prev.left} placed={need} empty={empty} what="online" more="Select more labs, or continue and they sit offline." />
            <p className="muted small"><button className="link" onClick={() => setSlot(p, s.id, { labs: labsOn.map(roomKey) })}>Select all labs</button> · <button className="link" onClick={() => setSlot(p, s.id, { labs: [] })}>Clear</button></p>
            {blds.map((b) => {
              const on = groups[b].filter((l) => sel.has(roomKey(l)));
              const seats = on.reduce((a, l) => a + Number(l.capacity), 0);
              const filled = on.reduce((a, l) => a + (fillOf[roomKey(l)]?.filled || 0), 0);
              return (
                <div className="lab-b" key={b}>
                  <div className="lab-bh"><b>{b}</b><span className="muted small">{on.length ? `${on.length} selected · ${seats} seats · ${filled} filled · ${seats - filled} empty` : `${pl(groups[b].length, 'lab')}`}</span></div>
                  <div className="lab-list">
                    {groups[b].map((l) => {
                      const k = roomKey(l), f = fillOf[k];
                      return (
                        <label key={k} className={`lab-row ${sel.has(k) ? 'on' : ''}`}>
                          <input type="checkbox" checked={sel.has(k)} onChange={() => toggle(k)} />
                          <span className="lab-name">{l.room}</span>
                          <span className="muted small">{l.capacity} seats{l.girls === false ? ' · no girls' : ''}</span>
                          {used[k] && <span className="small warnt">{used[k].filled} used by {used[k].by.join(', ')} at this time · {free(l)} free</span>}
                          {f && <span className={`small ${f.empty ? 'muted' : 'good'}`}>{f.filled} filled · {f.empty} empty</span>}
                        </label>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        );
      })}
      <div className="ctabar">
        <div><b>Labs</b><span className="muted">{`${new Set(p.slots.flatMap((x) => x.labs || [])).size} labs for all papers of this test`}</span></div>
        <div className="cta-row">
          <button className="btn ghost" onClick={back}>Back</button>
          <button className="btn cta" onClick={next}>{label || 'Continue'} →</button>
        </div>
      </div>
    </section>
  );
}

function RoomStep({ p, idx, settings, setSettings, setSlot: setOne, planOf, roomsOn, onRoomFile, adjustCap, setRoomCap, back, next, label }) {
  const setSlot = (pp, sid, patch) => setOne(pp, '*', patch);
  const editRoom = (kind, i, patch) => setSettings((s) => ({ ...s, [kind]: s[kind].map((r, j) => (j === i ? { ...r, ...patch } : r)) }));
  const move = (i, d) => setSettings((s) => {
    const o = s.buildingOrder.slice(); const j = i + d;
    if (j < 0 || j >= o.length) return s;
    [o[i], o[j]] = [o[j], o[i]]; return { ...s, buildingOrder: o };
  });
  const toggleGirls = (b) => setSettings((s) => ({ ...s, girlsOk: { ...s.girlsOk, [b]: s.girlsOk[b] === false } }));
  const roomsIn = (b) => roomsOn.filter((r) => r.building === b);
  const buildings = settings.buildingOrder.filter((b) => roomsIn(b).length);
  const slots = p.slots.filter((s) => s.batches.length && planOf(s).offNeed > 0).slice(0, 1);
  return (
    <section>
      <Head kicker={`Test ${idx + 1}: ${p.courses.join(', ')}`} title="Offline seating: buildings and rooms" lead="Tick the buildings for the offline students. Every room of a ticked building is used; open Choose rooms to leave a room out or change its capacity. All papers of this test are seated together. Seats another test fills at the same time are shown as used." />
      {slots.map((s) => {
        const pl0 = planOf(s);
        const need = pl0.offNeed;
        const used = pl0.taken.rooms;
        const usedIn = (r) => used[roomKey(r)]?.filled || 0;
        const free = (r) => Math.max(0, Number(r.capacity) - usedIn(r));
        const rooms = pl0.roomList.map((r) => ({ ...r, capacity: free(r) }));
        const order = settings.buildingOrder.filter((b) => sel.has(b));
        const prev = fillPreview(need, rooms, settings.buildingOrder);
        const fillOf = Object.fromEntries(prev.rooms.map((r) => [roomKey(r), r]));
        const empty = prev.rooms.reduce((a, r) => a + r.empty, 0);
        const sel = new Set(p.slots.flatMap((x) => x.buildings || []));
        const off = new Set(s.roomsOff);
        const toggleB = (b) => setSlot(p, s.id, { buildings: sel.has(b) ? [...sel].filter((x) => x !== b) : [...sel, b] });
        const toggleR = (k) => setSlot(p, s.id, { roomsOff: off.has(k) ? s.roomsOff.filter((x) => x !== k) : [...s.roomsOff, k] });
        return (
          <div className="card slotcard" key={s.id}>
            <h2>{testTitle(p)}</h2>
            <BigLine left={prev.left} placed={need} empty={empty} what="offline" more="Tick another building." />
            <div className="bld-chips">
              {buildings.map((b) => {
                const all = roomsIn(b);
                const seats = all.reduce((a, r) => a + Number(r.capacity), 0);
                const others = all.reduce((a, r) => a + usedIn(r), 0);
                const by = [...new Set(all.flatMap((r) => used[roomKey(r)]?.by || []))];
                const mine = sel.has(b) ? all.filter((r) => !off.has(roomKey(r))).reduce((a, r) => a + (fillOf[roomKey(r)]?.filled || 0), 0) : 0;
                return (
                  <label key={b} className={`bld-chip ${sel.has(b) ? 'on' : ''}`}>
                    <input type="checkbox" checked={sel.has(b)} onChange={() => toggleB(b)} />
                    <b>{sel.has(b) ? `${order.indexOf(b) + 1}. ` : ''}{b}</b>
                    <span>{seats} seats{others ? ` · ${others} used by ${by.join(', ')}` : ''}{sel.has(b) ? ` · ${mine} filled` : ''} · <b className={seats - others - mine ? 'good' : 'bad'}>{seats - others - mine} free</b></span>
                  </label>
                );
              })}
            </div>
            <CapacityTool buildings={buildings} roomsIn={roomsIn} adjustCap={adjustCap} />
            {sel.size > 0 && (
              <details className="fold inner">
                <summary>Choose rooms <span className="muted">· {rooms.length} rooms in use · capacity can be typed per room</span></summary>
                {order.map((b) => (
                  <div key={b} className="room-line">
                    <b>{b}</b>
                    <div className="room-chips">
                      {roomsIn(b).map((r) => {
                        const k = roomKey(r), f = fillOf[k], use = !off.has(k), u = usedIn(r);
                        return (
                          <div key={k} className={`room-chip ${use ? 'on' : ''} ${f && f.filled ? 'filled' : ''} ${u ? 'taken' : ''}`} title={u ? `${u} seats used by ${used[k].by.join(', ')}` : ''}>
                            <label><input type="checkbox" checked={use} onChange={() => toggleR(k)} /> <b>{r.room}</b></label>
                            <input className="cap sm" type="number" min="0" value={r.capacity} onChange={(e) => setRoomCap(k, e.target.value)} aria-label={`Capacity of ${b} ${r.room}`} />
                            <span>{u ? `${u} used · ` : ''}{use && f ? `${f.filled} filled` : ''}</span>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </details>
            )}
          </div>
        );
      })}
      <details className="fold">
        <summary>Building fill order, girls allowed, capacities and room list</summary>
        <ol className="order">
          {settings.buildingOrder.map((b, i) => (
            <li key={b}>
              <span className="order-n">{i + 1}</span>
              <span className="order-name">{b}</span>
              <span className="muted small">{roomsIn(b).reduce((a, r) => a + Number(r.capacity), 0)} seats</span>
              <label className="far"><input type="checkbox" checked={settings.girlsOk[b] !== false} onChange={() => toggleGirls(b)} /> Girls allowed</label>
              <button className="iconbtn" onClick={() => move(i, -1)} aria-label={`Move ${b} up`} disabled={i === 0}>▲</button>
              <button className="iconbtn" onClick={() => move(i, 1)} aria-label={`Move ${b} down`} disabled={i === settings.buildingOrder.length - 1}>▼</button>
            </li>
          ))}
        </ol>
        <RoomTable title="Classrooms (offline)" kind="rooms" list={settings.rooms} edit={editRoom} use={settings.buildingOrder} />
        <RoomTable title="Computer labs (online)" kind="labs" list={settings.labs} edit={editRoom} use={settings.buildingOrder} labs />
        <p className="muted">Upload a new room list only when rooms change. Uploads stay saved in this browser.</p>
        <div className="files"><FileCard title="Room list" hint="Sheets Room_Capacity and LAB_Capacity" onFile={onRoomFile} template={() => downloadTemplate('rooms')} /></div>
        <button className="btn ghost danger" onClick={() => { if (confirm('Reset buildings, rooms, capacities, girls settings and groups to the built-in lists?')) setSettings(freshSettings()); }}>{Ico.bin}Reset to built-in lists</button>
      </details>
      <div className="ctabar">
        <div><b>Buildings</b><span className="muted">{`All papers: ${[...new Set(p.slots.flatMap((x) => x.buildings || []))].join(', ') || 'none yet'}`}</span></div>
        <div className="cta-row">
          <button className="btn ghost" onClick={back}>Back</button>
          <button className="btn cta" onClick={next}>{label || 'Continue'} →</button>
        </div>
      </div>
    </section>
  );
}

// Change the room capacity of one building: by a percentage, or by a number of seats shared equally between its rooms
function CapacityTool({ buildings, roomsIn, adjustCap }) {
  const [b, setB] = useState('');
  const [pct, setPct] = useState('');
  const [num, setNum] = useState('');
  const cur = b || buildings[0] || '';
  const list = cur ? roomsIn(cur) : [];
  const seats = list.reduce((a, r) => a + Number(r.capacity), 0);
  const base = list.reduce((a, r) => a + Number(r.base ?? r.capacity), 0);
  if (!buildings.length) return null;
  return (
    <div className="captool">
      <b className="small">Change capacity</b>
      <select value={cur} onChange={(e) => setB(e.target.value)} aria-label="Building">{buildings.map((x) => <option key={x} value={x}>{x}</option>)}</select>
      <span className="muted small">{seats} seats{seats !== base ? ` (normally ${base})` : ''}</span>
      <label className="small">to <input className="cap sm" type="number" min="0" placeholder="80" value={pct} onChange={(e) => setPct(e.target.value)} aria-label="Percent of normal capacity" /> %</label>
      <button className="btn ghost sm" disabled={pct === ''} onClick={() => { adjustCap(cur, 'pct', Number(pct)); setPct(''); }}>Apply %</button>
      <label className="small">or by <input className="cap sm" type="number" placeholder="-20" value={num} onChange={(e) => setNum(e.target.value)} aria-label="Seats to add or remove" /> seats</label>
      <button className="btn ghost sm" disabled={num === ''} onClick={() => { adjustCap(cur, 'num', Number(num)); setNum(''); }}>Apply seats</button>
      {seats !== base && <button className="link small" onClick={() => adjustCap(cur, 'reset')}>Reset</button>}
    </div>
  );
}

// Seating rules (step 1, folded): seat order for the whole test, brothers and sisters apart, SMD groups
function SeatRules({ settings, setSettings, students }) {
  const COL = { name: ['Student_Name', 'student name'], father: ['Father_Name', "father's name"], dob: ['DOB', 'date of birth'], contact: ['Contact_Number', 'contact number'] };
  const groups = settings.cbGroups;
  const setGroups = (g) => setSettings((s) => ({ ...s, cbGroups: g }));
  const order = settings.seatOrder || 'mixed';
  return (
    <>
      <label className="field"><span>Seat order</span>
        <select value={order} onChange={(e) => setSettings((x) => ({ ...x, seatOrder: e.target.value }))} aria-label="Seat order">
          {Object.entries(SEAT_ORDERS).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
        </select>
      </label>
      <p className="muted small">{order === 'mixed' ? 'Students of all papers are shuffled together, and two students of the same batch never sit side by side. Reshuffle on the result page gives a new mix.' : `Every student of the test is seated in ${SEAT_ORDERS[order].toLowerCase()} order: the first room fills first, seat 1 onwards, then the next room.`}{['father', 'dob', 'contact'].includes(order) ? ' Needs that column in the admission list; students without it come last.' : ''}</p>
      {COL[settings.seatOrder] && students.length > 0 && (() => {
        const [k, w] = COL[settings.seatOrder];
        const n = students.filter((x) => String(x[k] || '').trim()).length;
        return n < students.length ? <p className="small bad">{n === 0 ? `The admission list has no ${w} column (or it is empty), so seats would follow roll numbers. Add a ${k} column.` : `${students.length - n} students have no ${w}; they are seated last.`}</p> : <p className="muted small">{w[0].toUpperCase() + w.slice(1)} found for all {students.length} students.</p>;
      })()}
      <label className="small"><input type="checkbox" checked={!!settings.apart} onChange={(e) => setSettings((x) => ({ ...x, apart: e.target.checked }))} /> Keep brothers and sisters apart (same father's name or same contact number never side by side)</label>
      <p className="muted small"><b>SMD sheet groups (CB column)</b> · {groups.length ? pl(groups.length, 'group') : 'by first letter of the batch'}. Batches not in a group are counted by the first letter of the batch (X-Phase, Y-Phase…).</p>
      {groups.map((g, i) => (
        <div className="group-row" key={i}>
          <input type="text" placeholder="Group name, e.g. Momentum XA" value={g.name} onChange={(e) => setGroups(groups.map((y, j) => (j === i ? { ...y, name: e.target.value } : y)))} aria-label="Group name" />
          <input type="text" className="grow" placeholder="Batches, e.g. XA1, XA2, XA3" value={g.batches} onChange={(e) => setGroups(groups.map((y, j) => (j === i ? { ...y, batches: e.target.value } : y)))} aria-label="Batches in this group" />
          <button className="iconbtn danger" onClick={() => setGroups(groups.filter((_, j) => j !== i))} aria-label="Remove group">{Ico.bin}</button>
        </div>
      ))}
      <button type="button" className="addslot" onClick={() => setGroups([...groups, { name: '', batches: '' }])}>{Ico.plus}Add a group</button>
    </>
  );
}

// Roll number box: keeps what the operator types (commas, spaces, new lines) and reads the roll numbers from it
function RollBox({ rolls, onChange, label }) {
  const [text, setText] = useState(rolls.join(', '));
  useEffect(() => { if (parseRolls(text).join(',') !== rolls.join(',')) setText(rolls.join(', ')); }, [rolls]); // a file was uploaded
  return <textarea rows="2" placeholder="Paste roll numbers, or upload a file" value={text} onChange={(e) => { setText(e.target.value); onChange(parseRolls(e.target.value)); }} aria-label={label} />;
}

function RollUpload({ onFile }) {
  const inp = useRef(null);
  return (
    <div className="roll-up">
      <button type="button" className="btn ghost sm" onClick={() => inp.current.click()}>{Ico.upload}Upload</button>
      <input ref={inp} hidden type="file" accept=".xlsx" onChange={(e) => { onFile(e.target.files[0]); e.target.value = ''; }} />
      <button type="button" className="tpl" onClick={() => downloadTemplate('rolls')}>{Ico.down}Template</button>
    </div>
  );
}

// Rooms kept for monitoring or due fee students: pick a building, then its rooms
function PinPicker({ choices, value, onChange, label }) {
  const blds = [...new Set(choices.map((r) => r.building))];
  const [b, setB] = useState('');
  const picked = value.filter((k) => choices.some((r) => roomKey(r) === k));
  const seats = choices.filter((r) => picked.includes(roomKey(r))).reduce((a, r) => a + Number(r.capacity || 0), 0);
  const cur = b && blds.includes(b) ? b : blds[0];
  return (
    <div className="pinpick">
      <p className="muted small">1. Building</p>
      <div className="chips">
        {blds.map((x) => {
          const n = choices.filter((r) => r.building === x && picked.includes(roomKey(r))).length;
          return <button key={x} type="button" className={`chip-btn ${cur === x ? 'on' : ''}`} onClick={() => setB(x)}>{x}{n ? ` · ${n} picked` : ''}</button>;
        })}
      </div>
      {cur && (
        <>
          <p className="muted small">2. Rooms in {cur} for {label.toLowerCase()} students</p>
          <div className="room-chips">
            {choices.filter((r) => r.building === cur).map((r) => {
              const k = roomKey(r), on = picked.includes(k);
              return (
                <label key={k} className={`room-chip ${on ? 'on filled' : ''}`}>
                  <input type="checkbox" checked={on} onChange={() => onChange(on ? picked.filter((x) => x !== k) : [...picked, k])} />
                  <b>{r.kind === 'Lab' ? `Lab ${r.room}` : r.room}</b><span>{r.capacity}</span>
                </label>
              );
            })}
          </div>
        </>
      )}
      {picked.length > 0 && <p className="small">Kept: {picked.map((k) => k.replace('|', ' · ')).join(', ')} · {seats} seats</p>}
    </div>
  );
}

function Meter({ label, need, have }) {
  const pct = have ? Math.min(100, Math.round((need / have) * 100)) : need ? 100 : 0;
  const bad = need > have;
  return (
    <div className={`meter ${bad ? 'bad' : ''}`}>
      <div className="meter-top"><span>{label}</span><b>{need} / {have}</b></div>
      <div className="meter-bar"><span style={{ width: `${pct}%` }} /></div>
      <span className="muted small">{bad ? `${need - have} short` : `${have - need} spare`}</span>
    </div>
  );
}

function RoomTable({ title, kind, list, edit, use, labs }) {
  const groups = {};
  list.forEach((r, i) => (groups[r.building] ||= []).push({ ...r, i }));
  return (
    <>
      <h2>{title}</h2>
      {Object.entries(groups).map(([b, rooms]) => {
        const on = rooms.filter((r) => r.enabled);
        const seats = on.reduce((a, r) => a + Number(r.capacity || 0), 0);
        return (
          <details key={b} className="bldg">
            <summary><span className="bldg-name">{b}</span><span className="muted">{on.length} of {rooms.length} {labs ? 'labs' : 'rooms'} on · {seats} seats{use.includes(b) ? '' : ' · building not used today'}</span></summary>
            <table className="rooms">
              <thead><tr><th>Use</th><th>{labs ? 'Lab' : 'Room'}</th><th>Capacity</th>{labs && <th>Girls allowed</th>}<th>Note</th></tr></thead>
              <tbody>
                {rooms.map((r) => (
                  <tr key={r.i} className={r.enabled ? '' : 'off'}>
                    <td><input type="checkbox" checked={r.enabled} onChange={(e) => edit(kind, r.i, { enabled: e.target.checked })} aria-label={`Use ${b} ${r.room}`} /></td>
                    <td>{r.room}</td>
                    <td><input className="cap" type="number" min="0" value={r.capacity} onChange={(e) => edit(kind, r.i, { capacity: Math.max(0, Number(e.target.value)) })} aria-label={`Capacity of ${b} ${r.room}`} /></td>
                    {labs && <td><input type="checkbox" checked={r.girls !== false} onChange={(e) => edit(kind, r.i, { girls: e.target.checked })} aria-label={`Girls allowed in ${b} ${r.room}`} /></td>}
                    <td className="muted">{r.note}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </details>
        );
      })}
    </>
  );
}

function ResultStep({ results, view, setView, metaFor, multi, order, reshuffle, download, downloadBoard, back }) {
  const [menu, setMenu] = useState(false);
  const x = results[view] || results[0];
  const result = x.result, meta = metaFor(x);
  const byBuilding = {};
  for (const r of result.usedRooms.slice().sort((a, b) => order.indexOf(a.building) - order.indexOf(b.building) || (a.kind === 'Online') - (b.kind === 'Online'))) (byBuilding[r.building] ||= []).push(r);
  const special = result.assigned.filter((s) => s.Flag).length;
  const girls = result.assigned.filter(isGirl).length;
  const unseatedAll = results.reduce((a, y) => a + y.result.unseated.length, 0);
  const files = results.filter((y) => y.batches.length).length;
  return (
    <section>
      <Head kicker={multi ? `Seating ready · ${pl(results.length, 'seating')}` : 'Seating ready'} title={meta.title}
        lead={`${fmtDateDMY(meta.dateISO)} · ${paperTime(x.paper)} · ${meta.modeLabel}${meta.batchesLabel ? ` · Batches ${meta.batchesLabel}` : ''}`} />
      {multi && (
        <div className="slottabs" role="tablist">
          {results.map((y, i) => (
            <button key={y.paper.id} role="tab" aria-selected={i === view} className={`${i === view ? 'on' : ''} ${y.result.unseated.length ? 'warn' : ''}`} onClick={() => setView(i)}>
              <b>{paperName(y.paper, y.index)}</b>
              <span>{paperTime(y.paper)} · {y.batches.length ? `${y.result.assigned.length} seated${y.result.unseated.length ? ` · ${y.result.unseated.length} no seat` : ''}` : 'no batches'}</span>
            </button>
          ))}
        </div>
      )}
      {unseatedAll > 0 && (
        <div className="alert bad"><span><b>{unseatedAll} students have no seat{multi ? ' across the papers' : ''}.</b> They are listed in the Not_Seated sheet with the reason.</span><button className="btn sm" onClick={back}>Fix buildings</button></div>
      )}
      {x.overlapsWith.length > 0 && (
        <div className="alert info"><span>This test overlaps {x.overlapsWith.join(', ')} (or starts less than {GAP} minutes after it). It uses the seats left free in shared rooms, so no seat is used twice at once.</span></div>
      )}
      {!x.batches.length && <div className="alert warn"><span>No batches are in this paper. Add them on the Batches step.</span></div>}
      <div className="actions">
        <button className="btn cta" onClick={download} disabled={!files}>{Ico.down}Download Excel{multi ? ' (all papers, one file)' : ''}</button>
        <div className="menu">
          <button className="btn ghost" onClick={() => setMenu(!menu)} aria-expanded={menu}>More ▾</button>
          {menu && (
            <div className="menu-list" onMouseLeave={() => setMenu(false)}>
              <button onClick={() => { setMenu(false); downloadBoard(); }}>Notice board only<small>One sheet for the building entrance</small></button>
              <button onClick={() => { setMenu(false); reshuffle(); }}>Reshuffle seats<small>Every student gets a new seat</small></button>
              <button onClick={() => { setMenu(false); back(); }}>Change buildings<small>Buildings, rooms, girls, special rooms</small></button>
              <button onClick={() => { setMenu(false); window.print(); }}>Print this page<small>The seat fill summary</small></button>
            </div>
          )}
        </div>
      </div>
      <div className="kpis">
        <div><b>{result.assigned.length}</b><span>of {x.students} seated</span></div>
        <div><b>{result.usedRooms.length}</b><span>rooms used</span></div>
        <div><b>{girls}</b><span>girls seated</span></div>
        <div><b>{special}</b><span>monitoring, due fee, top</span></div>
        <div><b>{shiftsOf(x.paper).length}</b><span>attendance {x.paper.two ? 'sheets (same seats)' : 'sheet'}</span></div>
      </div>
      {result.warnings.map((w) => <div key={w} className="alert warn"><span>{w}</span></div>)}

      <details className="fold" open>
        <summary>Rooms <span className="muted">· {result.usedRooms.length} used</span></summary>
        <p className="muted small">Green: students of this test (all its papers). Amber: monitoring, due fee or top batch students.{x.overlapsWith.length ? ' Grey: seats used by an overlapping test.' : ''}</p>
        {Object.entries(byBuilding).map(([b, rooms]) => (
          <div className="rbldg" key={b}>
            <div className="rb-head"><b>{b}</b><span className="muted">{pl(rooms.reduce((a, r) => a + r.used, 0), 'student')} · {pl(rooms.length, 'room')}</span></div>
            <div className="rcards">
              {rooms.map((r) => {
                const sp = r.occupants.filter((s) => s.Flag).length;
                const g = r.occupants.filter(isGirl).length;
                const w = (n) => `${(n / r.capacity) * 100}%`;
                return (
                  <div className="rcard" key={`${r.kind}${r.room}`} title={`${b} ${r.room}: ${r.used} students, ${g} girls${sp ? `, ${sp} special` : ''}${r.taken ? `, seats 1–${r.taken} used by an overlapping paper` : ''}`}>
                    <div className="rc-top"><b>{r.room}</b><span>{r.used}/{r.capacity}</span></div>
                    <div className="rc-bar"><span className="taken" style={{ width: w(r.taken) }} /><span className="home" style={{ width: w(r.used - sp) }} /><span className="moved" style={{ width: w(sp) }} /></div>
                    <span className="muted small">Seats {r.taken + 1}–{r.taken + r.used} · {pl(g, 'girl')}{sp ? ` · ${sp} special` : ''}</span>
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </details>
      <p className="muted small">One Excel file{multi ? ' for every paper' : ''}: Raw_Data, Student_list, Seating Arrangement, Attendance (every room{multi ? ', paper' : ''} and shift), SMD, Room_Layout, Master_List{unseatedAll ? ', Not_Seated' : ''}.{multi ? ' Each sheet has a Paper column or a block per paper.' : ''}</p>
    </section>
  );
}
