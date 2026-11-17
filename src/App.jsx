import React, { useEffect, useMemo, useRef, useState } from 'react';
import defaults from './data/defaults.json';
import { allocate, capacity, naturalCompare, isGirl, roomKey, FLAGS, SEAT_ORDERS } from './lib/engine.js';
import { buildWorkbookAll, buildNoticeBoard, fmtDateDMY, fmtTime } from './lib/exporter.js';
import { readAdmissions, readRooms, readRollList, parseRolls } from './lib/readers.js';
import { parsePhaseList, groupPhases, seatSummary, fillPreview, rptFill } from './lib/plan.js';

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
      // Capacity changes now belong to one test (paper.cap / paper.labCap): the saved room list keeps the original capacity
      const orig = (r) => { const { base, ...x } = r; return base != null ? { ...x, capacity: Number(base) } : x; };
      const labs = (s.labs || f.labs).map((l) => orig({ girls: true, ...l }));
      const rooms = s.rooms.map(orig);
      const order = allBuildingsOf(rooms, labs, s.buildingOrder || f.buildingOrder);
      return { ...f, ...s, rooms, labs, buildingOrder: order, buildingsOn: s.buildingsOn || order.slice() };
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
// Papers of a test that sit at the same time (or less than GAP minutes apart) are seated together in one seating;
// papers at different times each get their own seating on the same labs and rooms (seats are used again).
const shiftsOfSlot = (s) => (s.two ? [{ start: s.start, end: s.end }, s.s2] : [{ start: s.start, end: s.end }]);
const slotsClash = (a, b) => shiftsOfSlot(a).some((x) => shiftsOfSlot(b).some((y) => clash(x, y)));
export function timeGroupsOf(p) {
  const used = p.slots.filter((s) => s.batches.length).sort((a, b) => mins(a.start) - mins(b.start));
  const out = [];
  for (const s of used) {
    // join every group this paper clashes with (a paper can link two groups)
    const hit = out.filter((g) => g.some((x) => slotsClash(x, s)));
    const merged = hit.flat().concat([s]);
    for (const g of hit) out.splice(out.indexOf(g), 1);
    out.push(merged);
  }
  return out.map((g) => g.sort((a, b) => mins(a.start) - mins(b.start))).sort((a, b) => mins(a[0].start) - mins(b[0].start));
}
const newPaper = () => ({ id: uid(), name: '', courses: [], slots: [newSlot()], done: false });
function initialTest() {
  const last = loadLast();
  if (last && Array.isArray(last.papers) && last.papers.length) return { title: last.title || '', kind: last.kind || 'PRT', date: '', papers: last.papers.map((p) => ({ ...newPaper(), ...p, top: [], topLabs: [], topBlds: [], slots: (p.slots || [newSlot()]).map((s) => ({ ...newSlot(), ...s })), done: false })) };
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
const clash = (x, y) => mins(x.start) < mins(y.end) + GAP && mins(y.start) < mins(x.end) + GAP;
export const overlaps = (a, b) => a.id === b.id || shiftsOf(a).some((x) => shiftsOf(b).some((y) => clash(x, y)));
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

// In-app confirm dialog: askConfirm(text, okLabel, danger) resolves true or false
let confirmHost = null;
const askConfirm = (text, ok = 'Continue', danger = false) => new Promise((res) => (confirmHost ? confirmHost({ text, ok, danger, res }) : res(window.confirm(text))));
function ConfirmHost() {
  const [q, setQ] = useState(null);
  const okRef = useRef(null);
  useEffect(() => { confirmHost = setQ; return () => { confirmHost = null; }; }, []);
  useEffect(() => { if (q && okRef.current) okRef.current.focus(); }, [q]);
  if (!q) return null;
  const done = (v) => { q.res(v); setQ(null); };
  return (
    <div className="modal-bg" onClick={() => done(false)} onKeyDown={(e) => { if (e.key === 'Escape') done(false); }}>
      <div className="modal" role="alertdialog" aria-modal="true" aria-labelledby="modal-t" onClick={(e) => e.stopPropagation()}>
        <p id="modal-t">{q.text}</p>
        <div className="modal-act">
          <button className="btn secondary" onClick={() => done(false)}>Cancel</button>
          <button ref={okRef} className={`btn primary ${q.danger ? 'danger' : ''}`} onClick={() => done(true)}>{q.ok}</button>
        </div>
      </div>
    </div>
  );
}

const STEPS = [
  { id: 'students', label: 'Admission list' },
  { id: 'test', label: 'Test & papers' },
  { id: 'mode', label: 'Online / Offline' },
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
  const [lists, setLists] = useState({ Monitoring: [], 'Due fee': [] });
  const [seed, setSeed] = useState(() => Math.floor(Math.random() * 1e9));
  const [results, setResults] = useState(null); // [{ paper (one time slot), index, result, batches, students, overlapsWith }]
  const [view, setView] = useState(0);
  const [started, setStarted] = useState(false);
  const [sumOpen, setSumOpen] = useState(false);

  useEffect(() => {
    try { localStorage.setItem(STORE, JSON.stringify(settings)); } catch { /* ignore */ }
  }, [settings]);
  useEffect(() => {
    window.scrollTo(0, 0); setSumOpen(false);
    // phones: keep the current step visible in the step bar
    const on = document.querySelector('.stepper .st.on');
    if (on && on.parentElement) on.parentElement.scrollLeft = on.offsetLeft - 16;
  }, [step]);

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
    for (const k of Object.keys(m)) m[k].sort((a, b) => naturalCompare(String(a.Roll_Number), String(b.Roll_Number)));
    return m;
  }, [adm]);
  // Students of one batch at one time slot: all of them, or the part given by s.parts[batch] = [from, to) (RPT splits a
  // batch across times when a time has fewer lab seats than students)
  const stuOf = (s, b) => { const all = studentsByBatch[b] || [], pt = s.parts && s.parts[b]; return pt ? all.slice(pt[0], pt[1]) : all; };
  const courses = useMemo(() => [...new Set(batches.map((b) => b.course).filter((c) => String(c || '').trim()))].sort(naturalCompare), [batches]);
  const phaseList = useMemo(() => parsePhaseList(settings.phaseText || ''), [settings.phaseText]);
  const phases = useMemo(() => groupPhases(batches, phaseList.phases), [batches, phaseList]);

  const papers = test.papers;
  const curP = papers.find((p) => p.id === cur) || null;
  const curIdx = papers.findIndex((p) => p.id === cur);
  // Any change to a paper means its seating must be prepared again
  const setPaper = (id, patch) => setTest((t) => ({ ...t, papers: t.papers.map((p) => (p.id === id ? { ...p, ...patch, done: false } : p)) }));
  const setSlot = (p, sid, patch) => setPaper(p.id, { slots: p.slots.map((s) => (sid === '*' || s.id === sid || (Array.isArray(sid) && sid.includes(s.id)) ? { ...s, ...patch } : s)) });

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
  // Capacity changed for one test only: p.cap (classrooms) and p.labCap (labs) map a room key to its seats for that test.
  // base keeps the original capacity, shown in the Excel file. Other tests always start from the original room list.
  const withCap = (list, m) => (m ? list.map((r) => (m[roomKey(r)] != null ? { ...r, base: Number(r.capacity), capacity: Math.max(0, Number(m[roomKey(r)]) || 0) } : r)) : list);
  const labsOf = (p) => withCap(labsOn, p && p.labCap);
  const roomsOf = (p) => withCap(roomsOn, p && p.cap);
  // Until the operator picks labs and buildings for a time slot, every lab and every building in use counts for the estimate
  // Rooms and labs kept for monitoring or due fee students are always part of the seating (and of every seat count),
  // even when their building is not ticked, so the counts match what the seating really uses.
  const pinKeys = new Set(['Monitoring', 'Due fee'].filter((f) => lists[f].length).flatMap((f) => settings.pins[f] || []));
  // Labs and buildings picked for the test's top batches always belong to its seats
  const topLabsOf = (p) => (p && (p.top || []).length ? p.topLabs || [] : []);
  const topBldsOf = (p) => (p && (p.top || []).length ? p.topBlds || [] : []);
  const slotLabs = (s, p) => (s.labs ? labsOf(p).filter((l) => s.labs.includes(roomKey(l)) || pinKeys.has(roomKey(l)) || topLabsOf(p).includes(roomKey(l))) : labsOf(p));
  // Labs used for offline students (Offline seating step): never a lab the same test uses for online students
  const slotOffLabs = (s, p) => { const on = new Set(slotLabs(s, p).map(roomKey)); return s.labs ? labsOf(p).filter((l) => (s.offLabs || []).includes(roomKey(l)) && !on.has(roomKey(l))) : []; };
  const slotBuildings = (s) => s.buildings || settings.buildingsOn;
  const slotRooms = (s, p) => roomsOf(p).filter((r) => (slotBuildings(s).includes(r.building) && !s.roomsOff.includes(roomKey(r))) || pinKeys.has(roomKey(r)) || topBldsOf(p).includes(r.building));
  const sumSeats = (list) => list.reduce((a, r) => a + Number(r.capacity || 0), 0);
  const countSlot = (p, s) => {
    const c = { online: 0, offline: 0 };
    for (const b of s.batches) for (const st of stuOf(s, b)) c[((s.mode === 'Both' && studentMode[st.Roll_Number]) || modeIn(p, s, b)) === 'Online' ? 'online' : 'offline']++;
    return c;
  };
  const sumSlot = (p, s) => seatSummary({ ...countSlot(p, s), labSeats: sumSeats(slotLabs(s, p)), roomSeats: sumSeats(slotRooms(s, p)) + sumSeats(slotOffLabs(s, p)) });
  const addUp = (list) => list.reduce((a, x) => { for (const k of Object.keys(x)) a[k] = (a[k] || 0) + x[k]; return a; }, {});
  const paperOk = (p) => !!(p.courses.length && p.slots.length && p.slots.every((s) => slotOk(s) && s.mode && s.phases.length && s.batches.length));
  // Test step done (everything but online or offline, which has its own step)
  const setupReady = (p) => ((test.kind || 'PRT') === 'RPT' ? !!(p.courses.length && p.rpt && p.rpt.phases.length) : !!(p.courses.length && p.slots.length && p.slots.every((s) => slotOk(s) && s.phases.length && s.batches.length)));

  // Special students: Monitoring and Due fee from the lists on step 1, Top batch from the Top ticks
  const flags = useMemo(() => {
    const f = {};
    // A student on both lists is seated once, with the due fee students
    for (const r of lists.Monitoring) f[r] = 'Monitoring';
    for (const r of lists['Due fee']) f[r] = 'Due fee';
    return f;
  }, [lists]);
  // Top batch students of one test (top ticks belong to the test and are never saved); monitoring and due fee win
  const flagsOf = (p) => { const f = {}; for (const b of (p && p.top) || []) for (const x of studentsByBatch[b] || []) f[x.Roll_Number] = 'Top batch'; return { ...f, ...flags }; };
  const appearingRolls = new Set(papers.flatMap((p) => p.slots.flatMap((s) => s.batches.flatMap((b) => stuOf(s, b).map((x) => x.Roll_Number)))));
  const appearing = appearingRolls.size;
  const flagCount = (f) => [...appearingRolls].filter((r) => flags[r] === f).length;
  const curChoices = [...new Map([...labsOn.map((r) => ({ ...r, kind: 'Lab' })), ...roomsOn.map((r) => ({ ...r, kind: 'Room' }))].map((r) => [roomKey(r), r])).values()];
  const curKeys = new Set(curChoices.map(roomKey));
  const special = ['Monitoring', 'Due fee'].map((f) => ({ flag: f, n: flagCount(f), kept: (settings.pins[f] || []).some((k) => curKeys.has(k)) })).filter((x) => x.n);
  // One unit per test: all its papers are seated together, in one seating, on the same labs and rooms.
  // Tests are grouped by time: tests that overlap share the labs and rooms.
  const allSlots = papers.flatMap((p, i) => timeGroupsOf(p).map((used, gi, all) => {
    const w = { start: hhmm(Math.min(...used.map((x) => mins(x.start)))), end: hhmm(Math.max(...used.map((x) => mins(x.two ? x.s2.end : x.end)))) };
    const shifts = used.flatMap(shiftsOfSlot);
    const labs = new Map(), rooms = new Map();
    for (const x of used) { for (const l of slotLabs(x, p)) labs.set(roomKey(l), l); for (const r of slotRooms(x, p)) rooms.set(roomKey(r), r); for (const l of slotOffLabs(x, p)) rooms.set(roomKey(l), l); }
    return {
      p, s: used[0], used, w, shifts, name: `${(p.name || '').trim() || `Test ${i + 1}`}${all.length > 1 ? ` · ${used.map((x) => slotName(p, x)).join(' + ')}` : ''}`,
      c: addUp(used.map((x) => countSlot(p, x)).concat([{ online: 0, offline: 0 }])),
      labs: [...labs.values()], rooms: [...rooms.values()],
      labsSet: used.some((x) => x.labs && x.labs.length), roomsSet: used.some((x) => x.buildings && x.buildings.length),
    };
  })).sort((a, b) => mins(a.w.start) - mins(b.w.start));
  const timeGroups = [];
  for (const x of allSlots) {
    const g = timeGroups[timeGroups.length - 1];
    if (g && g.items.some((y) => x.shifts.some((u) => y.shifts.some((v) => clash(u, v))))) { g.items.push(x); if (mins(x.w.end) > mins(g.end)) g.end = x.w.end; } else timeGroups.push({ start: x.w.start, end: x.w.end, items: [x] });
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
  const plan = {};
  for (const x of allSlots) {
    const taken = { labs: {}, rooms: {} };
    for (const y of allSlots) {
      if (y === x) break;
      if (!x.shifts.some((a) => y.shifts.some((b) => clash(a, b)))) continue;
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
    const one = { taken, labs: lab.rooms, rooms: room.rooms, offNeed, c, roomList: x.rooms, first: x.s.id, ids: x.used.map((y) => y.id), title: `${x.used.map((y) => `${slotName(x.p, y)} (${slotLabel(y)})`).join(' + ')} · ${pl(x.used.reduce((a, y) => a + y.batches.length, 0), 'batch')}` };
    for (const y of x.used) plan[y.id] = one;
  }
  const planOf = (s) => plan[s.id] || { taken: { labs: {}, rooms: {} }, labs: [], rooms: [], offNeed: 0, c: { online: 0, offline: 0 }, roomList: [] };
  // Capacity changes are kept on the test (p.cap for classrooms, p.labCap for labs), never in the room list, so the next
  // test starts from the original capacity. A room set back to its original capacity is dropped from the map.
  const setCaps = (pid, field, list, fn) => setTest((t) => ({ ...t, papers: t.papers.map((p) => {
    if (p.id !== pid) return p;
    const m = { ...(p[field] || {}) };
    fn(m, (r) => (m[roomKey(r)] != null ? Number(m[roomKey(r)]) : Number(r.capacity)));
    for (const r of list) if (m[roomKey(r)] != null && Number(m[roomKey(r)]) === Number(r.capacity)) delete m[roomKey(r)];
    return { ...p, [field]: m, done: false };
  }) }));
  // Percent: +10 means 10% more than the original capacity (-10 means 10% less). Seats: added to what each room has now,
  // in every room given (the rooms in use in that building).
  const pctOf = (r, v) => Math.max(0, Math.round((Number(r.capacity) * (100 + v)) / 100));
  const adjustCap = (p, b, how, v, keys) => {
    const list = roomsOn.filter((r) => r.building === b);
    const use = keys ? list.filter((r) => keys.includes(roomKey(r))) : list;
    setCaps(p.id, 'cap', roomsOn, (m, cur) => {
      if (how === 'reset') { for (const r of list) delete m[roomKey(r)]; return; }
      // Seats: the same number added to every room in use (like the percent, which changes every room)
      use.forEach((r) => { m[roomKey(r)] = how === 'pct' ? pctOf(r, v) : Math.max(0, cur(r) + Math.trunc(v)); });
    });
  };
  const setRoomCap = (p, k, v) => setCaps(p.id, 'cap', roomsOn, (m) => { m[k] = Math.max(0, Number(v) || 0); });
  // Lab seats for one test: every lab by a percentage or by a number of seats per lab, or one lab by hand
  const adjustLabs = (p, how, v) => setCaps(p.id, 'labCap', labsOn, (m, cur) => {
    if (how === 'reset') { for (const k of Object.keys(m)) delete m[k]; return; }
    for (const l of labsOn) m[roomKey(l)] = how === 'pct' ? pctOf(l, v) : Math.max(0, cur(l) + Math.trunc(v));
  });
  const setLabCap = (p, k, v) => setCaps(p.id, 'labCap', labsOn, (m) => { m[k] = Math.max(0, Number(v) || 0); });
  const labSeats = sumSeats(labsOn);
  const labSeatsOf = (p) => sumSeats(labsOf(p));

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
    } catch (e) { setError(`Could not read “${file.name}”: ${e.message}`); }
    setBusy('');
  }
  async function removeAdmission() {
    if (!(await askConfirm('Remove the admission list from this page? Your rooms and settings stay.', 'Remove', true))) return;
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
        return { ...s, rooms, labs: newLabs, roomFile: file.name, buildingOrder: order, buildingsOn: [...new Set([...s.buildingsOn, ...order.filter((b) => !s.buildingOrder.includes(b))])] };
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
  function autoPick(p, used) {
    const gc = [addUp(used.map((s) => countSlot(p, s)).concat([{ online: 0, offline: 0 }]))];
    const c = { online: Math.max(0, ...gc.map((x) => x.online)) };
    const taken = { labs: {}, rooms: {} };
    for (const s of used) for (const kind of ['labs', 'rooms']) for (const [k, v] of Object.entries(planOf(s).taken[kind])) taken[kind][k] = Math.max(taken[kind][k] || 0, v.filled);
    const free = (r, kind) => Math.max(0, Number(r.capacity) - (taken[kind][roomKey(r)] || 0));
    const rank = (b) => { const i = settings.buildingOrder.indexOf(b); return i < 0 ? 1e9 : i; };
    // Rooms and labs kept for monitoring or due fee students are used anyway: count them first
    const labs = [];
    const labsOnP = labsOf(p), roomsOnP = roomsOf(p);
    let labSeats = labsOnP.filter((l) => pinKeys.has(roomKey(l))).reduce((a, l) => a + free(l, 'labs'), 0);
    for (const l of labsOnP.map((r, i) => ({ r, i })).sort((a, b) => rank(a.r.building) - rank(b.r.building) || a.i - b.i).map((x) => x.r)) {
      if (labSeats >= c.online) break;
      if (free(l, 'labs') > 0 && !pinKeys.has(roomKey(l))) { labs.push(roomKey(l)); labSeats += free(l, 'labs'); }
    }
    const offNeed = Math.max(0, ...gc.map((x) => x.offline + Math.max(0, x.online - labSeats)));
    const buildings = [];
    let roomSeats = roomsOnP.filter((r) => pinKeys.has(roomKey(r))).reduce((a, r) => a + free(r, 'rooms'), 0);
    for (const b of settings.buildingOrder.filter((x) => settings.buildingsOn.includes(x))) {
      if (roomSeats >= offNeed) break;
      const seats = roomsOnP.filter((r) => r.building === b && !pinKeys.has(roomKey(r))).reduce((a, r) => a + free(r, 'rooms'), 0);
      if (seats > 0) { buildings.push(b); roomSeats += seats; }
    }
    return { labs, buildings };
  }
  function startPaper(p) {
    const patch = {};
    for (const g of timeGroupsOf(p)) {
      const labs = [...new Set(g.flatMap((s) => s.labs || []))], buildings = [...new Set(g.flatMap((s) => s.buildings || []))];
      const online = g.reduce((a, s) => a + countSlot(p, s).online, 0);
      const pick = (online && !labs.length) || !buildings.length ? autoPick(p, g) : null;
      const sel = { labs: online && !labs.length ? pick.labs : labs, buildings: buildings.length ? buildings : pick.buildings };
      for (const s of g) patch[s.id] = sel;
    }
    const np = { ...p, done: true, slots: p.slots.map((s) => (patch[s.id] ? { ...s, ...patch[s.id] } : { ...s, labs: s.labs || [], buildings: s.buildings || [] })) };
    const list = test.papers.map((x) => (x.id === p.id ? np : x));
    setCur(p.id);
    setTest((t) => ({ ...t, papers: list }));
    // PRT: labs and buildings are picked automatically, then the operator checks the Online and Offline seating steps
    // and moves on with Prepare seating. RPT seats at once.
    if ((test.kind || 'PRT') === 'RPT') { run(seed, list); return; }
    setStep(np.slots.some((s) => countSlot(np, s).online > 0) ? 'labs' : 'rooms');
  }

  // One seating per test: the students of all its papers are seated together (mixed by the shuffle rules), in one
  // allocation, then shown and exported as an Online seating (labs) and an Offline seating (classrooms).
  // A test that overlaps an earlier test (or starts within GAP minutes of it) uses the seats that test left free.
  const shiftsOfTest = (p, slots) => slots.flatMap((s) => {
    const n = slotName(p, s);
    return s.two ? [{ label: `${n} shift 1`, paper: n, start: s.start, end: s.end }, { label: `${n} shift 2`, paper: n, start: s.s2.start, end: s.s2.end }] : [{ label: n, paper: n, start: s.start, end: s.end }];
  });
  // Students and seating rules for one test (the papers in `used` sit together). Labs used for offline students join the
  // classrooms; labs and buildings picked for top batches are kept for them (seated there first).
  function unitInput(p, used, taken, seedN, cbOf = {}) {
    const batchMode = {}, pairs = [];
    for (const s of used) for (const b of s.batches) {
      if (!batchMap[b]) continue;
      pairs.push([s, b]); batchMode[b] = modeIn(p, s, b);
    }
    const shiftLabel = (x) => `${fmtTime(x.start)}${x.end ? ` to ${fmtTime(x.end)}` : ''}`;
    const st = pairs.flatMap(([s, b]) => stuOf(s, b).map((x) => ({ ...x, Paper: slotName(p, s), _shifts: (s.two ? [s, s.s2] : [s]).map(shiftLabel) })));
    const labKeys = new Set(used.flatMap((s) => slotLabs(s, p).map(roomKey)));
    const roomKeys = new Set(used.flatMap((s) => slotRooms(s, p).map(roomKey)));
    const offKeys = new Set(used.flatMap((s) => slotOffLabs(s, p).map(roomKey)));
    const studentModeAll = used.some((s) => s.mode === 'Both') ? studentMode : {};
    const topKeys = [...topLabsOf(p), ...settings.rooms.filter((r) => topBldsOf(p).includes(r.building)).map(roomKey)];
    const pins = Object.fromEntries(Object.entries(settings.pins || {}).filter(([f]) => f !== 'Top batch' && (!(f in lists) || lists[f].length)));
    if (topKeys.length) pins['Top batch'] = topKeys;
    const cfg = {
      rooms: [...withCap(settings.rooms, p.cap).map((r) => ({ ...r, enabled: !!r.enabled && roomKeys.has(roomKey(r)) })),
        ...withCap(settings.labs, p.labCap).filter((l) => offKeys.has(roomKey(l))).map((l) => ({ ...l, enabled: !!l.enabled }))],
      labs: withCap(settings.labs, p.labCap).map((l) => ({ ...l, enabled: !!l.enabled && labKeys.has(roomKey(l)) && !offKeys.has(roomKey(l)) })),
      buildingOrder: settings.buildingOrder, girlsOk: settings.girlsOk,
      batchMode, studentMode: studentModeAll, flags: flagsOf(p), pins, cbOf,
      seatOrder: settings.seatOrder || 'mixed', seatOrder2: settings.seatOrder2 || '', apart: !!settings.apart, seed: seedN, taken,
    };
    return [st, cfg];
  }
  // Girls on the Online and Offline seating steps: a trial seating of the current choice, counted per room and lab
  function girlsOf(p, s) {
    const g = timeGroupsOf(p).find((gr) => gr.some((x) => x.id === s.id));
    if (!g || !g.some((x) => x.batches.length)) return null;
    const taken = {};
    const t = planOf(s).taken;
    for (const kind of ['labs', 'rooms']) for (const [k, v] of Object.entries(t[kind])) taken[k] = Math.max(taken[k] || 0, v.filled);
    const r = allocate(...unitInput(p, g.map((x) => ({ ...x, labs: x.labs || [], buildings: x.buildings || [] })), taken, seed));
    const inRoom = {}, filled = {};
    for (const room of r.usedRooms) { inRoom[room.kind + roomKey(room)] = room.occupants.filter(isGirl).length; filled[room.kind + roomKey(room)] = room.used; }
    const count = (m) => r.assigned.filter((x) => isGirl(x) && (x.Mode === 'Online') === (m === 'Online')).length;
    return { online: count('Online'), offline: count('Offline'), left: r.unseated.filter(isGirl).length, inRoom, filled };
  }
  function run(newSeed, list) {
    const cbOf = {};
    for (const g of settings.cbGroups) for (const b of parseRolls(g.batches)) if (g.name.trim()) cbOf[b] = g.name.trim();
    const units = list.filter((p) => p.done && p.slots.some((s) => s.batches.length))
      .flatMap((p) => timeGroupsOf(p).map((used, gi, all) => ({ p, used, unit: { id: `${p.id}-${gi}`, name: `${(p.name || '').trim() || `Test ${list.indexOf(p) + 1}`}${all.length > 1 ? ` · ${used.map((s) => slotName(p, s)).join(' + ')}` : ''}`, mode: '', two: false, shifts: shiftsOfTest(p, used) } })))
      .sort((a, b) => mins(windowOf(a.unit).start) - mins(windowOf(b.unit).start));
    const done = [], out = [];
    units.forEach(({ p, used, unit }, i) => {
      const taken = {};
      const over = done.filter((d) => overlaps(d.unit, unit));
      for (const d of over) for (const room of d.result.usedRooms) {
        const k = roomKey(room);
        taken[k] = Math.max(taken[k] || 0, room.taken + room.used);
      }
      const result = allocate(...unitInput(p, used, taken, newSeed + i, cbOf));
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
    try { localStorage.setItem(LAST, JSON.stringify({ title: test.title, kind: test.kind, papers: list.map((p) => ({ ...p, done: false, top: [], topLabs: [], topBlds: [] })) })); } catch { /* ignore */ }
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
  const canOpen = (id) => (id === 'students' ? true : id === 'test' ? !!adm : id === 'mode' ? !!adm && !!test.date && papers.some(setupReady) : id === 'result' ? !!results : setupOk);
  const doneMark = (id) => (id === 'students' ? !!adm : id === 'test' ? !!test.date && papers.some(setupReady) : id === 'mode' ? papers.some((p) => p.done) : false);
  const labsNeeded = curP ? curP.slots.some((s) => countSlot(curP, s).online > 0) : false;
  const offlineNeeded = curP ? curP.slots.some((s) => s.batches.length && planOf(s).offNeed > 0) : false;

  const landing = !started && !adm;
  const shown = STEPS.filter((s) => s.id !== 'labs' || labsNeeded || step === 'labs');
  const at = Math.max(0, shown.findIndex((s) => s.id === step));
  const showSide = !landing && !!adm && step !== 'students';
  return (
    <div className="app">
      <header className="top">
        <div className="top-in">
          <div className="logo">
            <img src="/brand/matrix-logo-green.png" alt="Matrix Academy" className="logo-img" />
            <span className="sep" aria-hidden="true" />
            <div className="brand"><b>Test Seating</b>{!landing && <span className="chip">{test.date ? `${fmtDateDMY(test.date)} · ${pl(papers.reduce((a, p) => a + p.slots.length, 0), 'paper')}` : 'No test date yet'}</span>}</div>
          </div>
          {!landing && <nav className="stepper" aria-label="Steps">
            {shown.map((s, i) => {
              const done = doneMark(s.id) && step !== s.id;
              return (
                <button key={s.id} type="button" disabled={!canOpen(s.id)} aria-current={step === s.id ? 'step' : undefined}
                  className={`st ${step === s.id ? 'on' : ''} ${done ? 'done' : ''}`} onClick={() => setStep(s.id)}>
                  <i>{done ? '✓' : i + 1}</i><span>{s.label}</span>
                </button>
              );
            })}
          </nav>}
        </div>
        {!landing && <div className="progress" role="progressbar" aria-label="Progress" aria-valuemin={1} aria-valuemax={shown.length} aria-valuenow={at + 1}><span style={{ width: `${((at + 1) / shown.length) * 100}%` }} /></div>}
      </header>

      {landing && <Landing go={() => { setStarted(true); setStep('students'); }} />}
      {!landing && <div className={`shell ${showSide ? 'with-side' : ''}`}>
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
            <FileCard title="Room and lab list" hint="Sheets Room_Capacity and LAB_Capacity · saved in this browser, upload again only when rooms change"
              loaded={<><b>{settings.roomFile || 'Built-in list'}</b><span>{pl(settings.rooms.length, 'room')} · {pl(settings.labs.length, 'lab')} · {labSeats} lab seats</span></>}
              onFile={onRoomFile} template={() => downloadTemplate('rooms')} />
            <details className="fold">
              <summary>Monitoring and due fee students <span className="muted">· optional · {lists.Monitoring.length} monitoring, {lists['Due fee'].length} due fee</span></summary>
              <p className="muted">Upload or paste their roll numbers, then pick the rooms where they sit. They are seated in these rooms first; other students fill the free seats only after every other room is full. A student on both lists is seated once, with the due fee students.</p>
              {(() => { const dueSet = new Set(lists['Due fee']); const both = lists.Monitoring.filter((r) => dueSet.has(r)).length; return both ? <p className="small warnt">{pl(both, 'student')} on both lists: seated with the due fee students only.</p> : null; })()}
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
              <summary>Seating rules <span className="muted">· {SEAT_ORDERS[settings.seatOrder || 'mixed']}{!['mixed', 'roll'].includes(settings.seatOrder || 'mixed') && settings.seatOrder2 && settings.seatOrder2 !== settings.seatOrder ? `, then ${SEAT_ORDERS[settings.seatOrder2].toLowerCase()}` : ''}{settings.apart ? ' · brothers and sisters apart' : ''}</span></summary>
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

        {(step === 'test' || step === 'mode') && (
          <PaperStep key={step} part={step === 'mode' ? 'mode' : 'setup'} setupReady={setupReady} goMode={() => setStep('mode')} goTest={() => setStep('test')} test={test} setTest={setTest} setPaper={setPaper} setSlot={setSlot} courses={courses} phases={phases} ownerOf={ownerOf} slotOf={slotOf}
            batchMap={batchMap} phaseBatches={phaseBatches} studentsByBatch={studentsByBatch} studentMode={studentMode} setStudentMode={setStudentMode}
            sumSlot={sumSlot} paperOk={paperOk} start={startPaper} back={() => setStep(step === 'mode' ? 'test' : 'students')} labSeatsOf={labSeatsOf} labsOf={labsOf} adjustLabs={adjustLabs} setLabCap={setLabCap}
            blds={settings.buildingOrder.filter((b) => roomsOn.some((r) => r.building === b))} />
        )}

        {step === 'labs' && curP && (
          <LabStep p={curP} idx={curIdx} labsOn={labsOf(curP)} order={settings.buildingOrder} setSlot={setSlot} countSlot={countSlot} planOf={planOf} girlsOf={(s) => girlsOf(curP, s)} topLabs={topLabsOf(curP)}
            back={() => setStep('mode')} next={() => (offlineNeeded ? setStep('rooms') : prepare())} label={offlineNeeded ? 'Continue' : 'Prepare seating'} />
        )}

        {step === 'rooms' && curP && (
          <RoomStep p={curP} idx={curIdx} settings={settings} setSettings={setSettings} setSlot={setSlot} planOf={planOf}
            roomsOn={roomsOf(curP)} labsOn={labsOf(curP)} girlsOf={(s) => girlsOf(curP, s)} topBlds={topBldsOf(curP)} onRoomFile={onRoomFile} adjustCap={(...a) => adjustCap(curP, ...a)} setRoomCap={(...a) => setRoomCap(curP, ...a)}
            back={() => setStep(labsNeeded ? 'labs' : 'mode')} next={prepare} label="Prepare seating" />
        )}

        {step === 'result' && results && (
          <ResultStep results={results} view={view} setView={setView} metaFor={metaFor} multi={results.length > 1} order={settings.buildingOrder}
            reshuffle={async () => { if (await askConfirm('Reshuffle every seat? The current seating is replaced.', 'Reshuffle')) run(Math.floor(Math.random() * 1e9), test.papers); }}
            download={download} downloadBoard={downloadBoard} back={() => setStep('rooms')} />
        )}
        {step === 'result' && results && papers.some((p) => !p.done) && (
          <div className="alert info"><span>{papers.map((p, i) => (p.done ? '' : `Test ${i + 1}`)).filter(Boolean).join(', ')}: seating not made yet.</span><button className="btn sm" onClick={() => setStep('test')}>Set up next paper</button></div>
        )}
      </main>
      {showSide && (
        <aside className={`side ${sumOpen ? 'open' : ''}`} aria-label="Summary">
          <button type="button" className="side-toggle" aria-expanded={sumOpen} onClick={() => setSumOpen(!sumOpen)}>
            <b>Summary</b><span>{Number(appearing).toLocaleString('en-IN')} appearing · {pl(adm.students.length, 'student')}</span><i aria-hidden="true">{sumOpen ? '▾' : '▴'}</i>
          </button>
          <div className="side-body">
            <SummaryBar adm={adm} papers={papers} appearing={appearing} groups={timeGroups} special={special}
              results={step === 'result' ? results : null} goRules={() => setStep('students')} />
          </div>
        </aside>
      )}
      </div>}
      <footer className="foot">Matrix Academy, Sikar · Operations · Student data stays in this browser</footer>
      <ConfirmHost />
    </div>
  );
}

// Start screen: one headline, one line, one button
function Landing({ go }) {
  return (
    <section className="landing">
      <span className="kicker">Matrix Academy · Operations</span>
      <h1 className="hero">Test seating, ready in minutes.</h1>
      <p className="lead">Upload the admission list, set up the test, and download the finished seating sheet. Student data never leaves this browser.</p>
      <button className="btn primary lg" onClick={go} autoFocus>Create Seating Plan</button>
      <ol className="land-steps" aria-label="How it works">
        {['Admission list', 'Test & papers', 'Online / Offline', 'Labs and rooms', 'Download'].map((x, i) => <li key={x}><i>{i + 1}</i>{x}</li>)}
      </ol>
    </section>
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
      <div><span className="kicker">First time here?</span><h2>Seating sheet in four steps</h2></div>
      <ol>
        <li><i>1</i><b>Admission list</b><span>Upload the admission list. Optional: monitoring and due fee students with their rooms, and seating rules.</span></li>
        <li><i>2</i><b>Test and papers</b><span>Date, courses, then each paper: name, shifts, phases and batches.</span></li>
        <li><i>3</i><b>Online / Offline</b><span>Online or offline for each paper or batch, and the top batches. Create seating: labs and buildings are picked for you.</span></li>
        <li><i>4</i><b>Seating</b><span>Check the labs and rooms, then download the Excel file.</span></li>
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
      {next && <button className="btn primary" disabled={!nextOk} onClick={next}>{nextLabel}</button>}
    </div>
  );
}

function PaperStep({ part = 'setup', setupReady, goMode, goTest, test, setTest, setPaper, setSlot, courses, phases, ownerOf, slotOf, phaseBatches, studentsByBatch, studentMode, setStudentMode, sumSlot, paperOk, start, back, labSeatsOf, labsOf, adjustLabs, setLabCap, blds }) {
  const topOf = (p) => new Set(p.top || []);
  const setTopOf = (p) => (fn) => setPaper(p.id, { top: [...fn(new Set(p.top || []))] });
  const addPaper = () => setTest({ ...test, papers: [...test.papers, newPaper()] });
  const kind = test.kind || 'PRT';
  const setKind = async (k) => {
    if (k === kind) return;
    if (test.papers.some((p) => p.courses.length) && !(await askConfirm('Switch the test type? The tests set up below are cleared.', 'Switch', true))) return;
    setTest((t) => ({ ...t, kind: k, papers: [newPaper()] }));
  };
  const cardOf = (p, i) => (kind === 'RPT' ? (
    <RptCard key={p.id} part={part} p={p} i={i} n={test.papers.length} setPaper={setPaper} remove={() => removePaper(p.id)} courses={courses} phases={phases} sumSlot={sumSlot} ok={paperOk(p)} start={() => start(p)} top={topOf(p)} setTop={setTopOf(p)} blds={blds}
      labSeats={labSeatsOf(p)} labsOn={labsOf(p)} adjustLabs={(how, v) => adjustLabs(p, how, v)} setLabCap={(k, v) => setLabCap(p, k, v)} />
  ) : (
    <PaperCard key={p.id} part={part} p={p} i={i} n={test.papers.length} setPaper={setPaper} setSlot={setSlot} remove={() => removePaper(p.id)}
      courses={courses} phases={phases} ownerOf={ownerOf} slotOf={slotOf} phaseBatches={phaseBatches}
      studentsByBatch={studentsByBatch} studentMode={studentMode} setStudentMode={setStudentMode}
      top={topOf(p)} setTop={setTopOf(p)} sumSlot={sumSlot} ok={paperOk(p)} start={() => start(p)} labs={labsOf(p)} blds={blds} />
  ));
  if (part === 'mode') {
    const ready = test.papers.map((p, i) => [p, i]).filter(([p]) => setupReady(p));
    const notYet = test.papers.map((p, i) => [p, i]).filter(([p]) => !setupReady(p));
    return (
      <section>
        <Head kicker="Step 3" title="Online or offline?" lead={kind === 'RPT' ? 'Choose online or offline, then the times and shifts. Tick ★ on a batch to make it a top batch. Create seating picks the labs and buildings for you.' : 'For each paper choose Online, Offline or Both. With Both, mark each batch, or single students. Tick Top for batches that sit first. Create seating picks the labs and buildings for you.'} />
        <div className="papers">{ready.map(([p, i]) => cardOf(p, i))}</div>
        {notYet.length > 0 && <div className="alert info"><span>{notYet.map(([p, i]) => `Test ${i + 1}`).join(', ')} {notYet.length > 1 ? 'are' : 'is'} not set up yet.</span><button className="btn sm" onClick={goTest}>Finish on Test & papers</button></div>}
        <Nav back={back} />
      </section>
    );
  }
  const removePaper = (id) => setTest({ ...test, papers: test.papers.filter((p) => p.id !== id) });
  return (
    <section>
      <Head kicker="Step 2" title="Test and papers" lead="Pick the date and the courses. Then set up each paper: name, 1 or 2 shifts, phases and batches. Online or offline comes next." />
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
      <div className="typebar">
        <b>Test type</b>
        <div className="seg" role="group" aria-label="Test type">
          <button className={kind === 'PRT' ? 'on' : ''} aria-pressed={kind === 'PRT'} onClick={() => setKind('PRT')}>PRT · set up each paper</button>
          <button className={kind === 'RPT' ? 'on' : ''} aria-pressed={kind === 'RPT'} onClick={() => setKind('RPT')}>RPT · spread over time slots automatically</button>
        </div>
      </div>
      {!test.date && <p className="muted small">Pick the test date first.</p>}
      {test.date && (
        <>
          <div className="papers">{test.papers.map(cardOf)}</div>
          <button type="button" className="addslot" onClick={addPaper}>{Ico.plus}Add another test (other courses)</button>
        </>
      )}
      <Nav back={back} next={goMode} nextOk={!!test.date && test.papers.some(setupReady)}
        hint={!test.date ? 'Pick the test date first' : test.papers.some(setupReady) ? '' : kind === 'RPT' ? 'Pick a course and its phases' : 'Each paper needs a time, phases and batches'} />
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

// ---------- RPT: one course, its phases, online or offline, then the batches are spread over the time slots automatically
// 1 shift: four times (6-9, 10-1, 2-5, 5:30-8:30). 2 shifts: two times, each sat in two shifts with the same seats.
const rptTimes = (two) => (two
  ? [{ start: '06:00', end: '09:00', s2: { start: '14:00', end: '17:00' } }, { start: '10:00', end: '13:00', s2: { start: '17:30', end: '20:30' } }]
  : [{ start: '06:00', end: '09:00' }, { start: '10:00', end: '13:00' }, { start: '14:00', end: '17:00' }, { start: '17:30', end: '20:30' }]);
const rptDefault = () => ({ phases: [], mode: '', two: false, times: rptTimes(false), noGirlsLast: true });
// The time that ends last (its last shift ends latest): no girls there when "No girls in the last time" is on
export const lastTimeOf = (r) => {
  const endOf = (t) => mins(r.two && t.s2 ? t.s2.end : t.end);
  return r.times.length > 1 ? r.times.reduce((best, t, i) => (endOf(t) > endOf(r.times[best]) ? i : best), 0) : -1;
};
// Which batches sit at each time: 'labs' (default for online) fills each time up to the lab seats and moves the rest to the
// next time, splitting a batch where needed; 'equal' keeps batches whole with about the same number at each time.
// The operator's own changes (r.pins: batches put at a time, r.outs: batches taken out of a time) always win.
export function rptPlan(r, phases, cap) {
  const batches = phases.filter((g) => r.phases.includes(g.key)).flatMap((g) => g.batches);
  const fill = r.mode === 'Offline' ? 'equal' : (r.fill || 'labs');
  return rptFill({ batches, n: r.times.length, cap, fill, noGirlsAt: r.noGirlsLast === false ? -1 : lastTimeOf(r), pins: r.pins || {}, outs: r.outs || {} });
}
export function rptSlots(p, r, phases, cap) {
  const { lists } = rptPlan(r, phases, cap);
  return r.times.map((t, i) => {
    const L = lists[i];
    if (!L || !L.batches.length) return null;
    const id = `${p.id}-t${i}`, old = p.slots.find((x) => x.id === id) || {};
    return { ...newSlot(t.start, t.end), ...old, id, name: `Time ${i + 1}`, mode: r.mode, phases: r.phases, start: t.start, end: t.end, two: !!r.two,
      s2: r.two ? (t.s2 || { start: hhmm(mins(t.end) + 120), end: hhmm(mins(t.end) + 120 + mins(t.end) - mins(t.start)) }) : (old.s2 || newSlot(t.start, t.end).s2),
      batches: L.batches, parts: L.parts, online: r.mode === 'Both' ? L.batches : [] };
  }).filter(Boolean);
}
// pins and outs are kept by time number: removing time k drops its entry and moves the later ones up
const dropTime = (m, k) => Object.fromEntries(Object.entries(m || {}).filter(([i]) => Number(i) !== k).map(([i, v]) => [Number(i) > k ? Number(i) - 1 : Number(i), v]));

function RptCard({ part = 'setup', p, i, n, setPaper, remove, courses, phases, sumSlot, ok, start, top, setTop, labSeats, labsOn, adjustLabs, setLabCap, blds }) {
  const r = p.rpt || rptDefault();
  const apply = (patch, extra = {}) => { const nr = { ...r, ...patch }; const np = { ...p, ...extra, rpt: nr }; setPaper(p.id, { ...extra, rpt: nr, slots: rptSlots(np, nr, phases, labSeats) }); };
  // Lab seats changed (or a new admission list): spread the batches again
  useEffect(() => {
    if (!r.phases.length || !r.mode) return;
    const ns = rptSlots(p, r, phases, labSeats);
    if (JSON.stringify(ns) !== JSON.stringify(p.slots)) setPaper(p.id, { slots: ns });
  }, [labSeats, phases]); // eslint-disable-line react-hooks/exhaustive-deps
  const mine = phases.filter((g) => p.courses.includes(g.course));
  const toggleCourse = (c) => {
    const cs = toggleIn(p.courses, c);
    const keep = r.phases.filter((k) => cs.includes(k.split('|')[0]));
    const added = cs.includes(c) ? phases.filter((g) => g.course === c).map((g) => g.key) : [];
    apply({ phases: [...keep, ...added] }, { courses: cs });
  };
  const setTime = (k, v) => apply({ times: r.times.map((t, j) => (j === k ? { ...t, ...v } : t)) });
  const addTime = () => {
    const last = r.times[r.times.length - 1];
    const st = hhmm(mins(r.two ? last.s2.end : last.end) + 30), len = mins(last.end) - mins(last.start);
    apply({ times: [...r.times, r.two ? { start: st, end: hhmm(mins(st) + len), s2: { start: hhmm(mins(st) + len + 60), end: hhmm(mins(st) + 2 * len + 60) } } : { start: st, end: hhmm(mins(st) + len) }] });
  };
  const removeTime = (k) => apply({ times: r.times.filter((_, j) => j !== k), pins: dropTime(r.pins, k), outs: dropTime(r.outs, k) });
  // Take a batch out of a time (it moves on to the next times), or put a whole batch at a time
  const takeOut = (k, b) => {
    const pins = r.pins || {}, outs = r.outs || {};
    if ((pins[k] || []).includes(b)) apply({ pins: { ...pins, [k]: pins[k].filter((x) => x !== b) } });
    else apply({ outs: { ...outs, [k]: [...(outs[k] || []), b] } });
  };
  const putAt = (k, b) => {
    const pins = Object.fromEntries(Object.entries(r.pins || {}).map(([j, v]) => [j, v.filter((x) => x !== b)]));
    const outs = r.outs || {};
    apply({ pins: { ...pins, [k]: [...(pins[k] || []), b] }, outs: { ...outs, [k]: (outs[k] || []).filter((x) => x !== b) } });
  };
  const manual = Object.values(r.pins || {}).some((v) => v.length) || Object.values(r.outs || {}).some((v) => v.length);
  const toggleTop = (b) => setTop((t) => { const x = new Set(t); x.has(b) ? x.delete(b) : x.add(b); return x; });
  const info = Object.fromEntries(mine.flatMap((g) => g.batches).map((b) => [b.batch, b]));
  const anyGirls = mine.some((g) => g.batches.some((b) => b.girls > 0));
  const last = r.noGirlsLast === false ? -1 : lastTimeOf(r);
  const byLabs = r.mode !== 'Offline' && (r.fill || 'labs') === 'labs';
  const plan = r.phases.length && r.mode ? rptPlan(r, phases, labSeats) : { lists: [], left: 0, leftBatches: [] };
  const partN = (s, b) => (s.parts && s.parts[b] ? s.parts[b][1] - s.parts[b][0] : info[b]?.count || 0);
  const count = (s) => s.batches.reduce((a, b) => a + partN(s, b), 0);
  const sel = mine.filter((g) => r.phases.includes(g.key));
  return (
    <div className="card paper">
      <div className="paper-top">
        <span className="slot-n">{i + 1}</span>
        <b className="grow">{n > 1 ? `RPT test ${i + 1}` : 'RPT test'}{p.courses.length ? ` · ${p.courses.join(', ')}` : ''}</b>
        {part === 'mode' ? <span className="muted small">{pl(sel.length, 'phase')}</span> : <button type="button" className="iconbtn danger" disabled={n === 1} onClick={remove} aria-label={`Remove test ${i + 1}`} title="Remove this test">{Ico.bin}</button>}
      </div>
      {part === 'setup' && <div className="q">
        <b className="q-l">1. Course</b>
        <div className="chips">{courses.map((c) => <button key={c} type="button" className={`chip-btn ${p.courses.includes(c) ? 'on' : ''}`} onClick={() => toggleCourse(c)}>{c}</button>)}</div>
      </div>}
      {part === 'setup' && p.courses.length > 0 && (
        <div className="q">
          <b className="q-l">2. Phases in this test</b>
          <p className="muted small">{r.phases.length} of {mine.length} phases · <button className="link" onClick={() => apply({ phases: mine.map((g) => g.key) })}>Select all</button> · <button className="link" onClick={() => apply({ phases: [] })}>Clear</button></p>
          <div className="chips">
            {mine.map((g) => (
              <button key={g.key} type="button" className={`chip-btn ${r.phases.includes(g.key) ? 'on' : ''}`} title={g.batches.map((b) => b.batch).join(', ')} onClick={() => apply({ phases: toggleIn(r.phases, g.key) })}>
                {p.courses.length > 1 ? `${g.course} · ` : ''}{g.phase} <small>{g.count}</small>
              </button>
            ))}
          </div>
        </div>
      )}
      {part === 'mode' && r.phases.length > 0 && (
        <div className="q">
          <b className="q-l">3. Online or offline?</b>
          <div className="seg" role="group" aria-label="Online or offline">
            {[['Online', 'Online'], ['Offline', 'Offline'], ['Both', 'Both (labs first, the rest offline)']].map(([m, l]) => (
              <button key={m} className={r.mode === m ? 'on' : ''} onClick={() => apply({ mode: m })}>{l}</button>
            ))}
          </div>
        </div>
      )}
      {part === 'mode' && r.phases.length > 0 && r.mode && r.mode !== 'Offline' && (
        <LabSeats labSeats={labSeats} labsOn={labsOn} adjustLabs={adjustLabs} setLabCap={setLabCap} />
      )}
      {part === 'mode' && r.phases.length > 0 && r.mode && (
        <div className="q">
          <b className="q-l">{r.mode !== 'Offline' ? '5' : '4'}. Times and shifts</b>
          <div className="seg sm" role="group" aria-label="Shifts">
            <button className={!r.two ? 'on' : ''} onClick={() => apply({ two: false, times: rptTimes(false), pins: {}, outs: {} })}>1 shift (4 times)</button>
            <button className={r.two ? 'on' : ''} onClick={() => apply({ two: true, times: rptTimes(true), pins: {}, outs: {} })}>2 shifts (same seat in both)</button>
          </div>
          {r.mode !== 'Offline' && (
            <div className="seg sm" role="group" aria-label="How to spread the batches">
              <button className={byLabs ? 'on' : ''} onClick={() => apply({ fill: 'labs' })}>Fill the lab seats, the rest to the next time</button>
              <button className={!byLabs ? 'on' : ''} onClick={() => apply({ fill: 'equal' })}>About equal students at each time</button>
            </div>
          )}
          <p className="muted small">{byLabs
            ? `Each time gets as many students as there are lab seats (${labSeats}), in phase and batch order. The students left over go to the next time; a batch is split between two times when it does not fit.`
            : 'Every batch is placed whole in one time, keeping the number of students at each time about equal.'}
            {' '}Take a batch out of a time with ×, or put a batch at a time with Add a batch.{r.mode !== 'Offline' ? ' Tick ★ on a batch to make it a top batch: its students take lab seats first.' : ''}</p>
          <label className="small"><input type="checkbox" checked={r.noGirlsLast !== false} onChange={(e) => apply({ noGirlsLast: e.target.checked })} /> No girls in the last time (batches with girls go to the earlier times)</label>
          {r.noGirlsLast !== false && !anyGirls && <p className="small bad">The admission list has no girls marked (no Gender values), so this rule cannot work.</p>}
          {manual && <p className="small">You moved batches by hand. <button className="link" onClick={() => apply({ pins: {}, outs: {} })}>Undo my changes</button></p>}
          {r.times.map((t, k) => {
            const s = p.slots.find((x) => x.id === `${p.id}-t${k}`);
            const over = byLabs && s ? count(s) - labSeats : 0;
            const here = new Set(s ? s.batches.filter((b) => !(s.parts && s.parts[b])) : []);
            return (
              <div className="rpt-time" key={k}>
                <TimeRow label={r.two ? `Time ${k + 1} · shift 1` : `Time ${k + 1}`} value={t} onChange={(v) => setTime(k, v)} onRemove={r.times.length > 1 ? () => removeTime(k) : null} />
                {r.two && <TimeRow label="Shift 2" value={t.s2} onChange={(v) => setTime(k, { s2: v })} />}
                {s ? (
                  <>
                    <p className="rpt-batches"><b>{pl(count(s), 'student')}</b>{byLabs ? ` of ${labSeats} lab seats` : ''} · {pl(s.batches.length, 'batch')}{k === last ? ' · no girls' : s.batches.some((b) => info[b]?.girls > 0) ? ' · has girls' : ''}</p>
                    <div className="chips">
                      {s.batches.map((b) => {
                        const pt = s.parts && s.parts[b];
                        const label = <>{top.has(b) ? '★ ' : ''}{b}{pt ? ` · ${pt[0] + 1} to ${pt[1]}` : ''} <small>{partN(s, b)}</small></>;
                        return (
                          <span key={b} className="chip-pair">
                            {r.mode === 'Offline'
                              ? <span className="chip-btn">{label}</span>
                              : <button type="button" className={`chip-btn ${top.has(b) ? 'on' : ''}`} title={top.has(b) ? 'Top batch: lab seats first' : 'Tick to make it a top batch'} onClick={() => toggleTop(b)}>{label}</button>}
                            <button type="button" className="chip-x" onClick={() => takeOut(k, b)} aria-label={`Take ${b} out of time ${k + 1}`} title="Take out of this time: its students go to the next time">×</button>
                          </span>
                        );
                      })}
                    </div>
                    {Object.keys(s.parts || {}).length > 0 && <p className="muted small">“1 to 10” means the first 10 students of the batch in roll-number order; the rest sit at another time.</p>}
                    {over > 0 && <p className="small bad">{pl(over, 'student')} more than the lab seats (batches added by hand). Add lab seats, or take a batch out.</p>}
                    <SeatBoard s={sumSlot(p, s)} mode={r.mode} />
                  </>
                ) : <p className="muted small">No batches at this time.</p>}
                <AddBatch groups={sel} skip={here} onAdd={(b) => putAt(k, b)} label={`Time ${k + 1}`} />
              </div>
            );
          })}
          {plan.left > 0 && (
            <p className="small bad">{pl(plan.left, 'student')} have no time yet: {plan.leftBatches.map((x) => `${x.batch} (${x.n})`).join(', ')}. Add another time, add lab seats{r.noGirlsLast !== false ? ', switch off “No girls in the last time”' : ''}, or put their batch at a time with Add a batch.</p>
          )}
          <button type="button" className="addslot sm" onClick={addTime}>{Ico.plus}Add another time</button>
        </div>
      )}
      {part === 'mode' && top.size > 0 && <TopSeats p={p} setPaper={setPaper} labs={labsOn} blds={blds} />}
      {part === 'mode' && ok && plan.left === 0 && (
        <div className="paper-cta">
          {p.done && <span className="pill">Seating ready ✓</span>}
          <button className="btn cta" onClick={start}>{Ico.spark}{p.done ? 'Change seating for this test' : 'Create seating for this test'}</button>
        </div>
      )}
    </div>
  );
}

// Put a whole batch at one time: pick the phase, then the batch
function AddBatch({ groups, skip, onAdd, label }) {
  const [g, setG] = useState('');
  const [b, setB] = useState('');
  const cur = groups.find((x) => x.key === g) || groups[0];
  if (!cur) return null;
  const list = cur.batches.filter((x) => !skip.has(x.batch));
  return (
    <div className="addbatch small">
      <span className="muted">Add a batch to {label}:</span>
      <select value={cur.key} onChange={(e) => { setG(e.target.value); setB(''); }} aria-label={`Phase for ${label}`}>{groups.map((x) => <option key={x.key} value={x.key}>{x.phase}</option>)}</select>
      <select value={b} onChange={(e) => setB(e.target.value)} aria-label={`Batch for ${label}`}>
        <option value="">Batch…</option>
        {list.map((x) => <option key={x.batch} value={x.batch}>{x.batch} ({x.count})</option>)}
      </select>
      <button type="button" className="btn ghost sm" disabled={!b} onClick={() => { onAdd(b); setB(''); }}>Add</button>
    </div>
  );
}

// Lab seats for an RPT test: the total at one time, and a way to raise (or lower) them: every lab by seats or %, or one lab
function LabSeats({ labSeats, labsOn, adjustLabs, setLabCap }) {
  const [num, setNum] = useState('');
  const [pct, setPct] = useState('');
  const base = labsOn.reduce((a, l) => a + Number(l.base ?? l.capacity), 0);
  return (
    <div className="q">
      <b className="q-l">4. Lab seats</b>
      <p className="small"><b>{labSeats} lab seats</b> at one time in {pl(labsOn.length, 'lab')}{labSeats !== base ? ` (normally ${base}; changed for this test only)` : ''}.{!labsOn.length ? ' No labs in the room list: upload it on step 1.' : ''}</p>
      {labsOn.length > 0 && (
        <div className="captool">
          <label className="small">Add <input className="cap sm" type="number" placeholder="5" value={num} onChange={(e) => setNum(e.target.value)} aria-label="Seats to add to every lab" /> seats to every lab</label>
          <button className="btn ghost sm" disabled={num === ''} onClick={() => { adjustLabs('num', Number(num)); setNum(''); }}>Apply</button>
          <label className="small">or increase every lab by <input className="cap sm" type="number" placeholder="10" value={pct} onChange={(e) => setPct(e.target.value)} aria-label="Percent to add to normal lab capacity" /> %</label>
          <button className="btn ghost sm" disabled={pct === ''} onClick={() => { adjustLabs('pct', Number(pct)); setPct(''); }}>Apply %</button>
          {labSeats !== base && <button className="link small" onClick={() => adjustLabs('reset')}>Reset</button>}
        </div>
      )}
      {labsOn.length > 0 && (
        <details className="fold">
          <summary>Change one lab <span className="muted">· {pl(labsOn.length, 'lab')}</span></summary>
          <div className="lab-list">
            {labsOn.map((l) => (
              <label key={roomKey(l)} className="lab-row on">
                <span className="lab-name">{l.building} · {l.room}</span>
                <input className="cap sm" type="number" min="0" value={l.capacity} onChange={(e) => setLabCap(roomKey(l), e.target.value)} aria-label={`Seats in ${l.building} ${l.room}`} />
                {Number(l.base ?? l.capacity) !== Number(l.capacity) && <span className="muted small">normally {l.base}</span>}
              </label>
            ))}
          </div>
        </details>
      )}
    </div>
  );
}

function PaperCard({ part = 'setup', p, i, n, setPaper, setSlot, remove, courses, phases, ownerOf, slotOf, phaseBatches, studentsByBatch, studentMode, setStudentMode, top, setTop, sumSlot, ok, start, labs, blds }) {
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
        {part === 'mode' ? <span className="muted small">{pl(p.slots.length, 'paper')}</span> : <button type="button" className="iconbtn danger" disabled={n === 1} onClick={remove} aria-label={`Remove test ${i + 1}`} title="Remove this test">{Ico.bin}</button>}
      </div>

      {part === 'mode' && p.slots.map((s, k) => (
        <ModeBox key={s.id} p={p} s={s} k={k} setSlot={setSlot} studentsByBatch={studentsByBatch} studentMode={studentMode} setStudentMode={setStudentMode} top={top} setTop={setTop} sum={sumSlot(p, s)} />
      ))}

      {part === 'setup' && <div className="q">
        <b className="q-l">1. Courses</b>
        <div className="chips">
          {courses.map((c) => <button key={c} type="button" className={`chip-btn ${p.courses.includes(c) ? 'on' : ''}`} onClick={() => toggleCourse(c)}>{c}</button>)}
        </div>
      </div>}

      {part === 'setup' && p.courses.length > 0 && (
        <div className="q">
          <b className="q-l">2. Papers</b>
          <p className="muted small">For each paper: its name, 1 or 2 shifts, its phases, then the batches of those phases.</p>
          {p.slots.map((s, k) => (
            <PaperBox key={s.id} p={p} s={s} k={k} setPaper={setPaper} setSlot={setSlot} phases={phases} ownerOf={ownerOf} slotOf={slotOf}
              pb={phaseBatches(s)} studentsByBatch={studentsByBatch} studentMode={studentMode} setStudentMode={setStudentMode}
              top={top} setTop={setTop} sum={sumSlot(p, s)} />
          ))}
          <button type="button" className="addslot sm" onClick={addSlot}>{Ico.plus}Add another paper</button>
        </div>
      )}

      {part === 'mode' && top.size > 0 && p.courses.length > 0 && <TopSeats p={p} setPaper={setPaper} labs={labs} blds={blds} />}
      {part === 'mode' && ok && (
        <div className="paper-cta">
          {p.done && <span className="pill">Seating ready ✓</span>}
          <button className="btn cta" onClick={start}>{Ico.spark}{p.done ? 'Change seating for this test' : 'Create seating for this test'}</button>
        </div>
      )}
      {part === 'mode' && !ok && <p className="muted small paper-cta">Choose online or offline for every paper to create the seating.</p>}
    </div>
  );
}

// Online / Offline step, one paper: Online, Offline or Both; with Both each batch (or single students); Top ticks
function ModeBox({ p, s, k, setSlot, studentsByBatch, studentMode, setStudentMode, top, setTop, sum }) {
  const [open, setOpen] = useState('');
  const set = (patch) => setSlot(p, s.id, patch);
  const name = (s.name || '').trim() || `Paper ${k + 1}`;
  const toggleTop = (b) => setTop((t) => { const x = new Set(t); x.has(b) ? x.delete(b) : x.add(b); return x; });
  return (
    <div className="timebox paperbox">
      <div className="mode-h"><b>{name}</b><span className="muted small">{slotLabel(s)} · {pl(s.batches.length, 'batch')}</span></div>
      <div className="seg" role="group" aria-label={`Online or offline for ${name}`}>
        {[['Online', 'Online'], ['Offline', 'Offline'], ['Both', 'Both online and offline']].map(([m, l]) => (
          <button key={m} className={s.mode === m ? 'on' : ''} aria-pressed={s.mode === m} onClick={() => set({ mode: m })}>{l}</button>
        ))}
      </div>
      {!s.mode && <p className="muted small">Choose one to continue.</p>}
      <div className="batch-grid">
        {s.batches.map((b) => {
          const on = s.online.includes(b);
          return (
            <div key={b} className="batch on">
              <div className="batch-top"><span className="batch-name">{top.has(b) ? '★ ' : ''}{b}</span><span className="batch-count">{(studentsByBatch[b] || []).length}</span></div>
              <div className="batch-opts">
                {s.mode === 'Both'
                  ? <div className="seg sm" role="group" aria-label={`${b} online or offline`}>{['Online', 'Offline'].map((m) => <button key={m} className={(on ? 'Online' : 'Offline') === m ? 'on' : ''} aria-pressed={(on ? 'Online' : 'Offline') === m} onClick={() => set({ online: m === 'Online' ? [...new Set([...s.online, b])] : s.online.filter((x) => x !== b) })}>{m}</button>)}</div>
                  : <span className="small muted">{s.mode || '–'}</span>}
                {s.mode === 'Both' && <button className="link small" onClick={() => setOpen(open === b ? '' : b)}>Students</button>}
                <label className="small top-tick" title="Top batch: seated first (lab seats first, or in the labs and buildings picked for top batches)"><input type="checkbox" checked={top.has(b)} onChange={() => toggleTop(b)} /> Top</label>
              </div>
            </div>
          );
        })}
      </div>
      {open && s.batches.includes(open) && (
        <StudentModes batch={open} students={studentsByBatch[open] || []} batchMode={s.online.includes(open) ? 'Online' : 'Offline'}
          studentMode={studentMode} setStudentMode={setStudentMode} close={() => setOpen('')} />
      )}
      {s.mode && s.batches.length > 0 && <SeatBoard s={sum} mode={s.mode} />}
    </div>
  );
}

// Top batches of a test: the labs (online students) and buildings (offline students) where they sit first.
// Other students get the seats left there. Belongs to this test only and is never saved.
function TopSeats({ p, setPaper, labs, blds }) {
  const tl = p.topLabs || [], tb = p.topBlds || [];
  const tog = (field, list, k) => setPaper(p.id, { [field]: list.includes(k) ? list.filter((x) => x !== k) : [...list, k] });
  return (
    <div className="q topseats">
      <b className="q-l">Top batches sit first in</b>
      <p className="muted small">Top batches of this test: {(p.top || []).join(', ')}. Pick labs for their online students and buildings for their offline students. They are seated there before anyone else; other students fill the seats left. Nothing picked: top batch students simply take lab seats first.</p>
      {labs.length > 0 && <div className="chips"><span className="small muted">Labs</span>{labs.map((l) => { const k = roomKey(l); return <button key={k} type="button" className={`chip-btn ${tl.includes(k) ? 'on' : ''}`} onClick={() => tog('topLabs', tl, k)}>{l.building} · {l.room} <small>{l.capacity}</small></button>; })}</div>}
      <div className="chips"><span className="small muted">Buildings</span>{blds.map((b) => <button key={b} type="button" className={`chip-btn ${tb.includes(b) ? 'on' : ''}`} onClick={() => tog('topBlds', tb, b)}>{b}</button>)}</div>
    </div>
  );
}

// One paper: name → shifts and times → online or offline → phases → batches of those phases
function PaperBox({ p, s, k, setPaper, setSlot, phases, ownerOf, slotOf, pb, studentsByBatch, studentMode, setStudentMode, top, setTop, sum }) {
  const set = (patch) => setSlot(p, s.id, patch);
  const label = `Paper ${k + 1}`;
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
          <button className={!s.two ? 'on' : ''} aria-pressed={!s.two} onClick={() => set({ two: false })}>1 shift</button>
          <button className={s.two ? 'on' : ''} aria-pressed={s.two} onClick={() => set({ two: true })}>2 shifts</button>
        </div>
        <div className="times">
          <TimeRow label={s.two ? 'Shift 1' : 'Time'} value={s} onChange={(v) => set(v)} />
          {s.two && <TimeRow label="Shift 2" value={s.s2} onChange={(v) => set({ s2: v })} />}
          {s.two && mins(s.s2.start) < mins(s.end) && <p className="muted small bad">Shift 2 must start after Shift 1 ends.</p>}
          {s.two && <p className="note small">Same students, same seat in both shifts. The attendance sheet gets a signature column for each shift.</p>}
        </div>
      </div>

      {(() => {
        // A phase stays open for this paper while some of its batches are not in another paper yet: a phase with 2 of
        // its 3 batches in Paper 1 shows here with the 3rd batch free
        const elsewhere = (b) => slotOf[b.batch] && slotOf[b.batch].slot !== s.id;
        const freeIn = (g) => g.batches.filter((b) => !elsewhere(b));
        const shown = phases.filter((g) => p.courses.includes(g.course));
        return (
          <div className="q">
            <b className="q-l">b. Phases in {s.name.trim() || label}</b>
            <p className="muted small"><button className="link" onClick={() => set(keepBatches(s, phases, shown.filter((g) => s.phases.includes(g.key) || freeIn(g).length).map((g) => g.key)))}>Select all</button> · <button className="link" onClick={() => set(keepBatches(s, phases, []))}>Clear</button></p>
            <div className="chips">
              {shown.map((g) => {
                const left = freeIn(g), part = left.length < g.batches.length;
                const full = !left.length && !s.phases.includes(g.key);
                return (
                  <button key={g.key} type="button" className={`chip-btn ${s.phases.includes(g.key) ? 'on' : ''}`} disabled={full}
                    title={full ? 'All its batches are in another paper' : part ? `Free here: ${left.map((b) => b.batch).join(', ')}` : g.batches.map((b) => b.batch).join(', ')} onClick={() => set(keepBatches(s, phases, toggleIn(s.phases, g.key)))}>
                    {p.courses.length > 1 ? `${g.course} · ` : ''}{g.phase} <small>{part ? `${left.reduce((a, b) => a + b.count, 0)} of ${g.count}` : g.count}</small>
                  </button>
                );
              })}
            </div>
            {shown.some((g) => freeIn(g).length && freeIn(g).length < g.batches.length) && <p className="muted small">“x of y”: the phase has batches in another paper; only its free batches can be ticked here.</p>}
          </div>
        );
      })()}

      {s.phases.length > 0 && (
        <div className="q">
          <b className="q-l">c. Batches</b>
          <p className="muted small">Batches of the phases picked above. <button className="link" onClick={() => set({ batches: free.map((b) => b.batch) })}>Tick all</button> · <button className="link" onClick={() => set({ batches: [], online: [] })}>Clear</button></p>
          <div className="batch-grid">
            {pb.map((b) => {
              const where = slotOf[b.batch];
              const here = s.batches.includes(b.batch);
              const elsewhere = where && where.slot !== s.id;
              return (
                <div key={b.batch} className={`batch ${here ? 'on' : ''} ${elsewhere ? 'other' : ''}`}>
                  <label className="batch-top">
                    <input type="checkbox" checked={here} onChange={() => tickBatch(b.batch)} disabled={!!elsewhere} />
                    <span className="batch-name">{b.batch}</span>
                    {elsewhere ? <span className="pill muted-pill">Other paper</span> : <span className="batch-count">{b.count}</span>}
                  </label>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {s.batches.length > 0 && <p className="small paper-count"><b>{pl(s.batches.reduce((a, b) => a + ((studentsByBatch[b] || []).length), 0), 'student')}</b> in {pl(s.batches.length, 'batch')}</p>}
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
// Girls of this test on a step: how many sit online or offline, and how many have no seat a girl may take
function GirlsLine({ g, kind }) {
  const n = kind === 'online' ? g.online : g.offline;
  return <p className={`small ${g.left ? 'bad' : 'muted'}`}>♀ <b>{n} girls</b> seated {kind}{g.left ? ` · ${g.left} girls have no seat in a room or lab that allows girls: tick another building or lab with Girls allowed` : ' · every girl is in a room or lab that allows girls'}.</p>;
}

function BigLine({ left, placed, empty, what, more }) {
  const n = (v) => Number(v || 0).toLocaleString('en-IN');
  return left > 0
    ? <div className="bigline bad"><b>{n(left)}</b> {what} students still need a seat. {more}</div>
    : <div className="bigline good">All {n(placed)} {what} students have a seat · <b>{n(empty)}</b> seats empty</div>;
}

// One card for the whole test (all its papers share one seating); a change applies to every paper of the test
function LabStep({ p, idx, labsOn, order, setSlot: setOne, planOf, girlsOf, topLabs = [], back, next, label }) {
  const setSlot = (pp, sid, patch) => setOne(pp, planOf({ id: sid }).ids || [sid], patch);
  const slots = p.slots.filter((s) => s.batches.length && planOf(s).first === s.id && planOf(s).c.online > 0);
  const groups = {};
  for (const l of labsOn) (groups[l.building] ||= []).push(l);
  const blds = Object.keys(groups).sort((a, b) => order.indexOf(a) - order.indexOf(b));
  return (
    <section>
      <Head kicker={`Test ${idx + 1}: ${p.courses.join(', ')}`} title="Online seating: which labs?" lead="Tick the computer labs for the online batches. Labs fill in the order shown, one lab completely before the next. Top batch students are seated first. Online students left without a lab seat sit offline." />
      {slots.map((s) => {
        const need = planOf(s).c.online;
        const mine = p.slots.filter((x) => (planOf(s).ids || []).includes(x.id)).flatMap((x) => x.labs || []);
        const sel = new Set([...mine, ...topLabs]);
        const used = planOf(s).taken.labs;
        const free = (l) => Math.max(0, Number(l.capacity) - (used[roomKey(l)]?.filled || 0));
        const chosen = labsOn.filter((l) => sel.has(roomKey(l))).map((l) => ({ ...l, capacity: free(l) }));
        const prev = fillPreview(need, chosen, order);
        const fillOf = Object.fromEntries(prev.rooms.map((r) => [roomKey(r), r]));
        const empty = prev.rooms.reduce((a, r) => a + r.empty, 0);
        const toggle = (k) => setSlot(p, s.id, { labs: mine.includes(k) ? mine.filter((x) => x !== k) : [...mine, k] });
        const gi = girlsOf(s);
        if (gi) for (const l of chosen) { const k = roomKey(l), n = gi.filled['Online' + k] || 0; fillOf[k] = { filled: n, empty: Math.max(0, Number(l.capacity) - n) }; }
        return (
          <div className="card slotcard" key={s.id}>
            <h2>{planOf(s).title}</h2>
            <BigLine left={prev.left} placed={need} empty={empty} what="online" more="Select more labs, or continue and they sit offline." />
            {gi && <GirlsLine g={gi} kind="online" />}
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
                          <input type="checkbox" checked={sel.has(k)} disabled={topLabs.includes(k)} title={topLabs.includes(k) ? 'Kept for top batches (change it on the test)' : ''} onChange={() => toggle(k)} />
                          <span className="lab-name">{l.room}</span>
                          <span className="muted small">{l.capacity} seats{l.girls === false ? ' · no girls' : ''}{topLabs.includes(k) ? ' · ★ top batches first' : ''}</span>
                          {used[k] && <span className="small warnt">{used[k].filled} used by {used[k].by.join(', ')} at this time · {free(l)} free</span>}
                          {f && <span className={`small ${f.empty ? 'muted' : 'good'}`}>{f.filled} filled · {f.empty} empty</span>}
                          {gi && sel.has(k) && <span className="small girls">♀ {gi.inRoom['Online' + k] || 0} girls</span>}
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
          <button className="btn cta" onClick={next}>{label || 'Continue'}</button>
        </div>
      </div>
    </section>
  );
}

function RoomStep({ p, idx, settings, setSettings, setSlot: setOne, planOf, roomsOn, labsOn = [], girlsOf, topBlds = [], onRoomFile, adjustCap, setRoomCap, back, next, label }) {
  const setSlot = (pp, sid, patch) => setOne(pp, planOf({ id: sid }).ids || [sid], patch);
  const editRoom = (kind, i, patch) => setSettings((s) => ({ ...s, [kind]: s[kind].map((r, j) => (j === i ? { ...r, ...patch } : r)) }));
  const move = (i, d) => setSettings((s) => {
    const o = s.buildingOrder.slice(); const j = i + d;
    if (j < 0 || j >= o.length) return s;
    [o[i], o[j]] = [o[j], o[i]]; return { ...s, buildingOrder: o };
  });
  const toggleGirls = (b) => setSettings((s) => ({ ...s, girlsOk: { ...s.girlsOk, [b]: s.girlsOk[b] === false } }));
  const roomsIn = (b) => roomsOn.filter((r) => r.building === b);
  const buildings = settings.buildingOrder.filter((b) => roomsIn(b).length);
  const slots = p.slots.filter((s) => s.batches.length && planOf(s).first === s.id && planOf(s).offNeed > 0);
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
        const prev = fillPreview(need, rooms, settings.buildingOrder);
        const fillOf = Object.fromEntries(prev.rooms.map((r) => [roomKey(r), r]));
        const empty = prev.rooms.reduce((a, r) => a + r.empty, 0);
        const mineB = p.slots.filter((x) => (planOf(s).ids || []).includes(x.id)).flatMap((x) => x.buildings || []);
        const sel = new Set([...mineB, ...topBlds]);
        const order = settings.buildingOrder.filter((b) => sel.has(b));
        const off = new Set(s.roomsOff);
        const toggleB = (b) => setSlot(p, s.id, { buildings: mineB.includes(b) ? mineB.filter((x) => x !== b) : [...mineB, b] });
        const toggleR = (k) => setSlot(p, s.id, { roomsOff: off.has(k) ? s.roomsOff.filter((x) => x !== k) : [...s.roomsOff, k] });
        const group = p.slots.filter((x) => (planOf(s).ids || []).includes(x.id));
        const onlineLabs = new Set(group.flatMap((x) => x.labs || []));
        const offSel = new Set(group.flatMap((x) => x.offLabs || []).filter((k) => !onlineLabs.has(k)));
        const spare = labsOn.filter((l) => !onlineLabs.has(roomKey(l)));
        const toggleL = (k) => setSlot(p, s.id, { offLabs: offSel.has(k) ? [...offSel].filter((x) => x !== k) : [...offSel, k] });
        const gi = girlsOf(s);
        const gIn = (r) => (gi ? gi.inRoom['Offline' + roomKey(r)] || 0 : 0);
        // Filled = the trial seating, so rooms kept for top batches show their students
        if (gi) for (const r of rooms) { const k = roomKey(r), n = gi.filled['Offline' + k] || 0; fillOf[k] = { filled: n, empty: Math.max(0, Number(r.capacity) - n) }; }
        return (
          <div className="card slotcard" key={s.id}>
            <h2>{planOf(s).title}</h2>
            <BigLine left={prev.left} placed={need} empty={empty} what="offline" more="Tick another building, or use a lab for offline students below." />
            {gi && <GirlsLine g={gi} kind="offline" />}
            <div className="bld-chips">
              {buildings.map((b) => {
                const all = roomsIn(b);
                const seats = all.reduce((a, r) => a + Number(r.capacity), 0);
                const others = all.reduce((a, r) => a + usedIn(r), 0);
                const by = [...new Set(all.flatMap((r) => used[roomKey(r)]?.by || []))];
                const mine = sel.has(b) ? all.filter((r) => !off.has(roomKey(r))).reduce((a, r) => a + (fillOf[roomKey(r)]?.filled || 0), 0) : 0;
                return (
                  <label key={b} className={`bld-chip ${sel.has(b) ? 'on' : ''}`}>
                    <input type="checkbox" checked={sel.has(b)} disabled={topBlds.includes(b)} title={topBlds.includes(b) ? 'Kept for top batches (change it on the test)' : ''} onChange={() => toggleB(b)} />
                    <b>{sel.has(b) ? `${order.indexOf(b) + 1}. ` : ''}{b}</b>
                    <span>{seats} seats{others ? ` · ${others} used by ${by.join(', ')}` : ''}{sel.has(b) ? ` · ${mine} filled` : ''} · <b className={seats - others - mine ? 'good' : 'bad'}>{seats - others - mine} free</b></span>
                    <span className="small girls">{settings.girlsOk[b] === false ? 'Boys only' : sel.has(b) ? `♀ ${all.reduce((a, r) => a + gIn(r), 0)} girls` : 'Girls allowed'}{topBlds.includes(b) ? ' · ★ top batches first' : ''}</span>
                  </label>
                );
              })}
            </div>
            <CapacityTool buildings={buildings} roomsIn={roomsIn} inUse={(r) => !off.has(roomKey(r))} adjustCap={adjustCap} />
            {spare.length > 0 && (
              <details className="fold inner" open={offSel.size > 0}>
                <summary>Labs used for offline students <span className="muted">· {offSel.size ? `${pl(offSel.size, 'lab')} chosen` : 'none'} · when classrooms run short; labs this test uses online are not listed</span></summary>
                <div className="lab-list">
                  {spare.map((l) => {
                    const k = roomKey(l), on = offSel.has(k), f = fillOf[k], u = usedIn(l);
                    return (
                      <label key={k} className={`lab-row ${on ? 'on' : ''}`}>
                        <input type="checkbox" checked={on} onChange={() => toggleL(k)} />
                        <span className="lab-name">{l.building} · {l.room}</span>
                        <span className="muted small">{l.capacity} seats{l.girls === false || settings.girlsOk[l.building] === false ? ' · no girls' : ''}{u ? ` · ${u} used by another test` : ''}</span>
                        {on && f && <span className={`small ${f.empty ? 'muted' : 'good'}`}>{f.filled} filled · {f.empty} empty</span>}
                        {on && gi && <span className="small girls">♀ {gi.inRoom['Offline' + k] || 0} girls</span>}
                      </label>
                    );
                  })}
                </div>
              </details>
            )}
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
                            <span>{r.base != null && r.base !== Number(r.capacity) ? `normally ${r.base} · ` : ''}{u ? `${u} used · ` : ''}{use && f ? `${f.filled} filled` : ''}{use && gIn(r) ? ` · ♀ ${gIn(r)}` : ''}</span>
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
        <p className="muted">To load a new room list, go back to step 1.</p>
        <button className="btn ghost danger" onClick={async () => { if (await askConfirm('Reset buildings, rooms, capacities, girls settings and groups to the built-in lists?', 'Reset', true)) setSettings(freshSettings()); }}>{Ico.bin}Reset to built-in lists</button>
      </details>
      <div className="ctabar">
        <div><b>Buildings</b><span className="muted">{`All papers: ${[...new Set(p.slots.flatMap((x) => x.buildings || []))].join(', ') || 'none yet'}`}</span></div>
        <div className="cta-row">
          <button className="btn ghost" onClick={back}>Back</button>
          <button className="btn cta" onClick={next}>{label || 'Continue'}</button>
        </div>
      </div>
    </section>
  );
}

// Change the room capacity for this test, for one or several buildings at once: by a percentage of the original
// capacity, or by a number of seats added to every room in use (the same way the percentage changes every room)
function CapacityTool({ buildings, roomsIn, inUse, adjustCap }) {
  const [sel, setSel] = useState([]);
  const [pct, setPct] = useState('');
  const [num, setNum] = useState('');
  const pick = sel.filter((b) => buildings.includes(b));
  const list = pick.flatMap((b) => roomsIn(b));
  const seats = list.reduce((a, r) => a + Number(r.capacity), 0);
  const base = list.reduce((a, r) => a + Number(r.base ?? r.capacity), 0);
  const rooms = list.filter(inUse).length;
  if (!buildings.length) return null;
  const tog = (b) => setSel((x) => (x.includes(b) ? x.filter((y) => y !== b) : [...x, b]));
  const each = (how, v) => pick.forEach((b) => adjustCap(b, how, v, roomsIn(b).filter(inUse).map(roomKey)));
  return (
    <div className="captool multi">
      <b className="small">Change capacity for this test</b>
      <div className="chips">
        {buildings.map((b) => <button key={b} type="button" className={`chip-btn sm ${pick.includes(b) ? 'on' : ''}`} onClick={() => tog(b)}>{b}</button>)}
        <button type="button" className="link small" onClick={() => setSel(pick.length === buildings.length ? [] : buildings.slice())}>{pick.length === buildings.length ? 'Clear' : 'Select all'}</button>
      </div>
      {pick.length > 0 ? <span className="muted small">{pl(pick.length, 'building')}: {seats} seats in {pl(rooms, 'room')} in use{seats !== base ? ` (normally ${base})` : ''}</span> : <span className="muted small">Pick one or more buildings.</span>}
      <div className="row">
        <label className="small">increase by <input className="cap sm" type="number" placeholder="10" value={pct} onChange={(e) => setPct(e.target.value)} aria-label="Percent to add to normal capacity" /> %</label>
        <button className="btn ghost sm" disabled={pct === '' || !pick.length} onClick={() => { each('pct', Number(pct)); setPct(''); }}>Apply %</button>
        <label className="small">or add <input className="cap sm" type="number" placeholder="2" value={num} onChange={(e) => setNum(e.target.value)} aria-label="Seats to add to every room (minus to remove)" /> seats to every room</label>
        <button className="btn ghost sm" disabled={num === '' || !pick.length} onClick={() => { each('num', Number(num)); setNum(''); }}>Add seats</button>
        {pick.length > 0 && seats !== base && <button className="link small" onClick={() => each('reset')}>Reset</button>}
      </div>
      <span className="muted small">A minus number removes seats. Other tests keep the normal capacity.</span>
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
      {!['mixed', 'roll'].includes(order) && (
        <label className="field"><span>Then by <em className="muted">(students equal on the first order)</em></span>
          <select value={settings.seatOrder2 || ''} onChange={(e) => setSettings((x) => ({ ...x, seatOrder2: e.target.value }))} aria-label="Second seat order">
            <option value="">Roll number</option>
            {Object.entries(SEAT_ORDERS).filter(([k]) => !['mixed', 'roll', order].includes(k)).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
          </select>
        </label>
      )}
      <p className="muted small">{order === 'mixed' ? 'Students of all papers are shuffled together, and two students of the same batch never sit side by side. Reshuffle on the result page gives a new mix.' : `Every student of the test is seated in ${SEAT_ORDERS[order].toLowerCase()} order${!['mixed', 'roll'].includes(order) && settings.seatOrder2 && settings.seatOrder2 !== order ? `, then ${SEAT_ORDERS[settings.seatOrder2].toLowerCase()}` : ''}: the first room fills first, seat 1 onwards, then the next room.`}{['father', 'dob', 'contact'].includes(order) ? ' Needs that column in the admission list; students without it come last.' : ''}</p>
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
  const [look, setLook] = useState('cards');
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
          <button className="btn secondary" onClick={() => setMenu(!menu)} aria-expanded={menu}>More ▾</button>
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
        <div><b>{shiftsOf(x.paper).length}</b><span>{shiftsOf(x.paper).length > 1 ? 'attendance sheets (one per shift)' : 'attendance sheet'}</span></div>
      </div>
      {result.warnings.map((w) => <div key={w} className="alert warn"><span>{w}</span></div>)}

      <details className="fold" open>
        <summary>Rooms <span className="muted">· {result.usedRooms.length} used</span></summary>
        <div className="seg sm viewseg" role="group" aria-label="Show rooms as">
          {[['cards', 'Cards'], ['table', 'Table']].map(([k, l]) => <button key={k} className={look === k ? 'on' : ''} aria-pressed={look === k} onClick={() => setLook(k)}>{l}</button>)}
        </div>
        {look === 'table' && (
          <div className="tablewrap">
            <table className="rtable">
              <thead><tr><th scope="col">Building</th><th scope="col">Room</th><th scope="col">Type</th><th scope="col" className="num">Students</th><th scope="col" className="num">Capacity</th><th scope="col">Seats</th><th scope="col" className="num">Girls</th><th scope="col" className="num">Special</th></tr></thead>
              <tbody>
                {Object.entries(byBuilding).flatMap(([b, rooms]) => rooms.map((r) => (
                  <tr key={`${b}${r.kind}${r.room}`}>
                    <td>{b}</td><td><b>{r.room}</b></td><td className="muted">{r.kind === 'Online' ? 'Lab' : 'Classroom'}</td>
                    <td className="num">{r.used}</td><td className="num">{r.capacity}</td><td className="muted">{r.taken + 1}–{r.taken + r.used}</td>
                    <td className="num">{r.occupants.filter(isGirl).length}</td><td className="num">{r.occupants.filter((s) => s.Flag).length || ''}</td>
                  </tr>
                )))}
              </tbody>
            </table>
          </div>
        )}
        {look === 'cards' && <p className="muted small">Green: students of this test (all its papers). Amber: monitoring, due fee or top batch students.{x.overlapsWith.length ? ' Grey: seats used by an overlapping test.' : ''}</p>}
        {look === 'cards' && Object.entries(byBuilding).map(([b, rooms]) => (
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
