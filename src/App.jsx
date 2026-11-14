import React, { useEffect, useMemo, useRef, useState } from 'react';
import defaults from './data/defaults.json';
import { capacity, naturalCompare, isGirl, roomKey, FLAGS, planTest, prtModes, onlineUnits, checkOfflineTiming, slotOk,
  TEST_TYPES, FORMATS, ANTI_CHEAT, defaultSlots, DEFAULT_TOP_SLOTS } from './lib/engine.js';
import { buildWorkbook, buildNoticeBoard, fmtDateDMY, fmtTime } from './lib/exporter.js';
import { readAdmissions, readRooms, readRollList, parseRolls, readSeatingFile } from './lib/readers.js';

const loadExcel = () => import('exceljs').then((m) => m.default || m);
const STORE = 'seating-settings-v1';
const LAST = 'seating-last-test-v2';
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
const newPaper = (start = '07:00', end = '10:00', extra = {}) => ({ id: uid(), name: '', course: '', phase: '', extra: [], s1: { start, end }, s2: { start: hhmm(mins(end) + 30), end: hhmm(mins(end) + 30 + (mins(end) - mins(start))) }, ...extra });
const testDefaults = (type = 'PRT') => ({ type, format: 'MAIN', days: 1, slots: defaultSlots(type), topSlots: DEFAULT_TOP_SLOTS.map((s) => ({ ...s })), topShare: false });
function initialTest() {
  const last = loadLast();
  if (last && Array.isArray(last.papers) && last.papers.length) {
    const papers = last.papers.map((p) => ({ course: '', phase: '', extra: [], ...p }));
    // Older saved tests: a paper in two shifts is the JEE Advanced format (two papers, same seat)
    const format = last.format || (papers.some((p) => p.two) ? 'ADV' : 'MAIN');
    return { ...testDefaults(last.type || 'PRT'), ...(last.testSet || {}), format, title: last.title || '', date: '', papers };
  }
  return { ...testDefaults(), title: '', date: '', papers: [newPaper()] };
}

const pl = (n, w) => `${n.toLocaleString('en-IN')} ${w}${n === 1 ? '' : /(ch|sh|s|x)$/.test(w) ? 'es' : 's'}`;
const mins = (t) => { const [h, m] = String(t || '0:0').split(':').map(Number); return h * 60 + m; };
const hhmm = (m) => `${String(Math.floor(Math.min(m, 1439) / 60)).padStart(2, '0')}:${String(Math.min(m, 1439) % 60).padStart(2, '0')}`;
// A paper's shifts. Two shifts: the same students and the same seats, an attendance sheet for each shift.
// JEE Advanced: Paper 1 and Paper 2, same students, same seats, an attendance sheet for each.
const shiftsOf = (p) => (p.two ? [{ label: 'Paper 1', ...p.s1 }, { label: 'Paper 2', ...p.s2 }] : [{ label: '', ...p.s1 }]);
const windowOf = (p) => { const sh = shiftsOf(p); return { start: sh[0].start, end: sh[sh.length - 1].end }; };
export const overlaps = (a, b) => {
  if (a.id === b.id) return true;
  const x = windowOf(a), y = windowOf(b);
  return mins(x.start) < mins(y.end) + GAP && mins(y.start) < mins(x.end) + GAP;
};
const paperName = (p, i) => p.name || [p.course, p.phase].filter(Boolean).join(' ').trim() || `Paper ${i + 1}`;
const range = (s) => `${fmtTime(s.start)}${s.end ? ` to ${fmtTime(s.end)}` : ''}`;
const paperTime = (p) => shiftsOf(p).map((s) => (s.label ? `${s.label} ${range(s)}` : range(s))).join(' · ');
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
  admission: { file: 'Admission_list_template.xlsx', sheets: [['Admission list', ['Roll_Number', 'Student_Name', 'Gender', 'Course', 'Batch', 'Medium', 'Status']]] },
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
  { id: 'test', label: 'Test & phases' },
  { id: 'batches', label: 'Batches' },
  { id: 'rooms', label: 'Buildings & rooms' },
  { id: 'result', label: 'Seating' },
];

export default function App() {
  const last0 = useMemo(loadLast, []);
  const [step, setStep] = useState('students');
  const [settings, setSettings] = useState(loadSettings);
  const [adm, setAdm] = useState(null);
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [test, setTest] = useState(initialTest);
  const [selected, setSelected] = useState(() => new Set());
  const [mode, setMode] = useState({});
  const [studentMode, setStudentMode] = useState({});
  const [paperOf, setPaperOf] = useState({});
  const [top, setTop] = useState(() => new Set(last0?.top || []));
  const [lists, setLists] = useState({ Monitoring: [], 'Due fee': [] });
  const [seed, setSeed] = useState(() => Math.floor(Math.random() * 1e9));
  const [results, setResults] = useState(null); // sittings from planTest: [{ key, name, time, slots, dateISO, result, batches, students, ... }]
  const [planWarnings, setPlanWarnings] = useState([]);
  const [rotation, setRotation] = useState({ on: false, files: [] }); // PRT: previous seating files for the online turns
  const [antiCheat, setAntiCheat] = useState('');
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
  const studentsByBatch = useMemo(() => {
    const m = {};
    for (const s of adm?.students || []) (m[s.Batch] ||= []).push(s);
    return m;
  }, [adm]);

  const adv = test.format === 'ADV';
  const prt = test.type === 'PRT';
  const papers = useMemo(() => test.papers.map((p) => ({ ...p, two: adv })), [test.papers, adv]);
  const multi = papers.length > 1;
  const paperIds = papers.map((p) => p.id);
  const paperFor = (batch) => (paperIds.includes(paperOf[batch]) ? paperOf[batch] : paperIds[0]);
  const chosen = batches.filter((b) => selected.has(b.batch));
  const studentsOf = (bs) => bs.flatMap((b) => studentsByBatch[b.batch] || []);
  const useOrder = settings.buildingOrder.filter((b) => settings.buildingsOn.includes(b));
  const base = { rooms: settings.rooms, labs: settings.labs, buildingOrder: useOrder, girlsOk: settings.girlsOk };
  // Previous tests for the PRT rotation: files of the same date are one test, oldest first
  const history = useMemo(() => {
    const by = {};
    rotation.files.forEach((f, i) => { (by[f.date || `file-${i}`] ||= []).push(...f.online); });
    return Object.keys(by).sort().map((k) => by[k]);
  }, [rotation.files]);
  const unitCfg = { ...base, testType: test.type, format: test.format, days: test.days, slots: test.slots, topSlots: test.topSlots };
  const prtPlan = useMemo(() => (prt ? prtModes(studentsOf(chosen), { ...unitCfg, topBatches: [...top], studentMode, batchMode: mode, rotation: { on: rotation.on, history } }) : { mode: {}, summary: [], seats: 0 }),
    [prt, chosen, settings, test.slots, test.format, top, studentMode, mode, rotation.on, history]);
  const modeOf = (s) => (!prt ? 'Online' : prtPlan.mode[s.Roll_Number] || studentMode[s.Roll_Number] || mode[s.Batch] || 'Offline');

  // Special students: Monitoring and Due fee from lists, Top batch from ticked top batches
  const flags = useMemo(() => {
    const f = {};
    for (const b of top) for (const s of studentsByBatch[b] || []) f[s.Roll_Number] = 'Top batch';
    for (const r of lists['Due fee']) f[r] = 'Due fee';
    for (const r of lists.Monitoring) f[r] = 'Monitoring';
    return f;
  }, [top, lists, studentsByBatch]);

  const needOf = (bs) => {
    const n = { offline: 0, online: 0, girlsOffline: 0, girlsOnline: 0 };
    for (const s of studentsOf(bs)) {
      const on = modeOf(s) === 'Online';
      n[on ? 'online' : 'offline']++;
      if (isGirl(s)) n[on ? 'girlsOnline' : 'girlsOffline']++;
    }
    return n;
  };
  const perPaper = papers.map((p) => ({ paper: p, list: chosen.filter((b) => !multi || paperFor(b.batch) === p.id) })).map((x) => ({ ...x, ...needOf(x.list) }));
  // Offline seats needed at the busiest moment: a paper plus every paper that overlaps it.
  // Online: every online student needs one seat in one slot (or one pair of slots for JEE Advanced).
  const need = perPaper.reduce((best, x) => {
    const over = perPaper.filter((y) => overlaps(x.paper, y.paper));
    const sum = (k) => over.reduce((a, y) => a + y[k], 0);
    const n = { offline: sum('offline'), girlsOffline: sum('girlsOffline') };
    return n.offline > best.offline ? n : best;
  }, { offline: 0, girlsOffline: 0 });
  Object.assign(need, { online: perPaper.reduce((a, x) => a + x.online, 0), girlsOnline: perPaper.reduce((a, x) => a + x.girlsOnline, 0) });

  const units = onlineUnits(unitCfg);
  const topStudents = prt ? 0 : studentsOf(chosen.filter((b) => top.has(b.batch))).length;
  const topCap = units.top.reduce((a, u) => a + u.cap, 0);
  const topGirlCap = units.top.reduce((a, u) => a + u.gcap, 0);
  const capOff = capacity(base, 'Offline');
  const have = {
    offline: prt ? capOff.seats : 0, girlsOffline: prt ? capOff.girlSeats : 0,
    online: units.general.reduce((a, u) => a + u.cap, 0) + (test.topShare ? topCap : Math.min(topCap, topStudents)),
    girlsOnline: units.general.reduce((a, u) => a + u.gcap, 0) + (test.topShare ? topGirlCap : Math.min(topGirlCap, topStudents)),
  };
  const timing = prt ? checkOfflineTiming(papers.map((p, i) => ({ ...p, name: paperName(p, i) })), adv) : [];
  const seatsShort = Math.max(0, need.offline - have.offline) + Math.max(0, need.online - have.online);
  const girlsShort = Math.max(0, need.girlsOffline - have.girlsOffline) + Math.max(0, need.girlsOnline - have.girlsOnline);
  const appearing = chosen.reduce((a, b) => a + b.count, 0);
  const slotsOk = test.slots.some(slotOk) && test.slots.every((s) => !s.start || slotOk(s));
  const papersOk = papers.length > 0 && (!prt || papers.every((p) => shiftOk(p.s1) && (!adv || (shiftOk(p.s2) && mins(p.s2.start) >= mins(p.s1.end))) && (adv || (p.extra || []).every(shiftOk)))) && (slotsOk || (prt && !test.slots.length));
  const emptyPapers = multi ? perPaper.filter((x) => !x.list.length) : [];

  const ready = {
    students: !!adm,
    test: !!test.date && papersOk,
    batches: chosen.length > 0,
    rooms: useOrder.length > 0,
    result: !!results,
  };
  const canCreate = ready.students && ready.test && ready.batches && ready.rooms && !!antiCheat && !timing.length;

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
      // Remember the last test: tick the same batches, modes and papers again when they exist in the new list
      const last = loadLast(); const names = new Set(r.students.map((s) => s.Batch)); const rolls = new Set(r.students.map((s) => s.Roll_Number));
      const keep = (o) => Object.fromEntries(Object.entries(o || {}).filter(([b]) => names.has(b)));
      const again = (last?.batches || []).filter((b) => names.has(b));
      setSelected(new Set(again)); setMode(keep(last?.mode)); setPaperOf(keep(last?.paperOf));
      setStudentMode(Object.fromEntries(Object.entries(last?.studentMode || {}).filter(([k]) => rolls.has(k))));
      setTop(new Set((last?.top || []).filter((b) => names.has(b))));
      if (again.length) setNotice(`${pl(again.length, 'batch')} from the last test ticked again. Change them on the Batches step.`);
      setStep('test');
    } catch (e) { setError(`Could not read “${file.name}”: ${e.message}`); }
    setBusy('');
  }
  function removeAdmission() {
    if (!confirm('Remove the admission list from this page? Your rooms and settings stay.')) return;
    setAdm(null); setSelected(new Set()); setMode({}); setStudentMode({}); setPaperOf({}); setResults(null); setNotice(''); setStep('students');
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

  async function onSeatingFiles(files) {
    if (!files || !files.length) return;
    setError(''); setBusy('Reading the previous seating…');
    const got = [];
    try {
      const ExcelJS = await loadExcel();
      for (const file of files) got.push({ name: file.name, ...(await readSeatingFile(ExcelJS, await file.arrayBuffer())) });
      setRotation((r) => ({ on: true, files: [...r.files.filter((f) => !got.some((g) => g.name === f.name)), ...got] }));
      setNotice(`${pl(got.length, 'previous seating file')} read: ${got.map((g) => `${g.online.length} online${g.date ? ` on ${fmtDateDMY(g.date)}` : ''}`).join('; ')}.`);
    } catch (e) { setError(`Could not read the previous seating file: ${e.message}`); }
    setBusy('');
  }

  // Plan the whole test: online slots (PRT shifts, RPT slots, Test series days) and offline sittings of each paper.
  // Sittings that overlap in time share rooms without using a seat twice. JEE Advanced: one seating for both papers.
  function run(newSeed = seed) {
    const cbOf = {};
    for (const g of settings.cbGroups) for (const b of parseRolls(g.batches)) if (g.name.trim()) cbOf[b] = g.name.trim();
    const plan = planTest(studentsOf(chosen), {
      ...unitCfg, dateISO: test.date, topShare: test.topShare, topBatches: [...top],
      papers: papers.map((p, i) => ({ ...p, name: paperName(p, i), batches: chosen.filter((b) => !multi || paperFor(b.batch) === p.id).map((b) => b.batch) })),
      batchMode: Object.fromEntries(chosen.map((b) => [b.batch, mode[b.batch] || 'Offline'])), studentMode, flags, pins: settings.pins, cbOf,
      rotation: { on: prt && rotation.on, history }, antiCheat, seed: newSeed,
    });
    if (plan.blocked) { setError(plan.warnings.join(' ')); return; }
    setResults(plan.units); setPlanWarnings(plan.warnings); setSeed(newSeed); setView(0); setStep('result');
    try {
      localStorage.setItem(LAST, JSON.stringify({ title: test.title, papers: test.papers, type: test.type, format: test.format,
        testSet: { days: test.days, slots: test.slots, topSlots: test.topSlots, topShare: test.topShare }, batches: [...selected], mode, studentMode, paperOf, top: [...top] }));
    } catch { /* ignore */ }
  }

  const courseOf = Object.fromEntries(batches.map((b) => [b.batch, b.course]));
  const metaFor = (x) => {
    const courses = [...new Set(x.batches.map((b) => courseOf[b]).filter(Boolean))];
    const n = { online: x.result.assigned.filter((s) => s.Mode === 'Online').length, offline: x.result.assigned.filter((s) => s.Mode !== 'Online').length };
    const head = test.title || `${courses.join(' & ') || 'Test'} Seating Arrangement`;
    return {
      title: results && results.length > 1 ? `${head} - ${x.name}` : head,
      dateISO: x.dateISO || test.date,
      shifts: x.slots.map((s) => s.start),
      slots: x.slots,
      batchesLabel: [...new Set(x.batches.map((b) => b.slice(0, 2)))].sort().join(', '),
      modeLabel: n.online && n.offline ? 'Offline + Online Test' : n.online ? 'Online Test' : 'Offline Test',
    };
  };

  const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  async function download(list) {
    const todo = list.filter((x) => x.students || x.result.unseated.length);
    setBusy(todo.length > 1 ? `Building ${todo.length} Excel files…` : 'Building the Excel file…');
    try {
      const ExcelJS = await loadExcel();
      for (const x of todo) {
        const meta = metaFor(x);
        const buf = await buildWorkbook(ExcelJS, x.result, meta).xlsx.writeBuffer();
        saveBlob(new Blob([buf], { type: XLSX }), `${meta.title.replace(/[^\w]+/g, '_')}_${fmtDateDMY(test.date)}.xlsx`);
        await new Promise((r) => setTimeout(r, 400));
      }
    } catch (e) { setError(`Could not build the file: ${e.message}`); }
    setBusy('');
  }
  async function downloadBoard() {
    setBusy('Building the notice board…');
    try {
      const ExcelJS = await loadExcel();
      const dayTitle = test.title || 'Test Seating Arrangement';
      const wb = buildNoticeBoard(ExcelJS, results.filter((x) => x.students).map((x) => ({ result: x.result, meta: metaFor(x) })), { dateISO: test.date, dayTitle });
      saveBlob(new Blob([await wb.xlsx.writeBuffer()], { type: XLSX }), `Notice_board_${fmtDateDMY(test.date)}.xlsx`);
    } catch (e) { setError(`Could not build the notice board: ${e.message}`); }
    setBusy('');
  }

  const canOpen = (i) => (STEPS[i].id === 'result' ? !!results : STEPS.slice(0, i).every((s) => ready[s.id] || s.id === 'rooms'));

  return (
    <div className="app">
      <header className="top">
        <div className="top-in">
          <div className="logo">
            <img src="/brand/matrix-logo-green.png" alt="Matrix Academy" className="logo-img" />
            <span className="sep" aria-hidden="true" />
            <div className="brand"><b>Test <em>Seating</em></b><span className="chip">{test.date ? `${(TEST_TYPES.find((t) => t.id === test.type) || {}).name} · ${fmtDateDMY(test.date)} · ${pl(papers.length, 'phase')}` : 'No test date yet'}</span></div>
          </div>
          <nav className="stepper" aria-label="Steps">
            {STEPS.map((s, i) => {
              const done = ready[s.id] && step !== s.id && s.id !== 'result';
              return (
                <button key={s.id} type="button" disabled={!canOpen(i)} aria-current={step === s.id ? 'step' : undefined}
                  className={`st ${step === s.id ? 'on' : ''} ${done ? 'done' : ''}`} onClick={() => setStep(s.id)}>
                  <i>{done ? '✓' : i + 1}</i><span>{s.label}</span>
                </button>
              );
            })}
          </nav>
        </div>
      </header>

      {adm && step !== 'students' && (
        <div className="sumbar" role="status">
          <span><b>{adm.students.length.toLocaleString('en-IN')}</b> active students</span>
          <span><b>{chosen.length}</b> {chosen.length === 1 ? 'batch' : 'batches'} ticked</span>
          <span><b>{appearing.toLocaleString('en-IN')}</b> appearing</span>
          <span><b>{papers.length}</b> {papers.length === 1 ? 'phase' : 'phases'}</span>
          <span><b>{useOrder.length}</b> {useOrder.length === 1 ? 'building' : 'buildings'}</span>
          <span className={seatsShort || girlsShort ? 'bad' : chosen.length ? 'good' : ''}><b>{(have.offline + have.online).toLocaleString('en-IN')}</b> seats{seatsShort ? ` · ${seatsShort} short` : girlsShort ? ` · ${girlsShort} girls short` : chosen.length ? ' · enough' : ''}</span>
        </div>
      )}

      <main className="main">
        {error && <div className="alert bad" role="alert"><span>{error}</span><button className="link" onClick={() => setError('')}>Dismiss</button></div>}
        {notice && !error && <div className="alert info"><span>{notice}</span><button className="link" onClick={() => setNotice('')}>OK</button></div>}
        {busy && <div className="busy" role="status"><span className="spin" aria-hidden="true" />{busy}</div>}

        {step === 'students' && (
          <section>
            <Head kicker="Step 1 of 4" title="Upload the admission list" lead="The Excel file with every student. Only students marked Active are used. The file is read inside this browser and is never uploaded anywhere." />
            {!adm && <Guide />}
            <FileCard title="Admission list" hint="Columns: Roll_Number, Student_Name, Gender, Course, Batch, Medium, Status"
              loaded={adm && <><b>{adm.fileName}</b><span>{pl(adm.students.length, 'active student')} · {pl(batches.length, 'batch')}{adm.inactive ? ` · ${adm.inactive} inactive left out` : ''}{adm.duplicates ? ` · ${adm.duplicates} duplicate rolls skipped` : ''}</span></>}
              onFile={onAdmission} onRemove={removeAdmission} template={() => downloadTemplate('admission')} big />
            {adm && <Nav next={() => setStep('test')} />}
          </section>
        )}

        {step === 'test' && (
          <PaperStep test={test} setTest={setTest} papersOk={papersOk} timing={timing} units={units}
            batchNames={new Set(batches.map((b) => b.batch))} setSelected={setSelected} setPaperOf={setPaperOf} setNotice={setNotice}
            back={() => setStep('students')} next={() => setStep('batches')} nextOk={ready.test} />
        )}

        {step === 'batches' && (
          <BatchStep
            batches={batches} studentsByBatch={studentsByBatch} selected={selected} setSelected={setSelected}
            mode={mode} setMode={setMode} studentMode={studentMode} setStudentMode={setStudentMode} top={top} setTop={setTop}
            papers={papers} multi={multi} paperFor={paperFor} setPaperOf={setPaperOf} perPaper={perPaper}
            need={need} emptyPapers={emptyPapers} appearing={appearing} prt={prt} type={test.type}
            rotation={rotation} setRotation={setRotation} onSeatingFiles={onSeatingFiles} rotSummary={prtPlan.summary} onlineSeats={have.online}
            back={() => setStep('test')} next={() => setStep('rooms')} nextOk={ready.batches}
          />
        )}

        {step === 'rooms' && (
          <RoomStep settings={settings} setSettings={setSettings} need={need} have={have} chosen={chosen} appearing={appearing}
            lists={lists} setLists={setLists} onRollFile={onRollFile} flags={flags} chosenRolls={new Set(studentsOf(chosen).map((s) => s.Roll_Number))}
            onRoomFile={onRoomFile} canCreate={canCreate} useOrder={useOrder} test={test} setTest={setTest} prt={prt}
            topStudents={topStudents} topCap={topCap} units={units} antiCheat={antiCheat} setAntiCheat={setAntiCheat} timing={timing}
            back={() => setStep('batches')} create={() => run()} hasResult={!!results} />
        )}

        {step === 'result' && results && (
          <ResultStep results={results} view={view} setView={setView} metaFor={metaFor} multi={results.length > 1} order={useOrder} planWarnings={planWarnings} seed={seed}
            antiCheatName={(ANTI_CHEAT.find((o) => o.id === antiCheat) || {}).name}
            reshuffle={() => { if (confirm('Reshuffle every seat? The current seating is replaced.')) run(Math.floor(Math.random() * 1e9)); }}
            download={download} downloadBoard={downloadBoard} back={() => setStep('rooms')} />
        )}
      </main>
      <footer className="foot">Matrix Academy, Sikar · Operations · Student data stays in this browser</footer>
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
        <li><i>1</i><b>Admission list</b><span>Upload the latest admission list.</span></li>
        <li><i>2</i><b>Date and papers</b><span>Add each paper, in one shift or two.</span></li>
        <li><i>3</i><b>Batches</b><span>Tick each paper's batches, Online or Offline.</span></li>
        <li><i>4</i><b>Buildings</b><span>Choose the buildings and their order, then Create seating.</span></li>
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

// Online slots editor: start, end and a "No girls" switch per slot
function SlotList({ title, hint, slots, onChange, girlsSwitch = true, max = 8 }) {
  const set = (i, patch) => onChange(slots.map((s, j) => (j === i ? { ...s, ...patch } : s)));
  const setStart = (i, v) => { const s = slots[i]; const len = slotOk(s) ? mins(s.end) - mins(s.start) : 180; set(i, { start: v, end: v ? hhmm(mins(v) + len) : s.end }); };
  const add = () => { const lastS = slots[slots.length - 1]; const start = lastS ? hhmm(mins(lastS.end) + 30) : '09:00'; onChange([...slots, { start, end: hhmm(mins(start) + 180), noGirls: false }]); };
  return (
    <div className="card paper">
      <div className="paper-top"><b className="grow">{title}</b><span className="muted small">{hint}</span></div>
      {slots.map((s, i) => (
        <div className={`slot ${slotOk(s) ? '' : 'bad'}`} key={i}>
          <b className="shift-l">Slot {i + 1}</b>
          <label className="field"><span>Starts</span><input type="time" value={s.start} onChange={(e) => setStart(i, e.target.value)} /></label>
          <label className="field"><span>Ends</span><input type="time" value={s.end} onChange={(e) => set(i, { end: e.target.value })} /></label>
          {girlsSwitch && <label className="far small"><input type="checkbox" checked={!!s.noGirls} onChange={(e) => set(i, { noGirls: e.target.checked })} /> No girls</label>}
          <span className="slot-time muted">{slotOk(s) ? range(s) : 'Set the end time'}</span>
          <button type="button" className="iconbtn danger" onClick={() => onChange(slots.filter((_, j) => j !== i))} aria-label={`Remove slot ${i + 1}`} title="Remove this slot">{Ico.bin}</button>
        </div>
      ))}
      {slots.length < max && <button type="button" className="addslot" onClick={add}>{Ico.plus}Add a slot</button>}
    </div>
  );
}

function PaperStep({ test, setTest, papersOk, timing, units, batchNames, setSelected, setPaperOf, setNotice, back, next, nextOk }) {
  const [paste, setPaste] = useState('');
  const prt = test.type === 'PRT';
  const adv = test.format === 'ADV';
  const setPaper = (id, patch) => setTest({ ...test, papers: test.papers.map((p) => (p.id === id ? { ...p, ...patch } : p)) });
  const setTime = (p, which, key, v) => {
    const s = { ...p[which] };
    if (key === 'start') { const len = s.end && s.start ? mins(s.end) - mins(s.start) : 180; s.start = v; s.end = v ? hhmm(mins(v) + (len > 0 ? len : 180)) : s.end; } else s.end = v;
    setPaper(p.id, { [which]: s });
  };
  const setExtra = (p, i, key, v) => {
    const extra = (p.extra || []).map((s, j) => {
      if (j !== i) return s;
      if (key === 'start') { const len = slotOk(s) ? mins(s.end) - mins(s.start) : 180; return { start: v, end: v ? hhmm(mins(v) + len) : s.end }; }
      return { ...s, end: v };
    });
    setPaper(p.id, { extra });
  };
  const addPaper = () => {
    const lastP = test.papers[test.papers.length - 1];
    setTest({ ...test, papers: [...test.papers, newPaper(lastP ? lastP.s1.start : '07:00', lastP ? lastP.s1.end : '10:00')] });
  };
  const removePaper = (id) => setTest({ ...test, papers: test.papers.filter((p) => p.id !== id) });
  const setType = (type) => setTest({ ...test, type, slots: type === test.type ? test.slots : defaultSlots(type), days: type === 'SERIES' ? Math.max(1, test.days || 1) : 1 });
  // "Course | Phase | Batches", one phase per line (the same list as the Counselling Planner)
  const makePhases = () => {
    const lines = paste.split(/\r?\n/).map((l) => l.split(/\t|\|/).map((x) => x.trim())).filter((c) => c.length >= 3 && c[2]);
    if (!lines.length) { setNotice('Nothing to add. Write one phase per line: Course | Phase | Batches'); return; }
    const t0 = test.papers[0] || newPaper();
    const made = lines.map(([course, phase]) => newPaper(t0.s1.start, t0.s1.end, { course, phase, s2: { ...t0.s2 } }));
    const unknown = [];
    const assign = {};
    lines.forEach((c, i) => parseRolls(c[2].replace(/\//g, ' ')).forEach((b) => { if (batchNames.has(b)) assign[b] = made[i].id; else unknown.push(b); }));
    setTest({ ...test, papers: made });
    setSelected(new Set(Object.keys(assign)));
    setPaperOf(assign);
    setNotice(`${pl(made.length, 'phase')} made with ${pl(Object.keys(assign).length, 'batch')}.${unknown.length ? ` Not in the admission list: ${unknown.join(', ')}.` : ''}`);
    setPaste('');
  };
  const onlineSittings = units.general.length + units.top.length;
  const hint = !test.date ? 'Pick the test date to continue.' : timing.length ? 'Fix the offline timing first.' : !papersOk ? (prt ? 'Each time needs a start and an end; Paper 2 starts after Paper 1 ends.' : 'Add at least one slot with a start and an end time.') : '';
  return (
    <section>
      <Head kicker="Step 2 of 4" title="Which test, and when?" lead="Pick the test type, the paper format and the date. Then add each phase: a course, a phase name and its batches. Each phase sits one paper." />
      <div className="card form">
        <label className="field">
          <span>Test date{test.type === 'SERIES' && test.days > 1 ? ' (day 1)' : ''}</span>
          <input type="date" value={test.date} onChange={(e) => setTest({ ...test, date: e.target.value })} autoFocus />
        </label>
        <label className="field grow">
          <span>Sheet title <em className="muted">(optional)</em></span>
          <input type="text" placeholder="e.g. Momentum Seating Arrangement" value={test.title} onChange={(e) => setTest({ ...test, title: e.target.value })} />
        </label>
      </div>

      <div className="slots-head"><h2>Test type</h2></div>
      <div className="choices">
        {TEST_TYPES.map((t) => (
          <label key={t.id} className={`choice ${test.type === t.id ? 'on' : ''}`}>
            <input type="radio" name="ttype" checked={test.type === t.id} onChange={() => setType(t.id)} />
            <b>{t.name}</b><span className="muted small">{t.text}</span>
          </label>
        ))}
      </div>
      <div className="slots-head"><h2>Paper format</h2></div>
      <div className="choices">
        {FORMATS.map((f) => (
          <label key={f.id} className={`choice ${test.format === f.id ? 'on' : ''}`}>
            <input type="radio" name="tformat" checked={test.format === f.id} onChange={() => setTest({ ...test, format: f.id })} />
            <b>{f.name}</b><span className="muted small">{f.text}</span>
          </label>
        ))}
      </div>
      {test.type === 'SERIES' && (
        <div className="card form">
          <span className="field"><span>Spread this test over</span>
            <span className="seg" role="group" aria-label="Number of days">
              {[1, 2, 3].map((d) => <button key={d} className={test.days === d ? 'on' : ''} onClick={() => setTest({ ...test, days: d })}>{d} {d === 1 ? 'day' : 'days'}</button>)}
            </span>
          </span>
          <p className="muted small grow">Days in a row from the test date. Each student sits once, on one day. Day 1 fills first, then day 2, then day 3.</p>
        </div>
      )}

      <div className="slots-head"><h2>{prt ? 'Online shifts (computer labs)' : 'Online slots (computer labs)'}</h2></div>
      <SlotList title={prt ? 'Shifts for online students' : 'Slots for all students'} slots={test.slots} onChange={(slots) => setTest({ ...test, slots })}
        hint={adv ? 'JEE Advanced: two slots in a row make one sitting (Paper 1 + Paper 2, same seat).' : prt ? 'Each online student sits one shift. Lab seats are used again in every shift.' : 'Each student sits one slot. Girls are never put in a slot marked No girls.'} />
      {!prt && (
        <SlotList title="Slots for top batches" slots={test.topSlots} onChange={(topSlots) => setTest({ ...test, topSlots })}
          hint="Top batches sit in the lab marked “Kept for top batches” (Buildings & rooms step)." />
      )}
      {units.unpaired.length > 0 && <div className="alert warn"><span>JEE Advanced needs two slots that do not overlap. {units.unpaired.map((s) => range(s)).join(', ')} has no partner and will not be used.</span></div>}
      {onlineSittings > 0 && <p className="muted small">{pl(onlineSittings, 'online sitting')}{test.days > 1 && test.type === 'SERIES' ? ` over ${test.days} days` : ''}.</p>}

      <div className="slots-head"><h2>Phases <span className="count">{test.papers.length}</span></h2></div>
      <details className="fold">
        <summary>Paste the phase list <span className="muted">· Course | Phase | Batches, one phase per line</span></summary>
        <p className="muted small">The same list as the Counselling Planner. Example: <code>Momentum | XA phase | XA1, XA2, XA3</code>. This replaces the phases below and ticks their batches.</p>
        <textarea rows="4" className="grow" value={paste} onChange={(e) => setPaste(e.target.value)} placeholder={'Momentum | XA phase | XA1, XA2\nMomentum | XB phase | XB1, XB2'} aria-label="Phase list" />
        <button type="button" className="btn primary sm" onClick={makePhases} disabled={!paste.trim()}>Make phases from this list</button>
      </details>
      <div className="papers">
        {test.papers.map((p, i) => {
          const times = adv ? [['s1', 'Paper 1'], ['s2', 'Paper 2']] : [['s1', (p.extra || []).length ? 'Sitting 1' : 'Time']];
          return (
            <div className="card paper" key={p.id}>
              <div className="paper-top">
                <span className="slot-n">{i + 1}</span>
                <label className="field"><span>Course</span>
                  <input type="text" placeholder="e.g. Momentum" value={p.course || ''} onChange={(e) => setPaper(p.id, { course: e.target.value })} /></label>
                <label className="field"><span>Phase</span>
                  <input type="text" placeholder="e.g. XA phase" value={p.phase || ''} onChange={(e) => setPaper(p.id, { phase: e.target.value })} /></label>
                <label className="field grow"><span>Paper name <em className="muted">(optional)</em></span>
                  <input type="text" placeholder={paperName({ ...p, name: '' }, i)} value={p.name} onChange={(e) => setPaper(p.id, { name: e.target.value })} /></label>
                <button type="button" className="iconbtn danger" disabled={test.papers.length === 1} onClick={() => removePaper(p.id)} aria-label={`Remove phase ${i + 1}`} title="Remove this phase">{Ico.bin}</button>
              </div>
              {prt && times.map(([k, label]) => {
                const s = p[k];
                const bad = !shiftOk(s) || (k === 's2' && mins(s.start) < mins(p.s1.end));
                return (
                  <div className={`slot ${bad ? 'bad' : ''}`} key={k}>
                    <b className="shift-l">Offline {label.toLowerCase()}</b>
                    <label className="field"><span>Starts</span><input type="time" value={s.start} onChange={(e) => setTime(p, k, 'start', e.target.value)} /></label>
                    <label className="field"><span>Ends</span><input type="time" value={s.end} onChange={(e) => setTime(p, k, 'end', e.target.value)} /></label>
                    <span className="slot-time muted">{!bad ? `${range(s)} · ${Math.round((mins(s.end) - mins(s.start)) / 6) / 10} h` : k === 's2' && shiftOk(s) ? 'Starts before Paper 1 ends' : 'Set the end time'}</span>
                  </div>
                );
              })}
              {prt && !adv && (p.extra || []).map((s, j) => (
                <div className={`slot ${shiftOk(s) ? '' : 'bad'}`} key={`x${j}`}>
                  <b className="shift-l">Offline sitting {j + 2}</b>
                  <label className="field"><span>Starts</span><input type="time" value={s.start} onChange={(e) => setExtra(p, j, 'start', e.target.value)} /></label>
                  <label className="field"><span>Ends</span><input type="time" value={s.end} onChange={(e) => setExtra(p, j, 'end', e.target.value)} /></label>
                  <span className="slot-time muted">{shiftOk(s) ? range(s) : 'Set the end time'}</span>
                  <button type="button" className="iconbtn danger" onClick={() => setPaper(p.id, { extra: p.extra.filter((_, x) => x !== j) })} aria-label="Remove this sitting">{Ico.bin}</button>
                </div>
              ))}
              {prt && !adv && (p.extra || []).length < 3 && (
                <button type="button" className="link small" onClick={() => setPaper(p.id, { extra: [...(p.extra || []), { start: hhmm(mins(p.s1.start) + 60), end: hhmm(mins(p.s1.end) + 60) }] })}>+ Add another offline sitting of this paper</button>
              )}
              {prt && adv && <p className="muted small">Both papers use the same seating: every student keeps the same seat. The Excel file has an attendance sheet for each paper.</p>}
              {!prt && <p className="muted small">Everyone in this phase sits online, in the slots above.</p>}
            </div>
          );
        })}
      </div>
      <button type="button" className="addslot" onClick={addPaper}>{Ico.plus}Add another phase</button>
      {prt && <p className="muted small">Offline rule: every offline sitting of a paper must start before its first sitting ends (7–10 and 8–11 is fine, 7–10 and 10–1 is not). Online shifts can be spread through the day.</p>}
      {timing.length > 0 && <div className="alert bad" role="alert"><span><b>Offline timing is not allowed.</b> {timing.join(' ')}</span></div>}
      <Nav back={back} next={next} nextOk={nextOk && !timing.length} hint={hint} />
    </section>
  );
}

// PRT: who sits online this time, from the previous seating files
function RotationPanel({ rotation, setRotation, onSeatingFiles, summary, onlineSeats }) {
  const inp = useRef(null);
  const total = summary.reduce((a, x) => a + x.onlineNow, 0);
  return (
    <div className="card special">
      <div className="special-head">
        <label className="far"><input type="checkbox" checked={rotation.on} onChange={(e) => setRotation((r) => ({ ...r, on: e.target.checked }))} /> <b>Choose online students by turn</b></label>
        <span className="muted small">Top batches always sit online. In other batches, students online last time sit offline now; a student is online again only after the whole batch had a turn.</span>
      </div>
      {rotation.on && (
        <>
          <div className="special-list">
            <span className="muted small grow">
              {rotation.files.length
                ? <>Previous seating: {rotation.files.map((f) => `${f.name}${f.date ? ` (${fmtDateDMY(f.date)})` : ''}`).join(', ')}. </>
                : 'Upload the seating Excel file of the last test (made by this planner). Add older files too, to follow the full turn. '}
            </span>
            <button type="button" className="btn ghost sm" onClick={() => inp.current.click()}>{Ico.upload}{rotation.files.length ? 'Add older files' : 'Upload previous seating'}</button>
            <input ref={inp} hidden type="file" accept=".xlsx" multiple onChange={(e) => { onSeatingFiles([...e.target.files]); e.target.value = ''; }} />
            {rotation.files.length > 0 && <button type="button" className="link small" onClick={() => setRotation((r) => ({ ...r, files: [] }))}>Clear files</button>}
          </div>
          {summary.length > 0 && (
            <>
              <p className="muted small">{total} of {onlineSeats} online seats go to these batches, shared by batch size. When the seats run out partway through a batch, only the rest of that batch sits offline.</p>
              <table className="rooms">
                <thead><tr><th>Batch</th><th>Students</th><th>Online now</th><th>Online last time</th><th>Still to get a turn</th></tr></thead>
                <tbody>{summary.map((x) => <tr key={x.batch}><td>{x.batch}</td><td>{x.size}</td><td>{x.onlineNow}</td><td>{x.onlineLast}</td><td>{x.stillAfter}</td></tr>)}</tbody>
              </table>
            </>
          )}
        </>
      )}
    </div>
  );
}

function BatchStep({ batches, studentsByBatch, selected, setSelected, mode, setMode, studentMode, setStudentMode, top, setTop, papers, multi, paperFor, setPaperOf, perPaper, need, emptyPapers, appearing, prt, rotation, setRotation, onSeatingFiles, rotSummary, onlineSeats, back, next, nextOk }) {
  const [q, setQ] = useState('');
  const [cur, setCur] = useState(papers[0].id);
  const [openBatch, setOpenBatch] = useState('');
  const curId = papers.some((p) => p.id === cur) ? cur : papers[0].id;
  const curIdx = papers.findIndex((p) => p.id === curId);
  const inOther = (b) => multi && selected.has(b) && paperFor(b) !== curId;
  const isOn = (b) => selected.has(b) && (!multi || paperFor(b) === curId);
  const tick = (b) => {
    if (isOn(b)) { setSelected((s) => { const n = new Set(s); n.delete(b); return n; }); return; }
    setSelected((s) => new Set(s).add(b));
    if (multi) setPaperOf((o) => ({ ...o, [b]: curId }));
  };
  const setAll = (list, on) => {
    const free = list.filter((b) => !inOther(b.batch)).map((b) => b.batch);
    setSelected((s) => { const n = new Set(s); free.forEach((b) => (on ? n.add(b) : n.delete(b))); return n; });
    if (on && multi) setPaperOf((o) => ({ ...o, ...Object.fromEntries(free.map((b) => [b, curId])) }));
  };
  // Course-wise Online / Offline for the ticked batches of a course
  const setCourseMode = (list, m) => setMode((o) => ({ ...o, ...Object.fromEntries(list.filter((b) => isOn(b.batch)).map((b) => [b.batch, m])) }));
  const setBatchMode = (b, m) => {
    setMode((o) => ({ ...o, [b]: m }));
    // students moved one by one go back to the batch's choice
    setStudentMode((o) => Object.fromEntries(Object.entries(o).filter(([roll]) => !(studentsByBatch[b] || []).some((s) => s.Roll_Number === roll))));
  };
  const changedIn = (b) => (studentsByBatch[b] || []).filter((s) => studentMode[s.Roll_Number] && studentMode[s.Roll_Number] !== (mode[b] || 'Offline')).length;
  const byCourse = {};
  for (const b of batches) if (!q || b.batch.toLowerCase().includes(q.toLowerCase())) (byCourse[b.course] ||= []).push(b);
  const curP = perPaper[curIdx];
  const nextP = papers[curIdx + 1];

  return (
    <section>
      <Head kicker="Step 3 of 4" title={multi ? 'Which batches are in each phase?' : 'Which batches are appearing?'}
        lead={`${multi ? 'Pick a phase, then tick its batches. A batch can be in one phase only. ' : ''}${prt ? 'Offline sits in classrooms, Online in computer labs. Set it for a whole course, a batch, or single students.' : 'Everyone sits online in this test type. Tick Top for top batches.'}`} />
      {multi && (
        <div className="slottabs" role="tablist">
          {perPaper.map((x, i) => (
            <button key={x.paper.id} role="tab" aria-selected={x.paper.id === curId} className={x.paper.id === curId ? 'on' : ''} onClick={() => setCur(x.paper.id)}>
              <b>{paperName(x.paper, i)}</b><span>{prt ? `${paperTime(x.paper)} · ` : ''}{x.list.length ? `${pl(x.list.length, 'batch')}, ${x.offline + x.online} students` : 'no batches yet'}</span>
            </button>
          ))}
        </div>
      )}
      {prt && <RotationPanel rotation={rotation} setRotation={setRotation} onSeatingFiles={onSeatingFiles} summary={rotSummary} onlineSeats={onlineSeats} />}
      <div className="toolbar">
        <input className="search" type="search" placeholder="Find a batch, e.g. XB" value={q} onChange={(e) => setQ(e.target.value)} />
        <div className="tally">{multi ? <>This paper: <b>{curP.offline}</b> offline · <b>{curP.online}</b> online</> : <><b>{need.offline}</b> offline · <b>{need.online}</b> online</>}</div>
      </div>
      {Object.entries(byCourse).map(([course, list]) => {
        const free = list.filter((b) => !inOther(b.batch));
        const allOn = free.length > 0 && free.every((b) => isOn(b.batch));
        const n = list.filter((b) => isOn(b.batch)).length;
        return (
          <div className="course" key={course}>
            <div className="course-head">
              <h2>{course} <span className="count">{n}/{list.length}</span></h2>
              {n > 0 && prt && !rotation.on && <span className="course-mode muted small">Ticked batches: <button className="link" onClick={() => setCourseMode(list, 'Offline')}>all Offline</button> · <button className="link" onClick={() => setCourseMode(list, 'Online')}>all Online</button></span>}
              {free.length > 0 && <button className="link" onClick={() => setAll(list, !allOn)}>{allOn ? 'Clear all' : 'Select all'}</button>}
            </div>
            <div className="batch-grid">
              {list.map((b) => {
                const on = isOn(b.batch);
                const other = inOther(b.batch);
                const otherIdx = other ? papers.findIndex((p) => p.id === paperFor(b.batch)) : -1;
                const changed = on ? changedIn(b.batch) : 0;
                return (
                  <div key={b.batch} className={`batch ${on ? 'on' : ''} ${other ? 'other' : ''}`}>
                    <label className="batch-top" title={other ? `In ${paperName(papers[otherIdx], otherIdx)}. Tick to move it to this paper.` : ''}>
                      <input type="checkbox" checked={on} onChange={() => tick(b.batch)} />
                      <span className="batch-name">{b.batch}</span>
                      {other ? <span className="pill muted-pill">{paperName(papers[otherIdx], otherIdx)}</span> : <span className="batch-count">{b.count}</span>}
                    </label>
                    {on && (
                      <div className="batch-opts">
                        {!prt || top.has(b.batch) ? <span className="pill">Online</span>
                          : rotation.on ? <span className="pill" title="Online or offline by turn">By turn</span>
                            : (
                              <div className="seg">
                                {['Offline', 'Online'].map((m) => (
                                  <button key={m} className={(mode[b.batch] || 'Offline') === m ? 'on' : ''} onClick={() => setBatchMode(b.batch, m)}>{m}</button>
                                ))}
                              </div>
                            )}
                        {prt && !top.has(b.batch) && <button className="link small" onClick={() => setOpenBatch(openBatch === b.batch ? '' : b.batch)}>{changed ? `${changed} moved` : 'Students'}</button>}
                        <label className="small top-tick" title="Top batch students can be kept in chosen rooms (Buildings & rooms step)">
                          <input type="checkbox" checked={top.has(b.batch)} onChange={() => setTop((t) => { const x = new Set(t); x.has(b.batch) ? x.delete(b.batch) : x.add(b.batch); return x; })} /> Top
                        </label>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
            {prt && list.some((b) => b.batch === openBatch && isOn(b.batch)) && (
              <StudentModes batch={openBatch} students={studentsByBatch[openBatch] || []} batchMode={mode[openBatch] || 'Offline'}
                studentMode={studentMode} setStudentMode={setStudentMode} close={() => setOpenBatch('')} />
            )}
          </div>
        );
      })}

      <div className="ctabar">
        <div>
          <b>{!nextOk ? 'Tick at least one batch' : 'Next: choose the buildings'}</b>
          <span className="muted">
            {appearing} students{multi ? ` in ${pl(papers.length, 'paper')}` : ''}.
            {emptyPapers.length > 0 && ` ${emptyPapers.map((x) => paperName(x.paper, papers.indexOf(x.paper))).join(', ')} ${emptyPapers.length === 1 ? 'has' : 'have'} no batches yet.`}
          </span>
        </div>
        <div className="cta-row">
          {multi && nextP && <button className="btn ghost" onClick={() => setCur(nextP.id)}>Next: {paperName(nextP, curIdx + 1)} →</button>}
          <button className="btn cta" onClick={next} disabled={!nextOk}>Buildings & rooms →</button>
        </div>
      </div>
      <div className="nav"><button className="btn ghost" onClick={back}>Back</button></div>
    </section>
  );
}

// Online / Offline for single students inside a batch
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

function RoomStep({ settings, setSettings, need, have, chosen, appearing, lists, setLists, onRollFile, flags, chosenRolls, onRoomFile, canCreate, useOrder, test, setTest, prt, topStudents, topCap, units, antiCheat, setAntiCheat, timing, back, create, hasResult }) {
  const editRoom = (kind, i, patch) => setSettings((s) => ({ ...s, [kind]: s[kind].map((r, j) => (j === i ? { ...r, ...patch } : r)) }));
  const move = (i, d) => setSettings((s) => {
    const o = s.buildingOrder.slice(); const j = i + d;
    if (j < 0 || j >= o.length) return s;
    [o[i], o[j]] = [o[j], o[i]]; return { ...s, buildingOrder: o };
  });
  const toggleUse = (b) => setSettings((s) => ({ ...s, buildingsOn: s.buildingsOn.includes(b) ? s.buildingsOn.filter((x) => x !== b) : [...s.buildingsOn, b] }));
  const toggleGirls = (b) => setSettings((s) => ({ ...s, girlsOk: { ...s.girlsOk, [b]: s.girlsOk[b] === false } }));
  const offShort = Math.max(0, need.offline - have.offline), onShort = Math.max(0, need.online - have.online);
  const gOffShort = Math.max(0, need.girlsOffline - have.girlsOffline), gOnShort = Math.max(0, need.girlsOnline - have.girlsOnline);
  const short = offShort > 0 || onShort > 0 || gOffShort > 0 || gOnShort > 0;
  const seatsIn = (b) => settings.rooms.filter((r) => r.building === b && r.enabled).reduce((a, r) => a + Number(r.capacity || 0), 0);
  const labSeatsIn = (b) => settings.labs.filter((r) => r.building === b && r.enabled).reduce((a, r) => a + Number(r.capacity || 0), 0);
  const roomsOn = settings.rooms.filter((r) => r.enabled && useOrder.includes(r.building)).length;
  const labsOn = settings.labs.filter((r) => r.enabled && useOrder.includes(r.building)).length;
  // rooms that can be kept for special groups: switched-on rooms and labs in the chosen buildings
  const pinChoices = [...settings.rooms.map((r) => ({ ...r, kind: 'Room' })), ...settings.labs.map((r) => ({ ...r, kind: 'Lab' }))]
    .filter((r) => r.enabled && useOrder.includes(r.building))
    .sort((a, b) => useOrder.indexOf(a.building) - useOrder.indexOf(b.building) || naturalCompare(a.room, b.room));
  const flagCount = (f) => [...chosenRolls].filter((r) => flags[r] === f).length;
  const setPins = (f, keys) => setSettings((s) => ({ ...s, pins: { ...s.pins, [f]: keys } }));
  const groups = settings.cbGroups;
  const setGroups = (g) => setSettings((s) => ({ ...s, cbGroups: g }));

  return (
    <section>
      <Head kicker="Step 4 of 4" title="Buildings and rooms" lead="Tick the buildings to use today and put them in fill order. Seats fill the first building completely, then the next. Changes are saved in this browser." />
      <div className="meters">
        {prt && <Meter label="Classroom seats" need={need.offline} have={have.offline} />}
        <Meter label={`Computer lab seats (${pl(units.general.length + (test.topShare ? units.top.length : 0), 'sitting')})`} need={need.online} have={have.online} />
        {prt && <Meter label="Seats for girls (classrooms)" need={need.girlsOffline} have={have.girlsOffline} />}
        {need.online > 0 && <Meter label="Seats for girls (labs)" need={need.girlsOnline} have={have.girlsOnline} />}
      </div>
      {short && (
        <div className="alert bad"><span>
          {offShort > 0 && <><b>{offShort}</b> offline students more than classroom seats. </>}
          {onShort > 0 && <><b>{onShort}</b> online students more than lab seats. </>}
          {gOffShort > 0 && <><b>{gOffShort}</b> girls more than classroom seats where girls are allowed. </>}
          {gOnShort > 0 && <><b>{gOnShort}</b> girls more than lab seats where girls are allowed. </>}
          Tick more buildings, switch on rooms, or allow girls in another building; otherwise they go to a Not_Seated sheet.
        </span></div>
      )}

      <div className="card">
        <h2>Buildings to use, in fill order <span className="count">{useOrder.length}</span></h2>
        <ol className="order">
          {settings.buildingOrder.map((b, i) => {
            const on = settings.buildingsOn.includes(b);
            return (
              <li key={b} className={on ? '' : 'off'}>
                <input type="checkbox" checked={on} onChange={() => toggleUse(b)} aria-label={`Use ${b}`} />
                <span className="order-n">{on ? useOrder.indexOf(b) + 1 : '–'}</span>
                <span className="order-name">{b}</span>
                <span className="muted small">{seatsIn(b)} seats{labSeatsIn(b) ? ` · lab ${labSeatsIn(b)}` : ''}</span>
                <label className="far"><input type="checkbox" checked={settings.girlsOk[b] !== false} onChange={() => toggleGirls(b)} /> Girls allowed</label>
                <button className="iconbtn" onClick={() => move(i, -1)} aria-label={`Move ${b} up`} disabled={i === 0}>▲</button>
                <button className="iconbtn" onClick={() => move(i, 1)} aria-label={`Move ${b} down`} disabled={i === settings.buildingOrder.length - 1}>▼</button>
              </li>
            );
          })}
        </ol>
      </div>

      <details className="fold" open={offShort > 0 || onShort > 0}>
        <summary>Rooms on or off, and capacities <span className="muted">· {roomsOn} classrooms and {labsOn} labs on in the chosen buildings</span></summary>
        <RoomTable title="Classrooms (offline)" kind="rooms" list={settings.rooms} edit={editRoom} use={useOrder} />
        <RoomTable title="Computer labs (online)" kind="labs" list={settings.labs} edit={editRoom} use={useOrder} labs topCol={!prt} />
      </details>

      {!prt && (
        <div className="card special">
          <div className="special-head"><b>Lab kept for top batches</b><span className="muted small">Mark it under Rooms on or off → Computer labs → Kept for top batches.</span></div>
          {units.top.length ? (
            <>
              <p className="muted small">{topStudents} top batch students · {topCap} seats over {pl(units.top.length, 'sitting')} ({test.topSlots.map((sl) => range(sl)).join(', ')}) · <b>{Math.max(0, topCap - topStudents)} empty seats</b></p>
              <label className="far"><input type="checkbox" checked={!!test.topShare} onChange={(e) => setTest({ ...test, topShare: e.target.checked })} /> Let other students fill the empty seats in this lab</label>
            </>
          ) : <p className="muted small">No lab is marked yet{topStudents ? `, so the ${topStudents} top batch students sit with everyone else` : ''}.</p>}
        </div>
      )}

      <div className="card">
        <h2>Stop copying: pick one</h2>
        <div className="choices">
          {ANTI_CHEAT.map((o) => (
            <label key={o.id} className={`choice ${antiCheat === o.id ? 'on' : ''}`}>
              <input type="radio" name="anticheat" checked={antiCheat === o.id} onChange={() => setAntiCheat(o.id)} />
              <b>{o.name}</b><span className="muted small">{o.text}</span>
            </label>
          ))}
        </div>
      </div>

      <details className="fold">
        <summary>Monitoring, due fee and top batch students <span className="muted">· {FLAGS.map((f) => `${f} ${flagCount(f)}`).join(' · ')}</span></summary>
        <p className="muted">Keep rooms for these students. They are seated there first; others fill those rooms only after every other room is full. A student in two lists counts once, in the first: Monitoring, then Due fee, then Top batch.</p>
        {FLAGS.map((f) => (
          <div className="card special" key={f}>
            <div className="special-head"><b>{f}</b><span className="muted small">{flagCount(f)} appearing today</span></div>
            {f !== 'Top batch' ? (
              <div className="special-list">
                <textarea rows="2" placeholder="Paste roll numbers, or upload a file" value={lists[f].join(', ')} onChange={(e) => setLists((l) => ({ ...l, [f]: parseRolls(e.target.value) }))} aria-label={`${f} roll numbers`} />
                <RollUpload onFile={(file) => onRollFile(f, file)} />
              </div>
            ) : <p className="muted small">Mark top batches with Top on the Batches step.</p>}
            <PinPicker choices={pinChoices} value={settings.pins[f] || []} onChange={(keys) => setPins(f, keys)} label={f} />
          </div>
        ))}
      </details>

      <details className="fold">
        <summary>SMD sheet groups (CB column) <span className="muted">· {groups.length ? pl(groups.length, 'group') : 'by first letter of the batch'}</span></summary>
        <p className="muted">The SMD sheet counts students per room for each group. Batches not in a group are counted by the first letter of the batch (X-Phase, Y-Phase…).</p>
        {groups.map((g, i) => (
          <div className="group-row" key={i}>
            <input type="text" placeholder="Group name, e.g. Momentum XA" value={g.name} onChange={(e) => setGroups(groups.map((x, j) => (j === i ? { ...x, name: e.target.value } : x)))} aria-label="Group name" />
            <input type="text" className="grow" placeholder="Batches, e.g. XA1, XA2, XA3" value={g.batches} onChange={(e) => setGroups(groups.map((x, j) => (j === i ? { ...x, batches: e.target.value } : x)))} aria-label="Batches in this group" />
            <button className="iconbtn danger" onClick={() => setGroups(groups.filter((_, j) => j !== i))} aria-label="Remove group">{Ico.bin}</button>
          </div>
        ))}
        <button type="button" className="addslot" onClick={() => setGroups([...groups, { name: '', batches: chosen.map((b) => b.batch).join(', ') }])}>{Ico.plus}Add a group</button>
      </details>

      <details className="fold">
        <summary>Room list <span className="muted">· {settings.rooms.length} rooms, {settings.labs.length} labs</span></summary>
        <p className="muted">Built in already. Upload a new file only when rooms change. Uploads stay saved in this browser.</p>
        <div className="files">
          <FileCard title="Room list" hint="Sheets Room_Capacity and LAB_Capacity" onFile={onRoomFile} template={() => downloadTemplate('rooms')} />
        </div>
        <button className="btn ghost danger" onClick={() => { if (confirm('Reset buildings, rooms, capacities, girls settings and groups to the built-in lists?')) setSettings(freshSettings()); }}>{Ico.bin}Reset to built-in lists</button>
      </details>

      <div className="ctabar">
        <div><b>{!useOrder.length ? 'Tick at least one building' : timing.length ? 'Fix the offline timing (Test & phases step)' : !antiCheat ? 'Pick a way to stop copying' : short ? 'Some students will have no seat' : 'Ready to create the seating'}</b><span className="muted">{appearing} students in {chosen.length} batches · {useOrder.join(' → ') || 'no buildings'}</span></div>
        <div className="cta-row">
          <button className="btn ghost" onClick={back}>Back</button>
          <button className="btn cta" onClick={create} disabled={!canCreate}>{Ico.spark}{hasResult ? 'Create again' : 'Create seating'}</button>
        </div>
      </div>
    </section>
  );
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

function PinPicker({ choices, value, onChange, label }) {
  const keyOf = (r) => roomKey(r);
  const picked = value.filter((k) => choices.some((r) => keyOf(r) === k));
  const seats = choices.filter((r) => picked.includes(keyOf(r))).reduce((a, r) => a + Number(r.capacity || 0), 0);
  return (
    <div className="pins">
      <select value="" onChange={(e) => e.target.value && onChange([...picked, e.target.value])} aria-label={`Add a room for ${label}`}>
        <option value="">Keep a room for {label.toLowerCase()}…</option>
        {choices.filter((r) => !picked.includes(keyOf(r))).map((r) => <option key={keyOf(r)} value={keyOf(r)}>{r.building} · {r.kind} {r.room} · {r.capacity} seats</option>)}
      </select>
      {picked.map((k) => <span className="pill" key={k}>{k.replace('|', ' · ')} <button className="link" onClick={() => onChange(picked.filter((x) => x !== k))} aria-label={`Remove ${k}`}>×</button></span>)}
      {picked.length > 0 && <span className="muted small">{seats} seats kept</span>}
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

function RoomTable({ title, kind, list, edit, use, labs, topCol }) {
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
              <thead><tr><th>Use</th><th>{labs ? 'Lab' : 'Room'}</th><th>Capacity</th>{labs && <th>Girls allowed</th>}{topCol && <th>Kept for top batches</th>}<th>Note</th></tr></thead>
              <tbody>
                {rooms.map((r) => (
                  <tr key={r.i} className={r.enabled ? '' : 'off'}>
                    <td><input type="checkbox" checked={r.enabled} onChange={(e) => edit(kind, r.i, { enabled: e.target.checked })} aria-label={`Use ${b} ${r.room}`} /></td>
                    <td>{r.room}</td>
                    <td><input className="cap" type="number" min="0" value={r.capacity} onChange={(e) => edit(kind, r.i, { capacity: Math.max(0, Number(e.target.value)) })} aria-label={`Capacity of ${b} ${r.room}`} /></td>
                    {labs && <td><input type="checkbox" checked={r.girls !== false} onChange={(e) => edit(kind, r.i, { girls: e.target.checked })} aria-label={`Girls allowed in ${b} ${r.room}`} /></td>}
                    {topCol && <td><input type="checkbox" checked={!!r.top} onChange={(e) => edit(kind, r.i, { top: e.target.checked })} aria-label={`${b} ${r.room} kept for top batches`} /></td>}
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

function ResultStep({ results, view, setView, metaFor, multi, order, planWarnings, seed, antiCheatName, reshuffle, download, downloadBoard, back }) {
  const [menu, setMenu] = useState(false);
  const x = results[view] || results[0];
  const result = x.result, meta = metaFor(x);
  const byBuilding = {};
  for (const r of result.usedRooms.slice().sort((a, b) => order.indexOf(a.building) - order.indexOf(b.building) || (a.kind === 'Online') - (b.kind === 'Online'))) (byBuilding[r.building] ||= []).push(r);
  const special = result.assigned.filter((s) => s.Flag).length;
  const girls = result.assigned.filter(isGirl).length;
  const unseatedAll = results.reduce((a, y) => a + y.result.unseated.length, 0);
  const files = results.filter((y) => y.students).length;
  return (
    <section>
      <Head kicker={multi ? `Seating ready · ${pl(results.length, 'sitting')}` : 'Seating ready'} title={meta.title}
        lead={`${fmtDateDMY(meta.dateISO)} · ${x.time} · ${meta.modeLabel}${meta.batchesLabel ? ` · Batches ${meta.batchesLabel}` : ''}${antiCheatName ? ` · ${antiCheatName} (seed ${seed})` : ''}`} />
      {planWarnings.map((w) => <div key={w} className="alert warn"><span>{w}</span></div>)}
      {multi && (
        <div className="slottabs" role="tablist">
          {results.map((y, i) => (
            <button key={y.key} role="tab" aria-selected={i === view} className={`${i === view ? 'on' : ''} ${y.result.unseated.length ? 'warn' : ''}`} onClick={() => setView(i)}>
              <b>{y.name}</b>
              <span>{y.time}{y.dateISO && y.day > 1 ? ` · ${fmtDateDMY(y.dateISO)}` : ''} · {y.students ? `${y.result.assigned.length} seated${y.result.unseated.length ? ` · ${y.result.unseated.length} no seat` : ''}` : 'nobody yet'}{y.emptySeats ? ` · ${y.emptySeats} empty` : ''}</span>
            </button>
          ))}
        </div>
      )}
      {unseatedAll > 0 && (
        <div className="alert bad"><span><b>{unseatedAll} students have no seat{multi ? ' across the papers' : ''}.</b> They are listed in the Not_Seated sheet with the reason.</span><button className="btn sm" onClick={back}>Fix buildings</button></div>
      )}
      {x.overlapsWith.length > 0 && (
        <div className="alert info"><span>This sitting overlaps {x.overlapsWith.join(', ')} (or starts less than {GAP} minutes after it). It uses the seats left free in shared rooms, so no seat is used twice at once.</span></div>
      )}
      {x.tag === 'top' && <div className="alert info"><span>Lab kept for top batches: {x.topCount || 0} top batch students{x.students > (x.topCount || 0) ? `, ${x.students - x.topCount} other students` : ''} · <b>{x.emptySeats} empty seats</b>.</span></div>}
      {x.noGirls && <div className="alert info"><span>No girls in this slot.</span></div>}
      {!x.students && !x.result.unseated.length && <div className="alert warn"><span>Nobody sits in this sitting.</span></div>}
      <div className="actions">
        {multi
          ? <><button className="btn cta" onClick={() => download(results)} disabled={!files}>{Ico.down}Download all ({files} {files === 1 ? 'file' : 'files'})</button>
              <button className="btn primary" onClick={() => download([x])} disabled={!x.students}>{Ico.down}Only this sitting</button></>
          : <button className="btn cta" onClick={() => download([x])}>{Ico.down}Download Excel</button>}
        <button className="btn ghost" onClick={downloadBoard} disabled={!files}>{Ico.down}Notice board{multi ? ' (all papers)' : ''}</button>
        <div className="menu">
          <button className="btn ghost" onClick={() => setMenu(!menu)} aria-expanded={menu}>More ▾</button>
          {menu && (
            <div className="menu-list" onMouseLeave={() => setMenu(false)}>
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
        <div><b>{x.slots.length}</b><span>attendance {x.slots.length > 1 ? 'sheets (same seats)' : 'sheet'}</span></div>
      </div>
      {result.warnings.map((w) => <div key={w} className="alert warn"><span>{w}</span></div>)}

      <details className="fold" open>
        <summary>Rooms <span className="muted">· {result.usedRooms.length} used</span></summary>
        <p className="muted small">Green: students of this {multi ? 'sitting' : 'test'}. Amber: monitoring, due fee or top batch students.{x.overlapsWith.length ? ' Grey: seats used by an overlapping paper.' : ''}</p>
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
      <p className="muted small">{multi ? 'Each sitting file has' : 'The Excel file has'}: Raw_Data, Student_list, Seating Arrangement, {x.slots.length > 1 ? 'an attendance sheet per paper' : 'Attendance'}, SMD, Room_Layout, Master_List{result.unseated.length ? ', Not_Seated' : ''}. The notice board file puts every sitting on one sheet. Keep the files: next time, upload them to choose online students by turn.</p>
    </section>
  );
}
