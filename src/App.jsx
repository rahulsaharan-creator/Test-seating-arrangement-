import React, { useEffect, useMemo, useRef, useState } from 'react';
import defaults from './data/defaults.json';
import { allocate, capacity, naturalCompare, isGirl, roomKey, FLAGS } from './lib/engine.js';
import { buildWorkbook, buildNoticeBoard, fmtDateDMY, fmtTime } from './lib/exporter.js';
import { readAdmissions, readRooms, readRollList, parseRolls } from './lib/readers.js';
import { parsePhaseList, groupPhases, seatSummary, fillPreview } from './lib/plan.js';

const loadExcel = () => import('exceljs').then((m) => m.default || m);
const STORE = 'seating-settings-v1';
const LAST = 'seating-last-test-v3';
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
// A paper: name, phases, mode (Online / Offline / Both), shifts (1 or 2), times, online batches (Both), labs and removed rooms
const newPaper = (start = '09:00', end = '12:00') => ({ id: uid(), name: '', phases: [], mode: '', shifts: 0, onlineBatches: [], labs: null, roomsOff: [], done: false, two: false, s1: { start, end }, online: { start, end }, s2: { start: hhmm(mins(end) + 30), end: hhmm(mins(end) + 30 + (mins(end) - mins(start))) } });
function initialTest() {
  const last = loadLast();
  if (last && Array.isArray(last.papers) && last.papers.length) return { title: last.title || '', date: '', papers: last.papers.map((p) => ({ ...newPaper(), ...p, done: false })) };
  return { title: '', date: '', papers: [newPaper()] };
}

const pl = (n, w) => `${n.toLocaleString('en-IN')} ${w}${n === 1 ? '' : /(ch|sh|s|x)$/.test(w) ? 'es' : 's'}`;
const mins = (t) => { const [h, m] = String(t || '0:0').split(':').map(Number); return h * 60 + m; };
const hhmm = (m) => `${String(Math.floor(Math.min(m, 1439) / 60)).padStart(2, '0')}:${String(Math.min(m, 1439) % 60).padStart(2, '0')}`;
// A paper's shifts. Two shifts: the same students and the same seats, an attendance sheet for each shift.
const shiftsOf = (p) => (p.two ? [{ label: 'Shift 1', ...p.s1 }, { label: 'Shift 2', ...p.s2 }] : [{ label: '', ...p.s1 }]);
// Online time of a paper: online students may sit at a different time from offline students (default: same as Shift 1)
const onlineOf = (p) => (p.mode === 'Both' && p.online && p.online.start ? p.online : p.s1);
const windowOf = (p) => {
  const sh = shiftsOf(p), on = onlineOf(p);
  return { start: hhmm(Math.min(mins(sh[0].start), mins(on.start))), end: hhmm(Math.max(mins(sh[sh.length - 1].end), mins(on.end))) };
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
  { id: 'test', label: 'Test & papers' },
  { id: 'labs', label: 'Labs' },
  { id: 'rooms', label: 'Buildings & rooms' },
  { id: 'rules', label: 'Shuffle rules' },
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
  const [results, setResults] = useState(null); // [{ paper, index, result, batches, students, overlapsWith }]
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
  const phaseList = useMemo(() => parsePhaseList(settings.phaseText || ''), [settings.phaseText]);
  const phases = useMemo(() => groupPhases(batches, phaseList.phases), [batches, phaseList]);

  const papers = test.papers;
  const multi = papers.length > 1;
  const curP = papers.find((p) => p.id === cur) || null;
  const curIdx = papers.findIndex((p) => p.id === cur);
  // Any change to a paper means its seating must be prepared again
  const setPaper = (id, patch) => setTest((t) => ({ ...t, papers: t.papers.map((p) => (p.id === id ? { ...p, ...patch, done: false } : p)) }));

  const ownerOf = {};
  for (const p of papers) for (const k of p.phases) ownerOf[k] ||= p.id;
  const batchesOf = (p) => phases.filter((g) => p.phases.includes(g.key)).flatMap((g) => g.batches);
  const batchModeOf = (p, b) => (p.mode === 'Online' ? 'Online' : p.mode === 'Offline' ? 'Offline' : p.onlineBatches.includes(b) ? 'Online' : 'Offline');
  const studentModeOf = (p, s) => (p.mode === 'Both' && studentMode[s.Roll_Number]) || batchModeOf(p, s.Batch);
  const labsOn = settings.labs.filter((l) => l.enabled && Number(l.capacity) > 0);
  const orderOn = settings.buildingOrder.filter((b) => settings.buildingsOn.includes(b));
  // Labs: the operator picks them per paper (until then, every lab counts for the estimate). Rooms: all rooms of the
  // buildings in use, minus the rooms the operator removed for this paper.
  const chosenLabs = (p) => (p.labs ? labsOn.filter((l) => p.labs.includes(roomKey(l))) : labsOn);
  const chosenRooms = (p) => settings.rooms.filter((r) => r.enabled && Number(r.capacity) > 0 && orderOn.includes(r.building) && !p.roomsOff.includes(roomKey(r)));
  const sumSeats = (list) => list.reduce((a, r) => a + Number(r.capacity || 0), 0);
  const countOf = (p) => {
    const c = { online: 0, offline: 0 };
    for (const b of batchesOf(p)) for (const s of studentsByBatch[b.batch] || []) c[studentModeOf(p, s) === 'Online' ? 'online' : 'offline']++;
    return c;
  };
  const summaryOf = (p) => seatSummary({ ...countOf(p), labSeats: sumSeats(chosenLabs(p)), roomSeats: sumSeats(chosenRooms(p)) });
  const paperOk = (p) => !!(p.name.trim() && p.phases.length && p.mode && p.shifts && shiftOk(p.s1)
    && (p.shifts === 1 || (shiftOk(p.s2) && mins(p.s2.start) >= mins(p.s1.end))) && (p.mode !== 'Both' || shiftOk(onlineOf(p))));
  const curSum = curP ? summaryOf(curP) : null;

  // Special students: Monitoring and Due fee from lists, Top batch from ticked top batches
  const flags = useMemo(() => {
    const f = {};
    for (const b of top) for (const s of studentsByBatch[b] || []) f[s.Roll_Number] = 'Top batch';
    for (const r of lists['Due fee']) f[r] = 'Due fee';
    for (const r of lists.Monitoring) f[r] = 'Monitoring';
    return f;
  }, [top, lists, studentsByBatch]);
  const appearingRolls = new Set(papers.flatMap((p) => batchesOf(p).flatMap((b) => (studentsByBatch[b.batch] || []).map((s) => s.Roll_Number))));
  const appearing = appearingRolls.size;
  const flagCount = (f) => [...appearingRolls].filter((r) => flags[r] === f).length;
  // Monitoring / Due fee students with no room kept for them yet
  const unassigned = ['Monitoring', 'Due fee'].map((f) => ({ flag: f, n: (settings.pins[f] || []).length ? 0 : flagCount(f) })).filter((x) => x.n);

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
      setStep('test');
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

  // "Create seating for this test": labs first (if anyone sits online), then rooms, then shuffle rules
  function startPaper(p) {
    setCur(p.id);
    if (!p.labs) setPaper(p.id, { labs: [] });
    setStep(countOf(p).online > 0 ? 'labs' : 'rooms');
  }

  // One seating per prepared paper, in time order. A paper that overlaps an earlier one (or starts within GAP minutes
  // of it) starts each room's seat numbers after the seats that paper already uses, so no seat is used twice at once.
  function run(newSeed, list) {
    const cbOf = {};
    for (const g of settings.cbGroups) for (const b of parseRolls(g.batches)) if (g.name.trim()) cbOf[b] = g.name.trim();
    const order = list.map((p) => ({ p, i: list.indexOf(p) })).filter(({ p }) => p.done)
      .sort((a, b) => mins(windowOf(a.p).start) - mins(windowOf(b.p).start) || a.i - b.i);
    const done = [];
    for (const { p, i } of order) {
      const bs = batchesOf(p);
      const taken = {};
      const over = done.filter((d) => overlaps(d.paper, p));
      for (const d of over) for (const room of d.result.usedRooms) {
        const k = roomKey(room);
        taken[k] = Math.max(taken[k] || 0, room.taken + room.used);
      }
      const st = bs.flatMap((b) => studentsByBatch[b.batch] || []);
      const labKeys = new Set(chosenLabs(p).map(roomKey)), roomKeys = new Set(chosenRooms(p).map(roomKey));
      const cfg = {
        rooms: settings.rooms.map((r) => ({ ...r, enabled: roomKeys.has(roomKey(r)) })),
        labs: settings.labs.map((l) => ({ ...l, enabled: labKeys.has(roomKey(l)) })),
        buildingOrder: settings.buildingOrder, girlsOk: settings.girlsOk,
        batchMode: Object.fromEntries(bs.map((b) => [b.batch, batchModeOf(p, b.batch)])),
        studentMode: p.mode === 'Both' ? studentMode : {}, flags, pins: settings.pins, cbOf,
        seatOrder: settings.seatOrder || 'mixed', seed: newSeed + i, taken,
      };
      done.push({ paper: p, index: i, result: allocate(st, cfg), batches: bs, students: st.length, overlapsWith: over.map((d) => d.index) });
    }
    setResults(done.sort((a, b) => a.index - b.index)); setSeed(newSeed); setView(0); setStep('result');
    try { localStorage.setItem(LAST, JSON.stringify({ title: test.title, papers: list.map((p) => ({ ...p, done: false })), top: [...top] })); } catch { /* ignore */ }
  }
  function prepare() {
    const list = test.papers.map((p) => (p.id === cur ? { ...p, done: true } : p));
    setTest((t) => ({ ...t, papers: list }));
    run(seed, list);
  }

  const metaFor = (x) => {
    const courses = [...new Set(x.batches.map((b) => b.course))];
    const n = { online: x.result.assigned.filter((s) => s.Mode === 'Online').length, offline: x.result.assigned.filter((s) => s.Mode !== 'Online').length };
    const head = test.title || `${courses.join(' & ') || 'Test'} Seating Arrangement`;
    const sh = shiftsOf(x.paper);
    return {
      title: multi ? `${head} - ${paperName(x.paper, x.index)}` : head,
      dateISO: test.date,
      shifts: sh.map((s) => s.start),
      slots: sh,
      batchesLabel: [...new Set(x.batches.map((b) => b.batch.slice(0, 2)))].sort().join(', '),
      modeLabel: n.online && n.offline ? `Offline + Online Test (Online ${range(onlineOf(x.paper))})` : n.online ? `Online Test (${range(onlineOf(x.paper))})` : 'Offline Test',
    };
  };

  const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  async function download(list) {
    const todo = list.filter((x) => x.batches.length);
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
      const wb = buildNoticeBoard(ExcelJS, results.filter((x) => x.batches.length).map((x) => ({ result: x.result, meta: metaFor(x) })), { dateISO: test.date, dayTitle });
      saveBlob(new Blob([await wb.xlsx.writeBuffer()], { type: XLSX }), `Notice_board_${fmtDateDMY(test.date)}.xlsx`);
    } catch (e) { setError(`Could not build the notice board: ${e.message}`); }
    setBusy('');
  }

  const setupOk = !!adm && !!test.date && !!curP && paperOk(curP);
  const canOpen = (id) => (id === 'students' ? true : id === 'test' ? !!adm : id === 'result' ? !!results : setupOk);
  const doneMark = (id) => (id === 'students' ? !!adm : id === 'test' ? papers.some((p) => p.done) : false);
  const labsNeeded = curSum ? curSum.online > 0 : false;
  const afterLabs = () => setStep(curSum && curSum.offline > 0 ? 'rooms' : 'rules');

  return (
    <div className="app">
      <header className="top">
        <div className="top-in">
          <div className="logo">
            <img src="/brand/matrix-logo-green.png" alt="Matrix Academy" className="logo-img" />
            <span className="sep" aria-hidden="true" />
            <div className="brand"><b>Test <em>Seating</em></b><span className="chip">{test.date ? `${fmtDateDMY(test.date)} · ${pl(papers.length, 'paper')}` : 'No test date yet'}</span></div>
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
          <SummaryBar adm={adm} papers={papers} appearing={appearing} curP={curP} curIdx={curIdx} sum={curSum} unassigned={unassigned}
            results={step === 'result' ? results : null} goRules={setupOk ? () => setStep('rules') : null} />
        )}
      </header>

      <main className="main">
        {error && <div className="alert bad" role="alert"><span>{error}</span><button className="link" onClick={() => setError('')}>Dismiss</button></div>}
        {notice && !error && <div className="alert info"><span>{notice}</span><button className="link" onClick={() => setNotice('')}>OK</button></div>}
        {busy && <div className="busy" role="status"><span className="spin" aria-hidden="true" />{busy}</div>}

        {step === 'students' && (
          <section>
            <Head kicker="Step 1" title="Upload the admission list" lead="The Excel file with every student. Only students marked Active are used. The file is read inside this browser and is never uploaded anywhere." />
            {!adm && <Guide />}
            <FileCard title="Admission list" hint="Columns: Roll_Number, Student_Name, Gender, Course, Batch, Medium, Status"
              loaded={adm && <><b>{adm.fileName}</b><span>{pl(adm.students.length, 'active student')} · {pl(batches.length, 'batch')}{adm.inactive ? ` · ${adm.inactive} inactive left out` : ''}{adm.duplicates ? ` · ${adm.duplicates} duplicate rolls skipped` : ''}</span></>}
              onFile={onAdmission} onRemove={removeAdmission} template={() => downloadTemplate('admission')} big />
            {adm && <Nav next={() => setStep('test')} />}
          </section>
        )}

        {step === 'test' && (
          <PaperStep test={test} setTest={setTest} setPaper={setPaper} phases={phases} phaseIssues={phaseList.issues} ownerOf={ownerOf}
            studentsByBatch={studentsByBatch} studentMode={studentMode} setStudentMode={setStudentMode} top={top} setTop={setTop}
            summaryOf={summaryOf} paperOk={paperOk} start={startPaper} settings={settings} setSettings={setSettings}
            back={() => setStep('students')} />
        )}

        {step === 'labs' && curP && (
          <LabStep p={curP} idx={curIdx} labsOn={labsOn} order={settings.buildingOrder} setPaper={setPaper} sum={curSum}
            back={() => setStep('test')} next={afterLabs} />
        )}

        {step === 'rooms' && curP && (
          <RoomStep p={curP} idx={curIdx} settings={settings} setSettings={setSettings} setPaper={setPaper} sum={curSum}
            rooms={chosenRooms(curP)} orderOn={orderOn} onRoomFile={onRoomFile}
            back={() => setStep(labsNeeded ? 'labs' : 'test')} next={() => setStep('rules')} />
        )}

        {step === 'rules' && curP && (
          <RulesStep p={curP} idx={curIdx} settings={settings} setSettings={setSettings} lists={lists} setLists={setLists}
            onRollFile={onRollFile} flagCount={flagCount} batches={batchesOf(curP)}
            choices={[...chosenLabs(curP).map((r) => ({ ...r, kind: 'Lab' })), ...chosenRooms(curP).map((r) => ({ ...r, kind: 'Room' }))]}
            back={() => setStep(curSum && curSum.offline > 0 ? 'rooms' : 'labs')} prepare={prepare} />
        )}

        {step === 'result' && results && (
          <ResultStep results={results} view={view} setView={setView} metaFor={metaFor} multi={results.length > 1} order={settings.buildingOrder}
            reshuffle={() => { if (confirm('Reshuffle every seat? The current seating is replaced.')) run(Math.floor(Math.random() * 1e9), test.papers); }}
            download={download} downloadBoard={downloadBoard} back={() => setStep('rooms')} />
        )}
        {step === 'result' && results && papers.some((p) => !p.done) && (
          <div className="alert info"><span>{papers.filter((p) => !p.done).map((p) => paperName(p, papers.indexOf(p))).join(', ')}: seating not made yet.</span><button className="btn sm" onClick={() => setStep('test')}>Set up next paper</button></div>
        )}
      </main>
      <footer className="foot">Matrix Academy, Sikar · Operations · Student data stays in this browser</footer>
    </div>
  );
}

// Always-visible summary under the steps: the current paper's online and offline seats, monitoring and due fee
// students still without a room, and after seating, what was seated.
function SummaryBar({ adm, papers, appearing, curP, curIdx, sum, unassigned, results, goRules }) {
  const n = (v) => Number(v || 0).toLocaleString('en-IN');
  const warn = unassigned.map((u) => (
    <span key={u.flag} className="bad"><b>{u.flag}: {n(u.n)}</b> not assigned a seat{goRules && <> · <button className="link" onClick={goRules}>Choose rooms</button></>}</span>
  ));
  if (results) {
    const all = results.flatMap((x) => x.result.assigned);
    const on = all.filter((s) => s.Mode === 'Online').length;
    const moved = results.reduce((a, x) => a + (x.result.movedOffline || 0), 0);
    const notSeated = results.reduce((a, x) => a + x.result.unseated.length, 0);
    return (
      <div className="sumbar" role="status">
        <span><b>{n(all.length + notSeated)}</b> in {pl(results.length, 'paper')}</span>
        <span className="good">Online <b>{n(on)}</b></span>
        <span>Offline <b>{n(all.length - on)}</b>{moved ? ` (${n(moved)} moved from online, labs full)` : ''}</span>
        <span className={notSeated ? 'bad' : 'good'}><b>{n(notSeated)}</b> not seated</span>
        {warn}
      </div>
    );
  }
  return (
    <div className="sumbar" role="status">
      <span><b>{n(adm.students.length)}</b> active · <b>{n(appearing)}</b> appearing · <b>{papers.length}</b> {papers.length === 1 ? 'paper' : 'papers'}</span>
      {curP && sum && <>
        <span><b>{paperName(curP, curIdx)}</b></span>
        {curP.mode !== 'Offline' && <span className={sum.onlineShort ? 'warnb' : 'good'}>Online <b>{n(sum.online - sum.overflow)}</b> / {n(sum.onlineSeats)} lab seats · {sum.onlineShort ? `${n(sum.onlineShort)} go offline (labs full)` : `${n(sum.onlineLeft)} empty`}</span>}
        {(curP.mode !== 'Online' || sum.overflow > 0) && <span className={sum.offlineShort ? 'bad' : 'good'}>Offline <b>{n(sum.offline)}</b> / {n(sum.offlineSeats)} seats · {sum.offlineShort ? `${n(sum.offlineShort)} need a seat` : `${n(sum.offlineLeft)} empty`}</span>}
      </>}
      {warn}
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
      <div><span className="kicker">First time here?</span><h2>Seating sheet in five steps</h2></div>
      <ol>
        <li><i>1</i><b>Admission list</b><span>Upload the latest admission list.</span></li>
        <li><i>2</i><b>Test and papers</b><span>Each paper: name, phases, online or offline, shifts, online batches.</span></li>
        <li><i>3</i><b>Labs</b><span>Tick labs until every online student has a seat.</span></li>
        <li><i>4</i><b>Buildings and rooms</b><span>All rooms are selected; untick any you do not want.</span></li>
        <li><i>5</i><b>Shuffle rules</b><span>Seat order, monitoring and due fee rooms, then Prepare seating.</span></li>
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

function PaperStep({ test, setTest, setPaper, phases, phaseIssues, ownerOf, studentsByBatch, studentMode, setStudentMode, top, setTop, summaryOf, paperOk, start, settings, setSettings, back }) {
  const addPaper = () => {
    const lastP = test.papers[test.papers.length - 1];
    const start0 = lastP ? lastP.s1.start : '09:00';
    setTest({ ...test, papers: [...test.papers, newPaper(start0, hhmm(mins(start0) + 180))] });
  };
  const removePaper = (id) => setTest({ ...test, papers: test.papers.filter((p) => p.id !== id) });
  return (
    <section>
      <Head kicker="Step 2" title="Test and papers" lead="Pick the date. For each paper answer the questions in order: name, phases, online or offline, shifts, and (for both) which batches sit online. The seat summary updates as you click." />
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
              <PaperCard key={p.id} p={p} i={i} n={test.papers.length} setPaper={setPaper} remove={() => removePaper(p.id)}
                phases={phases} ownerOf={ownerOf} studentsByBatch={studentsByBatch} studentMode={studentMode} setStudentMode={setStudentMode}
                top={top} setTop={setTop} sum={summaryOf(p)} ok={paperOk(p)} start={() => start(p)} />
            ))}
          </div>
          <button type="button" className="addslot" onClick={addPaper}>{Ico.plus}Add another paper</button>
          <details className="fold">
            <summary>Phase list <span className="muted">· {settings.phaseText ? 'pasted list in use' : 'phases from batch codes (XA1, XA2 → XA)'}</span></summary>
            <p className="muted">Paste the course and phase list from the Counselling Planner, one phase per line: Course | Phase | Batches. Batches not in the list are grouped by their batch code.</p>
            <textarea rows="5" placeholder={'Momentum | XA | XA1, XA2, XA3\nMomentum | XB | XB1, XB2'} value={settings.phaseText || ''}
              onChange={(e) => setSettings((s) => ({ ...s, phaseText: e.target.value }))} aria-label="Phase list" />
            {phaseIssues.map((x) => <p key={x} className="muted small bad">{x}</p>)}
          </details>
        </>
      )}
      <Nav back={back} />
    </section>
  );
}

function TimeRow({ label, value, onChange, note }) {
  const bad = !shiftOk(value);
  const set = (key, v) => {
    const o = { ...value };
    if (key === 'start') { const len = o.end && o.start ? mins(o.end) - mins(o.start) : 180; o.start = v; o.end = v ? hhmm(mins(v) + (len > 0 ? len : 180)) : o.end; } else o.end = v;
    onChange(o);
  };
  return (
    <div className={`slot ${bad ? 'bad' : ''}`}>
      <b className="shift-l">{label}</b>
      <label className="field"><span>Starts</span><input type="time" value={value.start} onChange={(e) => set('start', e.target.value)} /></label>
      <label className="field"><span>Ends</span><input type="time" value={value.end} onChange={(e) => set('end', e.target.value)} /></label>
      <span className="slot-time muted">{bad ? 'Set the end time' : note || `${range(value)} · ${Math.round((mins(value.end) - mins(value.start)) / 6) / 10} h`}</span>
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

function PaperCard({ p, i, n, setPaper, remove, phases, ownerOf, studentsByBatch, studentMode, setStudentMode, top, setTop, sum, ok, start }) {
  const [open, setOpen] = useState('');
  const bs = phases.filter((g) => p.phases.includes(g.key)).flatMap((g) => g.batches);
  const togglePhase = (k) => setPaper(p.id, { phases: p.phases.includes(k) ? p.phases.filter((x) => x !== k) : [...p.phases, k] });
  const toggleOnline = (b) => setPaper(p.id, { onlineBatches: p.onlineBatches.includes(b) ? p.onlineBatches.filter((x) => x !== b) : [...p.onlineBatches, b] });
  const toggleTop = (b) => setTop((t) => { const x = new Set(t); x.has(b) ? x.delete(b) : x.add(b); return x; });
  const named = !!p.name.trim();
  return (
    <div className="card paper">
      <div className="paper-top">
        <span className="slot-n">{i + 1}</span>
        <label className="field grow"><span>1. Paper name</span>
          <input type="text" placeholder="e.g. PRT-3 Paper 1" value={p.name} onChange={(e) => setPaper(p.id, { name: e.target.value })} /></label>
        <button type="button" className="iconbtn danger" disabled={n === 1} onClick={remove} aria-label={`Remove paper ${i + 1}`} title="Remove this paper">{Ico.bin}</button>
      </div>

      {named && (
        <div className="q">
          <b className="q-l">2. Phases in this paper</b>
          <div className="chips">
            {phases.map((g) => {
              const other = ownerOf[g.key] && ownerOf[g.key] !== p.id;
              return (
                <button key={g.key} type="button" className={`chip-btn ${p.phases.includes(g.key) ? 'on' : ''}`} disabled={other}
                  title={other ? 'Already in another paper' : g.batches.map((b) => b.batch).join(', ')} onClick={() => togglePhase(g.key)}>
                  {g.course} · {g.phase} <small>{g.count}</small>
                </button>
              );
            })}
          </div>
          {bs.length > 0 && <p className="muted small">Batches: {bs.map((b) => b.batch).join(', ')}</p>}
        </div>
      )}

      {named && p.phases.length > 0 && (
        <div className="q">
          <b className="q-l">3. Online or offline?</b>
          <div className="seg" role="group" aria-label="Online or offline">
            {[['Online', 'Online'], ['Offline', 'Offline'], ['Both', 'Both online and offline']].map(([m, l]) => (
              <button key={m} className={p.mode === m ? 'on' : ''} onClick={() => setPaper(p.id, { mode: m })}>{l}</button>
            ))}
          </div>
        </div>
      )}

      {named && p.phases.length > 0 && p.mode && (
        <div className="q">
          <b className="q-l">4. One shift or two?</b>
          <div className="seg" role="group" aria-label="Shifts">
            <button className={p.shifts === 1 ? 'on' : ''} onClick={() => setPaper(p.id, { shifts: 1, two: false })}>1 shift</button>
            <button className={p.shifts === 2 ? 'on' : ''} onClick={() => setPaper(p.id, { shifts: 2, two: true })}>2 shifts</button>
          </div>
          {p.shifts > 0 && (
            <div className="times">
              <TimeRow label={p.shifts === 2 ? 'Shift 1' : p.mode === 'Both' ? 'Offline time' : 'Time'} value={p.s1} onChange={(v) => setPaper(p.id, { s1: v })} />
              {p.shifts === 2 && <TimeRow label="Shift 2" value={p.s2} onChange={(v) => setPaper(p.id, { s2: v })} />}
              {p.mode === 'Both' && <TimeRow label="Online time" value={onlineOf(p)} onChange={(v) => setPaper(p.id, { online: v })} note={`${range(onlineOf(p))} · computer labs`} />}
              {p.shifts === 2 && <p className="muted small">Both shifts use the same seating: a student keeps one seat. The Excel file has an attendance sheet for each shift.</p>}
            </div>
          )}
        </div>
      )}

      {named && p.mode === 'Both' && p.shifts > 0 && (
        <div className="q">
          <b className="q-l">5. Which batches sit online?</b>
          <p className="muted small">Ticked batches go to the computer labs (top batches first). The other batches sit offline. If the labs fill up, the remaining online students sit offline. <button className="link" onClick={() => setPaper(p.id, { onlineBatches: bs.map((b) => b.batch) })}>All online</button> · <button className="link" onClick={() => setPaper(p.id, { onlineBatches: [] })}>All offline</button></p>
          <div className="batch-grid">
            {bs.map((b) => {
              const on = p.onlineBatches.includes(b.batch);
              return (
                <div key={b.batch} className={`batch ${on ? 'on' : ''}`}>
                  <label className="batch-top">
                    <input type="checkbox" checked={on} onChange={() => toggleOnline(b.batch)} />
                    <span className="batch-name">{b.batch}</span>
                    <span className="batch-count">{b.count}</span>
                  </label>
                  <div className="batch-opts">
                    <span className={`small ${on ? 'good' : 'muted'}`}>{on ? 'Online' : 'Offline'}</span>
                    <button className="link small" onClick={() => setOpen(open === b.batch ? '' : b.batch)}>Students</button>
                    <label className="small top-tick" title="Top batch students take lab seats first"><input type="checkbox" checked={top.has(b.batch)} onChange={() => toggleTop(b.batch)} /> Top</label>
                  </div>
                </div>
              );
            })}
          </div>
          {open && bs.some((b) => b.batch === open) && (
            <StudentModes batch={open} students={studentsByBatch[open] || []} batchMode={p.onlineBatches.includes(open) ? 'Online' : 'Offline'}
              studentMode={studentMode} setStudentMode={setStudentMode} close={() => setOpen('')} />
          )}
        </div>
      )}

      {named && p.phases.length > 0 && p.mode && <SeatBoard s={sum} mode={p.mode} />}

      {ok && (
        <div className="paper-cta">
          {p.done && <span className="pill">Seating ready ✓</span>}
          <button className="btn cta" onClick={start}>{Ico.spark}{p.done ? 'Change seating for this test' : 'Create seating for this test'} →</button>
        </div>
      )}
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

function LabStep({ p, idx, labsOn, order, setPaper, sum, back, next }) {
  const sel = new Set(p.labs || []);
  const chosen = labsOn.filter((l) => sel.has(roomKey(l)));
  const prev = fillPreview(sum.online, chosen, order);
  const fillOf = Object.fromEntries(prev.rooms.map((r) => [roomKey(r), r]));
  const empty = prev.rooms.reduce((a, r) => a + r.empty, 0);
  const toggle = (k) => setPaper(p.id, { labs: sel.has(k) ? [...sel].filter((x) => x !== k) : [...sel, k] });
  const groups = {};
  for (const l of labsOn) (groups[l.building] ||= []).push(l);
  const blds = Object.keys(groups).sort((a, b) => order.indexOf(a) - order.indexOf(b));
  return (
    <section>
      <Head kicker={`Paper ${idx + 1}: ${p.name}`} title="Which labs for this test?" lead={`Tick the computer labs for the ${sum.online} online students. Labs fill in the order shown, one lab completely before the next. Top batch students are seated first.`} />
      <BigLine left={prev.left} placed={sum.online} empty={empty} what="online" more="Select more labs, or continue and they sit offline." />
      <p className="muted small"><button className="link" onClick={() => setPaper(p.id, { labs: labsOn.map(roomKey) })}>Select all labs</button> · <button className="link" onClick={() => setPaper(p.id, { labs: [] })}>Clear</button></p>
      {blds.map((b) => {
        const list = groups[b];
        const on = list.filter((l) => sel.has(roomKey(l)));
        const seats = on.reduce((a, l) => a + Number(l.capacity), 0);
        const filled = on.reduce((a, l) => a + (fillOf[roomKey(l)]?.filled || 0), 0);
        return (
          <div className="card lab-b" key={b}>
            <div className="lab-bh"><b>{b}</b><span className="muted small">{on.length ? `${on.length} selected · ${seats} seats · ${filled} filled · ${seats - filled} empty` : `${list.length} ${list.length === 1 ? 'lab' : 'labs'}`}</span></div>
            <div className="lab-list">
              {list.map((l) => {
                const k = roomKey(l), f = fillOf[k];
                return (
                  <label key={k} className={`lab-row ${sel.has(k) ? 'on' : ''}`}>
                    <input type="checkbox" checked={sel.has(k)} onChange={() => toggle(k)} />
                    <span className="lab-name">{l.room}</span>
                    <span className="muted small">{l.capacity} seats{l.girls === false ? ' · no girls' : ''}</span>
                    {f && <span className={`small ${f.empty ? 'muted' : 'good'}`}>{f.filled} filled · {f.empty} empty</span>}
                  </label>
                );
              })}
            </div>
          </div>
        );
      })}
      <div className="ctabar">
        <div><b>{prev.left ? `${prev.left} online students without a lab seat` : 'Labs ready'}</b><span className="muted">{chosen.length} labs · {chosen.reduce((a, l) => a + Number(l.capacity), 0)} seats</span></div>
        <div className="cta-row">
          <button className="btn ghost" onClick={back}>Back</button>
          <button className="btn cta" onClick={next} disabled={!chosen.length}>Continue →</button>
        </div>
      </div>
    </section>
  );
}

function RoomStep({ p, idx, settings, setSettings, setPaper, sum, rooms, orderOn, onRoomFile, back, next }) {
  const editRoom = (kind, i, patch) => setSettings((s) => ({ ...s, [kind]: s[kind].map((r, j) => (j === i ? { ...r, ...patch } : r)) }));
  const move = (i, d) => setSettings((s) => {
    const o = s.buildingOrder.slice(); const j = i + d;
    if (j < 0 || j >= o.length) return s;
    [o[i], o[j]] = [o[j], o[i]]; return { ...s, buildingOrder: o };
  });
  const toggleUse = (b) => setSettings((s) => ({ ...s, buildingsOn: s.buildingsOn.includes(b) ? s.buildingsOn.filter((x) => x !== b) : [...s.buildingsOn, b] }));
  const toggleGirls = (b) => setSettings((s) => ({ ...s, girlsOk: { ...s.girlsOk, [b]: s.girlsOk[b] === false } }));
  const off = new Set(p.roomsOff);
  const toggleRoom = (k) => setPaper(p.id, { roomsOff: off.has(k) ? p.roomsOff.filter((x) => x !== k) : [...p.roomsOff, k] });
  const prev = fillPreview(sum.offline, rooms, orderOn);
  const fillOf = Object.fromEntries(prev.rooms.map((r) => [roomKey(r), r]));
  const empty = prev.rooms.reduce((a, r) => a + r.empty, 0);
  const roomsIn = (b) => settings.rooms.filter((r) => r.building === b && r.enabled && Number(r.capacity) > 0);
  const buildings = settings.buildingOrder.filter((b) => roomsIn(b).length);
  return (
    <section>
      <Head kicker={`Paper ${idx + 1}: ${p.name}`} title="Buildings and rooms" lead={`Seats for the ${sum.offline} offline students${sum.overflow ? ` (including ${sum.overflow} from online, labs full)` : ''}. Every room is selected; untick a room to leave it out. Buildings fill in the order shown, one room completely before the next.`} />
      <BigLine left={prev.left} placed={sum.offline} empty={empty} what="offline" more="Tick another building or room." />
      {buildings.map((b) => {
        const i = settings.buildingOrder.indexOf(b);
        const on = settings.buildingsOn.includes(b);
        const list = roomsIn(b);
        const usedList = list.filter((r) => on && !off.has(roomKey(r)));
        const seats = usedList.reduce((a, r) => a + Number(r.capacity), 0);
        const filled = usedList.reduce((a, r) => a + (fillOf[roomKey(r)]?.filled || 0), 0);
        return (
          <div className={`card lab-b ${on ? '' : 'off'}`} key={b}>
            <div className="lab-bh">
              <label className="bld-use"><input type="checkbox" checked={on} onChange={() => toggleUse(b)} /> <b>{on ? `${orderOn.indexOf(b) + 1}. ` : ''}{b}</b></label>
              <span className="muted small">{on ? `${usedList.length} of ${list.length} rooms · ${seats} seats · ${filled} filled · ${seats - filled} empty` : 'not used'}</span>
              <label className="far small"><input type="checkbox" checked={settings.girlsOk[b] !== false} onChange={() => toggleGirls(b)} /> Girls allowed</label>
              <button className="iconbtn" onClick={() => move(i, -1)} aria-label={`Move ${b} up`} disabled={i === 0}>▲</button>
              <button className="iconbtn" onClick={() => move(i, 1)} aria-label={`Move ${b} down`} disabled={i === settings.buildingOrder.length - 1}>▼</button>
            </div>
            {on && (
              <div className="room-chips">
                {list.map((r) => {
                  const k = roomKey(r), f = fillOf[k], use = !off.has(k);
                  return (
                    <label key={k} className={`room-chip ${use ? 'on' : ''} ${f && f.filled ? 'filled' : ''}`} title={f ? `${f.filled} filled, ${f.empty} empty` : 'Not used'}>
                      <input type="checkbox" checked={use} onChange={() => toggleRoom(k)} />
                      <b>{r.room}</b><span>{use && f ? `${f.filled}/${r.capacity}` : r.capacity}</span>
                    </label>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}
      <details className="fold">
        <summary>Room capacities and the room list <span className="muted">· {settings.rooms.length} rooms, {settings.labs.length} labs</span></summary>
        <RoomTable title="Classrooms (offline)" kind="rooms" list={settings.rooms} edit={editRoom} use={orderOn} />
        <RoomTable title="Computer labs (online)" kind="labs" list={settings.labs} edit={editRoom} use={settings.buildingOrder} labs />
        <p className="muted">Upload a new room list only when rooms change. Uploads stay saved in this browser.</p>
        <div className="files"><FileCard title="Room list" hint="Sheets Room_Capacity and LAB_Capacity" onFile={onRoomFile} template={() => downloadTemplate('rooms')} /></div>
        <button className="btn ghost danger" onClick={() => { if (confirm('Reset buildings, rooms, capacities, girls settings and groups to the built-in lists?')) setSettings(freshSettings()); }}>{Ico.bin}Reset to built-in lists</button>
      </details>
      <div className="ctabar">
        <div><b>{prev.left ? `${prev.left} offline students without a seat` : 'Rooms ready'}</b><span className="muted">{rooms.length} rooms · {rooms.reduce((a, r) => a + Number(r.capacity), 0)} seats · {orderOn.join(' → ')}</span></div>
        <div className="cta-row">
          <button className="btn ghost" onClick={back}>Back</button>
          <button className="btn cta" onClick={next}>Continue →</button>
        </div>
      </div>
    </section>
  );
}

function RulesStep({ p, idx, settings, setSettings, lists, setLists, onRollFile, flagCount, batches, choices, back, prepare }) {
  const setPins = (f, keys) => setSettings((s) => ({ ...s, pins: { ...s.pins, [f]: keys } }));
  const groups = settings.cbGroups;
  const setGroups = (g) => setSettings((s) => ({ ...s, cbGroups: g }));
  const order = settings.seatOrder || 'mixed';
  return (
    <section>
      <Head kicker={`Paper ${idx + 1}: ${p.name}`} title="Shuffle rules" lead="Choose how students are shuffled to stop copying, and where monitoring and due fee students sit. Then prepare the seating." />
      <div className="card seatorder">
        <h2>Seat order inside each room</h2>
        <div className="seg" role="group" aria-label="Seat order inside each room">
          <button className={order === 'mixed' ? 'on' : ''} onClick={() => setSettings((s) => ({ ...s, seatOrder: 'mixed' }))}>Mixed batches</button>
          <button className={order === 'roll' ? 'on' : ''} onClick={() => setSettings((s) => ({ ...s, seatOrder: 'roll' }))}>Roll number order</button>
        </div>
        <p className="muted small">{order === 'mixed' ? 'Students are shuffled across rooms, and two students of the same batch never sit side by side. Reshuffle on the result page gives a new mix.' : 'Students are shuffled across rooms; inside a room, seats follow roll numbers, so one batch sits together.'}</p>
      </div>

      {['Monitoring', 'Due fee'].map((f) => {
        const n = flagCount(f), pinned = (settings.pins[f] || []).length;
        return (
          <div className={`card special ${n && !pinned ? 'need' : ''}`} key={f}>
            <div className="special-head"><b>{f} students</b><span className={`small ${n && !pinned ? 'bad' : 'muted'}`}>{n} appearing{n && !pinned ? ' · not assigned a seat: choose rooms below' : ''}</span></div>
            <div className="special-list">
              <RollBox rolls={lists[f]} onChange={(rolls) => setLists((l) => ({ ...l, [f]: rolls }))} label={`${f} roll numbers`} />
              <RollUpload onFile={(file) => onRollFile(f, file)} />
            </div>
            <PinPicker choices={choices} value={settings.pins[f] || []} onChange={(keys) => setPins(f, keys)} label={f} />
          </div>
        );
      })}

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
        <button type="button" className="addslot" onClick={() => setGroups([...groups, { name: '', batches: batches.map((b) => b.batch).join(', ') }])}>{Ico.plus}Add a group</button>
      </details>

      <div className="ctabar">
        <div><b>Ready to prepare the seating</b><span className="muted">{p.name} · {batches.length} batches</span></div>
        <div className="cta-row">
          <button className="btn ghost" onClick={back}>Back</button>
          <button className="btn cta" onClick={prepare}>{Ico.spark}Prepare seating</button>
        </div>
      </div>
    </section>
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
      <Head kicker={multi ? `Seating ready · ${pl(results.length, 'paper')}` : 'Seating ready'} title={meta.title}
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
        <div className="alert info"><span>This paper overlaps {x.overlapsWith.map((j) => paperName(results[j].paper, j)).join(', ')} (or starts less than {GAP} minutes after it). It uses the seats left free in shared rooms, so no seat is used twice at once.</span></div>
      )}
      {!x.batches.length && <div className="alert warn"><span>No batches are in this paper. Add them on the Batches step.</span></div>}
      <div className="actions">
        {multi
          ? <><button className="btn cta" onClick={() => download(results)} disabled={!files}>{Ico.down}Download all ({files} {files === 1 ? 'file' : 'files'})</button>
              <button className="btn primary" onClick={() => download([x])} disabled={!x.batches.length}>{Ico.down}Only {paperName(x.paper, x.index)}</button></>
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
        <div><b>{shiftsOf(x.paper).length}</b><span>attendance {x.paper.two ? 'sheets (same seats)' : 'sheet'}</span></div>
      </div>
      {result.warnings.map((w) => <div key={w} className="alert warn"><span>{w}</span></div>)}

      <details className="fold" open>
        <summary>Rooms <span className="muted">· {result.usedRooms.length} used</span></summary>
        <p className="muted small">Green: students of this {multi ? 'paper' : 'test'}. Amber: monitoring, due fee or top batch students.{x.overlapsWith.length ? ' Grey: seats used by an overlapping paper.' : ''}</p>
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
      <p className="muted small">{multi ? 'Each paper file has' : 'The Excel file has'}: Raw_Data, Student_list, Seating Arrangement, {x.paper.two ? 'an attendance sheet per shift' : 'Attendance'}, SMD, Room_Layout, Master_List{result.unseated.length ? ', Not_Seated' : ''}. The notice board file puts every paper on one sheet.</p>
    </section>
  );
}
