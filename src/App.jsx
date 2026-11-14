import React, { useEffect, useMemo, useRef, useState } from 'react';
import defaults from './data/defaults.json';
import { allocate, capacity, naturalCompare, isGirl, roomKey, FLAGS } from './lib/engine.js';
import { buildWorkbook, buildNoticeBoard, fmtDateDMY, fmtTime } from './lib/exporter.js';
import { readAdmissions, readRooms, readRollList, parseRolls } from './lib/readers.js';

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
const newPaper = (start = '09:00', end = '12:00') => ({ id: uid(), name: '', two: false, s1: { start, end }, online: { start, end }, s2: { start: hhmm(mins(end) + 30), end: hhmm(mins(end) + 30 + (mins(end) - mins(start))) } });
function initialTest() {
  const last = loadLast();
  if (last && Array.isArray(last.papers) && last.papers.length) return { title: last.title || '', date: '', papers: last.papers };
  return { title: '', date: '', papers: [newPaper()] };
}

const pl = (n, w) => `${n.toLocaleString('en-IN')} ${w}${n === 1 ? '' : /(ch|sh|s|x)$/.test(w) ? 'es' : 's'}`;
const mins = (t) => { const [h, m] = String(t || '0:0').split(':').map(Number); return h * 60 + m; };
const hhmm = (m) => `${String(Math.floor(Math.min(m, 1439) / 60)).padStart(2, '0')}:${String(Math.min(m, 1439) % 60).padStart(2, '0')}`;
// A paper's shifts. Two shifts: the same students and the same seats, an attendance sheet for each shift.
const shiftsOf = (p) => (p.two ? [{ label: 'Shift 1', ...p.s1 }, { label: 'Shift 2', ...p.s2 }] : [{ label: '', ...p.s1 }]);
// Online time of a paper: online students may sit at a different time from offline students (default: same as Shift 1)
const onlineOf = (p) => (p.online && p.online.start ? p.online : p.s1);
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
const paperTime = (p) => `Offline ${shiftsOf(p).map((s) => (s.label ? `${s.label} ${range(s)}` : range(s))).join(' · ')} · Online ${range(onlineOf(p))}`;
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
  { id: 'test', label: 'Date & papers' },
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

  const papers = test.papers;
  const multi = papers.length > 1;
  const paperIds = papers.map((p) => p.id);
  const paperFor = (batch) => (paperIds.includes(paperOf[batch]) ? paperOf[batch] : paperIds[0]);
  const chosen = batches.filter((b) => selected.has(b.batch));
  const modeOf = (s) => studentMode[s.Roll_Number] || mode[s.Batch] || 'Offline';
  const studentsOf = (bs) => bs.flatMap((b) => studentsByBatch[b.batch] || []);

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
  // Seats needed at the busiest moment: a paper plus every paper that overlaps it
  const need = perPaper.reduce((best, x) => {
    const over = perPaper.filter((y) => overlaps(x.paper, y.paper));
    const sum = (k) => over.reduce((a, y) => a + y[k], 0);
    const n = { offline: sum('offline'), online: sum('online'), girlsOffline: sum('girlsOffline'), girlsOnline: sum('girlsOnline') };
    return n.offline + n.online > best.offline + best.online ? n : best;
  }, { offline: 0, online: 0, girlsOffline: 0, girlsOnline: 0 });

  const useOrder = settings.buildingOrder.filter((b) => settings.buildingsOn.includes(b));
  const base = { rooms: settings.rooms, labs: settings.labs, buildingOrder: useOrder, girlsOk: settings.girlsOk };
  const capOff = capacity(base, 'Offline'), capOn = capacity(base, 'Online');
  const have = { offline: capOff.seats, online: capOn.seats, girlsOffline: capOff.girlSeats, girlsOnline: capOn.girlSeats };
  // Online first: online students beyond the lab seats are seated offline
  const overflow = Math.max(0, need.online - have.online);
  const girlsOverflow = Math.max(0, need.girlsOnline - have.girlsOnline);
  need.online -= overflow; need.offline += overflow; need.girlsOnline -= girlsOverflow; need.girlsOffline += girlsOverflow;
  need.overflow = overflow;
  const seatsShort = Math.max(0, need.offline - have.offline);
  const girlsShort = Math.max(0, need.girlsOffline - have.girlsOffline);
  const appearing = chosen.reduce((a, b) => a + b.count, 0);
  const papersOk = papers.length > 0 && papers.every((p) => shiftOk(p.s1) && shiftOk(onlineOf(p)) && (!p.two || (shiftOk(p.s2) && mins(p.s2.start) >= mins(p.s1.end))));
  const emptyPapers = multi ? perPaper.filter((x) => !x.list.length) : [];

  const ready = {
    students: !!adm,
    test: !!test.date && papersOk,
    batches: chosen.length > 0,
    rooms: useOrder.length > 0,
    result: !!results,
  };
  const canCreate = ready.students && ready.test && ready.batches && ready.rooms;

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

  // One seating per paper, in time order. A paper that overlaps an earlier one (or starts within GAP minutes of it)
  // starts each room's seat numbers after the seats that paper already uses, so no seat is used twice at once.
  // A paper with two shifts has one seating: the same seats in both shifts, an attendance sheet for each.
  function run(newSeed = seed) {
    const cbOf = {};
    for (const g of settings.cbGroups) for (const b of parseRolls(g.batches)) if (g.name.trim()) cbOf[b] = g.name.trim();
    const cfg = { ...base, batchMode: Object.fromEntries(chosen.map((b) => [b.batch, mode[b.batch] || 'Offline'])), studentMode, flags, pins: settings.pins, cbOf, seatOrder: settings.seatOrder || 'mixed' };
    const order = papers.map((p, i) => ({ p, i })).sort((a, b) => mins(windowOf(a.p).start) - mins(windowOf(b.p).start) || a.i - b.i);
    const done = [];
    for (const { p, i } of order) {
      const bs = chosen.filter((b) => !multi || paperFor(b.batch) === p.id);
      const taken = {};
      const over = done.filter((d) => overlaps(d.paper, p));
      for (const d of over) for (const room of d.result.usedRooms) {
        const k = roomKey(room);
        taken[k] = Math.max(taken[k] || 0, room.taken + room.used);
      }
      const st = studentsOf(bs);
      done.push({ paper: p, index: i, result: allocate(st, { ...cfg, seed: newSeed + i, taken }), batches: bs, students: st.length, overlapsWith: over.map((d) => d.index) });
    }
    setResults(done.sort((a, b) => a.index - b.index)); setSeed(newSeed); setView(0); setStep('result');
    try { localStorage.setItem(LAST, JSON.stringify({ title: test.title, papers, batches: [...selected], mode, studentMode, paperOf, top: [...top] })); } catch { /* ignore */ }
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

  const canOpen = (i) => (STEPS[i].id === 'result' ? !!results : STEPS.slice(0, i).every((s) => ready[s.id] || s.id === 'rooms'));

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
        {adm && step !== 'students' && <SummaryBar adm={adm} chosen={chosen} appearing={appearing} papers={papers} useOrder={useOrder}
          need={need} have={have} seatsShort={seatsShort} girlsShort={girlsShort} results={step === 'result' ? results : null} />}
      </header>

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
          <PaperStep test={test} setTest={setTest} papersOk={papersOk}
            back={() => setStep('students')} next={() => setStep('batches')} nextOk={ready.test} />
        )}

        {step === 'batches' && (
          <BatchStep
            batches={batches} studentsByBatch={studentsByBatch} selected={selected} setSelected={setSelected}
            mode={mode} setMode={setMode} studentMode={studentMode} setStudentMode={setStudentMode} top={top} setTop={setTop}
            papers={papers} multi={multi} paperFor={paperFor} setPaperOf={setPaperOf} perPaper={perPaper}
            need={need} emptyPapers={emptyPapers} appearing={appearing}
            back={() => setStep('test')} next={() => setStep('rooms')} nextOk={ready.batches}
          />
        )}

        {step === 'rooms' && (
          <RoomStep settings={settings} setSettings={setSettings} need={need} have={have} chosen={chosen} appearing={appearing}
            lists={lists} setLists={setLists} onRollFile={onRollFile} flags={flags} chosenRolls={new Set(studentsOf(chosen).map((s) => s.Roll_Number))}
            onRoomFile={onRoomFile} canCreate={canCreate} useOrder={useOrder}
            back={() => setStep('batches')} create={() => run()} hasResult={!!results} />
        )}

        {step === 'result' && results && (
          <ResultStep results={results} view={view} setView={setView} metaFor={metaFor} multi={multi} order={useOrder}
            reshuffle={() => { if (confirm('Reshuffle every seat? The current seating is replaced.')) run(Math.floor(Math.random() * 1e9)); }}
            download={download} downloadBoard={downloadBoard} back={() => setStep('rooms')} />
        )}
      </main>
      <footer className="foot">Matrix Academy, Sikar · Operations · Student data stays in this browser</footer>
    </div>
  );
}

// Always-visible summary under the steps: what is planned before Create, what was seated after.
function SummaryBar({ adm, chosen, appearing, papers, useOrder, need, have, seatsShort, girlsShort, results }) {
  const n = (v) => Number(v || 0).toLocaleString('en-IN');
  const tail = <span><b>{papers.length}</b> {papers.length === 1 ? 'paper' : 'papers'} · <b>{useOrder.length}</b> {useOrder.length === 1 ? 'building' : 'buildings'}</span>;
  if (results) {
    const all = results.flatMap((x) => x.result.assigned);
    const on = all.filter((s) => s.Mode === 'Online').length;
    const moved = results.reduce((a, x) => a + (x.result.movedOffline || 0), 0);
    const notSeated = results.reduce((a, x) => a + x.result.unseated.length, 0);
    return (
      <div className="sumbar" role="status">
        <span><b>{n(appearing)}</b> appearing</span>
        <span className="good">Online <b>{n(on)}</b></span>
        <span>Offline <b>{n(all.length - on)}</b>{moved ? ` (${n(moved)} moved from online, labs full)` : ''}</span>
        <span className={notSeated ? 'bad' : 'good'}><b>{n(notSeated)}</b> not seated</span>
        {tail}
      </div>
    );
  }
  return (
    <div className="sumbar" role="status">
      <span><b>{n(adm.students.length)}</b> active students</span>
      <span><b>{chosen.length}</b> {chosen.length === 1 ? 'batch' : 'batches'} · <b>{n(appearing)}</b> appearing</span>
      <span className={need.online ? 'good' : ''}>Online <b>{n(need.online)}</b> / {n(have.online)} lab seats</span>
      <span className={seatsShort ? 'bad' : ''}>Offline <b>{n(need.offline)}</b> / {n(have.offline)} classroom seats{need.overflow ? ` (${n(need.overflow)} from online, labs full)` : ''}</span>
      <span className={seatsShort || girlsShort ? 'bad' : chosen.length ? 'good' : ''}>{seatsShort ? <><b>{n(seatsShort)}</b> seats short</> : girlsShort ? <><b>{n(girlsShort)}</b> girls' seats short</> : chosen.length ? <b>Enough seats</b> : 'No batches yet'}</span>
      {tail}
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
        <li><i>2</i><b>Date and papers</b><span>Add each paper with its offline and online time.</span></li>
        <li><i>3</i><b>Batches</b><span>Tick each paper's batches. Online batches fill the labs first; the rest sit offline.</span></li>
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

function PaperStep({ test, setTest, papersOk, back, next, nextOk }) {
  const setPaper = (id, patch) => setTest({ ...test, papers: test.papers.map((p) => (p.id === id ? { ...p, ...patch } : p)) });
  const setTime = (p, which, key, v) => {
    const s = { ...p[which] };
    if (key === 'start') { const len = s.end && s.start ? mins(s.end) - mins(s.start) : 180; s.start = v; s.end = v ? hhmm(mins(v) + (len > 0 ? len : 180)) : s.end; } else s.end = v;
    setPaper(p.id, { [which]: s });
  };
  const addPaper = () => {
    const lastP = test.papers[test.papers.length - 1];
    const end = lastP ? windowOf(lastP).end : '07:00';
    const start = hhmm(mins(end) + 30);
    setTest({ ...test, papers: [...test.papers, newPaper(start, hhmm(mins(start) + 180))] });
  };
  const removePaper = (id) => setTest({ ...test, papers: test.papers.filter((p) => p.id !== id) });
  const clash = test.papers.length > 1 ? test.papers.filter((a) => test.papers.some((b) => b.id !== a.id && overlaps(a, b))).length : 0;
  const hint = !test.date ? 'Pick the test date to continue.' : !papersOk ? 'Each shift needs a start and an end time; Shift 2 starts after Shift 1 ends.' : '';
  return (
    <section>
      <Head kicker="Step 2 of 4" title="When is the test?" lead="Pick the date and add each paper. Set the offline time (classrooms) and the online time (computer labs). Offline can be one shift, or two shifts with the same seating." />
      <div className="card form">
        <label className="field">
          <span>Test date</span>
          <input type="date" value={test.date} onChange={(e) => setTest({ ...test, date: e.target.value })} autoFocus />
        </label>
        <label className="field grow">
          <span>Sheet title <em className="muted">(optional)</em></span>
          <input type="text" placeholder="e.g. Momentum Seating Arrangement" value={test.title} onChange={(e) => setTest({ ...test, title: e.target.value })} />
        </label>
      </div>
      <div className="slots-head"><h2>Papers <span className="count">{test.papers.length}</span></h2></div>
      <div className="papers">
        {test.papers.map((p, i) => (
          <div className="card paper" key={p.id}>
            <div className="paper-top">
              <span className="slot-n">{i + 1}</span>
              <label className="field grow"><span>Paper name <em className="muted">(optional)</em></span>
                <input type="text" placeholder={`Paper ${i + 1}, e.g. Physics`} value={p.name} onChange={(e) => setPaper(p.id, { name: e.target.value })} /></label>
              <div className="seg" role="group" aria-label={`Shifts for paper ${i + 1}`}>
                <button className={!p.two ? 'on' : ''} onClick={() => setPaper(p.id, { two: false })}>1 shift</button>
                <button className={p.two ? 'on' : ''} onClick={() => setPaper(p.id, { two: true })}>2 shifts</button>
              </div>
              <button type="button" className="iconbtn danger" disabled={test.papers.length === 1} onClick={() => removePaper(p.id)} aria-label={`Remove paper ${i + 1}`} title="Remove this paper">{Ico.bin}</button>
            </div>
            {(p.two ? [['s1', 'Offline shift 1'], ['s2', 'Offline shift 2']] : [['s1', 'Offline time']]).map(([k, label]) => {
              const s = p[k];
              const bad = !shiftOk(s) || (k === 's2' && mins(s.start) < mins(p.s1.end));
              return (
                <div className={`slot ${bad ? 'bad' : ''}`} key={k}>
                  <b className="shift-l">{label}</b>
                  <label className="field"><span>Starts</span><input type="time" value={s.start} onChange={(e) => setTime(p, k, 'start', e.target.value)} /></label>
                  <label className="field"><span>Ends</span><input type="time" value={s.end} onChange={(e) => setTime(p, k, 'end', e.target.value)} /></label>
                  <span className="slot-time muted">{!bad ? `${range(s)} · ${Math.round((mins(s.end) - mins(s.start)) / 6) / 10} h` : k === 's2' && shiftOk(s) ? 'Starts before Shift 1 ends' : 'Set the end time'}</span>
                </div>
              );
            })}
            {(() => {
              const on = onlineOf(p);
              const bad = !shiftOk(on);
              const setOn = (key, v) => {
                const o = { ...on };
                if (key === 'start') { const len = o.end && o.start ? mins(o.end) - mins(o.start) : 180; o.start = v; o.end = v ? hhmm(mins(v) + (len > 0 ? len : 180)) : o.end; } else o.end = v;
                setPaper(p.id, { online: o });
              };
              return (
                <div className={`slot online ${bad ? 'bad' : ''}`}>
                  <b className="shift-l">Online time</b>
                  <label className="field"><span>Starts</span><input type="time" value={on.start} onChange={(e) => setOn('start', e.target.value)} /></label>
                  <label className="field"><span>Ends</span><input type="time" value={on.end} onChange={(e) => setOn('end', e.target.value)} /></label>
                  <span className="slot-time muted">{!bad ? `${range(on)} · computer labs` : 'Set the end time'}</span>
                </div>
              );
            })()}
            {p.two && <p className="muted small">Both shifts use the same seating: every student keeps the Shift 1 seat. The Excel file has an attendance sheet for each shift.</p>}
          </div>
        ))}
      </div>
      <button type="button" className="addslot" onClick={addPaper}>{Ico.plus}Add another paper</button>
      {clash > 0 && <p className="muted small">Some papers overlap or are less than {GAP} minutes apart. They get different seats in shared rooms, so no seat is used twice at once.</p>}
      <Nav back={back} next={next} nextOk={nextOk} hint={hint} />
    </section>
  );
}

function BatchStep({ batches, studentsByBatch, selected, setSelected, mode, setMode, studentMode, setStudentMode, top, setTop, papers, multi, paperFor, setPaperOf, perPaper, need, emptyPapers, appearing, back, next, nextOk }) {
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
      <Head kicker="Step 3 of 4" title={multi ? 'Which batches sit each paper?' : 'Which batches are appearing?'}
        lead={`${multi ? 'Pick a paper, then tick its batches. A batch can be in one paper only. ' : ''}Online batches fill the computer labs first, top batches before others. When the labs are full, the remaining students sit offline in classrooms. Set it for a whole course, a batch, or single students.`} />
      {multi && (
        <div className="slottabs" role="tablist">
          {perPaper.map((x, i) => (
            <button key={x.paper.id} role="tab" aria-selected={x.paper.id === curId} className={x.paper.id === curId ? 'on' : ''} onClick={() => setCur(x.paper.id)}>
              <b>{paperName(x.paper, i)}</b><span>{paperTime(x.paper)} · {x.list.length ? `${pl(x.list.length, 'batch')}, ${x.offline + x.online} students` : 'no batches yet'}</span>
            </button>
          ))}
        </div>
      )}
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
              {n > 0 && <span className="course-mode muted small">Ticked batches: <button className="link" onClick={() => setCourseMode(list, 'Offline')}>all Offline</button> · <button className="link" onClick={() => setCourseMode(list, 'Online')}>all Online</button></span>}
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
                        <div className="seg">
                          {['Offline', 'Online'].map((m) => (
                            <button key={m} className={(mode[b.batch] || 'Offline') === m ? 'on' : ''} onClick={() => setBatchMode(b.batch, m)}>{m}</button>
                          ))}
                        </div>
                        <button className="link small" onClick={() => setOpenBatch(openBatch === b.batch ? '' : b.batch)}>{changed ? `${changed} moved` : 'Students'}</button>
                        <label className="small top-tick" title="Top batch students can be kept in chosen rooms (Buildings & rooms step)">
                          <input type="checkbox" checked={top.has(b.batch)} onChange={() => setTop((t) => { const x = new Set(t); x.has(b.batch) ? x.delete(b.batch) : x.add(b.batch); return x; })} /> Top
                        </label>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
            {list.some((b) => b.batch === openBatch && isOn(b.batch)) && (
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

function RoomStep({ settings, setSettings, need, have, chosen, appearing, lists, setLists, onRollFile, flags, chosenRolls, onRoomFile, canCreate, useOrder, back, create, hasResult }) {
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
        <Meter label="Classroom seats" need={need.offline} have={have.offline} />
        <Meter label="Computer lab seats" need={need.online} have={have.online} />
        <Meter label="Seats for girls (classrooms)" need={need.girlsOffline} have={have.girlsOffline} />
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
        <RoomTable title="Computer labs (online)" kind="labs" list={settings.labs} edit={editRoom} use={useOrder} labs />
      </details>

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

      <div className="card seatorder">
        <h2>Seat order inside each room</h2>
        <div className="seg" role="group" aria-label="Seat order inside each room">
          <button className={(settings.seatOrder || 'mixed') === 'mixed' ? 'on' : ''} onClick={() => setSettings((s) => ({ ...s, seatOrder: 'mixed' }))}>Mixed batches</button>
          <button className={settings.seatOrder === 'roll' ? 'on' : ''} onClick={() => setSettings((s) => ({ ...s, seatOrder: 'roll' }))}>Roll number order</button>
        </div>
        <p className="muted small">{(settings.seatOrder || 'mixed') === 'mixed' ? 'Students of the same batch do not sit side by side, so copying is harder. Reshuffle on the result page gives a new mix.' : 'Seats follow roll numbers, so students of one batch sit together.'}</p>
      </div>

      <div className="ctabar">
        <div><b>{!useOrder.length ? 'Tick at least one building' : short ? 'Some students will have no seat' : 'Ready to create the seating'}</b><span className="muted">{appearing} students in {chosen.length} batches · {useOrder.join(' → ') || 'no buildings'}</span></div>
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
