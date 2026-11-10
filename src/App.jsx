import React, { useEffect, useMemo, useRef, useState } from 'react';
import defaults from './data/defaults.json';
import { allocate, naturalCompare } from './lib/engine.js';
import { buildWorkbook, fmtDateDMY, fmtTime } from './lib/exporter.js';
import { readAdmissions, readRooms, readBatchMap } from './lib/readers.js';

const loadExcel = () => import('exceljs').then((m) => m.default || m);
const STORE = 'seating-settings-v1';
const LAST = 'seating-last-test-v1';

const freshSettings = () => JSON.parse(JSON.stringify({
  rooms: defaults.rooms, labs: defaults.labs, batchMap: defaults.batchMap,
  buildingOrder: defaults.buildingOrder, farBuildings: defaults.farBuildings,
}));
function loadSettings() {
  try {
    const s = JSON.parse(localStorage.getItem(STORE));
    if (s && s.rooms && s.batchMap) return s;
  } catch { /* storage unavailable or empty */ }
  return freshSettings();
}
function loadLast() {
  try { return JSON.parse(localStorage.getItem(LAST)) || null; } catch { return null; }
}
const newSlot = (start = '07:00', end = '10:00') => ({ id: Math.random().toString(36).slice(2, 8), label: '', start, end });
function initialTest() {
  const last = loadLast();
  if (last && Array.isArray(last.slots) && last.slots.length) return { title: last.title || '', date: '', slots: last.slots.map((s) => ({ ...newSlot(), ...s })) };
  return { title: '', date: '', slots: [newSlot()] };
}

const STEPS = [
  { id: 'students', label: 'Admission list' },
  { id: 'test', label: 'Date & slots' },
  { id: 'batches', label: 'Batches' },
  { id: 'rooms', label: 'Rooms' },
  { id: 'result', label: 'Seating' },
];
const pl = (n, w) => `${n.toLocaleString('en-IN')} ${w}${n === 1 ? '' : 's'}`;

const Ico = {
  upload: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M12 16V4M7 9l5-5 5 5M4 20h16" /></svg>,
  down: <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M12 4v12M7 11l5 5 5-5M4 20h16" /></svg>,
  bin: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14" /></svg>,
  plus: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" aria-hidden="true"><path d="M12 5v14M5 12h14" /></svg>,
  spark: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m12 3 1.9 5.8L20 10l-4.9 3.6L16.5 20 12 16.5 7.5 20l1.4-6.4L4 10l6.1-1.2z" /></svg>,
  file: <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z" /><path d="M14 3v6h6M8 13h8M8 17h5" /></svg>,
};

// Blank Excel templates, made in the browser
const TEMPLATES = {
  admission: { file: 'Admission_list_template.xlsx', sheets: [['Admission list', ['Roll_Number', 'Student_Name', 'Gender', 'Course', 'Batch', 'Medium', 'Status']]] },
  rooms: { file: 'Room_list_template.xlsx', sheets: [['Room_Capacity', ['Building', 'Room No.', 'Seating Capacity', 'Note']], ['LAB_Capacity', ['Building', 'Lab', 'Capacity', 'Total systems']]] },
  batches: { file: 'Batch_building_list_template.xlsx', sheets: [['Batch buildings', ['Building', 'Floor', 'Room', 'Morning', 'Evening']]] },
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

export default function App() {
  const [step, setStep] = useState('students');
  const [settings, setSettings] = useState(loadSettings);
  const [adm, setAdm] = useState(null);
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [test, setTest] = useState(initialTest);
  const [selected, setSelected] = useState(() => new Set());
  const [mode, setMode] = useState({});
  const [seed, setSeed] = useState(() => Math.floor(Math.random() * 1e9));
  const [result, setResult] = useState(null);

  useEffect(() => {
    try { localStorage.setItem(STORE, JSON.stringify(settings)); } catch { /* ignore */ }
  }, [settings]);
  useEffect(() => { window.scrollTo(0, 0); }, [step]);

  const allBuildings = useMemo(() => [...new Set([...settings.buildingOrder, ...settings.rooms.map((r) => r.building)])], [settings]);

  const batches = useMemo(() => {
    if (!adm) return [];
    const m = {};
    for (const s of adm.students) {
      const k = s.Batch;
      m[k] ||= { batch: k, course: s.Course, count: 0 };
      m[k].count++;
    }
    return Object.values(m).sort((a, b) => naturalCompare(a.batch, b.batch));
  }, [adm]);

  const chosen = batches.filter((b) => selected.has(b.batch));
  const missingBuilding = chosen.filter((b) => !settings.batchMap[b.batch]?.building);
  const need = {
    offline: chosen.filter((b) => mode[b.batch] !== 'Online').reduce((a, b) => a + b.count, 0),
    online: chosen.filter((b) => mode[b.batch] === 'Online').reduce((a, b) => a + b.count, 0),
  };
  const have = {
    offline: settings.rooms.filter((r) => r.enabled).reduce((a, r) => a + Number(r.capacity || 0), 0),
    online: settings.labs.filter((r) => r.enabled).reduce((a, r) => a + Number(r.capacity || 0), 0),
  };
  const slotsOk = test.slots.length > 0 && test.slots.every((s) => s.start && (!s.end || s.end > s.start));

  const ready = {
    students: !!adm,
    test: !!test.date && slotsOk,
    batches: chosen.length > 0 && missingBuilding.length === 0,
    rooms: true,
    result: !!result,
  };

  // ---------- handlers
  async function onAdmission(file) {
    if (!file) return;
    setError(''); setNotice(''); setBusy('Reading admission list…');
    try {
      const ExcelJS = await loadExcel();
      const r = await readAdmissions(ExcelJS, await file.arrayBuffer());
      if (!r.students.length) throw new Error(`no Active students found${r.sheet ? ` on sheet “${r.sheet}”` : ''}${r.inactive ? ` (${r.inactive} rows are not marked Active in the Status column)` : ''}. Check that the Status column says Active, or use the Template link for the expected layout`);
      setAdm({ ...r, fileName: file.name });
      setResult(null);
      // Remember the last test: tick the same batches and modes again when they exist in the new list
      const last = loadLast(); const names = new Set(r.students.map((s) => s.Batch));
      const again = (last?.batches || []).filter((b) => names.has(b));
      setSelected(new Set(again));
      setMode(Object.fromEntries(Object.entries(last?.mode || {}).filter(([b]) => names.has(b))));
      if (again.length) setNotice(`${pl(again.length, 'batch')} from the last test ticked again. Change them on the Batches step.`);
    } catch (e) { setError(`Could not read “${file.name}”: ${e.message}`); }
    setBusy('');
  }
  function removeAdmission() {
    if (!confirm('Remove the admission list from this page? Your rooms and settings stay.')) return;
    setAdm(null); setSelected(new Set()); setMode({}); setResult(null); setNotice(''); setStep('students');
  }
  async function onRoomFile(file) {
    if (!file) return;
    setError(''); setBusy('Reading room list…');
    try {
      const ExcelJS = await loadExcel();
      const { rooms, labs } = await readRooms(ExcelJS, await file.arrayBuffer());
      setSettings((s) => ({ ...s, rooms, labs: labs.length ? labs : s.labs }));
      setNotice(`Room list loaded: ${pl(rooms.length, 'room')}${labs.length ? `, ${pl(labs.length, 'lab')}` : ''}.`);
    } catch (e) { setError(`Could not read the room list: ${e.message}`); }
    setBusy('');
  }
  async function onBatchFile(file) {
    if (!file) return;
    setError(''); setBusy('Reading batch–building list…');
    try {
      const ExcelJS = await loadExcel();
      const batchMap = await readBatchMap(ExcelJS, await file.arrayBuffer());
      setSettings((s) => ({ ...s, batchMap }));
      setNotice(`Batch–building list loaded: ${pl(Object.keys(batchMap).length, 'batch')}.`);
    } catch (e) { setError(`Could not read the batch–building list: ${e.message}`); }
    setBusy('');
  }

  function run(newSeed = seed) {
    const students = adm.students.filter((s) => selected.has(s.Batch));
    const batchBuilding = Object.fromEntries(Object.entries(settings.batchMap).map(([k, v]) => [k, v.building]));
    const batchMode = Object.fromEntries(chosen.map((b) => [b.batch, mode[b.batch] || 'Offline']));
    const r = allocate(students, {
      rooms: settings.rooms, labs: settings.labs, batchBuilding, batchMode,
      buildingOrder: settings.buildingOrder, farBuildings: settings.farBuildings, seed: newSeed,
    });
    setResult(r); setSeed(newSeed); setStep('result');
    try { localStorage.setItem(LAST, JSON.stringify({ title: test.title, slots: test.slots, batches: [...selected], mode })); } catch { /* ignore */ }
  }

  const courses = [...new Set(chosen.map((b) => b.course))];
  const meta = {
    title: test.title || `${courses.join(' & ') || 'Test'} Seating Arrangement`,
    dateISO: test.date,
    shifts: test.slots.map((s) => s.start),
    slots: test.slots,
    batchesLabel: [...new Set(chosen.map((b) => b.batch.slice(0, 2)))].sort().join(', '),
    modeLabel: need.online && need.offline ? 'Offline + Online Test' : need.online ? 'Online Test' : 'Offline Test',
  };

  async function download() {
    setBusy('Building the Excel file…');
    try {
      const ExcelJS = await loadExcel();
      const wb = buildWorkbook(ExcelJS, result, meta);
      const buf = await wb.xlsx.writeBuffer();
      saveBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${meta.title.replace(/[^\w]+/g, '_')}_${fmtDateDMY(test.date)}.xlsx`);
    } catch (e) { setError(`Could not build the file: ${e.message}`); }
    setBusy('');
  }

  const idx = STEPS.findIndex((s) => s.id === step);
  const canOpen = (i) => (STEPS[i].id === 'result' ? !!result : STEPS.slice(0, i).every((s) => ready[s.id]));
  const seatsShort = Math.max(0, need.offline - have.offline) + Math.max(0, need.online - have.online);

  return (
    <div className="app">
      <header className="top">
        <div className="top-in">
          <div className="logo">
            <img src="/brand/matrix-logo-green.png" alt="Matrix Academy" className="logo-img" />
            <span className="sep" aria-hidden="true" />
            <div className="brand"><b>Test <em>Seating</em></b><span className="chip">{test.date ? `${fmtDateDMY(test.date)} · ${pl(test.slots.length, 'slot')}` : 'No test date yet'}</span></div>
          </div>
          <nav className="stepper" aria-label="Steps">
            {STEPS.map((s, i) => (
              <button key={s.id} type="button" disabled={!canOpen(i)} aria-current={step === s.id ? 'step' : undefined}
                className={`st ${step === s.id ? 'on' : ''} ${ready[s.id] && s.id !== 'rooms' && step !== s.id ? 'done' : ''}`}
                onClick={() => setStep(s.id)}>
                <i>{ready[s.id] && s.id !== 'rooms' && step !== s.id ? '✓' : i + 1}</i><span>{s.label}</span>
              </button>
            ))}
          </nav>
        </div>
      </header>

      {adm && (
        <div className="sumbar" role="status">
          <span><b>{adm.students.length.toLocaleString('en-IN')}</b> active students</span>
          <span><b>{chosen.length}</b> of {batches.length} batches ticked</span>
          <span><b>{(need.offline + need.online).toLocaleString('en-IN')}</b> appearing</span>
          <span><b>{test.slots.length}</b> {test.slots.length === 1 ? 'slot' : 'slots'}</span>
          <span className={seatsShort ? 'bad' : chosen.length ? 'good' : ''}><b>{(have.offline + have.online).toLocaleString('en-IN')}</b> seats on{seatsShort ? ` · ${seatsShort} short` : chosen.length ? ' · enough' : ''}</span>
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
            <Nav next={() => setStep('test')} nextOk={ready.students} hint={!adm ? 'Upload the admission list to continue.' : ''} />
          </section>
        )}

        {step === 'test' && (
          <SlotStep test={test} setTest={setTest} slotsOk={slotsOk}
            back={() => setStep('students')} next={() => setStep('batches')} nextOk={ready.test} />
        )}

        {step === 'batches' && (
          <BatchStep
            batches={batches} selected={selected} setSelected={setSelected} mode={mode} setMode={setMode}
            settings={settings} setSettings={setSettings} allBuildings={allBuildings}
            need={need} missing={missingBuilding}
            back={() => setStep('test')} next={() => setStep('rooms')} nextOk={ready.batches}
          />
        )}

        {step === 'rooms' && (
          <RoomStep settings={settings} setSettings={setSettings} need={need} have={have} chosen={chosen}
            onRoomFile={onRoomFile} onBatchFile={onBatchFile}
            back={() => setStep('batches')} create={() => run()} hasResult={!!result} />
        )}

        {step === 'result' && result && (
          <ResultStep result={result} meta={meta} settings={settings} total={need.online + need.offline}
            reshuffle={() => { if (confirm('Reshuffle every seat? The current seating is replaced.')) run(Math.floor(Math.random() * 1e9)); }}
            download={download} back={() => setStep('rooms')} />
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
      <div><span className="kicker">First time here?</span><h2>Make the seating sheet in four steps</h2></div>
      <ol>
        <li><i>1</i><b>Admission list</b><span>Upload the latest admission list.</span></li>
        <li><i>2</i><b>Date and slots</b><span>Pick the date and add as many test slots as you need.</span></li>
        <li><i>3</i><b>Batches</b><span>Tick the batches appearing, Offline or Online.</span></li>
        <li><i>4</i><b>Rooms</b><span>Check seats, then press Create seating.</span></li>
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
        {template && <button type="button" className="tpl" onClick={template}>{Ico.down}Template</button>}
        {loaded && onRemove && <button type="button" className="iconbtn danger" onClick={onRemove} aria-label={`Remove the ${title.toLowerCase()}`} title={`Remove the ${title.toLowerCase()}`}>{Ico.bin}</button>}
      </div>
    </div>
  );
}

function Nav({ back, next, nextOk = true, nextLabel = 'Continue', hint }) {
  return (
    <div className="nav">
      {back && <button className="btn ghost" onClick={back}>Back</button>}
      <span className="nav-hint muted">{hint}</span>
      {next && <button className="btn primary" disabled={!nextOk} onClick={next}>{nextLabel} →</button>}
    </div>
  );
}

function SlotStep({ test, setTest, slotsOk, back, next, nextOk }) {
  const setSlot = (id, patch) => setTest({ ...test, slots: test.slots.map((s) => (s.id === id ? { ...s, ...patch } : s)) });
  const addSlot = () => {
    const last = test.slots[test.slots.length - 1];
    const start = last?.end || '10:00';
    const [h, m] = start.split(':').map(Number);
    const end = `${String(Math.min(23, h + 3)).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    setTest({ ...test, slots: [...test.slots, newSlot(start, end)] });
  };
  const removeSlot = (id) => setTest({ ...test, slots: test.slots.filter((s) => s.id !== id) });
  const hint = !test.date ? 'Pick the test date to continue.' : !test.slots.length ? 'Add at least one slot.' : !slotsOk ? 'Each slot needs a start time, and its end must be after the start.' : '';
  return (
    <section>
      <Head kicker="Step 2 of 4" title="When is the test?" lead="Pick the date and add the test slots. Add as many slots as you need. Students keep the same seat in every slot, and each slot gets its own attendance sheet." />
      <div className="card form">
        <label className="field">
          <span>Test date</span>
          <input type="date" value={test.date} onChange={(e) => setTest({ ...test, date: e.target.value })} />
        </label>
        <label className="field grow">
          <span>Sheet title <em className="muted">(optional)</em></span>
          <input type="text" placeholder="e.g. Momentum Seating Arrangement" value={test.title} onChange={(e) => setTest({ ...test, title: e.target.value })} />
        </label>
      </div>
      <div className="slots-head">
        <h2>Test slots <span className="count">{test.slots.length}</span></h2>
        <button type="button" className="btn ghost" onClick={addSlot}>{Ico.plus}Add slot</button>
      </div>
      <div className="slots">
        {test.slots.map((s, i) => {
          const bad = s.end && s.start && s.end <= s.start;
          return (
            <div className={`slot ${bad ? 'bad' : ''}`} key={s.id}>
              <span className="slot-n">{i + 1}</span>
              <label className="field grow"><span>Name <em className="muted">(optional)</em></span>
                <input type="text" placeholder={`Slot ${i + 1}, e.g. Physics paper`} value={s.label} onChange={(e) => setSlot(s.id, { label: e.target.value })} /></label>
              <label className="field"><span>Starts</span><input type="time" value={s.start} onChange={(e) => setSlot(s.id, { start: e.target.value })} /></label>
              <label className="field"><span>Ends</span><input type="time" value={s.end} onChange={(e) => setSlot(s.id, { end: e.target.value })} /></label>
              <span className="slot-time muted">{s.start ? fmtTime(s.start) : '—'}{s.end ? ` to ${fmtTime(s.end)}` : ''}</span>
              <button type="button" className="iconbtn danger" disabled={test.slots.length === 1} onClick={() => removeSlot(s.id)} aria-label={`Remove slot ${i + 1}`} title="Remove this slot">{Ico.bin}</button>
            </div>
          );
        })}
      </div>
      <button type="button" className="addslot" onClick={addSlot}>{Ico.plus}Add another slot</button>
      <Nav back={back} next={next} nextOk={nextOk} hint={hint} />
    </section>
  );
}

function BatchStep({ batches, selected, setSelected, mode, setMode, settings, setSettings, allBuildings, need, missing, back, next, nextOk }) {
  const [q, setQ] = useState('');
  const byCourse = {};
  for (const b of batches) if (!q || b.batch.toLowerCase().includes(q.toLowerCase())) (byCourse[b.course] ||= []).push(b);
  const toggle = (b) => setSelected((s) => { const n = new Set(s); n.has(b) ? n.delete(b) : n.add(b); return n; });
  const setAll = (list, on) => setSelected((s) => { const n = new Set(s); list.forEach((b) => (on ? n.add(b.batch) : n.delete(b.batch))); return n; });
  const setBuilding = (batch, building) => setSettings((s) => ({ ...s, batchMap: { ...s.batchMap, [batch]: { ...(s.batchMap[batch] || {}), building } } }));

  return (
    <section>
      <Head kicker="Step 3 of 4" title="Which batches are appearing?" lead="Tick each batch taking the test and choose Offline (classroom) or Online (computer lab). Students start in the building where their class is held." />
      <div className="toolbar">
        <input className="search" type="search" placeholder="Find a batch, e.g. XB" value={q} onChange={(e) => setQ(e.target.value)} />
        <div className="tally"><b>{need.offline}</b> offline · <b>{need.online}</b> online</div>
        {selected.size > 0 && <button className="link" onClick={() => setSelected(new Set())}>Clear all</button>}
      </div>
      {missing.length > 0 && (
        <div className="alert warn"><span>Choose a building for <b>{missing.map((b) => b.batch).join(', ')}</b> before continuing. The choice is saved for next time.</span></div>
      )}
      {Object.entries(byCourse).map(([course, list]) => {
        const allOn = list.every((b) => selected.has(b.batch));
        const n = list.filter((b) => selected.has(b.batch)).length;
        return (
          <div className="course" key={course}>
            <div className="course-head">
              <h2>{course} <span className="count">{n}/{list.length}</span></h2>
              <button className="link" onClick={() => setAll(list, !allOn)}>{allOn ? 'Clear all' : 'Select all'}</button>
            </div>
            <div className="batch-grid">
              {list.map((b) => {
                const on = selected.has(b.batch);
                const home = settings.batchMap[b.batch]?.building;
                return (
                  <div key={b.batch} className={`batch ${on ? 'on' : ''}`}>
                    <label className="batch-top">
                      <input type="checkbox" checked={on} onChange={() => toggle(b.batch)} />
                      <span className="batch-name">{b.batch}</span>
                      <span className="batch-count">{b.count}</span>
                    </label>
                    {on && (
                      <div className="batch-opts">
                        {home && !missing.find((m) => m.batch === b.batch) ? (
                          <span className="muted small">{home}</span>
                        ) : (
                          <select value={home || ''} onChange={(e) => setBuilding(b.batch, e.target.value)} aria-label={`Building for ${b.batch}`}>
                            <option value="">Choose building</option>
                            {allBuildings.map((x) => <option key={x}>{x}</option>)}
                          </select>
                        )}
                        <div className="seg">
                          {['Offline', 'Online'].map((m) => (
                            <button key={m} className={(mode[b.batch] || 'Offline') === m ? 'on' : ''} onClick={() => setMode({ ...mode, [b.batch]: m })}>{m}</button>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
      <Nav back={back} next={next} nextOk={nextOk} hint={!selected.size ? 'Tick at least one batch.' : ''} />
    </section>
  );
}

function RoomStep({ settings, setSettings, need, have, chosen, onRoomFile, onBatchFile, back, create, hasResult }) {
  const homes = new Set(chosen.map((b) => settings.batchMap[b.batch]?.building));
  const editRoom = (kind, i, patch) => setSettings((s) => ({ ...s, [kind]: s[kind].map((r, j) => (j === i ? { ...r, ...patch } : r)) }));
  const move = (i, d) => setSettings((s) => {
    const o = s.buildingOrder.slice(); const j = i + d;
    if (j < 0 || j >= o.length) return s;
    [o[i], o[j]] = [o[j], o[i]]; return { ...s, buildingOrder: o };
  });
  const toggleFar = (b) => setSettings((s) => ({ ...s, farBuildings: s.farBuildings.includes(b) ? s.farBuildings.filter((x) => x !== b) : [...s.farBuildings, b] }));
  const offShort = Math.max(0, need.offline - have.offline), onShort = Math.max(0, need.online - have.online);
  const offRooms = settings.rooms.filter((r) => !r.enabled);
  const switchOn = (n, kind) => setSettings((s) => {
    let left = n; const list = s[kind].map((r) => {
      if (left > 0 && !r.enabled && !/^self/i.test(r.note || '')) { left -= Number(r.capacity || 0); return { ...r, enabled: true }; }
      return r;
    });
    return { ...s, [kind]: list };
  });

  return (
    <section>
      <Head kicker="Step 4 of 4" title="Check rooms and capacity" lead="Switch off rooms that are busy at test time, or raise a capacity if extra seats are put in. Changes are saved in this browser." />
      <div className="meters">
        <Meter label="Classroom seats" need={need.offline} have={have.offline} />
        <Meter label="Computer lab seats" need={need.online} have={have.online} />
      </div>
      {(offShort > 0 || onShort > 0) && (
        <div className="alert bad">
          <span>
            {offShort > 0 && <><b>{offShort}</b> offline students more than classroom seats. </>}
            {onShort > 0 && <><b>{onShort}</b> online students more than lab seats. </>}
            Switch on more rooms or raise capacities below; otherwise they go to a Not_Seated sheet.
          </span>
          {offShort > 0 && offRooms.length > 0 && <button className="btn sm" onClick={() => switchOn(offShort, 'rooms')}>Switch on rooms to fit</button>}
        </div>
      )}

      <h2>Order for moving students</h2>
      <p className="muted">When a building is full, its extra students go to the next building in this list. Far buildings are used only after every near building is full.</p>
      <ol className="order">
        {settings.buildingOrder.map((b, i) => (
          <li key={b}>
            <span className="order-n">{i + 1}</span>
            <span className="order-name">{b}</span>
            {homes.has(b) && <span className="pill">class building today</span>}
            <label className="far"><input type="checkbox" checked={settings.farBuildings.includes(b)} onChange={() => toggleFar(b)} /> far away</label>
            <button className="iconbtn" onClick={() => move(i, -1)} aria-label={`Move ${b} up`} disabled={i === 0}>▲</button>
            <button className="iconbtn" onClick={() => move(i, 1)} aria-label={`Move ${b} down`} disabled={i === settings.buildingOrder.length - 1}>▼</button>
          </li>
        ))}
      </ol>

      <RoomTable title="Classrooms (offline)" kind="rooms" list={settings.rooms} edit={editRoom} open={homes} />
      <RoomTable title="Computer labs (online)" kind="labs" list={settings.labs} edit={editRoom} open={new Set(need.online ? ['*'] : [])} />

      <details className="fold">
        <summary>Room list and batch–building list <span className="muted">· {settings.rooms.length} rooms, {settings.labs.length} labs, {Object.keys(settings.batchMap).length} batches mapped</span></summary>
        <p className="muted">Built in already. Upload a new file only when rooms or class allocations change. Uploads stay saved in this browser.</p>
        <div className="files">
          <FileCard title="Room list" hint="Sheets Room_Capacity and LAB_Capacity" onFile={onRoomFile} template={() => downloadTemplate('rooms')} />
          <FileCard title="Batch–building list" hint="Columns Building, Floor, Room, Morning, Evening" onFile={onBatchFile} template={() => downloadTemplate('batches')} />
        </div>
        <button className="btn ghost danger" onClick={() => { if (confirm('Reset rooms, capacities and batch buildings to the built-in lists?')) setSettings(freshSettings()); }}>{Ico.bin}Reset to built-in lists</button>
      </details>

      <div className="ctabar">
        <div><b>{offShort || onShort ? 'Some students will have no seat' : 'Ready to create the seating'}</b><span className="muted">{need.offline + need.online} students in {chosen.length} batches. The seating opens on the next screen.</span></div>
        <div className="cta-row">
          <button className="btn ghost" onClick={back}>Back</button>
          <button className="btn cta" onClick={create}>{Ico.spark}{hasResult ? 'Create seating again' : 'Create seating'}</button>
        </div>
      </div>
    </section>
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

function RoomTable({ title, kind, list, edit, open }) {
  const groups = {};
  list.forEach((r, i) => (groups[r.building] ||= []).push({ ...r, i }));
  return (
    <>
      <h2>{title}</h2>
      {Object.entries(groups).map(([b, rooms]) => {
        const on = rooms.filter((r) => r.enabled);
        const seats = on.reduce((a, r) => a + Number(r.capacity || 0), 0);
        return (
          <details key={b} className="bldg" open={open.has(b) || open.has('*')}>
            <summary><span className="bldg-name">{b}</span><span className="muted">{on.length} of {rooms.length} rooms on · {seats} seats</span></summary>
            <table className="rooms">
              <thead><tr><th>Use</th><th>Room</th><th>Capacity</th><th>Note</th></tr></thead>
              <tbody>
                {rooms.map((r) => (
                  <tr key={r.i} className={r.enabled ? '' : 'off'}>
                    <td><input type="checkbox" checked={r.enabled} onChange={(e) => edit(kind, r.i, { enabled: e.target.checked })} aria-label={`Use ${b} room ${r.room}`} /></td>
                    <td>{r.room}</td>
                    <td><input className="cap" type="number" min="0" value={r.capacity} onChange={(e) => edit(kind, r.i, { capacity: Math.max(0, Number(e.target.value)) })} aria-label={`Capacity of ${b} room ${r.room}`} /></td>
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

function ResultStep({ result, meta, settings, total, reshuffle, download, back }) {
  const [menu, setMenu] = useState(false);
  const byBuilding = {};
  for (const r of result.usedRooms) (byBuilding[r.building] ||= []).push(r);
  const movedCount = result.moved.reduce((a, m) => a + m.count, 0);
  const homeOf = (s) => settings.batchMap[s.Batch]?.building;
  const slotText = meta.slots.map((s, i) => `${s.label || `Slot ${i + 1}`} ${fmtTime(s.start)}${s.end ? `–${fmtTime(s.end)}` : ''}`).join(' · ');
  return (
    <section>
      <Head kicker="Seating ready" title={meta.title} lead={`${fmtDateDMY(meta.dateISO)} · ${slotText} · ${meta.modeLabel} · Batches ${meta.batchesLabel}`} />
      {result.unseated.length > 0 && (
        <div className="alert bad"><span><b>{result.unseated.length} students have no seat.</b> They are listed in the Not_Seated sheet. Go back to Rooms and switch on more rooms or raise capacities.</span><button className="btn sm" onClick={back}>Fix rooms</button></div>
      )}
      <div className="actions">
        <button className="btn cta" onClick={download}>{Ico.down}Download Excel</button>
        <div className="menu">
          <button className="btn ghost" onClick={() => setMenu(!menu)} aria-expanded={menu}>More ▾</button>
          {menu && (
            <div className="menu-list" onMouseLeave={() => setMenu(false)}>
              <button onClick={() => { setMenu(false); reshuffle(); }}>Reshuffle seats<small>Every student gets a new seat</small></button>
              <button onClick={() => { setMenu(false); back(); }}>Change rooms<small>Back to rooms and capacity</small></button>
              <button onClick={() => { setMenu(false); window.print(); }}>Print this page<small>The seat fill summary</small></button>
            </div>
          )}
        </div>
      </div>
      <div className="kpis">
        <div><b>{result.assigned.length}</b><span>of {total} seated</span></div>
        <div><b>{result.usedRooms.length}</b><span>rooms used</span></div>
        <div className={movedCount ? 'warn' : ''}><b>{movedCount}</b><span>moved to another building</span></div>
        <div><b>{meta.slots.length}</b><span>attendance {meta.slots.length === 1 ? 'sheet' : 'sheets'}</span></div>
      </div>
      {result.warnings.map((w) => <div key={w} className="alert warn"><span>{w}</span></div>)}
      {result.moved.length > 0 && <p className="muted">{result.moved.map((m) => `${m.count} from ${m.from} → ${m.to}`).join(' · ')}</p>}

      <h2>Rooms</h2>
      <p className="muted">Each card is a room. Amber shows students moved in from another building.</p>
      {Object.entries(byBuilding).map(([b, rooms]) => (
        <div className="rbldg" key={b}>
          <div className="rb-head"><b>{b}</b><span className="muted">{rooms.reduce((a, r) => a + r.used, 0)} students · {rooms.length} rooms</span></div>
          <div className="rcards">
            {rooms.map((r) => {
              const movedIn = r.occupants.filter((s) => s.Mode === 'Offline' && homeOf(s) !== b).length;
              const pct = Math.round((r.used / r.capacity) * 100);
              return (
                <div className="rcard" key={r.room} title={`${b} room ${r.room}: ${r.used}/${r.capacity}${movedIn ? `, ${movedIn} moved in` : ''}`}>
                  <div className="rc-top"><b>{r.room}</b><span>{r.used}/{r.capacity}</span></div>
                  <div className="rc-bar"><span className="home" style={{ width: `${((r.used - movedIn) / r.capacity) * 100}%` }} /><span className="moved" style={{ width: `${(movedIn / r.capacity) * 100}%` }} /></div>
                  <span className="muted small">{pct}% full{movedIn ? ` · ${movedIn} moved in` : ''}</span>
                </div>
              );
            })}
          </div>
        </div>
      ))}
      <p className="muted small">The Excel file has these sheets: Raw_Data, Student_list, Seating Arrangement, one attendance sheet per slot, SMD, Room_Layout, Master_List{result.unseated.length ? ', Not_Seated' : ''}.</p>
    </section>
  );
}
