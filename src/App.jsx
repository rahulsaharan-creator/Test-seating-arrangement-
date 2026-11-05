import React, { useEffect, useMemo, useState } from 'react';
import defaults from './data/defaults.json';
import { allocate, naturalCompare } from './lib/engine.js';
import { buildWorkbook, fmtDateDMY, fmtTime } from './lib/exporter.js';
import { readAdmissions, readRooms, readBatchMap } from './lib/readers.js';

const loadExcel = () => import('exceljs').then((m) => m.default || m);
const STORE = 'seating-settings-v1';

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

const STEPS = [
  { id: 'students', label: 'Admission list' },
  { id: 'test', label: 'Test date & shifts' },
  { id: 'batches', label: 'Batches & mode' },
  { id: 'rooms', label: 'Rooms & capacity' },
  { id: 'result', label: 'Generate sheet' },
];

export default function App() {
  const [step, setStep] = useState('students');
  const [settings, setSettings] = useState(loadSettings);
  const [adm, setAdm] = useState(null);
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [test, setTest] = useState({ title: '', date: '', shiftCount: 1, shifts: ['07:00', '10:00'] });
  const [selected, setSelected] = useState(() => new Set());
  const [mode, setMode] = useState({});
  const [seed, setSeed] = useState(() => Math.floor(Math.random() * 1e9));
  const [result, setResult] = useState(null);

  useEffect(() => {
    try { localStorage.setItem(STORE, JSON.stringify(settings)); } catch { /* ignore */ }
  }, [settings]);

  const allBuildings = useMemo(() => {
    const s = new Set([...settings.buildingOrder, ...settings.rooms.map((r) => r.building)]);
    return [...s];
  }, [settings]);

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

  const ready = {
    students: !!adm,
    test: !!test.date && !!test.shifts[0] && (test.shiftCount === 1 || !!test.shifts[1]),
    batches: chosen.length > 0 && missingBuilding.length === 0,
  };

  // ---------- handlers
  async function onAdmission(file) {
    if (!file) return;
    setError(''); setBusy('Reading admission list…');
    try {
      const ExcelJS = await loadExcel();
      const r = await readAdmissions(ExcelJS, await file.arrayBuffer());
      setAdm({ ...r, fileName: file.name });
      setSelected(new Set()); setMode({}); setResult(null);
    } catch (e) { setError(`Could not read “${file.name}”: ${e.message}`); }
    setBusy('');
  }
  async function onRoomFile(file) {
    if (!file) return;
    setError(''); setBusy('Reading room list…');
    try {
      const ExcelJS = await loadExcel();
      const { rooms, labs } = await readRooms(ExcelJS, await file.arrayBuffer());
      setSettings((s) => ({ ...s, rooms, labs: labs.length ? labs : s.labs }));
    } catch (e) { setError(`Could not read room list: ${e.message}`); }
    setBusy('');
  }
  async function onBatchFile(file) {
    if (!file) return;
    setError(''); setBusy('Reading batch list…');
    try {
      const ExcelJS = await loadExcel();
      const batchMap = await readBatchMap(ExcelJS, await file.arrayBuffer());
      setSettings((s) => ({ ...s, batchMap }));
    } catch (e) { setError(`Could not read batch list: ${e.message}`); }
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
  }

  const courses = [...new Set(chosen.map((b) => b.course))];
  const meta = {
    title: test.title || `${courses.join(' & ') || 'Test'} Seating Arrangement`,
    dateISO: test.date,
    shifts: test.shifts.slice(0, test.shiftCount),
    batchesLabel: [...new Set(chosen.map((b) => b.batch.slice(0, 2)))].sort().join(', '),
    modeLabel: need.online && need.offline ? 'Offline + Online Test' : need.online ? 'Online Test' : 'Offline Test',
  };

  async function download() {
    setBusy('Building Excel file…');
    try {
      const ExcelJS = await loadExcel();
      const wb = buildWorkbook(ExcelJS, result, meta);
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `${meta.title.replace(/[^\w]+/g, '_')}_${fmtDateDMY(test.date)}.xlsx`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
    } catch (e) { setError(`Could not build the file: ${e.message}`); }
    setBusy('');
  }

  const canOpen = (i) => STEPS.slice(0, i).every((s) => ready[s.id] ?? true);

  return (
    <div className="shell">
      <aside className="rail">
        <div className="brand">
          <span className="brand-mark" aria-hidden>▦</span>
          <div>
            <div className="brand-name">Test Seating</div>
            <div className="brand-sub">Matrix seating arrangement</div>
          </div>
        </div>
        <ol className="steps">
          {STEPS.map((s, i) => (
            <li key={s.id}>
              <button
                className={`step ${step === s.id ? 'is-current' : ''} ${ready[s.id] ? 'is-done' : ''}`}
                disabled={!canOpen(i)}
                onClick={() => (s.id === 'result' ? run() : setStep(s.id))}
              >
                <span className="step-num">{ready[s.id] && step !== s.id ? '✓' : i + 1}</span>
                <span>{s.label}</span>
              </button>
            </li>
          ))}
        </ol>
        <p className="rail-note">Files are read inside this browser. Student data is not uploaded anywhere.</p>
      </aside>

      <main className="main">
        {error && <div className="alert alert-bad" role="alert">{error} <button className="link" onClick={() => setError('')}>Dismiss</button></div>}
        {busy && <div className="alert">{busy}</div>}

        {step === 'students' && (
          <section>
            <h1>Upload the admission list</h1>
            <p className="lead">The Excel file with every student. It needs the columns Roll_Number, Student_Name, Batch and Course. Only students marked Active are used.</p>
            <label className="drop">
              <input type="file" accept=".xlsx" onChange={(e) => onAdmission(e.target.files[0])} />
              <span className="drop-title">{adm ? adm.fileName : 'Choose admission list (.xlsx)'}</span>
              <span className="drop-sub">{adm ? 'Choose another file to replace it' : 'Click to browse'}</span>
            </label>
            {adm && (
              <div className="facts">
                <div><b>{adm.students.length.toLocaleString('en-IN')}</b> active students</div>
                <div><b>{batches.length}</b> batches</div>
                {adm.duplicates > 0 && <div className="warn"><b>{adm.duplicates}</b> duplicate roll numbers skipped</div>}
                {adm.inactive > 0 && <div><b>{adm.inactive}</b> inactive skipped</div>}
              </div>
            )}
            <details className="advanced">
              <summary>Update room list or batch–building list</summary>
              <p>The website already has your room list and batch–building list built in. Upload a new file only when rooms or class allocations change. Your changes stay saved in this browser.</p>
              <div className="row">
                <label className="btn btn-ghost">Upload room list<input hidden type="file" accept=".xlsx" onChange={(e) => onRoomFile(e.target.files[0])} /></label>
                <label className="btn btn-ghost">Upload batch–building list<input hidden type="file" accept=".xlsx" onChange={(e) => onBatchFile(e.target.files[0])} /></label>
                <button className="btn btn-ghost" onClick={() => { if (confirm('Reset rooms, capacities and batch buildings to the built-in lists?')) setSettings(freshSettings()); }}>Reset to built-in lists</button>
              </div>
              <p className="muted">{settings.rooms.length} rooms, {settings.labs.length} labs, {Object.keys(settings.batchMap).length} batches mapped.</p>
            </details>
            <Nav next={() => setStep('test')} nextOk={ready.students} />
          </section>
        )}

        {step === 'test' && (
          <section>
            <h1>When is the test?</h1>
            <div className="form">
              <label className="field">
                <span>Test date</span>
                <input type="date" value={test.date} onChange={(e) => setTest({ ...test, date: e.target.value })} />
              </label>
              <fieldset className="field">
                <legend>Does the test have one shift or two?</legend>
                <div className="seg">
                  {[1, 2].map((n) => (
                    <button key={n} className={test.shiftCount === n ? 'on' : ''} onClick={() => setTest({ ...test, shiftCount: n })}>
                      {n === 1 ? 'One shift' : 'Two shifts'}
                    </button>
                  ))}
                </div>
                {test.shiftCount === 2 && <p className="muted">The same students sit in the same seat in both shifts. A separate attendance sheet is made for each shift.</p>}
              </fieldset>
              <div className="row">
                <label className="field">
                  <span>Shift 1 time</span>
                  <input type="time" value={test.shifts[0]} onChange={(e) => setTest({ ...test, shifts: [e.target.value, test.shifts[1]] })} />
                </label>
                {test.shiftCount === 2 && (
                  <label className="field">
                    <span>Shift 2 time</span>
                    <input type="time" value={test.shifts[1]} onChange={(e) => setTest({ ...test, shifts: [test.shifts[0], e.target.value] })} />
                  </label>
                )}
              </div>
              <label className="field">
                <span>Sheet title <em className="muted">(optional)</em></span>
                <input type="text" placeholder="e.g. Momentum Seating Arrangement" value={test.title} onChange={(e) => setTest({ ...test, title: e.target.value })} />
              </label>
            </div>
            <Nav back={() => setStep('students')} next={() => setStep('batches')} nextOk={ready.test} hint={!test.date ? 'Pick the test date to continue.' : ''} />
          </section>
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
            back={() => setStep('batches')} next={() => run()} />
        )}

        {step === 'result' && result && (
          <ResultStep result={result} meta={meta} settings={settings} total={need.online + need.offline}
            reshuffle={() => run(Math.floor(Math.random() * 1e9))} download={download}
            back={() => setStep('rooms')} />
        )}
      </main>
    </div>
  );
}

function Nav({ back, next, nextOk = true, nextLabel = 'Continue', hint }) {
  return (
    <div className="nav">
      {back && <button className="btn btn-ghost" onClick={back}>Back</button>}
      {hint && <span className="muted">{hint}</span>}
      {next && <button className="btn" disabled={!nextOk} onClick={next}>{nextLabel}</button>}
    </div>
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
      <h1>Which batches are appearing?</h1>
      <p className="lead">Tick each batch taking the test and choose Offline (classroom) or Online (computer lab). Students start in the building where their class is held.</p>
      <div className="toolbar">
        <input className="search" type="search" placeholder="Find a batch, e.g. XB" value={q} onChange={(e) => setQ(e.target.value)} />
        <div className="tally"><b>{need.offline}</b> offline · <b>{need.online}</b> online</div>
      </div>
      {missing.length > 0 && (
        <div className="alert alert-warn">Choose a building for {missing.map((b) => b.batch).join(', ')} before continuing. The choice is saved for next time.</div>
      )}
      {Object.entries(byCourse).map(([course, list]) => {
        const allOn = list.every((b) => selected.has(b.batch));
        return (
          <div className="course" key={course}>
            <div className="course-head">
              <h2>{course}</h2>
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
                          <span className="muted">{home}</span>
                        ) : (
                          <select value={home || ''} onChange={(e) => setBuilding(b.batch, e.target.value)} aria-label={`Building for ${b.batch}`}>
                            <option value="">Choose building</option>
                            {allBuildings.map((x) => <option key={x}>{x}</option>)}
                          </select>
                        )}
                        <div className="seg seg-sm">
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
      <Nav back={back} next={next} nextOk={nextOk} />
    </section>
  );
}

function RoomStep({ settings, setSettings, need, have, chosen, back, next }) {
  const homes = new Set(chosen.map((b) => settings.batchMap[b.batch]?.building));
  const editRoom = (kind, i, patch) => setSettings((s) => ({ ...s, [kind]: s[kind].map((r, j) => (j === i ? { ...r, ...patch } : r)) }));
  const move = (i, d) => setSettings((s) => {
    const o = s.buildingOrder.slice(); const j = i + d;
    if (j < 0 || j >= o.length) return s;
    [o[i], o[j]] = [o[j], o[i]]; return { ...s, buildingOrder: o };
  });
  const toggleFar = (b) => setSettings((s) => ({ ...s, farBuildings: s.farBuildings.includes(b) ? s.farBuildings.filter((x) => x !== b) : [...s.farBuildings, b] }));

  const short = need.offline > have.offline || need.online > have.online;
  return (
    <section>
      <h1>Check rooms and capacity</h1>
      <p className="lead">Switch off rooms that are busy at test time, or raise a capacity if extra seats are put in. Changes are saved in this browser.</p>
      <div className="facts">
        <div className={need.offline > have.offline ? 'bad' : ''}><b>{need.offline}</b> offline students, <b>{have.offline}</b> classroom seats on</div>
        <div className={need.online > have.online ? 'bad' : ''}><b>{need.online}</b> online students, <b>{have.online}</b> lab seats on</div>
      </div>
      {short && <div className="alert alert-warn">There are not enough seats switched on. Students without a seat will be listed in a “Not_Seated” sheet.</div>}

      <h2>Order for moving students</h2>
      <p className="muted">When a building is full, its extra students go to the next building in this list. Far buildings are used only after every near building is full.</p>
      <ol className="order">
        {settings.buildingOrder.map((b, i) => (
          <li key={b}>
            <span className="order-name">{b}</span>
            {homes.has(b) && <span className="pill">class building today</span>}
            <label className="far"><input type="checkbox" checked={settings.farBuildings.includes(b)} onChange={() => toggleFar(b)} /> far away</label>
            <button className="icon" onClick={() => move(i, -1)} aria-label={`Move ${b} up`} disabled={i === 0}>▲</button>
            <button className="icon" onClick={() => move(i, 1)} aria-label={`Move ${b} down`} disabled={i === settings.buildingOrder.length - 1}>▼</button>
          </li>
        ))}
      </ol>

      <RoomTable title="Classrooms (offline)" kind="rooms" list={settings.rooms} edit={editRoom} open={homes} />
      <RoomTable title="Computer labs (online)" kind="labs" list={settings.labs} edit={editRoom} open={new Set(need.online ? ['*'] : [])} />
      <Nav back={back} next={next} nextLabel="Make seating arrangement" />
    </section>
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
  const byBuilding = {};
  for (const r of result.usedRooms) (byBuilding[r.building] ||= []).push(r);
  const movedCount = result.moved.reduce((a, m) => a + m.count, 0);
  const homeOf = (s) => settings.batchMap[s.Batch]?.building;
  return (
    <section>
      <h1>{meta.title}</h1>
      <p className="lead">{fmtDateDMY(meta.dateISO)} at {meta.shifts.map(fmtTime).join(' and ')} · {meta.modeLabel} · Batches {meta.batchesLabel}</p>
      <div className="facts">
        <div><b>{result.assigned.length}</b> of {total} seated</div>
        <div><b>{result.usedRooms.length}</b> rooms used</div>
        <div className={movedCount ? 'warn' : ''}><b>{movedCount}</b> moved to another building</div>
        {result.unseated.length > 0 && <div className="bad"><b>{result.unseated.length}</b> without a seat</div>}
      </div>
      {result.warnings.map((w) => <div key={w} className="alert alert-warn">{w}</div>)}
      {result.moved.length > 0 && (
        <p className="muted">{result.moved.map((m) => `${m.count} from ${m.from} → ${m.to}`).join(' · ')}</p>
      )}
      <div className="nav nav-top">
        <button className="btn" onClick={download}>Download Excel</button>
        <button className="btn btn-ghost" onClick={reshuffle}>Reshuffle seats</button>
        <button className="btn btn-ghost" onClick={back}>Change rooms</button>
      </div>

      <h2>Seat fill</h2>
      <p className="muted">Each bar is a room. Amber marks students moved in from another building.</p>
      <div className="fillmap">
        {Object.entries(byBuilding).map(([b, rooms]) => (
          <div className="fm-bldg" key={b}>
            <div className="fm-name">{b}</div>
            <div className="fm-rooms">
              {rooms.map((r) => {
                const movedIn = r.occupants.filter((s) => s.Mode === 'Offline' && homeOf(s) !== b).length;
                return (
                  <div className="fm-room" key={r.room} title={`${b} room ${r.room}: ${r.used}/${r.capacity}${movedIn ? `, ${movedIn} moved in` : ''}`}>
                    <div className="fm-bar">
                      <span className="fm-home" style={{ height: `${((r.used - movedIn) / r.capacity) * 100}%` }} />
                      <span className="fm-moved" style={{ height: `${(movedIn / r.capacity) * 100}%` }} />
                    </div>
                    <div className="fm-label">{r.room}</div>
                    <div className="fm-num">{r.used}/{r.capacity}</div>
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </div>
      <p className="muted">The Excel file has these sheets: Raw_Data, Student_list, Seating Arrangement, Attendance per shift, SMD, Room_Layout, Master_List{result.unseated.length ? ', Not_Seated' : ''}.</p>
    </section>
  );
}
