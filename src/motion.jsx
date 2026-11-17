// Motion and depth, in the spirit of Apple's site: sliding thumbs on segmented controls and the step bar, cards that
// rise in as they scroll into view, a light 3D tilt under the pointer, numbers that count up, a frosted header that
// gains a line once the page scrolls, and a 3D seat field on the start screen.
// Only transform and opacity move. Everything is off when the operator's system asks for reduced motion.
import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';

const reduced = () => typeof matchMedia !== 'undefined' && matchMedia('(prefers-reduced-motion: reduce)').matches;

// Sliding thumb: every .seg and the step bar get --x / --w of their selected button, a pseudo-element glides there
function placeThumbs() {
  for (const box of document.querySelectorAll('.seg, .stepper')) {
    const on = box.querySelector(':scope > button.on, :scope > .st.on');
    if (!on || !on.offsetWidth) { box.removeAttribute('data-thumb'); continue; }
    // measured against the control itself (also right when the control scrolls sideways, like the step bar)
    const x = on.getBoundingClientRect().left - box.getBoundingClientRect().left - box.clientLeft + box.scrollLeft;
    box.style.setProperty('--x', `${x}px`);
    box.style.setProperty('--w', `${on.offsetWidth}px`);
    if (!box.hasAttribute('data-thumb')) requestAnimationFrame(() => box.setAttribute('data-thumb', ''));
  }
}

let io = null;
function watchReveals() {
  if (reduced() || typeof IntersectionObserver === 'undefined') return;
  io ||= new IntersectionObserver((list) => {
    for (const e of list) if (e.isIntersecting) { e.target.dataset.rv = '1'; io.unobserve(e.target); }
  }, { rootMargin: '0px 0px -40px 0px' });
  let i = 0;
  for (const el of document.querySelectorAll('.card, .fold, .filecard, .rcard, .kpis > div, .paperbox, .slotcard, .rbldg')) {
    if (el.dataset.rv) continue;
    el.dataset.rv = '0';
    el.style.setProperty('--d', `${Math.min(i++, 8) * 40}ms`);
    io.observe(el);
  }
}

// Called by the app on every render: cheap, and keeps thumbs and reveals in step with what is on screen
export function useMotion() {
  useLayoutEffect(() => { placeThumbs(); });
  useEffect(() => { watchReveals(); });
  useEffect(() => {
    const onResize = () => placeThumbs();
    const onScroll = () => { const t = document.querySelector('.top'); if (t) t.toggleAttribute('data-scrolled', scrollY > 4); };
    // 3D tilt: elements marked .tilt lean towards the pointer (at most 4 degrees)
    const onMove = (e) => {
      if (reduced() || e.pointerType === 'touch') return;
      const el = e.target.closest && e.target.closest('.tilt');
      document.querySelectorAll('.tilt[data-tilting]').forEach((x) => { if (x !== el) { x.removeAttribute('data-tilting'); x.style.removeProperty('--rx'); x.style.removeProperty('--ry'); } });
      if (!el) return;
      const r = el.getBoundingClientRect(), px = (e.clientX - r.left) / r.width - 0.5, py = (e.clientY - r.top) / r.height - 0.5;
      el.setAttribute('data-tilting', '');
      el.style.setProperty('--rx', `${(-py * 8).toFixed(2)}deg`);
      el.style.setProperty('--ry', `${(px * 8).toFixed(2)}deg`);
    };
    addEventListener('resize', onResize); addEventListener('scroll', onScroll, { passive: true }); addEventListener('pointermove', onMove, { passive: true });
    document.fonts && document.fonts.ready.then(placeThumbs);
    return () => { removeEventListener('resize', onResize); removeEventListener('scroll', onScroll); removeEventListener('pointermove', onMove); };
  }, []);
}

// A number that counts to its new value (from 0 the first time when from0 is set)
export function Num({ v, from0 = false }) {
  const target = Number(v) || 0;
  const [d, setD] = useState(from0 && !reduced() ? 0 : target);
  const cur = useRef(from0 ? 0 : target);
  useEffect(() => {
    const a = cur.current, b = target;
    if (a === b || reduced()) { cur.current = b; setD(b); return undefined; }
    let raf = 0, t0 = 0;
    const tick = (t) => {
      t0 ||= t; const k = Math.min(1, (t - t0) / 700), e = 1 - Math.pow(1 - k, 3);
      const x = a + (b - a) * e; cur.current = x; setD(x);
      if (k < 1) raf = requestAnimationFrame(tick); else cur.current = b;
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [target]);
  return <>{Math.round(d).toLocaleString('en-IN')}</>;
}

// Start screen: a classroom of seats in 3D. Seats fill with students in a green wave; the room leans with the pointer.
export function Seats3D({ full = false, missing = 0 }) {
  const ref = useRef(null);
  useEffect(() => {
    const el = ref.current; if (!el || reduced()) return undefined;
    const move = (e) => {
      const r = el.getBoundingClientRect(), px = (e.clientX - (r.left + r.width / 2)) / innerWidth, py = (e.clientY - (r.top + r.height / 2)) / innerHeight;
      el.style.setProperty('--tx', `${(56 - py * 10).toFixed(2)}deg`);
      el.style.setProperty('--tz', `${(-32 + px * 14).toFixed(2)}deg`);
    };
    addEventListener('pointermove', move, { passive: true });
    return () => removeEventListener('pointermove', move);
  }, []);
  const rows = full ? 5 : 7, cols = full ? 7 : 9;
  const red = full ? Math.min(missing ? Math.max(1, Math.round(missing / 10)) : 0, rows * cols) : 0;
  return (
    <div className={`seats3d ${full ? 'full' : ''}`} ref={ref} aria-hidden="true">
      <div className="stage">
        <div className="floor" />
        {Array.from({ length: rows * cols }, (_, i) => {
          const r = Math.floor(i / cols), c = i % cols;
          return <span key={i} className={`seat ${i >= rows * cols - red ? 'none' : (r * 7 + c * 3) % 11 === 0 ? 'girl' : ''}`} style={{ '--r': r, '--c': c, '--i': r + c }} />;
        })}
        <div className="desk" />
      </div>
    </div>
  );
}

// A check mark (or, when someone has no seat, an exclamation mark) that draws itself
export const DoneMark = ({ bad = false }) => (
  <span className={`donemark ${bad ? 'bad' : ''}`} aria-hidden="true">
    <svg viewBox="0 0 52 52" width="56" height="56"><circle cx="26" cy="26" r="23" />{bad ? <path d="M26 14v15M26 37v.5" /> : <path d="M15 27l7 7 15-16" />}</svg>
  </span>
);

// ---------- One small 3D scene per step, about what that step does, and driven by its real numbers:
// students: an admission sheet whose rows fill (a tick when the file is read)
// test: a calendar showing the test date and a clock set to the first paper's time
// online / offline: students streaming to computer screens or to desks, in the real online : offline split
// online seating: a row of lab screens lighting up as far as the labs are filled
// offline seating: buildings rising, their windows lit as far as the rooms are filled
const clamp01 = (x) => Math.max(0, Math.min(1, Number.isFinite(x) ? x : 0));
function SceneSheet({ loaded }) {
  return (
    <div className="iso"><div className="iso-in">
      <div className="sheet s3" /><div className="sheet s2" />
      <div className={`sheet s1 ${loaded ? 'done' : ''}`}>
        <div className="sh-head" />
        {Array.from({ length: 6 }, (_, i) => <div key={i} className="sh-row" style={{ '--i': i }}><i /><i /><i /></div>)}
        {loaded && <span className="stamp">✓</span>}
      </div>
      {!loaded && <span className="uparrow">↑</span>}
    </div></div>
  );
}
function SceneCalendar({ date, time }) {
  const d = date ? new Date(`${date}T00:00:00`) : null;
  const [h, m] = (time || '09:00').split(':').map(Number);
  return (
    <div className="calscene">
      <div className="cal">
        <div className="cal-top"><i /><i /></div>
        <div className="cal-m">{d ? d.toLocaleDateString('en-IN', { month: 'short' }) : '—'}</div>
        <div className="cal-d" key={date || 'none'}>{d ? d.getDate() : '?'}</div>
        <div className="cal-w">{d ? d.toLocaleDateString('en-IN', { weekday: 'short' }) : ''}</div>
      </div>
      <div className="clock">
        {Array.from({ length: 12 }, (_, i) => <i key={i} style={{ '--a': `${i * 30}deg` }} />)}
        <span className="hand hh" style={{ '--to': `${(h % 12) * 30 + m / 2}deg` }} key={`h${time}`} />
        <span className="hand mm" style={{ '--to': `${m * 6}deg` }} key={`m${time}`} />
        <span className="pin" />
      </div>
    </div>
  );
}
function SceneSplit({ online, offline }) {
  const total = online + offline, N = 14;
  const nL = total ? Math.round((N * online) / total) : 0;
  return (
    <div className="split">
      <div className="spot lab"><span className="mon"><b /></span><small>{online.toLocaleString('en-IN')}</small></div>
      <div className="stream">{Array.from({ length: N }, (_, i) => <i key={i} className={!total ? 'idle' : i < nL ? 'l' : 'r'} style={{ '--i': i }} />)}</div>
      <div className="spot room"><span className="dsk"><b /></span><small>{offline.toLocaleString('en-IN')}</small></div>
    </div>
  );
}
function SceneLabs({ n, fill }) {
  // 7 screens stand for the seats of the chosen labs; as many light up as are filled
  const k = 7, lit = Math.max(n ? 1 : 0, Math.round(clamp01(fill) * k));
  return (
    <div className="labs3d"><div className="labs-in">
      {Array.from({ length: k }, (_, i) => <span key={i} className={`pc ${i < lit ? 'on' : ''}`} style={{ '--i': i }}><b /></span>)}
    </div></div>
  );
}
function SceneRooms({ n, fill }) {
  // 5 buildings: the chosen ones stand out; their windows light up as far as their rooms are filled
  const k = 5, used = Math.max(1, Math.min(k, n || 1)), f = clamp01(fill);
  return (
    <div className="iso"><div className="iso-in city">
      {Array.from({ length: k }, (_, i) => (
        <span key={i} className={`bld ${i < used ? 'used' : 'idle'}`} style={{ '--i': i, '--x': `${i * 58 - 18}px`, '--h': `${46 + ((i * 37) % 3) * 16}px`, '--lit': i < used ? (f > (i + 0.5) / used ? 0.95 : f > i / used ? 0.6 : 0.2) : 0.06 }}>
          <span className="f top" /><span className="f front" /><span className="f side" />
        </span>
      ))}
    </div></div>
  );
}
export function Scene({ kind, ...p }) {
  const body = { sheet: SceneSheet, calendar: SceneCalendar, split: SceneSplit, labs: SceneLabs, rooms: SceneRooms }[kind];
  return <div className={`scene scene-${kind}`} aria-hidden="true">{body ? React.createElement(body, p) : null}</div>;
}
