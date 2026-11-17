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
    box.style.setProperty('--x', `${on.offsetLeft}px`);
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
export function Seats3D() {
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
  const rows = 7, cols = 9;
  return (
    <div className="seats3d" ref={ref} aria-hidden="true">
      <div className="stage">
        <div className="floor" />
        {Array.from({ length: rows * cols }, (_, i) => {
          const r = Math.floor(i / cols), c = i % cols;
          return <span key={i} className={`seat ${(r * 7 + c * 3) % 11 === 0 ? 'girl' : ''}`} style={{ '--r': r, '--c': c, '--i': r + c }} />;
        })}
        <div className="desk" />
      </div>
    </div>
  );
}

// A check mark that draws itself when the seating is ready
export const DoneMark = () => (
  <span className="donemark" aria-hidden="true">
    <svg viewBox="0 0 52 52" width="44" height="44"><circle cx="26" cy="26" r="23" /><path d="M15 27l7 7 15-16" /></svg>
  </span>
);
