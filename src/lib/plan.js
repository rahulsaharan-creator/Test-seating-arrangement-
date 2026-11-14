// Planning helpers for the operator flow (pure JS, no UI): phases of a paper and live seat summaries.
import { naturalCompare } from './engine.js';

const S = (v) => String(v ?? '').trim();
export const normBatch = (v) => S(v).toUpperCase().replace(/\s+/g, '');
// Default phase of a batch: the batch code without its trailing digits (XA1, XA2 -> XA)
export const defaultPhase = (batch) => normBatch(batch).replace(/\d+$/, '') || normBatch(batch);
export const phaseKey = (course, phase) => `${course}|${phase}`;

/**
 * Pasted phase list, same format as the Counselling Planner: one phase per line, "Course | Phase | Batches"
 * (tabs or | between the parts; batches separated by commas or spaces). A header line "Course | Phase | Batches" is skipped.
 * Returns { phases: [{ course, phase, batches }], issues: ['Line 3: ...'] }
 */
export function parsePhaseList(text) {
  const out = [], issues = [];
  String(text || '').split(/\r?\n/).forEach((line, i) => {
    if (!S(line)) return;
    const parts = line.split(/\t|\|/).map(S);
    if (/^course$/i.test(parts[0])) return;
    if (parts.length < 3) { issues.push(`Line ${i + 1}: needs Course, Phase and Batches.`); return; }
    const batches = [...new Set(parts.slice(2).join(',').split(/[\s,;]+/).map(normBatch).filter(Boolean))];
    if (!parts[0] || !parts[1] || !batches.length) { issues.push(`Line ${i + 1}: course, phase or batches missing.`); return; }
    const o = out.find((x) => x.course === parts[0] && x.phase === parts[1]);
    if (o) o.batches = [...new Set(o.batches.concat(batches))]; else out.push({ course: parts[0], phase: parts[1], batches });
  });
  return { phases: out, issues };
}
export const phaseText = (list) => (list || []).map((x) => `${x.course} | ${x.phase} | ${x.batches.join(', ')}`).join('\n');

/**
 * Phases of the admission list.
 * batches: [{ batch, course, count }]; pasted: [{ course, phase, batches }] from parsePhaseList (optional).
 * A pasted list wins for the batches it names; every other batch is grouped by Course + batch code without trailing digits.
 * A batch belongs to one phase only (the first pasted line that names it).
 * Returns [{ key, course, phase, batches: [batch objects], count, pasted }], sorted by course then phase.
 */
export function groupPhases(batches, pasted = []) {
  const byCode = new Map(batches.map((b) => [normBatch(b.batch), b]));
  const used = new Set();
  const out = new Map();
  const add = (course, phase, b, fromList) => {
    const key = phaseKey(course, phase);
    if (!out.has(key)) out.set(key, { key, course, phase, batches: [], count: 0, pasted: fromList });
    const g = out.get(key);
    g.batches.push(b); g.count += Number(b.count || 0);
  };
  for (const p of pasted || []) {
    for (const code of p.batches) {
      const b = byCode.get(normBatch(code));
      if (!b || used.has(b.batch)) continue;
      used.add(b.batch);
      add(p.course, p.phase, b, true);
    }
  }
  for (const b of batches) if (!used.has(b.batch)) add(b.course || 'Other', defaultPhase(b.batch), b, false);
  const list = [...out.values()];
  for (const g of list) g.batches.sort((a, b) => naturalCompare(a.batch, b.batch));
  return list.sort((a, b) => naturalCompare(a.course, b.course) || naturalCompare(a.phase, b.phase));
}

/**
 * Live seat summary for one paper. Online first: online students beyond the lab seats sit offline.
 * Returns { online, onlineSeats, onlineLeft, onlineShort, overflow, offline, offlineSeats, offlineLeft, offlineShort }
 * where offline includes the overflow from the labs.
 */
export function seatSummary({ online = 0, offline = 0, labSeats = 0, roomSeats = 0 }) {
  const overflow = Math.max(0, online - labSeats);
  const off = offline + overflow;
  return {
    online, onlineSeats: labSeats, onlineLeft: Math.max(0, labSeats - online), onlineShort: overflow, overflow,
    offline: off, offlineSeats: roomSeats, offlineLeft: Math.max(0, roomSeats - off), offlineShort: Math.max(0, off - roomSeats),
  };
}

/**
 * Preview of how n students fill the chosen rooms: building order first, then the room list order,
 * one room completely before the next (the same order the engine uses).
 * rooms: [{ building, room, capacity }]; order: building names in fill order (buildings not listed come last).
 * Returns { rooms: [{ ...room, filled, empty }], left }
 */
export function fillPreview(n, rooms, order = []) {
  const rank = (b) => { const i = order.indexOf(b); return i < 0 ? Infinity : i; };
  let left = Math.max(0, n);
  const list = rooms.map((r, i) => ({ r, i }))
    .sort((a, b) => rank(a.r.building) - rank(b.r.building) || a.i - b.i)
    .map(({ r }) => {
      const cap = Math.max(0, Number(r.capacity || 0));
      const filled = Math.min(cap, left);
      left -= filled;
      return { ...r, capacity: cap, filled, empty: cap - filled };
    });
  return { rooms: list, left };
}
