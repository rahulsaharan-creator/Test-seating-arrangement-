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

// RPT: which batches (or part of a batch) sit at each time.
// batches: [{ batch, count, girls }] in phase order. n: number of times. cap: lab seats at one time.
// fill 'labs': each time takes students up to the lab seats, in phase and batch order; a batch that does not fit is split
// and the rest of it starts the next time. fill 'equal': whole batches, the biggest to the time with the fewest students.
// pins[i]: batches the operator put at time i (whole, placed first). outs[i]: batches the operator took out of time i.
// noGirlsAt: the time where batches with girls may not sit (-1: none).
// reserve[i]: lab seats blocked at time i (labs kept for top batches seated separately); the batches get only the rest.
// Returns lists[i] = { batches, parts: { batch: [from, to) }, load } and the students left without a time.
// RPT girls rule: no girl sits a time that starts at or after the cut-off ("late" times). The girls of a late time swap
// places with the same number of boys from the earlier times. The boys come from one batch: all its boys first, then
// the next batch only if more are needed. donor: the batch the operator picked to go first.
// lists[i] = { batches, parts }; studentsOf(batch) = [{ roll, girl }] in roll order; late = Set of late time numbers.
// Returns at[i] = { batch: [rolls] } after the swap, and what moved (girls, their batches, the boys' batches, girls left).
export function girlSwap({ lists, late, studentsOf, donor = '' }) {
  const n = lists.length;
  const at = lists.map((L) => Object.fromEntries(L.batches.map((b) => { const all = studentsOf(b), pt = L.parts && L.parts[b]; return [b, (pt ? all.slice(pt[0], pt[1]) : all).slice()]; })));
  const early = [...Array(n).keys()].filter((i) => !late.has(i));
  const lateList = [...late].filter((i) => i < n).sort((a, b) => a - b);
  const girls = lateList.flatMap((i) => Object.entries(at[i]).flatMap(([b, l]) => l.filter((s) => s.girl).map((s) => ({ s, b, from: i }))));
  // batches whose boys can move: sitting at an earlier time only (not also at a late time), latest earlier time first
  const atLate = new Set(lateList.flatMap((i) => lists[i].batches));
  const cand = [];
  for (const i of [...early].reverse()) for (const b of lists[i].batches) {
    if (atLate.has(b) || cand.some((c) => c.batch === b)) continue;
    const boys = early.reduce((a, k) => a + (at[k][b] || []).filter((s) => !s.girl).length, 0);
    if (boys) cand.push({ batch: b, boys });
  }
  const order = donor && cand.some((c) => c.batch === donor) ? [cand.find((c) => c.batch === donor), ...cand.filter((c) => c.batch !== donor)] : cand;
  const boys = [], donors = [];
  for (const c of order) {
    const need = girls.length - boys.length;
    if (need <= 0) break;
    const list = early.flatMap((k) => (at[k][c.batch] || []).filter((s) => !s.girl).map((s) => ({ s, b: c.batch, from: k }))).slice(0, need);
    donors.push({ batch: c.batch, n: list.length });
    boys.push(...list);
  }
  const k = boys.length;
  const move = (s, b, from, to) => { at[from][b] = at[from][b].filter((x) => x !== s); (at[to][b] = at[to][b] || []).push(s); };
  for (let j = 0; j < k; j++) { const g = girls[j], y = boys[j]; move(g.s, g.b, g.from, y.from); move(y.s, y.b, y.from, g.from); }
  return { at, girls: k, left: girls.length - k, girlBatches: [...new Set(girls.slice(0, k).map((g) => g.b))], donors, cand };
}
export function rptFill({ batches, n, cap = 0, fill = 'labs', noGirlsAt = -1, pins = {}, outs = {}, reserve = {} }) {
  const valid = new Map(batches.map((b) => [b.batch, b]));
  const pinnedTo = {};
  for (const [i, bs] of Object.entries(pins)) if (Number(i) < n) for (const b of bs || []) if (valid.has(b)) pinnedTo[b] = Number(i);
  const lists = Array.from({ length: n }, (_, i) => ({ batches: [], parts: {}, load: Number(reserve[i]) || 0 }));
  for (const [b, i] of Object.entries(pinnedTo)) { lists[i].batches.push(b); lists[i].load += valid.get(b).count; }
  const out = (i, b) => (outs[i] || []).includes(b);
  const girlsBlocked = (i, b) => b.girls > 0 && i === noGirlsAt;
  const rem = Object.fromEntries(batches.map((b) => [b.batch, pinnedTo[b.batch] != null ? 0 : b.count]));
  if (fill === 'labs' && cap > 0) {
    for (let i = 0; i < n; i++) {
      const L = lists[i];
      for (const b of batches) {
        if (L.load >= cap) break;
        if (!rem[b.batch] || out(i, b.batch) || girlsBlocked(i, b)) continue;
        const from = b.count - rem[b.batch], take = Math.min(rem[b.batch], cap - L.load);
        L.batches.push(b.batch);
        if (take < b.count) L.parts[b.batch] = [from, from + take];
        rem[b.batch] -= take; L.load += take;
      }
    }
  } else {
    const order = batches.filter((b) => rem[b.batch]).sort((a, b) => (b.girls > 0) - (a.girls > 0) || b.count - a.count || naturalCompare(a.batch, b.batch));
    for (const b of order) {
      const i = lists.reduce((best, L, k) => (!out(k, b.batch) && !girlsBlocked(k, b) && (best < 0 || L.load < lists[best].load) ? k : best), -1);
      if (i < 0) continue;
      lists[i].batches.push(b.batch); lists[i].load += b.count; rem[b.batch] = 0;
    }
  }
  const leftBatches = batches.filter((b) => rem[b.batch] > 0).map((b) => ({ batch: b.batch, n: rem[b.batch] }));
  return { lists, left: leftBatches.reduce((a, x) => a + x.n, 0), leftBatches };
}
