// Checks for the planning helpers (phases and live seat summaries). Run: node tests/plan.mjs
import { defaultPhase, parsePhaseList, phaseText, groupPhases, seatSummary, fillPreview } from '../src/lib/plan.js';

let fails = 0;
const ok = (c, m) => { console.log(`${c ? 'ok  ' : 'FAIL'} - ${m}`); if (!c) fails++; };

const batches = [
  { batch: 'XA1', course: 'Momentum', count: 40 }, { batch: 'XA2', course: 'Momentum', count: 35 },
  { batch: 'XA10', course: 'Momentum', count: 5 }, { batch: 'XB1', course: 'Momentum', count: 30 },
  { batch: 'ICB1', course: 'Impulse', count: 20 }, { batch: 'ICB2', course: 'Impulse', count: 25 },
];

console.log('1. Default phases');
ok(defaultPhase('XA1') === 'XA' && defaultPhase('XA10') === 'XA' && defaultPhase('icb2') === 'ICB' && defaultPhase('PRT') === 'PRT', 'phase = batch code without trailing digits');
const g = groupPhases(batches);
ok(g.map((x) => x.key).join(',') === 'Impulse|ICB,Momentum|XA,Momentum|XB', `grouped by course and phase (${g.map((x) => x.key).join(', ')})`);
const xa = g.find((x) => x.key === 'Momentum|XA');
ok(xa.batches.map((b) => b.batch).join(',') === 'XA1,XA2,XA10' && xa.count === 80, 'XA has XA1, XA2, XA10 (natural order) and 80 students');

console.log('\n2. Pasted phase list');
const text = 'Course | Phase | Batches\nMomentum | Phase 1 | XA1, XB1\nImpulse\tICB-All\tICB1 ICB2\nbad line\nMomentum | Phase 1 | xa2\nMomentum | Ghost | ZZ9';
const p = parsePhaseList(text);
ok(p.phases.length === 3 && p.issues.length === 1 && /Line 4/.test(p.issues[0]), `3 phases, 1 issue (${p.issues.join('; ')})`);
ok(p.phases[0].batches.join(',') === 'XA1,XB1,XA2', 'repeated phase lines are merged, batches upper-cased');
ok(phaseText(p.phases).split('\n')[1] === 'Impulse | ICB-All | ICB1, ICB2', 'phase list written back in the same format');
const g2 = groupPhases(batches, p.phases);
const ph1 = g2.find((x) => x.key === 'Momentum|Phase 1');
ok(ph1 && ph1.batches.map((b) => b.batch).join(',') === 'XA1,XA2,XB1' && ph1.pasted, 'pasted phase overrides the default grouping');
ok(g2.find((x) => x.key === 'Momentum|XA').batches.map((b) => b.batch).join(',') === 'XA10', 'batches not in the pasted list keep the default phase');
ok(!g2.some((x) => x.phase === 'Ghost'), 'a pasted phase with no students in the list is left out');
ok(g2.reduce((a, x) => a + x.batches.length, 0) === batches.length, 'every batch in exactly one phase');

console.log('\n3. Seat summary');
const s1 = seatSummary({ online: 100, offline: 50, labSeats: 80, roomSeats: 100 });
ok(s1.overflow === 20 && s1.offline === 70 && s1.offlineLeft === 30 && s1.onlineLeft === 0, 'labs full: 20 online students counted offline, 30 offline seats left');
const s2 = seatSummary({ online: 10, offline: 200, labSeats: 30, roomSeats: 150 });
ok(s2.onlineLeft === 20 && s2.offlineShort === 50, '20 lab seats empty, 50 offline students short');

console.log('\n4. Fill preview');
const f = fillPreview(70, [{ building: 'GP', room: '1', capacity: 30 }, { building: 'M1', room: '1', capacity: 50 }, { building: 'GP', room: '2', capacity: 30 }], ['GP', 'M1']);
ok(f.rooms.map((r) => `${r.building}${r.room}:${r.filled}`).join(',') === 'GP1:30,GP2:30,M11:10' && f.left === 0, 'fills in building order, one room before the next');
ok(fillPreview(100, [{ building: 'GP', room: '1', capacity: 30 }]).left === 70, '70 students left when seats run out');

console.log(fails ? `\n${fails} FAILED` : '\nAll plan checks passed');
process.exit(fails ? 1 : 0);
