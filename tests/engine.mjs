// Rule checks for the seating engine with made-up students. Run: node tests/engine.mjs
import { allocate, capacity, isGirl, roomKey, seatNeighbours, planTest, prtModes, rotationTurns, checkOfflineTiming, pairSlots, defaultSlots, toMin, ANTI_CHEAT } from '../src/lib/engine.js';

let fails = 0;
const ok = (c, m) => { console.log(`${c ? 'ok  ' : 'FAIL'} - ${m}`); if (!c) fails++; };

const mk = (n, batch, girlsEvery = 3, start = 1) => Array.from({ length: n }, (_, i) => ({
  Roll_Number: String(start + i), Student_Name: `Demo ${start + i}`, Gender: (girlsEvery && i % girlsEvery === 0) ? 'Female' : 'Male',
  Course: 'Demo', Batch: batch, Medium: 'English',
}));
const room = (building, r, capacity) => ({ building, room: String(r), capacity, enabled: true });
const basic = (res, n) => {
  ok(res.assigned.length + res.unseated.length === n, `everyone accounted for (${res.assigned.length} seated, ${res.unseated.length} not)`);
  ok(new Set(res.assigned.map((s) => s.Roll_Number)).size === res.assigned.length, 'no student seated twice');
  ok(new Set(res.assigned.map((s) => `${s.Building}|${s['Room No.']}|${s['Seat No.']}`)).size === res.assigned.length, 'no seat used twice');
  ok(res.usedRooms.every((r) => r.taken + r.used <= r.capacity), 'no room over capacity');
};

// 1. Building order chosen by Operations; buildings not chosen are not used; rooms fill one after another
{
  const st = [...mk(50, 'XA1'), ...mk(40, 'XB1', 3, 101)];
  const rooms = [room('GP', 1, 30), room('GP', 2, 30), room('M1', 1, 40), room('PC', 1, 100)];
  const res = allocate(st, { rooms, labs: [], buildingOrder: ['M1', 'GP'], seed: 7 });
  console.log('\n1. Building order');
  basic(res, 90);
  ok(!res.assigned.some((s) => s.Building === 'PC'), 'PC not chosen, so not used');
  ok(res.usedRooms[0].building === 'M1' && res.usedRooms[0].used === 40, 'M1 (first in order) filled first');
  ok(res.usedRooms.find((r) => r.building === 'GP' && r.room === '1').used === 30, 'GP room 1 full before room 2');
  ok(new Set(res.usedRooms[0].occupants.map((s) => s.Batch)).size === 2, 'batches mixed inside a room');
}

// 2. Girls allowed: no girls in a building marked No, and enough seats kept for them elsewhere
{
  const st = mk(20, 'XA1', 4); // 5 girls, 15 boys
  const rooms = [room('GP', 1, 10), room('PC', 1, 10)];
  const res = allocate(st, { rooms, labs: [], buildingOrder: ['GP', 'PC'], girlsOk: { PC: false }, seed: 3 });
  console.log('\n2. Girls allowed');
  basic(res, 20);
  ok(!res.assigned.some((s) => s.Building === 'PC' && isGirl(s)), 'no girl in PC');
  ok(res.unseated.length === 0, 'all 5 girls found a seat in GP');
  // tight case: GP 10 seats, 12 girls -> 2 girls cannot sit; boys still fill PC
  const st2 = [...mk(12, 'XA1', 1), ...mk(8, 'XB1', 0, 101)];
  const r2 = allocate(st2, { rooms, labs: [], buildingOrder: ['GP', 'PC'], girlsOk: { PC: false }, seed: 3 });
  ok(r2.unseated.length === 2 && r2.unseated.every(isGirl), '12 girls, 10 girl seats: 2 girls not seated, with a reason');
  ok(r2.warnings.some((w) => /girls/.test(w)), 'warning mentions girls');
  // girls-only reservation when a girls-allowed building comes first
  const st3 = mk(20, 'XA1', 2); // 10 girls
  const r3 = allocate(st3, { rooms: [room('GP', 1, 12), room('PC', 1, 12)], labs: [], buildingOrder: ['GP', 'PC'], girlsOk: { PC: false }, seed: 9 });
  ok(r3.unseated.length === 0 && r3.assigned.filter((s) => s.Building === 'GP').filter(isGirl).length === 10, 'GP first: keeps its seats for all 10 girls');
}

// 3. Online / Offline per batch, with single students moved
{
  const st = [...mk(10, 'XA1'), ...mk(10, 'XB1', 3, 101)];
  const labs = [{ building: 'M3', room: 'LAB-3', capacity: 30, enabled: true }];
  const res = allocate(st, { rooms: [room('M3', 1, 30)], labs, buildingOrder: ['M3'], batchMode: { XA1: 'Online' }, studentMode: { 1: 'Offline', 101: 'Online' }, seed: 1 });
  console.log('\n3. Online and Offline');
  basic(res, 20);
  const m = Object.fromEntries(res.assigned.map((s) => [s.Roll_Number, s.Mode]));
  ok(m['2'] === 'Online' && m['102'] === 'Offline', 'batch mode used');
  ok(m['1'] === 'Offline' && m['101'] === 'Online', 'a student\'s own choice wins over the batch');
  ok(res.assigned.filter((s) => s.Mode === 'Online').every((s) => s['Room No.'] === 'LAB-3'), 'online students only in the lab');
}

// 4. Special groups go to the rooms kept for them
{
  const st = mk(60, 'XA1');
  const flags = { 1: 'Monitoring', 2: 'Monitoring', 3: 'Due fee', 4: 'Due fee', 5: 'Due fee' };
  const rooms = [room('GP', 1, 30), room('GP', 2, 30), room('GP', 3, 10)];
  const res = allocate(st, { rooms, labs: [], buildingOrder: ['GP'], flags, pins: { Monitoring: ['GP|3'], 'Due fee': ['GP|3'] }, seed: 5 });
  console.log('\n4. Monitoring and due fee rooms');
  basic(res, 60);
  const where = Object.fromEntries(res.assigned.map((s) => [s.Roll_Number, `${s.Building}|${s['Room No.']}`]));
  ok(['1', '2', '3', '4', '5'].every((r) => where[r] === 'GP|3'), 'monitoring and due fee students in GP room 3');
  ok(res.assigned.find((s) => s.Roll_Number === '1').Flag === 'Monitoring', 'flag carried to the sheet');
  ok(res.usedRooms.find((r) => r.room === '1').used === 30, 'others fill normal rooms first');
  // kept room has free seats: others use them before moving to the next building
  const r2 = allocate(mk(64, 'XA1'), { rooms: [...rooms, room('M1', 1, 50)], labs: [], buildingOrder: ['GP', 'M1'], flags, pins: { Monitoring: ['GP|3'] }, seed: 5 });
  ok(!r2.assigned.some((s) => s.Building === 'M1') && r2.usedRooms.find((r) => r.building === 'GP' && r.room === '3').used === 4, 'free seats in a kept room are used before the next building');
}

// 5. SMD groups and overlapping papers
{
  const st = [...mk(5, 'XA1'), ...mk(5, 'ZB1', 3, 101)];
  const res = allocate(st, { rooms: [room('GP', 1, 20)], labs: [], buildingOrder: ['GP'], cbOf: { XA1: 'Group A' }, taken: { 'GP|1': 8 }, seed: 2 });
  console.log('\n5. SMD groups and taken seats');
  basic(res, 10);
  ok(res.assigned.find((s) => s.Batch === 'XA1').CB === 'Group A' && res.assigned.find((s) => s.Batch === 'ZB1').CB === 'Z-Phase', 'CB from batch groups, default by first letter');
  ok(Math.min(...res.assigned.map((s) => s['Seat No.'])) === 9, 'seats start after the 8 used by an overlapping paper');
  const cap = capacity({ rooms: [room('GP', 1, 20), room('PC', 1, 30)], buildingOrder: ['GP', 'PC'], girlsOk: { PC: false } }, 'Offline');
  ok(cap.seats === 50 && cap.girlSeats === 20, 'capacity counts girls-allowed seats separately');
}

// 6. Bigger run: 3,000 students, many rooms
{
  const st = Array.from({ length: 30 }, (_, b) => mk(100, `B${b}`, 3, b * 1000)).flat();
  const rooms = Array.from({ length: 60 }, (_, i) => room(['GP', 'M1', 'M2', 'PC'][i % 4], i, 50 + (i % 5) * 5));
  const t = Date.now();
  const res = allocate(st, { rooms, labs: [], buildingOrder: ['GP', 'M1', 'M2', 'PC'], girlsOk: { PC: false }, seed: 11 });
  console.log(`\n6. 3,000 students in ${Date.now() - t} ms`);
  basic(res, 3000);
  ok(!res.assigned.some((s) => s.Building === 'PC' && isGirl(s)), 'no girl in PC');
}

// ---------- spec v4: test types, paper format, rotation, offline timing, anti-cheating ----------
const lab = (building, r, capacity, extra = {}) => ({ building, room: String(r), capacity, enabled: true, girls: true, ...extra });
const allSeated = (plan) => plan.units.flatMap((u) => u.result.assigned);
const onceEach = (plan, n, label) => {
  const rolls = allSeated(plan).map((s) => s.Roll_Number);
  ok(rolls.length === n && new Set(rolls).size === n, `${label}: all ${n} students seated, each once`);
};
const neighbourClash = (assigned, rooms, key) => {
  let bad = 0;
  const by = {};
  for (const s of assigned) (by[`${s.Building}|${s['Room No.']}`] ||= {})[s['Seat No.']] = s;
  for (const [k, seats] of Object.entries(by)) {
    const room = rooms.find((r) => roomKey(r) === k);
    for (const [n, s] of Object.entries(seats)) for (const m of seatNeighbours(room, Number(n))) if (seats[m] && key(seats[m]) === key(s)) bad++;
  }
  return bad;
};

// 7. RPT: everyone online; no girls in the 5:30-8:30 slot
{
  console.log('\n7. RPT girls slot rule');
  const st = [...mk(75, 'XA1'), ...mk(75, 'XA2', 3, 201)]; // 50 girls, 100 boys
  const labs = [lab('M3', 'LAB-3', 25), lab('M1', 'LAB-1', 15)];
  const slots = defaultSlots('RPT');
  ok(slots[3].start === '17:30' && slots[3].noGirls && !slots[0].noGirls, 'default slots 6-9, 10-1, 2-5, 5:30-8:30; last slot "No girls" on');
  const plan = planTest(st, { testType: 'RPT', format: 'MAIN', slots, labs, rooms: [], buildingOrder: ['M3', 'M1'], papers: [{ id: 'p', name: 'XA', batches: ['XA1', 'XA2'] }], seed: 3, antiCheat: 'mix', dateISO: '2026-11-20' });
  onceEach(plan, 150, 'RPT');
  ok(allSeated(plan).every((s) => s.Mode === 'Online'), 'everyone online');
  const late = plan.units.find((u) => u.slots[0].start === '17:30');
  ok(late && late.students > 0 && !late.result.assigned.some(isGirl), `5:30-8:30 slot used (${late?.students}) with no girls`);
  // switch off: girls may sit late
  const s2 = slots.map((s) => ({ ...s, noGirls: false }));
  const p2 = planTest(mk(160, 'XA1', 1), { testType: 'RPT', slots: s2, labs, rooms: [], buildingOrder: ['M3', 'M1'], papers: [{ id: 'p', name: 'XA', batches: ['XA1'] }], seed: 3 });
  ok(p2.units.find((u) => u.slots[0].start === '17:30').result.assigned.some(isGirl), 'with the switch off, girls can sit the last slot');
  // top batches in their own lab, 9-12 and 2-5; empty seats shown; others may fill them when ticked
  const st3 = [...mk(30, 'TOP1', 0), ...mk(200, 'XB1', 0, 1001)];
  const labs3 = [lab('M3', 'LAB-3', 40), lab('M1', 'LAB-1', 20, { top: true })];
  const cfg3 = { testType: 'RPT', slots: defaultSlots('RPT'), labs: labs3, rooms: [], buildingOrder: ['M3', 'M1'], topBatches: ['TOP1'], papers: [{ id: 'p', name: 'X', batches: ['TOP1', 'XB1'] }], seed: 8 };
  const p3 = planTest(st3, cfg3);
  const topU = p3.units.filter((u) => u.tag === 'top');
  ok(topU.length === 2 && topU.map((u) => u.slots[0].start).join() === '09:00,14:00', 'top batches use 9-12 and 2-5 in the kept lab');
  ok(topU.every((u) => u.result.assigned.every((s) => s.Batch === 'TOP1')) && topU.reduce((a, u) => a + u.emptySeats, 0) === 10, 'kept lab: only top batches, 10 empty seats shown');
  ok(p3.units.filter((u) => u.tag === 'general').every((u) => !u.result.assigned.some((s) => s['Room No.'] === 'LAB-1')), 'other students stay out of the kept lab');
  const p4 = planTest(st3, { ...cfg3, topShare: true });
  ok(p4.units.filter((u) => u.tag === 'top').reduce((a, u) => a + u.emptySeats, 0) === 0, 'ticked: other students fill the empty seats');
}

// 8. JEE Advanced: two papers, two slots that do not overlap, same seat for both
{
  console.log('\n8. JEE Advanced two slots');
  const st = mk(100, 'XA1');
  const plan = planTest(st, { testType: 'RPT', format: 'ADV', slots: defaultSlots('RPT'), labs: [lab('M3', 'LAB-3', 60)], rooms: [], buildingOrder: ['M3'], papers: [{ id: 'p', name: 'XA', batches: ['XA1'] }], seed: 5 });
  onceEach(plan, 100, 'Advanced');
  ok(plan.units.every((u) => u.slots.length === 2 && toMin(u.slots[1].start) >= toMin(u.slots[0].end)), 'every sitting has Paper 1 and Paper 2, not overlapping');
  ok(!plan.units.some((u) => u.slots[1].start === '17:30' && u.result.assigned.some(isGirl)), 'no girls in a pair that uses the No-girls slot');
  const p = pairSlots([{ start: '09:00', end: '12:00' }, { start: '11:00', end: '14:00' }, { start: '14:00', end: '17:00' }]);
  ok(p.pairs.length === 1 && p.single.length === 1, 'overlapping slots are not paired');
  // PRT offline Advanced: Paper 2 must not start before Paper 1 ends
  ok(checkOfflineTiming([{ name: 'X', s1: { start: '09:00', end: '12:00' }, s2: { start: '11:00', end: '14:00' } }], true).length === 1, 'PRT Advanced: overlapping papers are flagged');
  const off = planTest(mk(40, 'XB1'), { testType: 'PRT', format: 'ADV', slots: defaultSlots('PRT'), labs: [], rooms: [room('GP', 1, 50)], buildingOrder: ['GP'], papers: [{ id: 'p', name: 'XB', batches: ['XB1'], s1: { start: '09:00', end: '12:00' }, s2: { start: '14:00', end: '17:00' } }], seed: 2 });
  ok(off.units.length === 1 && off.units[0].slots.length === 2 && off.units[0].result.assigned.length === 40, 'PRT Advanced offline: one seating for both papers (same seat)');
}

// 9. Test series: up to 3 days, day 1 first, each student once
{
  console.log('\n9. Test series day spread');
  const st = mk(250, 'XA1');
  const slots = [{ start: '09:00', end: '12:00' }, { start: '14:00', end: '17:00' }];
  const plan = planTest(st, { testType: 'SERIES', days: 3, slots, labs: [lab('M3', 'LAB-3', 50)], rooms: [], buildingOrder: ['M3'], papers: [{ id: 'p', name: 'XA', batches: ['XA1'] }], seed: 4, dateISO: '2026-11-20' });
  onceEach(plan, 250, 'Series');
  const perDay = [1, 2, 3].map((d) => plan.units.filter((u) => u.day === d).reduce((a, u) => a + u.students, 0));
  ok(perDay.join() === '100,100,50', `day 1 filled first, then 2, then 3 (${perDay.join(', ')})`);
  ok(plan.units.find((u) => u.day === 3).dateISO === '2026-11-22', 'day 3 is two days after the test date');
  const one = planTest(st, { testType: 'SERIES', days: 1, slots, labs: [lab('M3', 'LAB-3', 50)], rooms: [], buildingOrder: ['M3'], papers: [{ id: 'p', name: 'XA', batches: ['XA1'] }], seed: 4 });
  ok(one.units.reduce((a, u) => a + u.result.unseated.length, 0) === 150, 'one day only: 150 have no seat, with a reason');
}

// 10. PRT online rotation
{
  console.log('\n10. Rotation');
  const b = mk(10, 'XA1', 0);
  const r = (a, z) => Array.from({ length: z - a + 1 }, (_, i) => String(a + i));
  ok(rotationTurns(b, [r(1, 3)]).order.slice(0, 3).join() === '4,5,6', 'one file: last time online 1-3 go offline; 4-6 next');
  ok(rotationTurns(b, [r(1, 3), r(4, 6), r(7, 9)]).order.slice(0, 3).join() === '10,1,2', 'after 1-9 had a turn: 10 first, then a new round');
  const t = rotationTurns(b, [r(1, 3), r(4, 6), r(7, 9), ['10', '1', '2']]);
  ok(t.order.slice(0, 3).join() === '3,4,5' && !t.order.includes('1'), 'cycle complete: new round goes on, online last time stay offline');
  // planner: top batch always online; seats run out partway, only the rest go offline
  const st = [...mk(10, 'TOP1', 0, 101), ...mk(10, 'XA1', 0, 1), ...mk(10, 'XB1', 0, 201)];
  const labs = [lab('M3', 'LAB-3', 10)];
  const cfg = { testType: 'PRT', slots: [{ start: '09:00', end: '12:00' }, { start: '14:00', end: '17:00' }], labs, rooms: [room('GP', 1, 40)], buildingOrder: ['M3', 'GP'], topBatches: ['TOP1'],
    papers: [{ id: 'p', name: 'X', batches: ['TOP1', 'XA1', 'XB1'], s1: { start: '07:00', end: '10:00' } }], rotation: { on: true, history: [[...r(1, 3), '201', '202']] }, seed: 6 };
  const m = prtModes(st, cfg);
  ok(m.seats === 20 && st.filter((s) => s.Batch === 'TOP1').every((s) => m.mode[s.Roll_Number] === 'Online'), 'top batch always online (20 online seats over 2 shifts)');
  const xa = m.summary.find((x) => x.batch === 'XA1');
  ok(xa.onlineNow === 5 && xa.onlineLast === 3 && xa.stillBefore === 7 && xa.stillAfter === 2, `summary XA1: online now ${xa.onlineNow}, last time ${xa.onlineLast}, still to get a turn ${xa.stillAfter}`);
  ok(['1', '2', '3'].every((x) => m.mode[x] === 'Offline') && ['4', '5', '6', '7', '8'].every((x) => m.mode[x] === 'Online'), 'XA1: last time online go offline; next 5 online; the rest offline');
  const plan = planTest(st, cfg);
  onceEach(plan, 30, 'PRT rotation');
  ok(allSeated(plan).filter((s) => s.Mode === 'Online').length === 20 && plan.units.filter((u) => u.kind === 'Online').length === 2, 'online seats used in both shifts, each student one shift');
}

// 11. Offline timing rule
{
  console.log('\n11. Offline timing rule');
  const p = (extra) => [{ name: 'XA phase', s1: { start: '07:00', end: '10:00' }, extra }];
  ok(checkOfflineTiming(p([{ start: '08:00', end: '11:00' }])).length === 0, '7-10 and 8-11: fine');
  ok(checkOfflineTiming(p([{ start: '10:00', end: '13:00' }])).length === 1, '7-10 and 10-1: not allowed');
  const cfg = { testType: 'PRT', slots: [], labs: [], rooms: [room('GP', 1, 30), room('GP', 2, 30)], buildingOrder: ['GP'], papers: [{ id: 'p', name: 'XA', batches: ['XA1', 'XA2'], s1: { start: '07:00', end: '10:00' }, extra: [{ start: '10:00', end: '13:00' }] }], seed: 1 };
  const bad = planTest([...mk(20, 'XA1'), ...mk(20, 'XA2', 3, 101)], cfg);
  ok(bad.blocked && bad.units.length === 0 && /starts after the first sitting ends/.test(bad.warnings[0]), 'seating is blocked with a clear warning');
  cfg.papers[0].extra = [{ start: '08:00', end: '11:00' }];
  const good = planTest([...mk(20, 'XA1'), ...mk(20, 'XA2', 3, 101)], cfg);
  onceEach(good, 40, 'two offline sittings');
  const [a, b2] = good.units;
  ok(good.units.length === 2 && !a.result.assigned.some((s) => b2.result.assigned.some((t) => t.Building === s.Building && t['Room No.'] === s['Room No.'] && t['Seat No.'] === s['Seat No.'])), 'overlapping sittings never share a seat');
}

// 12. Anti-cheating options
{
  console.log('\n12. Anti-cheating');
  ok(ANTI_CHEAT.length === 3 && ANTI_CHEAT.every((o) => o.name && o.text), 'exactly 3 named options, each explained');
  const st = ['XA1', 'XA2', 'XA3', 'XB1'].flatMap((b, i) => mk(45, b, 3, i * 100 + 1));
  const rooms = [room('GP', 1, 60), room('GP', 2, 60), room('GP', 3, 60), room('M1', 1, 200)];
  for (const o of ANTI_CHEAT) {
    const res = allocate(st, { rooms, labs: [], buildingOrder: ['GP', 'M1'], seed: 21, antiCheat: o.id });
    basic(res, 180);
    ok(neighbourClash(res.assigned, rooms, (s) => s.Batch) === 0, `${o.name}: no two students of the same batch side by side`);
    const again = allocate(st, { rooms, labs: [], buildingOrder: ['GP', 'M1'], seed: 21, antiCheat: o.id });
    ok(JSON.stringify(again.assigned) === JSON.stringify(res.assigned), `${o.name}: same seed gives the same seating`);
    if (o.id === 'spread') {
      const used = res.usedRooms;
      ok(used.length === 3 && used.every((r) => ['XA1', 'XA2', 'XA3', 'XB1'].every((b) => r.occupants.filter((s) => s.Batch === b).length >= 14)), 'Spread batches: every batch in every room, evenly');
    }
    if (o.id === 'strict') {
      ok(res.usedRooms.every((r) => r.gap) && res.assigned.every((s) => (s['Seat No.'] - 1) % 2 === 0), 'Strict: an empty seat between students when seats allow');
    }
  }
  // Strict in a shared lab: neighbours write different phases' papers
  const two = [...mk(30, 'XA1', 3, 1), ...mk(30, 'ZB1', 3, 501)];
  const labs = [lab('M3', 'LAB-3', 60, { cols: 6 })];
  const plan = planTest(two, { testType: 'RPT', slots: [{ start: '09:00', end: '12:00' }], labs, rooms: [], buildingOrder: ['M3'], papers: [{ id: 'a', name: 'XA phase', batches: ['XA1'] }, { id: 'z', name: 'ZB phase', batches: ['ZB1'] }], seed: 9, antiCheat: 'strict' });
  const res = plan.units[0].result;
  ok(res.assigned.length === 60 && neighbourClash(res.assigned, labs, (s) => s.Phase) === 0, 'Strict: neighbours write a different phase’s paper');
  const tight = allocate(st, { rooms: [room('GP', 1, 100), room('GP', 2, 100)], labs: [], buildingOrder: ['GP'], seed: 2, antiCheat: 'strict' });
  ok(tight.unseated.length === 0 && tight.warnings.some((w) => /empty seat/.test(w)), 'Strict: no gaps when seats are short, with a note');
}

console.log(fails ? `\n${fails} FAILED` : '\nAll checks passed');
process.exit(fails ? 1 : 0);
