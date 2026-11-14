// Rule checks for the seating engine with made-up students. Run: node tests/engine.mjs
import { allocate, capacity, isGirl, sameBatchNeighbours, orderSeats, keepRelativesApart, relativeNeighbours, dobKey } from '../src/lib/engine.js';

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

// 7. Online first: labs fill first, top batch before others; the rest of the online students sit offline
{
  const st = [...mk(30, 'XA1', 0), ...mk(30, 'XB1', 0, 101), ...mk(20, 'XC1', 0, 201)];
  const labs = [{ ...room('GP', 'LAB', 40), girls: true }];
  const rooms = [room('GP', 1, 60)];
  const flags = Object.fromEntries(mk(30, 'XB1', 0, 101).map((s) => [s.Roll_Number, 'Top batch']));
  const res = allocate(st, { rooms, labs, buildingOrder: ['GP'], batchMode: { XA1: 'Online', XB1: 'Online', XC1: 'Offline' }, flags, seed: 5 });
  console.log('\n7. Online first');
  basic(res, 80);
  const on = res.assigned.filter((s) => s.Mode === 'Online');
  ok(on.length === 40, 'labs filled (40 online)');
  ok(on.filter((s) => s.Batch === 'XB1').length === 30, 'every top batch student online');
  ok(res.movedOffline === 20 && res.unseated.length === 0, '20 online students moved offline, none unseated');
  ok(res.assigned.filter((s) => s.Batch === 'XA1' && s.Mode === 'Online').length === 10, 'XA1 split: 10 online, the rest offline');
}

// 8. Mixed seat order: no two students of the same batch side by side when batches allow it
{
  const st = [...mk(20, 'XA1'), ...mk(20, 'XB1', 3, 101), ...mk(20, 'XC1', 3, 201)];
  const rooms = [room('GP', 1, 30), room('GP', 2, 30)];
  const mixed = allocate(st, { rooms, labs: [], buildingOrder: ['GP'], seed: 3 });
  const roll = allocate(st, { rooms, labs: [], buildingOrder: ['GP'], seed: 3, seatOrder: 'roll' });
  console.log('\n8. Seat order');
  basic(mixed, 60);
  ok(sameBatchNeighbours(mixed.assigned) === 0, 'mixed: no same-batch neighbours');
  ok(sameBatchNeighbours(roll.assigned) > 0, 'roll order: same-batch neighbours exist (old behaviour kept as an option)');
}

// 9. More seat orders and brothers and sisters apart
{
  console.log('\n9. Seat orders by name, father, date of birth; relatives apart');
  const st = [
    { Roll_Number: '3', Student_Name: 'Charu', Father_Name: 'Ram', DOB: '2008-05-01', Contact_Number: '9000000001', Batch: 'A' },
    { Roll_Number: '1', Student_Name: 'Amit', Father_Name: 'Ram', DOB: '31-12-2007', Contact_Number: '9000000001', Batch: 'A' },
    { Roll_Number: '2', Student_Name: 'Bina', Father_Name: 'Shyam', DOB: '15/06/2008', Contact_Number: '9000000002', Batch: 'A' },
    { Roll_Number: '4', Student_Name: 'Dev', Father_Name: 'Mohan', DOB: '2008-01-20', Contact_Number: '9000000003', Batch: 'A' },
  ];
  ok(orderSeats(st, 'name').map((s) => s.Student_Name).join() === 'Amit,Bina,Charu,Dev', 'by student name');
  ok(orderSeats(st, 'father')[0].Father_Name === 'Mohan', "by father's name");
  ok(orderSeats(st, 'dob')[0].Student_Name === 'Amit' && dobKey('31-12-2007') === 20071231, 'by date of birth, dd-mm-yyyy read correctly');
  const byFather = orderSeats(st, 'father');
  ok(relativeNeighbours(byFather) > 0 && relativeNeighbours(keepRelativesApart(byFather)) === 0, 'brothers and sisters moved apart');
}

console.log(fails ? `\n${fails} FAILED` : '\nAll checks passed');
process.exit(fails ? 1 : 0);
