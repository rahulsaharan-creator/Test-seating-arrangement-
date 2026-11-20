// RPT fill: lab seats per time, split batches, pins and outs, no girls in the last time (made-up batches)
import assert from 'node:assert';
import { rptFill } from '../src/lib/plan.js';

const B = [{ batch: 'XA1', count: 40, girls: 5 }, { batch: 'XA2', count: 40, girls: 0 }, { batch: 'XB1', count: 30, girls: 0 }, { batch: 'XB2', count: 25, girls: 3 }];

// Each time takes exactly the lab seats; the rest of a batch starts the next time
let r = rptFill({ batches: B, n: 4, cap: 50 });
assert.deepEqual(r.lists.map((l) => l.load), [50, 50, 35, 0]);
assert.deepEqual(r.lists[0].parts, { XA2: [0, 10] });
assert.deepEqual(r.lists[1].parts, { XA2: [10, 40], XB1: [0, 20] });
assert.deepEqual(r.lists[2].parts, { XB1: [20, 30] });
assert.equal(r.left, 0);

// Too few times: students left over are reported
r = rptFill({ batches: B, n: 2, cap: 50 });
assert.equal(r.left, 35);

// No girls in the last time: XB2 (girls) cannot go to time 3, so it is left over
r = rptFill({ batches: B, n: 3, cap: 50, noGirlsAt: 2 });
assert.ok(!r.lists[2].batches.includes('XB2'));
assert.deepEqual(r.leftBatches, [{ batch: 'XB2', n: 25 }]);

// The operator puts XB2 at time 2 (whole, first) and takes XA2 out of time 1
r = rptFill({ batches: B, n: 4, cap: 50, pins: { 1: ['XB2'] }, outs: { 0: ['XA2'] } });
assert.equal(r.lists[1].batches[0], 'XB2');
assert.ok(!r.lists[0].batches.includes('XA2'));
assert.equal(r.lists.reduce((a, l) => a + l.load, 0), 135);
assert.equal(r.left, 0);

// Equal spread keeps batches whole
r = rptFill({ batches: B, n: 4, fill: 'equal', noGirlsAt: 3 });
assert.ok(r.lists.every((l) => Object.keys(l.parts).length === 0));
assert.ok(!r.lists[3].batches.some((b) => B.find((x) => x.batch === b).girls > 0));

console.log('rpt: all checks passed');

// Labs blocked for top batches at time 2 (30 seats): time 2 takes only the other 20 seats
r = rptFill({ batches: B, n: 4, cap: 50, reserve: { 1: 30 } });
assert.equal(r.lists[1].load, 50);
assert.equal(r.lists[1].batches.reduce((a, b) => a + (r.lists[1].parts[b] ? r.lists[1].parts[b][1] - r.lists[1].parts[b][0] : B.find((x) => x.batch === b).count), 0), 20);
assert.equal(r.left, 0);
console.log('rpt blocked labs: ok');

// Girls rule: girls of the late time (3) swap with boys of one batch from the earlier times
import { girlSwap } from '../src/lib/plan.js';
const S = { A: [['a1', 0], ['a2', 0], ['a3', 0]], B: [['b1', 0], ['b2', 0]], C: [['c1', 1], ['c2', 0], ['c3', 1]] };
const studentsOf = (b) => (S[b] || []).map(([roll, g]) => ({ roll, girl: !!g }));
let g = girlSwap({ lists: [{ batches: ['A'], parts: {} }, { batches: ['B'], parts: {} }, { batches: ['C'], parts: {} }], late: new Set([2]), studentsOf });
assert.equal(g.girls, 2);
assert.deepEqual(g.donors, [{ batch: 'B', n: 2 }], 'boys of one batch (the latest earlier time first)');
assert.deepEqual(g.at[2].C.map((s) => s.roll), ['c2']);
assert.deepEqual(g.at[2].B.map((s) => s.roll), ['b1', 'b2']);
assert.deepEqual(g.at[1].C.map((s) => s.roll), ['c1', 'c3']);
assert.equal(g.at.reduce((a, t) => a + Object.values(t).flat().length, 0), 8, 'nobody lost');
g = girlSwap({ lists: [{ batches: ['A'], parts: {} }, { batches: ['B'], parts: {} }, { batches: ['C'], parts: {} }], late: new Set([2]), studentsOf, donor: 'A' });
assert.deepEqual(g.donors, [{ batch: 'A', n: 2 }], 'the batch the operator picked goes first');
console.log('rpt girls rule: ok');
