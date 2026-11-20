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
