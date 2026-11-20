// Notice board floors: Floor column first, else the first digit of a 3-digit room number
import assert from 'node:assert';
import { floorOf, noticeData, noticeFloors } from '../src/lib/noticePdf.js';

assert.equal(floorOf({ room: '101(A)' }), '1');
assert.equal(floorOf({ room: '404' }), '4');
assert.equal(floorOf({ room: '12' }), '');
assert.equal(floorOf({ room: 'LAB4-F1' }), '');
assert.equal(floorOf({ room: '7', floor: 'G' }), 'G');
const rooms = [{ building: 'PC', room: '101(A)' }, { building: 'PC', room: '201(A)' }, { building: 'GP', room: '1' }];
const assigned = [['1', 'PC', '101(A)'], ['2', 'PC', '201(A)'], ['3', 'GP', '1']].map(([r, b, n]) => ({ Roll_Number: r, Building: b, 'Room No.': n, 'Seat No.': 1 }));
const data = noticeData([{ result: { assigned, usedRooms: rooms }, meta: { title: 'T', slots: [] } }]);
assert.deepEqual(noticeFloors(data, 'PC'), ['1', '2']);
assert.deepEqual(noticeFloors(data, 'GP'), ['']);
console.log('notice: all checks passed');
