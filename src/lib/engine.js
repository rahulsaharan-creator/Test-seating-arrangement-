// Core seating allocation logic (pure JS, no UI). Used by the browser app and by tests.
//
// How a seating is made (one paper, one call to allocate):
//  1. Each student is Online (computer lab) or Offline (classroom): a student's own choice wins over the batch's.
//     Online means "online first": labs fill first (top batches before others); when the labs are full,
//     the remaining online students are seated offline, one by one, not the whole batch.
//  2. Special students (Monitoring, Due fee, Top batch) go first into the rooms chosen for their group.
//  3. Everyone else is shuffled together and fills the chosen buildings in the order Operations set,
//     one room completely before the next.
//  4. Girls only sit in buildings and labs marked "Girls allowed". Enough of those seats are kept for them.
//  5. Inside a room, seats are mixed so that two students of the same batch do not sit side by side
//     (seatOrder 'mixed', the default), or follow roll-number order (seatOrder 'roll').
//     CB (the SMD sheet group) comes from the batch groups.

export const phaseOf = (batch) => {
  const c = String(batch || '').trim().charAt(0).toUpperCase();
  return c ? `${c}-Phase` : 'Unknown';
};

// Natural sort for room numbers like 1, 2, 101(A), PC-LAB
export const naturalCompare = (a, b) =>
  String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });

// Small seeded RNG (mulberry32) so a shuffle can be reproduced from its seed number.
export function makeRng(seed) {
  let t = seed >>> 0;
  return () => {
    t += 0x6d2b79f5;
    let r = Math.imul(t ^ (t >>> 15), 1 | t);
    r ^= r + Math.imul(r ^ (r >>> 7), 61 | r);
    return ((r ^ (r >>> 14)) >>> 0) / 4294967296;
  };
}

export function shuffle(arr, rng) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

export const roomKey = (r) => `${r.building}|${r.room}`;
export const isGirl = (s) => /^(f|female|girl)$/i.test(String(s.Gender || '').trim());

// Special groups, highest first. A student in two lists counts in the first one.
export const FLAGS = ['Monitoring', 'Due fee', 'Top batch'];

/**
 * students: [{Roll_Number, Student_Name, Gender, Course, Batch, Medium}]
 * config: {
 *   rooms: [{building, room, capacity, enabled}],   // classrooms (Offline)
 *   labs:  [{building, room, capacity, enabled, girls}], // computer labs (Online); girls === false means no girls
 *   buildingOrder: ['GP', 'M1', ...],                 // buildings to use, in fill order; others are not used
 *   girlsOk: {GP: true, PC: false, ...},              // classrooms: may girls sit in this building? (default yes)
 *   batchMode: {BATCH: 'Offline' | 'Online'},         // default Offline
 *   studentMode: {ROLL: 'Offline' | 'Online'},        // overrides the batch
 *   flags: {ROLL: 'Monitoring' | 'Due fee' | 'Top batch'},
 *   pins: {'Monitoring': ['GP|1', ...], ...},          // rooms or labs kept for each special group
 *   cbOf: {BATCH: 'Group name'},                      // SMD sheet group; default first letter + "-Phase"
 *   taken: {'GP|1': 12},                              // seats already used by an overlapping paper
 *   seatOrder: 'mixed' | 'roll',                     // seat order inside a room (default 'mixed')
 *   seed: number
 * }
 */
export function allocate(students, config) {
  const rng = makeRng(config.seed || 1);
  const warnings = [];
  const order = config.buildingOrder || [];
  const rank = (b) => { const i = order.indexOf(b); return i < 0 ? Infinity : i; };
  const girlsOk = config.girlsOk || {};
  const taken = config.taken || {};

  // Seat pools in fill order: chosen buildings only, in the chosen order, then room order inside a building.
  const pool = (list, kind) => list
    .map((r, i) => ({ r, i }))
    .filter(({ r }) => r.enabled && Number(r.capacity) > 0 && rank(r.building) < Infinity)
    .sort((a, b) => rank(a.r.building) - rank(b.r.building) || a.i - b.i)
    .map(({ r }) => ({
      ...r, kind, capacity: Number(r.capacity), taken: Number(taken[roomKey(r)] || 0), used: 0, occupants: [],
      girls: kind === 'Online' ? r.girls !== false && girlsOk[r.building] !== false : girlsOk[r.building] !== false,
    }))
    .filter((r) => r.capacity - r.taken > 0);
  const rooms = pool(config.rooms || [], 'Offline');
  const labs = pool(config.labs || [], 'Online');
  const free = (room) => room.capacity - room.taken - room.used;
  const seat = (room, s) => { room.occupants.push({ ...s, Mode: room.kind }); room.used++; };

  // 1. Mode and special group of each student
  const modeOf = (s) => (config.studentMode || {})[s.Roll_Number] || (config.batchMode || {})[s.Batch] || 'Offline';
  const flagOf = (s) => (config.flags || {})[s.Roll_Number] || '';
  const list = students.map((s) => ({ ...s, Flag: flagOf(s) }));
  // Seat order other than "mixed" (roll number, name, father's name...) runs across all rooms: rooms fill in that order.
  // "Mixed" shuffles everyone together. Sorting starts from roll order so ties stay stable.
  const how = config.seatOrder || 'mixed';
  const arrange = (arr) => (how === 'mixed' ? shuffle(arr, rng) : orderSeats(arr.slice().sort((a, b) => naturalCompare(a.Roll_Number, b.Roll_Number)), how));

  /**
   * Fill rooms in order, one completely before the next, keeping enough girls-allowed seats for the girls.
   * Returns the students who did not fit.
   */
  const fill = (queue, roomList) => {
    // Keep the shuffled order: take the next student who may sit in this room.
    // Boys are kept for the boys-only rooms further on in the same way, so those rooms are not left half empty.
    let girlsLeft = queue.filter(isGirl).length;
    let boysLeft = queue.length - girlsLeft;
    const q = queue.slice();
    for (let k = 0; k < roomList.length && q.length; k++) {
      const room = roomList[k];
      const girlSeatsLater = roomList.slice(k + 1).filter((r) => r.girls).reduce((a, r) => a + free(r), 0);
      const boySeatsLater = roomList.slice(k + 1).filter((r) => !r.girls).reduce((a, r) => a + free(r), 0);
      while (free(room) > 0 && q.length) {
        let idx = 0;
        if (!room.girls) idx = q.findIndex((s) => !isGirl(s));
        else if (girlsLeft - girlSeatsLater >= free(room)) idx = q.findIndex(isGirl); // otherwise some girls would have no seat
        else if (girlsLeft > 0 && boysLeft <= boySeatsLater) idx = q.findIndex(isGirl); // otherwise a boys-only room would stay half empty
        if (idx < 0) break;
        const [s] = q.splice(idx, 1);
        seat(room, s);
        if (isGirl(s)) girlsLeft--; else boysLeft--;
      }
    }
    return q;
  };

  const unseated = [];
  const noRoomReason = (s, kind) => (isGirl(s) && !(kind === 'Online' ? labs : rooms).some((r) => r.girls)
    ? `No ${kind === 'Online' ? 'lab' : 'building'} allows girls`
    : `No ${kind === 'Online' ? 'lab' : 'classroom'} seat left`);

  let movedOffline = [];
  for (const kind of ['Online', 'Offline']) {
    const roomList = kind === 'Online' ? labs : rooms;
    let people = list.filter((s) => modeOf(s) === kind);
    if (kind === 'Offline') people = people.concat(movedOffline);
    if (!people.length) continue;
    if (!roomList.length && kind === 'Online') {
      movedOffline = people.map((s) => ({ ...s, Moved: 'Labs full' }));
      warnings.push(`No computer lab is switched on in the chosen buildings: ${people.length} online students are seated offline.`);
      continue;
    }
    if (!roomList.length) {
      warnings.push(`No ${kind === 'Online' ? 'computer lab' : 'classroom'} is switched on in the chosen buildings: ${people.length} ${kind.toLowerCase()} students not seated.`);
      unseated.push(...people.map((s) => ({ ...s, Mode: kind, Reason: `No ${kind === 'Online' ? 'lab' : 'classroom'} chosen` })));
      continue;
    }

    // 2. Special groups first, into the rooms kept for them
    for (const flag of FLAGS) {
      const keys = new Set((config.pins || {})[flag] || []);
      if (!keys.size) continue;
      const own = roomList.filter((r) => keys.has(roomKey(r)));
      if (!own.length) continue;
      const group = people.filter((s) => s.Flag === flag);
      if (!group.length) continue;
      const left = fill(arrange(group), own);
      const placed = new Set(group.filter((s) => !left.includes(s)));
      people = people.filter((s) => !placed.has(s));
      if (left.length) warnings.push(`${flag}: ${left.length} ${kind.toLowerCase()} students did not fit in the rooms kept for them and were seated with everyone else.`);
    }

    // 3. Everyone else, shuffled together, in building order. Inside a building, rooms kept for a group
    //    are used last, so their free seats are filled before anyone moves to the next building.
    const pinned = new Set(Object.values(config.pins || {}).flat());
    const ordered = roomList.map((r, i) => ({ r, i }))
      .sort((a, b) => rank(a.r.building) - rank(b.r.building) || pinned.has(roomKey(a.r)) - pinned.has(roomKey(b.r)) || a.i - b.i)
      .map(({ r }) => r);
    // Online first: top batch students take lab seats before everyone else.
    const queue = kind === 'Online'
      ? [...arrange(people.filter((s) => s.Flag === 'Top batch')), ...arrange(people.filter((s) => s.Flag !== 'Top batch'))]
      : arrange(people);
    const left = fill(queue, ordered);
    if (left.length && kind === 'Online') {
      movedOffline = left.map((s) => ({ ...s, Moved: 'Labs full' }));
      warnings.push(`Labs are full: ${left.length} online students are seated offline.`);
    } else if (left.length) {
      const g = left.filter(isGirl).length;
      warnings.push(`${kind === 'Online' ? 'Labs' : 'Classrooms'} are full: ${left.length} ${kind.toLowerCase()} students not seated${g ? ` (${g} girls)` : ''}. Switch on more ${kind === 'Online' ? 'labs' : 'rooms or buildings'}${g ? ', or allow girls in another building' : ''}.`);
      unseated.push(...left.map((s) => ({ ...s, Mode: kind, Reason: noRoomReason(s, kind) })));
    }
  }

  // 5. Seat numbers inside each room: batches mixed (default) or roll-number order
  const assigned = [];
  const all = [...labs, ...rooms];
  for (const room of all) {
    room.occupants.sort((a, b) => naturalCompare(a.Roll_Number, b.Roll_Number));
    room.occupants = orderSeats(room.occupants, how);
    if (config.apart) room.occupants = keepRelativesApart(room.occupants);
    room.occupants.forEach((s, idx) => {
      assigned.push({ ...s, Building: room.building, 'Room No.': room.room, 'Seat No.': room.taken + idx + 1, CB: (config.cbOf || {})[s.Batch] || phaseOf(s.Batch) });
    });
  }

  return { assigned, unseated, warnings, movedOffline: movedOffline.length, usedRooms: all.filter((r) => r.used > 0) };
}

// Date of birth as a sortable number: accepts 2008-05-31, 31-05-2008, 31/05/2008, 31.05.2008 and Excel dates
export function dobKey(v) {
  const t = String(v || '').trim();
  let m = t.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) return +m[1] * 10000 + +m[2] * 100 + +m[3];
  m = t.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (m) return +m[3] * 10000 + +m[2] * 100 + +m[1];
  const d = new Date(t);
  return isNaN(d) ? Infinity : d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
}

// Seat order inside a room (the list arrives in roll-number order)
export const SEAT_ORDERS = {
  mixed: 'Mixed batches', roll: 'Roll number', name: 'Student name (A to Z)', father: "Father's name (A to Z)", dob: 'Date of birth (oldest first)', contact: 'Contact number',
};
export function orderSeats(list, how) {
  const by = (f) => list.slice().sort((a, b) => f(a, b) || naturalCompare(a.Roll_Number, b.Roll_Number));
  const txt = (k) => (a, b) => String(a[k] || '~').localeCompare(String(b[k] || '~'), undefined, { sensitivity: 'base', numeric: true });
  if (how === 'roll') return list;
  if (how === 'name') return by(txt('Student_Name'));
  if (how === 'father') return by(txt('Father_Name'));
  if (how === 'contact') return by(txt('Contact_Number'));
  if (how === 'dob') return by((a, b) => dobKey(a.DOB) - dobKey(b.DOB) || 0);
  return mixBatches(list);
}

// Brothers and sisters (same father's name or same contact number) do not sit side by side, where it can be avoided
const related = (a, b) => !!a && !!b && ((a.Father_Name && String(a.Father_Name).trim().toLowerCase() === String(b.Father_Name).trim().toLowerCase())
  || (a.Contact_Number && String(a.Contact_Number).replace(/\D/g, '') === String(b.Contact_Number).replace(/\D/g, '') && String(a.Contact_Number).replace(/\D/g, '').length >= 6));
export function keepRelativesApart(list) {
  const out = list.slice();
  const bad = (k) => related(out[k - 1], out[k]) || related(out[k], out[k + 1]);
  for (let i = 1; i < out.length; i++) {
    if (!related(out[i - 1], out[i])) continue;
    // swap seat i with a later seat when that leaves both places without a brother or sister next door
    for (let j = i + 1; j < out.length; j++) {
      [out[i], out[j]] = [out[j], out[i]];
      if (!bad(i) && !bad(j)) break;
      [out[i], out[j]] = [out[j], out[i]];
    }
  }
  return out;
}
export const relativeNeighbours = (list) => list.slice(1).filter((s, i) => related(list[i], s)).length;

// Seat order that keeps two students of the same batch apart: take one student from each batch in turn,
// always from the batch with the most students left (and never the batch of the seat just filled, if another is left).
export function mixBatches(list) {
  const groups = {};
  for (const s of list) (groups[s.Batch] ||= []).push(s);
  const out = [];
  let prev = null;
  while (out.length < list.length) {
    const keys = Object.keys(groups).filter((k) => groups[k].length);
    keys.sort((a, b) => groups[b].length - groups[a].length || naturalCompare(a, b));
    const k = keys.find((x) => x !== prev) ?? keys[0];
    out.push(groups[k].shift());
    prev = k;
  }
  return out;
}

// Same-batch neighbours (seat n and seat n+1) inside the rooms of a result. Used by tests and the summary.
export function sameBatchNeighbours(assigned) {
  const byRoom = {};
  for (const s of assigned) (byRoom[`${s.Building}|${s['Room No.']}`] ||= []).push(s);
  let n = 0;
  for (const list of Object.values(byRoom)) {
    list.sort((a, b) => a['Seat No.'] - b['Seat No.']);
    for (let i = 1; i < list.length; i++) if (list[i].Batch === list[i - 1].Batch && list[i]['Seat No.'] === list[i - 1]['Seat No.'] + 1) n++;
  }
  return n;
}

// Seats free for one kind in the chosen buildings (girls-allowed seats separately).
export function capacity(config, kind) {
  const order = config.buildingOrder || [];
  const list = (kind === 'Online' ? config.labs : config.rooms) || [];
  const on = list.filter((r) => r.enabled && order.includes(r.building));
  const ok = (r) => (kind === 'Online' ? r.girls !== false : true) && (config.girlsOk || {})[r.building] !== false;
  return {
    seats: on.reduce((a, r) => a + Number(r.capacity || 0), 0),
    girlSeats: on.filter(ok).reduce((a, r) => a + Number(r.capacity || 0), 0),
  };
}
