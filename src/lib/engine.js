// Core seating allocation logic (pure JS, no UI). Used by the browser app and by tests.
//
// How a seating is made (one paper, one call to allocate):
//  1. Each student is Online (computer lab) or Offline (classroom): a student's own choice wins over the batch's.
//  2. Special students (Monitoring, Due fee, Top batch) go first into the rooms chosen for their group.
//  3. Everyone else is shuffled together and fills the chosen buildings in the order Operations set,
//     one room completely before the next.
//  4. Girls only sit in buildings and labs marked "Girls allowed". Enough of those seats are kept for them.
//  5. Inside a room, seat numbers follow roll-number order. CB (the SMD sheet group) comes from the batch groups.

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

  /**
   * Fill rooms in order, one completely before the next, keeping enough girls-allowed seats for the girls.
   * Returns the students who did not fit.
   */
  const fill = (queue, roomList) => {
    // Keep the shuffled order: take the next student who may sit in this room.
    let girlsLeft = queue.filter(isGirl).length;
    const q = queue.slice();
    for (let k = 0; k < roomList.length && q.length; k++) {
      const room = roomList[k];
      const girlSeatsLater = roomList.slice(k + 1).filter((r) => r.girls).reduce((a, r) => a + free(r), 0);
      while (free(room) > 0 && q.length) {
        let idx = 0;
        if (!room.girls) idx = q.findIndex((s) => !isGirl(s));
        else if (girlsLeft - girlSeatsLater >= free(room)) idx = q.findIndex(isGirl); // otherwise some girls would have no seat
        if (idx < 0) break;
        const [s] = q.splice(idx, 1);
        seat(room, s);
        if (isGirl(s)) girlsLeft--;
      }
    }
    return q;
  };

  const unseated = [];
  const noRoomReason = (s, kind) => (isGirl(s) && !(kind === 'Online' ? labs : rooms).some((r) => r.girls)
    ? `No ${kind === 'Online' ? 'lab' : 'building'} allows girls`
    : `No ${kind === 'Online' ? 'lab' : 'classroom'} seat left`);

  for (const kind of ['Online', 'Offline']) {
    const roomList = kind === 'Online' ? labs : rooms;
    let people = list.filter((s) => modeOf(s) === kind);
    if (!people.length) continue;
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
      const left = fill(shuffle(group, rng), own);
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
    const left = fill(shuffle(people, rng), ordered);
    if (left.length) {
      const g = left.filter(isGirl).length;
      warnings.push(`${kind === 'Online' ? 'Labs' : 'Classrooms'} are full: ${left.length} ${kind.toLowerCase()} students not seated${g ? ` (${g} girls)` : ''}. Switch on more ${kind === 'Online' ? 'labs' : 'rooms or buildings'}${g ? ', or allow girls in another building' : ''}.`);
      unseated.push(...left.map((s) => ({ ...s, Mode: kind, Reason: noRoomReason(s, kind) })));
    }
  }

  // 5. Seat numbers in roll-number order inside each room
  const assigned = [];
  const all = [...labs, ...rooms];
  for (const room of all) {
    room.occupants.sort((a, b) => naturalCompare(a.Roll_Number, b.Roll_Number));
    room.occupants.forEach((s, idx) => {
      assigned.push({ ...s, Building: room.building, 'Room No.': room.room, 'Seat No.': room.taken + idx + 1, CB: (config.cbOf || {})[s.Batch] || phaseOf(s.Batch) });
    });
  }

  return { assigned, unseated, warnings, usedRooms: all.filter((r) => r.used > 0) };
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
