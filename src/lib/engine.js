// Core seating allocation logic (pure JS, no UI). Used by the browser app and by tests.
//
// How a seating is made (one paper, one call to allocate):
//  1. Each student is Online (computer lab) or Offline (classroom): a student's own choice wins over the batch's.
//  2. Special students (Monitoring, Due fee, Top batch) go first into the rooms chosen for their group.
//  3. Everyone else is shuffled together and fills the chosen buildings in the order Operations set,
//     one room completely before the next.
//  4. Girls only sit in buildings and labs marked "Girls allowed". Enough of those seats are kept for them.
//  5. Inside a room, seat numbers follow roll-number order, or the anti-cheating option when one is chosen.
//     CB (the SMD sheet group) comes from the batch groups.
//
// A whole test (PRT, RPT or Test series; JEE Main or JEE Advanced) is planned by planTest further down,
// which splits students into sittings (time slots) and calls allocate once per sitting.

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
 *   mode: 'Online' | 'Offline',                       // optional: everyone in this kind of room
 *   antiCheat: 'mix' | 'spread' | 'strict',           // optional: see ANTI_CHEAT
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
  const free = (room) => (room.limit ?? room.capacity - room.taken) - room.used;
  const ac = config.antiCheat || '';
  const arrange = (arr) => (ac === 'spread' || ac === 'strict' ? spreadOrder(arr, rng) : shuffle(arr, rng));
  const seat = (room, s) => { room.occupants.push({ ...s, Mode: room.kind }); room.used++; };

  // 1. Mode and special group of each student
  const modeOf = (s) => config.mode || (config.studentMode || {})[s.Roll_Number] || (config.batchMode || {})[s.Batch] || 'Offline';
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
    // Strict: an empty seat between students when there are enough seats
    if (ac === 'strict') {
      const half = roomList.reduce((a, r) => a + Math.ceil((r.capacity - r.taken) / 2), 0);
      if (half >= people.length) roomList.forEach((r) => { r.limit = Math.ceil((r.capacity - r.taken) / 2); r.gap = true; });
      else warnings.push(`Not enough ${kind === 'Online' ? 'lab' : 'classroom'} seats to leave an empty seat between students, so no gaps were left.`);
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
    const left = fill(arrange(people), ordered);
    if (left.length) {
      const g = left.filter(isGirl).length;
      warnings.push(`${kind === 'Online' ? 'Labs' : 'Classrooms'} are full: ${left.length} ${kind.toLowerCase()} students not seated${g ? ` (${g} girls)` : ''}. Switch on more ${kind === 'Online' ? 'labs' : 'rooms or buildings'}${g ? ', or allow girls in another building' : ''}.`);
      unseated.push(...left.map((s) => ({ ...s, Mode: kind, Reason: noRoomReason(s, kind) })));
    }
  }

  // 5. Seat numbers: roll-number order, or arranged so neighbours differ (anti-cheating option)
  const assigned = [];
  const all = [...labs, ...rooms];
  const cbOf = (s) => (config.cbOf || {})[s.Batch] || phaseOf(s.Batch);
  for (const room of all) {
    const seatNo = (idx) => room.taken + (room.gap ? 2 * idx : idx) + 1;
    if (ac) room.occupants = arrangeRoom(room, room.occupants, seatNo, ac === 'strict', (s) => s.Phase || cbOf(s));
    else room.occupants.sort((a, b) => naturalCompare(a.Roll_Number, b.Roll_Number));
    room.occupants.forEach((s, idx) => {
      assigned.push({ ...s, Building: room.building, 'Room No.': room.room, 'Seat No.': seatNo(idx), CB: cbOf(s) });
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

// ======================================================================================
// Anti-cheating
// ======================================================================================

// Exactly three options. The operator picks one before creating the seating. Every option uses the seeded
// shuffle, so the same seed number gives the same seating again.
export const ANTI_CHEAT = [
  { id: 'mix', name: 'Shuffle and mix', text: 'Seats are shuffled and no two students of the same batch sit next to each other.' },
  { id: 'spread', name: 'Spread batches', text: 'Shuffle and mix, and every batch is spread evenly over all the rooms used.' },
  { id: 'strict', name: 'Strict', text: 'Spread batches, plus neighbours write a different phase’s paper where a room is shared, and an empty seat is left between students when there are enough seats.' },
];

// Shuffle each batch, then interleave the batches evenly, so every room gets its share of every batch.
export function spreadOrder(arr, rng) {
  const by = {};
  for (const s of shuffle(arr, rng)) (by[s.Batch] ||= []).push(s);
  const keyed = [];
  for (const list of Object.values(by)) list.forEach((s, i) => keyed.push({ s, k: (i + 0.5) / list.length, t: rng() }));
  return keyed.sort((a, b) => a.k - b.k || a.t - b.t).map((x) => x.s);
}

// Seat grid used by the Room_Layout printout: seats run down each column.
export const gridRows = (room) => {
  const cols = room.cols || Math.min(10, Math.max(2, Math.ceil(Number(room.capacity) / 12)));
  return Math.ceil(Number(room.capacity) / cols);
};
// Seats next to seat n: the seat in front (same column) and the seat beside it (column to the left).
export const seatNeighbours = (room, n) => {
  const rows = gridRows(room);
  const out = [];
  if ((n - 1) % rows !== 0) out.push(n - 1);
  if (n - rows >= 1) out.push(n - rows);
  return out;
};

// Put a room's students on its seats so that neighbours are from different batches (and, for Strict, write
// different phases' papers when the room has more than one). Greedy: on each seat, take the batch with the
// most students left that differs from the neighbours.
function arrangeRoom(room, occupants, seatNo, strict, phaseKey) {
  const queues = new Map();
  for (const s of occupants) { if (!queues.has(s.Batch)) queues.set(s.Batch, []); queues.get(s.Batch).push(s); }
  const phaseOfBatch = new Map([...queues].map(([b, list]) => [b, phaseKey(list[0])]));
  const multiPhase = strict && new Set(phaseOfBatch.values()).size > 1;
  const at = new Map(); // seat number -> student
  const out = [];
  for (let i = 0; i < occupants.length; i++) {
    const n = seatNo(i);
    const near = seatNeighbours(room, n).map((m) => at.get(m)).filter(Boolean);
    const nb = new Set(near.map((s) => s.Batch));
    const np = new Set(near.map((s) => phaseOfBatch.get(s.Batch)));
    const live = [...queues].filter(([, q]) => q.length);
    const best = (ok) => live.filter(ok).reduce((a, x) => (!a || x[1].length > a[1].length ? x : a), null);
    const pick = (multiPhase && best(([b]) => !nb.has(b) && !np.has(phaseOfBatch.get(b)))) || best(([b]) => !nb.has(b)) || best(() => true);
    const s = pick[1].shift();
    at.set(n, s);
    out.push(s);
  }
  // Repair: a seat whose neighbour is from the same batch swaps with a seat where both fit
  const rows = gridRows(room);
  const around = (n) => [...seatNeighbours(room, n), ...((n % rows !== 0) ? [n + 1] : []), n + rows].filter((m) => at.has(m));
  const clash = (n, s) => around(n).some((m) => at.get(m) !== s && at.get(m).Batch === s.Batch);
  const nums = occupants.map((_, i) => seatNo(i));
  for (let pass = 0; pass < 2; pass++) {
    for (const n of nums) {
      const s = at.get(n);
      if (!clash(n, s)) continue;
      for (const m of nums) {
        const t = at.get(m);
        if (t.Batch === s.Batch) continue;
        at.set(n, t); at.set(m, s);
        if (!clash(n, t) && !clash(m, s) && around(n).every((k) => !clash(k, at.get(k))) && around(m).every((k) => !clash(k, at.get(k)))) break;
        at.set(n, s); at.set(m, t);
      }
    }
  }
  return nums.map((n) => at.get(n));
}

// ======================================================================================
// Test types, paper formats, slots and offline timing
// ======================================================================================

export const TEST_TYPES = [
  { id: 'PRT', name: 'PRT', text: 'Online seats first, the rest sit offline in classrooms. Each student sits one shift.' },
  { id: 'RPT', name: 'RPT', text: 'Everyone sits online, in lab slots through the day. Top batches get their own lab.' },
  { id: 'SERIES', name: 'Test series', text: 'Like RPT, but one test can be spread over up to 3 days. Each student sits once.' },
];
export const FORMATS = [
  { id: 'MAIN', name: 'JEE Main', text: '1 paper of 3 hours. Each student needs one slot.' },
  { id: 'ADV', name: 'JEE Advanced', text: '2 papers of 3 hours on the same day. Each student needs two slots and keeps the same seat.' },
];
export const defaultSlots = (type) => (type === 'PRT'
  ? [{ start: '09:00', end: '12:00', noGirls: false }, { start: '14:00', end: '17:00', noGirls: false }]
  : [{ start: '06:00', end: '09:00', noGirls: false }, { start: '10:00', end: '13:00', noGirls: false },
    { start: '14:00', end: '17:00', noGirls: false }, { start: '17:30', end: '20:30', noGirls: true }]);
export const DEFAULT_TOP_SLOTS = [{ start: '09:00', end: '12:00', noGirls: false }, { start: '14:00', end: '17:00', noGirls: false }];

export const GAP = 15; // minutes: sittings closer than this share no seats (students leaving and arriving)
export const toMin = (t) => { const [h, m] = String(t || '0:0').split(':').map(Number); return (h || 0) * 60 + (m || 0); };
export const slotOk = (s) => !!(s && s.start && s.end && toMin(s.end) > toMin(s.start));
export const spanOf = (slots) => ({ start: slots[0].start, end: slots[slots.length - 1].end });
export const overlapping = (a, b, gap = 0) => toMin(a.start) < toMin(b.end) + gap && toMin(b.start) < toMin(a.end) + gap;
export const clock = (t) => { const m = toMin(t); const h = Math.floor(m / 60); return `${((h + 11) % 12) + 1}:${String(m % 60).padStart(2, '0')} ${h >= 12 ? 'PM' : 'AM'}`; };
const addDays = (iso, n) => {
  if (!iso) return iso;
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10);
};

// Offline sittings of one paper: the first time plus any extra sittings (JEE Main only).
export const offlineSittings = (p, adv) => (adv ? [[p.s1, p.s2]] : [p.s1, ...(p.extra || [])].map((s) => [s]));

/**
 * Offline timing rule: every offline sitting of a paper must start before the first sitting of that paper ends
 * (7-10 and 8-11 is fine; 7-10 and 10-1 is not), so no student can pass on the paper. Returns plain-English problems.
 */
export function checkOfflineTiming(papers, adv = false) {
  const problems = [];
  papers.forEach((p, i) => {
    const name = p.name || `Paper ${i + 1}`;
    if (adv) {
      if (slotOk(p.s1) && slotOk(p.s2) && toMin(p.s2.start) < toMin(p.s1.end)) problems.push(`${name}: Paper 2 starts before Paper 1 ends. The two papers must not overlap.`);
      return;
    }
    const list = [p.s1, ...(p.extra || [])].filter(slotOk);
    if (list.length < 2) return;
    const first = list.reduce((a, s) => (toMin(s.start) < toMin(a.start) ? s : a));
    for (const s of list) {
      if (s !== first && toMin(s.start) >= toMin(first.end)) {
        problems.push(`${name}: the offline sitting at ${clock(s.start)} starts after the first sitting ends (${clock(first.end)}). Every offline sitting must start before the first one ends.`);
      }
    }
  });
  return problems;
}

// JEE Advanced: pair slots in time order into (Paper 1, Paper 2) that do not overlap.
export function pairSlots(slots) {
  const list = slots.filter(slotOk).slice().sort((a, b) => toMin(a.start) - toMin(b.start));
  const pairs = [], single = [];
  let i = 0;
  while (i < list.length) {
    if (i + 1 < list.length && toMin(list[i + 1].start) >= toMin(list[i].end)) { pairs.push([list[i], list[i + 1]]); i += 2; } else { single.push(list[i]); i++; }
  }
  return { pairs, single };
}

const labsOn = (cfg) => (cfg.labs || []).filter((r) => r.enabled && Number(r.capacity) > 0 && (cfg.buildingOrder || []).includes(r.building));
const girlsLab = (cfg) => (r) => r.girls !== false && (cfg.girlsOk || {})[r.building] !== false;
const seatsOf = (list) => list.reduce((a, r) => a + Number(r.capacity), 0);

/**
 * Online sittings of a test (before students are placed). RPT and Test series keep labs marked "top" for top
 * batches in their own slots; PRT uses every lab in its shifts. JEE Advanced pairs slots, one sitting = two papers.
 */
export function onlineUnits(cfg) {
  const type = cfg.testType || 'PRT';
  const adv = cfg.format === 'ADV';
  const days = type === 'SERIES' ? Math.min(3, Math.max(1, Number(cfg.days) || 1)) : 1;
  const labs = labsOn(cfg);
  const ownTop = type !== 'PRT';
  const topLabs = ownTop ? labs.filter((l) => l.top) : [];
  const genLabs = ownTop ? labs.filter((l) => !l.top) : labs;
  const unpaired = [];
  const groups = (slots) => {
    if (!adv) return slots.filter(slotOk).map((s) => [s]);
    const { pairs, single } = pairSlots(slots);
    unpaired.push(...single);
    return pairs;
  };
  const make = (slots, list, tag) => {
    const g = groups(slots);
    const out = [];
    for (let day = 1; day <= days; day++) {
      for (const grp of g) {
        const noGirls = grp.some((s) => s.noGirls);
        out.push({ kind: 'Online', tag, day, slots: grp, labs: list, cap: seatsOf(list), gcap: noGirls ? 0 : seatsOf(list.filter(girlsLab(cfg))), noGirls, students: [] });
      }
    }
    return out;
  };
  const top = topLabs.length ? make(cfg.topSlots || DEFAULT_TOP_SLOTS, topLabs, 'top') : [];
  const general = genLabs.length ? make(cfg.slots || defaultSlots(type), genLabs, 'general') : [];
  return { top, general, days, unpaired };
}

const byTime = (a, b) => a.day - b.day || toMin(a.slots[0].start) - toMin(b.slots[0].start);

// Place students into sittings in order (day 1 first, then by time), one completely before the next,
// keeping enough girls-allowed seats for the girls. Returns the students who did not fit.
function assignToUnits(queue, units) {
  const q = queue.slice();
  let girlsLeft = q.filter(isGirl).length;
  const free = (u) => u.cap - u.students.length;
  const gfree = (u) => Math.max(0, Math.min(free(u), u.gcap - u.students.filter(isGirl).length));
  units.forEach((u, k) => {
    while (free(u) > 0 && q.length) {
      const later = units.slice(k + 1).reduce((a, x) => a + gfree(x), 0);
      let idx = 0;
      if (gfree(u) <= 0) idx = q.findIndex((s) => !isGirl(s));
      else if (girlsLeft - later >= free(u)) idx = q.findIndex(isGirl);
      if (idx < 0) break;
      const [s] = q.splice(idx, 1);
      u.students.push(s);
      if (isGirl(s)) girlsLeft--;
    }
  });
  return q;
}

// ======================================================================================
// PRT rotation: who sits online this time
// ======================================================================================

/**
 * One batch's turns. history: previous tests, oldest first, each a list of roll numbers that sat online.
 * A student goes online again only after the whole batch had a turn. Students online last time go offline.
 */
export function rotationTurns(students, history) {
  const ids = students.map((s) => s.Roll_Number).sort(naturalCompare);
  const inBatch = new Set(ids);
  let had = new Set();
  let last = new Set();
  for (const test of history) {
    const on = [...new Set(test.map(String))].filter((r) => inBatch.has(r));
    const again = on.filter((r) => had.has(r));
    on.forEach((r) => had.add(r));
    if (ids.length && ids.every((r) => had.has(r))) had = new Set(again); // cycle complete; repeats start the next one
    last = new Set(on);
  }
  const still = ids.filter((r) => !had.has(r) && !last.has(r));
  const next = ids.filter((r) => had.has(r) && !last.has(r));
  return { ids, last, still, order: [...still, ...next] };
}

/**
 * PRT Online / Offline for every student. Top batches always online. A student's own choice wins.
 * With rotation on, other batches share the remaining online seats in proportion to their size.
 * Returns { mode: {roll: 'Online'|'Offline'}, summary: [{batch, size, onlineNow, onlineLast, stillBefore, stillAfter}], seats }
 */
export function prtModes(students, cfg) {
  const top = new Set(cfg.topBatches || []);
  const own = cfg.studentMode || {};
  const rot = cfg.rotation && cfg.rotation.on;
  const mode = {};
  const pool = {};
  for (const s of students) {
    const r = s.Roll_Number;
    if (top.has(s.Batch)) mode[r] = 'Online';
    else if (own[r]) mode[r] = own[r];
    else if (rot) (pool[s.Batch] ||= []).push(s);
    else mode[r] = (cfg.batchMode || {})[s.Batch] || 'Offline';
  }
  const seats = onlineUnits({ ...cfg, testType: 'PRT' }).general.reduce((a, u) => a + u.cap, 0);
  const summary = [];
  if (rot) {
    let left = Math.max(0, seats - Object.values(mode).filter((m) => m === 'Online').length);
    const info = Object.keys(pool).sort(naturalCompare).map((batch) => ({ batch, size: pool[batch].length, q: 0, ...rotationTurns(pool[batch], cfg.rotation.history || []) }));
    while (left > 0) {
      const open = info.filter((x) => x.q < x.order.length);
      if (!open.length) break;
      const tot = open.reduce((a, x) => a + x.size, 0);
      let given = 0;
      for (const x of open) { const add = Math.min(x.order.length - x.q, Math.floor((left * x.size) / tot)); x.q += add; given += add; }
      if (!given) for (const x of open.slice().sort((a, b) => b.size - a.size)) { if (given >= left) break; x.q++; given++; }
      left -= given;
    }
    for (const x of info) {
      const on = new Set(x.order.slice(0, x.q));
      for (const r of x.ids) mode[r] = on.has(r) ? 'Online' : 'Offline';
      summary.push({ batch: x.batch, size: x.size, onlineNow: x.q, onlineLast: x.last.size, stillBefore: x.still.length, stillAfter: Math.max(0, x.still.length - x.q) });
    }
  }
  return { mode, summary, seats };
}

// ======================================================================================
// Whole test
// ======================================================================================

/**
 * Plan a whole test.
 * cfg (on top of the allocate config: rooms, labs, buildingOrder, girlsOk, studentMode, batchMode, flags, pins, cbOf, seed):
 *   testType: 'PRT' | 'RPT' | 'SERIES', format: 'MAIN' | 'ADV', days: 1-3 (Test series), dateISO,
 *   slots: [{start, end, noGirls}]      online slots (PRT shifts, RPT slots)
 *   topSlots: [{start, end, noGirls}]   RPT / Test series: slots of the lab(s) kept for top batches (labs with top: true)
 *   topShare: true                      let other students fill empty seats in the top-batch lab
 *   topBatches: ['XA1', ...]
 *   papers: [{id, name, batches: [...], s1, s2, extra: [{start, end}]}]   one paper per phase; offline times (PRT)
 *   rotation: {on, history: [[rolls online], ...]}                        PRT only
 *   antiCheat: 'mix' | 'spread' | 'strict'
 * Returns { units: [{key, kind, tag, day, dateISO, name, time, slots, paperIds, batches, students, result, overlapsWith, emptySeats}],
 *           warnings, blocked, summary }
 */
export function planTest(students, cfg) {
  const type = cfg.testType || 'PRT';
  const adv = cfg.format === 'ADV';
  const seed = cfg.seed || 1;
  const rng = makeRng(seed);
  const warnings = [];
  const papers = cfg.papers || [];
  const paperOf = {};
  papers.forEach((p) => (p.batches || []).forEach((b) => { paperOf[b] = p; }));
  const list = students.filter((s) => paperOf[s.Batch]).map((s) => ({ ...s, Phase: paperOf[s.Batch].name || paperOf[s.Batch].id }));
  const ac = cfg.antiCheat || '';
  const order = (arr) => (ac === 'spread' || ac === 'strict' ? spreadOrder(arr, rng) : shuffle(arr, rng));
  const top = new Set(cfg.topBatches || []);

  if (type === 'PRT') {
    const problems = checkOfflineTiming(papers, adv);
    if (problems.length) return { units: [], warnings: problems, blocked: true, summary: [] };
  }

  const { top: topUnits, general, days, unpaired } = onlineUnits(cfg);
  if (unpaired.length) warnings.push(`JEE Advanced needs two slots that do not overlap. ${unpaired.length === 1 ? 'This slot was' : 'These slots were'} left unused: ${[...new Set(unpaired.map((s) => `${clock(s.start)} to ${clock(s.end)}`))].join(', ')}.`);
  const units = [];
  const unseated = [];
  let summary = [];

  if (type === 'PRT') {
    const m = prtModes(list, cfg);
    summary = m.summary;
    const online = list.filter((s) => m.mode[s.Roll_Number] === 'Online');
    let offline = list.filter((s) => m.mode[s.Roll_Number] !== 'Online');
    const onUnits = general.slice().sort(byTime);
    const queue = [...order(online.filter((s) => top.has(s.Batch))), ...order(online.filter((s) => !top.has(s.Batch)))];
    const over = assignToUnits(queue, onUnits);
    if (over.length) {
      warnings.push(`Online seats are full: ${over.length} students moved to offline.`);
      offline = offline.concat(over);
    }
    units.push(...onUnits.filter((u) => u.students.length));
    // Offline: each paper in its sittings; whole batches are shared between sittings evenly
    for (const p of papers) {
      const mine = offline.filter((s) => paperOf[s.Batch] === p);
      if (!mine.length) continue;
      const sittings = offlineSittings(p, adv).filter((g) => g.every(slotOk));
      const parts = sittings.map((slots) => ({ kind: 'Offline', tag: 'paper', day: 1, slots, paper: p, students: [] }));
      const byBatch = {};
      for (const s of mine) (byBatch[s.Batch] ||= []).push(s);
      for (const b of Object.keys(byBatch).sort(naturalCompare)) {
        const target = parts.reduce((a, x) => (x.students.length < a.students.length ? x : a));
        target.students.push(...byBatch[b]);
      }
      parts.forEach((x, i) => { x.sitting = parts.length > 1 ? i + 1 : 0; });
      units.push(...parts.filter((x) => x.students.length));
    }
  } else {
    // RPT and Test series: everyone online
    const topStudents = list.filter((s) => top.has(s.Batch));
    let rest = list.filter((s) => !top.has(s.Batch));
    const tu = topUnits.slice().sort(byTime);
    if (topStudents.length && !tu.length) {
      warnings.push('No lab is marked “Kept for top batches”, so top batches sit with everyone else.');
      rest = list;
    } else if (topStudents.length) {
      const left = assignToUnits(order(topStudents), tu);
      if (left.length) { warnings.push(`The top-batch lab is full: ${left.length} top batch students sit in the other labs.`); rest = rest.concat(left); }
    }
    tu.forEach((u) => { u.topCount = u.students.length; });
    const pool = [...general, ...(cfg.topShare ? tu : [])].sort(byTime);
    const left = assignToUnits(order(rest), pool);
    if (left.length) {
      const g = left.filter(isGirl).length;
      warnings.push(`Lab seats are full in every slot: ${left.length} students not seated${g ? ` (${g} girls; girls cannot sit in a “No girls” slot)` : ''}. Add a slot${type === 'SERIES' ? ' or a day' : ''}, or switch on more labs.`);
      unseated.push(...left.map((s) => ({ ...s, Mode: 'Online', Reason: isGirl(s) && pool.some((u) => u.noGirls) ? 'No lab seat left in a slot open to girls' : 'No lab seat left in any slot' })));
    }
    units.push(...general.filter((u) => u.students.length), ...(topStudents.length ? tu : []));
  }

  // Seat every sitting, in time order. Sittings that overlap in time (or are less than GAP minutes apart)
  // on the same day share rooms: a later one starts after the seats the earlier one uses.
  units.sort(byTime);
  const done = [];
  const base = { rooms: cfg.rooms, labs: cfg.labs, buildingOrder: cfg.buildingOrder, girlsOk: cfg.girlsOk, flags: cfg.flags, pins: cfg.pins, cbOf: cfg.cbOf, antiCheat: ac };
  units.forEach((u, i) => {
    const taken = {};
    const over = done.filter((d) => d.day === u.day && overlapping(spanOf(d.slots), spanOf(u.slots), GAP));
    for (const d of over) for (const r of d.result.usedRooms) taken[roomKey(r)] = Math.max(taken[roomKey(r)] || 0, r.taken + r.used);
    const result = allocate(u.students, { ...base, mode: u.kind, rooms: u.kind === 'Offline' ? cfg.rooms : [], labs: u.kind === 'Online' ? u.labs : [], taken, seed: seed + i + 1 });
    const name = u.kind === 'Offline'
      ? `${u.paper.name || 'Paper'} · Offline${u.sitting ? ` sitting ${u.sitting}` : ''}`
      : `${u.tag === 'top' ? 'Top batches · ' : ''}Online${days > 1 ? ` · Day ${u.day}` : ''} · ${clock(u.slots[0].start)}`;
    done.push({
      key: `${u.kind}-${u.tag}-${u.day}-${u.slots[0].start}-${u.sitting || 0}-${u.paper ? u.paper.id : ''}`,
      kind: u.kind, tag: u.tag, day: u.day, dateISO: addDays(cfg.dateISO, u.day - 1), name,
      time: u.slots.map((s) => `${clock(s.start)} to ${clock(s.end)}`).join(' & '),
      slots: u.slots.map((s, j) => ({ label: u.slots.length > 1 ? `Paper ${j + 1}` : '', start: s.start, end: s.end })),
      noGirls: !!u.noGirls,
      paperIds: [...new Set(u.students.map((s) => paperOf[s.Batch].id))],
      batches: [...new Set(u.students.map((s) => s.Batch))].sort(naturalCompare),
      students: u.students.length, result, overlapsWith: [],
      overlapKeys: over.map((d) => d.key),
      topCount: u.topCount, emptySeats: u.tag === 'top' ? u.cap - u.students.length : undefined,
    });
  });
  done.forEach((d) => { d.overlapsWith = d.overlapKeys.map((k) => done.find((x) => x.key === k)?.name).filter(Boolean); });
  if (unseated.length) {
    const bucket = done.find((d) => d.kind === 'Online') || done[0];
    if (bucket) bucket.result.unseated.push(...unseated);
    else done.push({ key: 'none', kind: 'Online', tag: 'general', day: 1, dateISO: cfg.dateISO, name: 'Not seated', time: '', slots: [], paperIds: [], batches: [], students: 0, result: { assigned: [], unseated, warnings: [], usedRooms: [] }, overlapsWith: [], overlapKeys: [] });
  }
  return { units: done, warnings, blocked: false, summary };
}
