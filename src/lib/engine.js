// Core seating allocation logic (pure JS, no UI). Used by the browser app and by tests.

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

/**
 * Order in which buildings are tried for students whose class is in `home`.
 * Own building first. Then the "near" buildings in priority order.
 * Far buildings (PC) are used only after every near building is full.
 * Students of a far building overflow to near buildings first.
 */
export function overflowOrder(home, buildingOrder, farBuildings) {
  const far = new Set(farBuildings);
  const near = buildingOrder.filter((b) => !far.has(b) && b !== home);
  const farOthers = buildingOrder.filter((b) => far.has(b) && b !== home);
  return [home, ...near, ...farOthers];
}

/**
 * students: [{Roll_Number, Student_Name, Gender, Course, Batch, Medium}]
 * config: {
 *   rooms: [{building, room, capacity, enabled}],         // offline rooms
 *   labs:  [{building, room, capacity, enabled}],         // online (computer) rooms
 *   batchBuilding: {BATCH: building},                     // home building per batch
 *   batchMode: {BATCH: 'Offline' | 'Online'},
 *   buildingOrder: [...], farBuildings: [...], seed: number
 * }
 */
export function allocate(students, config) {
  const rng = makeRng(config.seed);
  const warnings = [];

  // Seat pools, in fill order. "used" counts seats already given.
  const pool = (list) =>
    list
      .filter((r) => r.enabled && Number(r.capacity) > 0)
      .map((r) => ({ ...r, capacity: Number(r.capacity), used: 0, occupants: [] }));
  const rooms = pool(config.rooms);
  const labs = pool(config.labs);
  const roomsByBuilding = {};
  for (const r of rooms) (roomsByBuilding[r.building] ||= []).push(r);

  const fill = (queue, roomList, mode) => {
    // Fill each room completely before moving on to the next one.
    const left = [];
    let i = 0;
    for (const room of roomList) {
      while (i < queue.length && room.used < room.capacity) {
        room.occupants.push({ ...queue[i], Mode: mode });
        room.used++;
        i++;
      }
      if (i >= queue.length) break;
    }
    while (i < queue.length) left.push(queue[i++]);
    return left;
  };

  const online = students.filter((s) => config.batchMode[s.Batch] === 'Online');
  const offline = students.filter((s) => config.batchMode[s.Batch] !== 'Online');

  // ---- Online: labs in the student's own building first, then near buildings, far ones last
  const labsByBuilding = {};
  for (const l of labs) (labsByBuilding[l.building] ||= []).push(l);
  const labOrder = [...config.buildingOrder, ...Object.keys(labsByBuilding).filter((b) => !config.buildingOrder.includes(b))];
  const onlineByHome = {};
  for (const s of online) (onlineByHome[config.batchBuilding[s.Batch] || '?'] ||= []).push(s);
  const onlineOver = {};
  for (const home of Object.keys(onlineByHome)) {
    const left = fill(shuffle(onlineByHome[home], rng), labsByBuilding[home] || [], 'Online');
    if (left.length) onlineOver[home] = left;
  }
  const onlineLeft = [];
  for (const home of Object.keys(onlineOver)) {
    let queue = onlineOver[home];
    for (const b of overflowOrder(home, labOrder, config.farBuildings).slice(1)) {
      if (!queue.length) break;
      queue = fill(queue, labsByBuilding[b] || [], 'Online');
    }
    onlineLeft.push(...queue);
  }

  // ---- Offline: step 1, each building takes its own students first
  const byHome = {};
  const noBuilding = [];
  for (const s of offline) {
    const b = config.batchBuilding[s.Batch];
    if (!b) noBuilding.push(s);
    else (byHome[b] ||= []).push(s);
  }
  if (noBuilding.length) {
    const bs = [...new Set(noBuilding.map((s) => s.Batch))].join(', ');
    warnings.push(`No building set for batch(es): ${bs} — ${noBuilding.length} students not seated.`);
  }

  const overflow = {}; // home building -> students who did not fit at home
  for (const home of Object.keys(byHome)) {
    const queue = shuffle(byHome[home], rng);
    const ownRooms = roomsByBuilding[home] || [];
    if (!ownRooms.length) warnings.push(`Building ${home} has no rooms switched on; its students move to other buildings.`);
    const left = fill(queue, ownRooms, 'Offline');
    if (left.length) overflow[home] = left;
  }

  // ---- Offline: step 2, students who did not fit go to the next building in order
  const offlineLeft = [];
  const moved = [];
  for (const home of Object.keys(overflow)) {
    let queue = overflow[home];
    const order = overflowOrder(home, config.buildingOrder, config.farBuildings).slice(1);
    for (const b of order) {
      if (!queue.length) break;
      const before = queue.length;
      queue = fill(queue, roomsByBuilding[b] || [], 'Offline');
      if (before > queue.length) moved.push({ from: home, to: b, count: before - queue.length });
    }
    offlineLeft.push(...queue);
  }

  // Seat numbers: inside each room, seats follow roll-number order (as in the example sheet)
  const assigned = [];
  for (const room of [...labs, ...rooms]) {
    room.occupants.sort((a, b) => naturalCompare(a.Roll_Number, b.Roll_Number));
    room.occupants.forEach((s, idx) => {
      assigned.push({ ...s, Building: room.building, 'Room No.': room.room, 'Seat No.': idx + 1, CB: phaseOf(s.Batch) });
    });
  }

  const unseated = [...onlineLeft, ...offlineLeft, ...noBuilding];
  if (onlineLeft.length) warnings.push(`Labs are full: ${onlineLeft.length} online students not seated. Increase lab capacity or switch on more labs.`);
  if (offlineLeft.length) warnings.push(`All rooms are full: ${offlineLeft.length} offline students not seated. Increase capacity or switch on more rooms.`);

  return {
    assigned,
    unseated,
    warnings,
    moved,
    usedRooms: [...labs, ...rooms].filter((r) => r.used > 0),
  };
}
