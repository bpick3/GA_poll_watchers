import { db } from './db.js';
import { randomUUID as uuid } from 'node:crypto';

export function seedIfEmpty() {
  const count = db.prepare('SELECT COUNT(*) c FROM people').get().c;
  if (count > 0) return;

  const insertPerson = db.prepare(
    'INSERT INTO people (id, name, status, isOrganizer, isGuestOfHonor) VALUES (?,?,?,?,?)'
  );

  const roster = [
    ['Brandon', 'confirmed', 1, 1],
    ['Rachel', 'confirmed', 0, 1],
    ['Lance', 'confirmed', 0, 1],
    ['Guest 1', 'confirmed', 0, 0],
    ['Guest 2', 'confirmed', 0, 0],
    ['Guest 3', 'confirmed', 0, 0],
    ['Guest 4', 'confirmed', 0, 0],
    ['Guest 5', 'confirmed', 0, 0],
    ['Guest 6', 'confirmed', 0, 0],
    ['Guest 7', 'confirmed', 0, 0],
    ['Yarkenda', 'maybe', 0, 0],
  ];
  const ids = {};
  for (const [name, status, org, goh] of roster) {
    const id = uuid();
    ids[name] = id;
    insertPerson.run(id, name, status, org, goh);
  }

  // settings
  const setSetting = db.prepare('INSERT INTO settings (key, value) VALUES (?,?)');
  setSetting.run('tripStart', '2026-10-02');
  setSetting.run('tripEnd', '2026-10-05');
  setSetting.run('address', '227 North Fork Road, East Keystone, Colorado');
  setSetting.run('rentalName', 'Summit County Mountain Retreats');
  setSetting.run('rentalLink', 'https://www.scmountainretreats.com/rentals/227-north-fork-road');
  setSetting.run('checkIn', '4:00 PM');
  setSetting.run('checkOut', '10:00 AM');
  setSetting.run('wifiPassword', 'CabinVibes2026');
  setSetting.run('houseRules', 'No pets. No large parties/events beyond our group. Please respect the neighbors.');
  setSetting.run('quietHours', '10:00 PM - 8:00 AM');
  setSetting.run('altitudeTips', "We're at ~9,000 ft. Hydrate constantly, go easy on alcohol the first night, bring sunscreen and chapstick, and take hikes slower than you'd expect.");
  setSetting.run('emergencyInfo', 'Nearest hospital: St. Anthony Summit Hospital, Frisco, CO. Local non-emergency: Summit County Sheriff (970) 453-2232. In a real emergency call 911.');
  setSetting.run('costPerPerson', '365');
  setSetting.run('lodgingCost', '315');
  setSetting.run('foodCost', '50');

  // days
  const days = [
    { id: uuid(), date: '2026-10-02', label: 'Fri 10/2', theme: 'Arrival & Fireside Welcome 🔥', order: 0 },
    { id: uuid(), date: '2026-10-03', label: 'Sat 10/3', theme: 'Birthday Bash 🎂', order: 1 },
    { id: uuid(), date: '2026-10-04', label: 'Sun 10/4', theme: 'Peak Cozy: Flannel & Fall Colors 🍁', order: 2 },
    { id: uuid(), date: '2026-10-05', label: 'Mon 10/5', theme: 'Slow Morning & Send-off ☕', order: 3 },
  ];
  const insertDay = db.prepare('INSERT INTO days (id, date, label, theme, sortOrder) VALUES (?,?,?,?,?)');
  for (const d of days) insertDay.run(d.id, d.date, d.label, d.theme, d.order);
  const [fri, sat, sun, mon] = days;

  setSetting.run('altThemes', JSON.stringify([
    'Y2K Night 💿',
    'Cabin Casino 🎰',
    'Murder Mystery Dinner 🔪',
    'White Lies Party 🤥',
    'Decades Night (choose an era) 🕺',
  ]));

  // blocks
  const insertBlock = db.prepare(
    'INSERT INTO blocks (id, dayId, time, title, type, location, ownerId, ownerText, notes, sortOrder) VALUES (?,?,?,?,?,?,?,?,?,?)'
  );
  const insertRsvp = db.prepare('INSERT INTO block_rsvps (blockId, personId, status) VALUES (?,?,?)');

  function addBlock({ dayId, time, title, type = 'custom', location = '', ownerId = null, ownerText = '', notes = '', order = 0, rsvps = [] }) {
    const id = uuid();
    insertBlock.run(id, dayId, time, title, type, location, ownerId, ownerText, notes, order);
    for (const [name, status] of rsvps) insertRsvp.run(id, ids[name], status);
    return id;
  }

  const trailNotes = 'Suggested trails — verify seasonal conditions & closures before hiking. Early-Oct afternoons at 9000ft can turn cold/windy fast; bring layers.';
  const trails = [
    'Sapphire Point Overlook — Easy, 15 min drive, big Dillon Reservoir views',
    'Lily Pad Lake Trail (Frisco) — Easy/Moderate, 20 min drive, 2.2mi loop',
    'Peaks Trail (Frisco to Breck) — Moderate, 20 min drive, out-and-back as far as you want',
    'Sylvan Lake Trail (Eagles Nest Wilderness) — Moderate, 30 min drive, alpine lake payoff',
    'Willowbrook / Keystone Gulch — Moderate, 10 min drive, close to the house',
  ].join(' | ');

  // FRI
  addBlock({ dayId: fri.id, time: '3:00 PM', title: 'Arrival & Check-in', type: 'custom', location: '227 North Fork Rd', ownerText: 'Everyone', order: 0,
    rsvps: [['Brandon','in'],['Rachel','in'],['Lance','in'],['Guest 1','in'],['Guest 2','in']] });
  addBlock({ dayId: fri.id, time: '5:30 PM', title: 'River Run Village Walk', type: 'custom', location: 'River Run Village', ownerId: ids['Guest 3'], order: 1,
    rsvps: [['Guest 3','in'],['Guest 4','in']] });
  addBlock({ dayId: fri.id, time: '7:00 PM', title: 'Arrival Taco Bar 🌮 (dinner)', type: 'custom', location: 'Kitchen', ownerText: 'TBD', order: 2,
    rsvps: [['Brandon','in'],['Rachel','in']] });
  addBlock({ dayId: fri.id, time: '9:00 PM', title: 'Movie + Fireplace Night', type: 'movie', location: 'Great Room', ownerText: 'Whoever wins the vote', notes: 'Nominate & vote below — locks at 8pm.', order: 3 });
  addBlock({ dayId: fri.id, time: '9:00 PM', title: "S'mores at the Fire Pit", type: 'custom', location: 'Covered Patio', ownerId: ids['Guest 5'], order: 4,
    rsvps: [['Guest 5','in'],['Guest 6','in'],['Yarkenda','in']] });

  // SAT
  addBlock({ dayId: sat.id, time: '8:30 AM', title: 'Yoga on the Patio', type: 'yoga', location: 'Covered Patio', ownerId: ids['Rachel'], notes: 'Optional, all levels.', order: 0,
    rsvps: [['Rachel','in'],['Guest 1','in']] });
  addBlock({ dayId: sat.id, time: '10:00 AM', title: 'Grocery Run (birthday dinner supplies)', type: 'custom', location: 'City Market, Silverthorne', ownerId: ids['Guest 2'], order: 1,
    rsvps: [['Guest 2','in']] });
  addBlock({ dayId: sat.id, time: '11:30 AM', title: 'Afternoon Hike', type: 'hike', location: 'TBD trail', ownerId: ids['Lance'], notes: trailNotes + ' | Options: ' + trails, order: 2,
    rsvps: [['Lance','in'],['Guest 3','in'],['Guest 4','in']] });
  addBlock({ dayId: sat.id, time: '3:00 PM', title: 'Ping Pong / Foosball Tournament', type: 'custom', location: 'Basement Game Room', ownerId: ids['Guest 5'], order: 3,
    rsvps: [['Guest 5','in'],['Guest 6','in'],['Guest 7','in']] });
  addBlock({ dayId: sat.id, time: '6:30 PM', title: 'Birthday Toast 🥂', type: 'custom', location: 'Great Room', ownerText: 'Everyone', order: 4,
    rsvps: [['Brandon','in'],['Rachel','in'],['Lance','in']] });
  addBlock({ dayId: sat.id, time: '7:00 PM', title: 'Big Birthday Dinner (dinner) + Dessert 🎂', type: 'custom', location: 'Kitchen / Dining', ownerText: 'TBD', order: 5,
    rsvps: [['Brandon','in'],['Rachel','in'],['Lance','in'],['Guest 1','in']] });
  addBlock({ dayId: sat.id, time: '9:30 PM', title: 'Movie + Fireplace Night', type: 'movie', location: 'Great Room', notes: 'Nominate & vote below — locks at 8pm.', order: 6 });

  // SUN
  addBlock({ dayId: sun.id, time: '7:15 AM', title: 'Sunrise Photos', type: 'custom', location: 'Private Balcony / Ridge', ownerId: ids['Guest 6'], order: 0,
    rsvps: [['Guest 6','in']] });
  addBlock({ dayId: sun.id, time: '9:00 AM', title: 'Yoga on the Patio', type: 'yoga', location: 'Covered Patio', ownerId: ids['Rachel'], notes: 'Optional, all levels.', order: 1,
    rsvps: [['Rachel','in']] });
  addBlock({ dayId: sun.id, time: '12:00 PM', title: 'Afternoon Hike', type: 'hike', location: 'TBD trail', ownerId: ids['Guest 7'], notes: trailNotes + ' | Options: ' + trails, order: 2,
    rsvps: [['Guest 7','in'],['Lance','in']] });
  addBlock({ dayId: sun.id, time: '4:00 PM', title: 'Game Tournament Finals', type: 'custom', location: 'Basement Game Room', ownerId: ids['Guest 5'], order: 3,
    rsvps: [['Guest 5','in'],['Guest 6','in']] });
  addBlock({ dayId: sun.id, time: '7:00 PM', title: 'Chili + Cornbread Fireside (dinner)', type: 'custom', location: 'Great Room', ownerText: 'TBD', order: 4,
    rsvps: [['Brandon','in'],['Guest 2','in']] });
  addBlock({ dayId: sun.id, time: '9:00 PM', title: 'Movie + Fireplace Night', type: 'movie', location: 'Great Room', notes: 'Nominate & vote below — locks at 8pm.', order: 5 });

  // MON
  addBlock({ dayId: mon.id, time: '8:00 AM', title: 'Leftovers + Pancakes (breakfast)', type: 'custom', location: 'Kitchen', ownerId: ids['Guest 1'], order: 0,
    rsvps: [['Guest 1','in'],['Brandon','in'],['Rachel','in']] });
  addBlock({ dayId: mon.id, time: '9:30 AM', title: 'Checkout Cleanup Sweep', type: 'custom', location: 'Whole House', ownerText: 'Everyone', order: 1,
    rsvps: [['Brandon','in'],['Rachel','in'],['Lance','in']] });
  addBlock({ dayId: mon.id, time: '10:00 AM', title: 'Send-off & Departure', type: 'custom', location: '227 North Fork Rd', ownerText: 'Everyone', order: 2 });

  // movie nominations (a couple pre-seeded on Fri and Sat)
  const insertNom = db.prepare('INSERT INTO movie_nominations (id, dayId, title, nominatedBy) VALUES (?,?,?,?)');
  const insertVote = db.prepare('INSERT INTO movie_votes (nominationId, personId) VALUES (?,?)');
  const nom1 = uuid(); insertNom.run(nom1, fri.id, 'Hocus Pocus', ids['Rachel']);
  insertVote.run(nom1, ids['Rachel']); insertVote.run(nom1, ids['Guest 1']);
  const nom2 = uuid(); insertNom.run(nom2, fri.id, 'The Shining', ids['Lance']);
  insertVote.run(nom2, ids['Lance']);
  const nom3 = uuid(); insertNom.run(nom3, sat.id, "Practical Magic", ids['Brandon']);
  insertVote.run(nom3, ids['Brandon']); insertVote.run(nom3, ids['Guest 3']); insertVote.run(nom3, ids['Guest 4']);

  // hot tub slots (evening onward, 45 min, per day)
  const insertSlot = db.prepare('INSERT INTO hottub_slots (id, dayId, startTime, label) VALUES (?,?,?,?)');
  const insertSignup = db.prepare('INSERT INTO hottub_signups (slotId, personId) VALUES (?,?)');
  const slotTimes = ['6:00 PM', '6:45 PM', '7:30 PM', '8:15 PM', '9:00 PM', '9:45 PM'];
  for (const day of [fri, sat, sun]) {
    for (const t of slotTimes) {
      const sid = uuid();
      insertSlot.run(sid, day.id, t, `${t} - 45 min`);
    }
  }
  // a couple pre-filled signups Saturday evening
  const satSlots = db.prepare('SELECT id FROM hottub_slots WHERE dayId=? ORDER BY rowid').all(sat.id);
  insertSignup.run(satSlots[0].id, ids['Brandon']);
  insertSignup.run(satSlots[0].id, ids['Rachel']);
  insertSignup.run(satSlots[2].id, ids['Guest 3']);
  insertSignup.run(satSlots[2].id, ids['Guest 4']);
  insertSignup.run(satSlots[2].id, ids['Guest 5']);

  // meals
  const insertMeal = db.prepare('INSERT INTO meals (id, dayId, mealType, plan, cooks, cleanup) VALUES (?,?,?,?,?,?)');
  function meal(dayId, mealType, plan, cooks = [], cleanup = []) {
    insertMeal.run(uuid(), dayId, mealType, plan, cooks.map(n => ids[n]).join(','), cleanup.map(n => ids[n]).join(','));
  }
  meal(fri.id, 'Breakfast', '', [], []); // travel day, no seed
  meal(fri.id, 'Lunch', 'On the road / grab something in town', [], []);
  meal(fri.id, 'Dinner', 'Arrival Taco Bar 🌮', [], []); // owner TBD - unclaimed on purpose
  meal(fri.id, 'Snacks', 'Chips, salsa, queso out on the counter', ['Guest 2'], []);

  meal(sat.id, 'Breakfast', 'Simple sandwich / bagel bar', ['Guest 1'], ['Guest 2']);
  meal(sat.id, 'Lunch', 'Sandwich bar + hike snacks (trail mix, fruit, jerky)', ['Guest 3'], []);
  meal(sat.id, 'Dinner', 'Big Birthday Dinner 🎉', [], []); // unclaimed on purpose
  meal(sat.id, 'Snacks', 'Birthday Dessert 🎂 — cake + candles', ['Rachel'], ['Lance']);

  meal(sun.id, 'Breakfast', 'Sandwich bar / bagel bar leftovers', [], []);
  meal(sun.id, 'Lunch', 'Hike snacks — trail mix, sandwiches to-go', ['Guest 4'], []);
  meal(sun.id, 'Dinner', 'Chili + Cornbread Fireside 🌶️', ['Brandon'], ['Guest 5', 'Guest 6']);
  meal(sun.id, 'Snacks', 'Popcorn + hot cocoa bar for movie night', ['Guest 7'], []);

  meal(mon.id, 'Breakfast', 'Leftovers + Pancakes before checkout', ['Guest 1'], ['Guest 2', 'Guest 3']);
  meal(mon.id, 'Lunch', '', [], []);
  meal(mon.id, 'Dinner', '', [], []);
  meal(mon.id, 'Snacks', '', [], []);

  // dietary notes
  const insertDiet = db.prepare('INSERT INTO dietary_notes (personId, note) VALUES (?,?)');
  insertDiet.run(ids['Rachel'], 'Vegetarian');
  insertDiet.run(ids['Guest 2'], 'Gluten-free');
  insertDiet.run(ids['Guest 5'], 'Nut allergy — please label anything with nuts');
  insertDiet.run(ids['Yarkenda'], 'No pork');

  // grocery list
  const insertGrocery = db.prepare('INSERT INTO grocery_items (id, category, item, qty, claimedBy, checked) VALUES (?,?,?,?,?,?)');
  const groceries = [
    ['Produce', 'Lettuce', '2 heads', 'Guest 2', 0],
    ['Produce', 'Tomatoes', '6', 'Guest 2', 0],
    ['Produce', 'Onions', '4', null, 0],
    ['Produce', 'Limes', '10', null, 0],
    ['Protein', 'Ground beef', '3 lbs', 'Brandon', 0],
    ['Protein', 'Chicken thighs', '3 lbs', null, 0],
    ['Protein', 'Black beans (chili)', '4 cans', 'Brandon', 0],
    ['Pantry', 'Taco shells + tortillas', '2 packs', 'Guest 2', 0],
    ['Pantry', 'Cornbread mix', '2 boxes', 'Brandon', 0],
    ['Pantry', 'Chili seasoning', '2 packets', null, 0],
    ['Pantry', 'Pancake mix', '1 large box', 'Guest 1', 0],
    ['Pantry', 'Birthday candles', '1 pack', 'Rachel', 1],
    ['Drinks', 'Coffee (whole bean)', '2 bags', null, 0],
    ['Drinks', 'Hot cocoa mix', '1 tub', 'Guest 7', 0],
    ['Drinks', 'Sparkling water', '2 cases', null, 0],
    ['Snacks', 'Trail mix', '3 bags', 'Guest 3', 0],
    ['Snacks', 'Popcorn', '1 box', 'Guest 7', 0],
    ['Paper Goods', 'Paper towels', '4 rolls', null, 0],
    ['Paper Goods', 'Plates + napkins', '1 set', null, 0],
  ];
  for (const g of groceries) insertGrocery.run(uuid(), ...g);

  // drinks & snacks board
  const insertDrink = db.prepare('INSERT INTO drinks_snacks (id, item, claimedBy) VALUES (?,?,?)');
  const drinks = [
    ['Local IPA 6-packs', 'Guest 4'],
    ['Bottle of tequila', 'Lance'],
    ['Sparkling cider (non-alc)', 'Rachel'],
    ['Chips & guac', null],
    ['Charcuterie board fixings', null],
    ['White claws', 'Guest 6'],
  ];
  for (const d of drinks) insertDrink.run(uuid(), ...d);

  // payments
  const insertPayment = db.prepare('INSERT INTO payments (id, personId, dueLabel, dueDate, amount, status) VALUES (?,?,?,?,?,?)');
  const confirmedPeople = roster.filter(r => r[1] === 'confirmed').map(r => r[0]);
  const statuses = ['sent', 'confirmed', 'sent', 'not_sent'];
  let si = 0;
  for (const name of confirmedPeople) {
    insertPayment.run(uuid(), ids[name], 'Payment 1', '2026-07-31', 182.5, name === 'Brandon' ? 'confirmed' : (si++ % 3 === 0 ? 'sent' : 'not_sent'));
    insertPayment.run(uuid(), ids[name], 'Payment 2', '2026-09-18', 182.5, 'not_sent');
  }

  // expenses
  const insertExpense = db.prepare('INSERT INTO expenses (id, description, amount, paidBy, date) VALUES (?,?,?,?,?)');
  insertExpense.run(uuid(), 'Costco run — drinks & paper goods', 96.40, ids['Brandon'], '2026-09-20');
  insertExpense.run(uuid(), 'Propane for fire pit', 28.00, ids['Guest 3'], '2026-09-22');
  insertExpense.run(uuid(), 'Firewood bundle', 40.00, ids['Lance'], '2026-09-25');

  // committees
  const insertCommittee = db.prepare('INSERT INTO committees (id, name, emoji) VALUES (?,?,?)');
  const committees = [
    { name: 'Food Planning & Cooking', emoji: '🍽️' },
    { name: 'Supplies & Shopping', emoji: '🛒' },
    { name: 'Games & Activities', emoji: '🎲' },
    { name: 'Vibes & Extras', emoji: '✨' },
  ];
  const committeeIds = {};
  for (const c of committees) {
    const id = uuid();
    committeeIds[c.name] = id;
    insertCommittee.run(id, c.name, c.emoji);
  }
  const insertMember = db.prepare('INSERT INTO committee_members (committeeId, personId) VALUES (?,?)');
  insertMember.run(committeeIds['Food Planning & Cooking'], ids['Brandon']);
  insertMember.run(committeeIds['Food Planning & Cooking'], ids['Guest 1']);
  insertMember.run(committeeIds['Supplies & Shopping'], ids['Guest 2']);
  insertMember.run(committeeIds['Supplies & Shopping'], ids['Guest 3']);
  insertMember.run(committeeIds['Games & Activities'], ids['Guest 5']);
  insertMember.run(committeeIds['Games & Activities'], ids['Lance']);
  insertMember.run(committeeIds['Vibes & Extras'], ids['Rachel']);
  insertMember.run(committeeIds['Vibes & Extras'], ids['Guest 6']);

  const insertCTask = db.prepare('INSERT INTO committee_tasks (id, committeeId, task, assigneeId, dueDate, done) VALUES (?,?,?,?,?,?)');
  insertCTask.run(uuid(), committeeIds['Food Planning & Cooking'], 'Finalize birthday dinner menu', ids['Brandon'], '2026-09-15', 0);
  insertCTask.run(uuid(), committeeIds['Food Planning & Cooking'], 'Order birthday cake', ids['Guest 1'], '2026-09-25', 0);
  insertCTask.run(uuid(), committeeIds['Supplies & Shopping'], 'Buy paper goods & ice', ids['Guest 2'], '2026-10-01', 0);
  insertCTask.run(uuid(), committeeIds['Supplies & Shopping'], 'Confirm grill propane level', ids['Guest 3'], '2026-10-01', 1);
  insertCTask.run(uuid(), committeeIds['Games & Activities'], 'Print tournament bracket sheet', ids['Guest 5'], '2026-09-28', 0);
  insertCTask.run(uuid(), committeeIds['Games & Activities'], 'Pack card games box', ids['Lance'], '2026-09-30', 0);
  insertCTask.run(uuid(), committeeIds['Vibes & Extras'], 'Make Spotify/Sonos playlist', ids['Rachel'], '2026-09-20', 1);
  insertCTask.run(uuid(), committeeIds['Vibes & Extras'], 'Bring string lights for patio', ids['Guest 6'], '2026-09-30', 0);

  const insertCNote = db.prepare('INSERT INTO committee_notes (id, committeeId, personId, text, ts) VALUES (?,?,?,?,?)');
  insertCNote.run(uuid(), committeeIds['Food Planning & Cooking'], ids['Brandon'], 'Thinking smoked brisket for the birthday dinner — thoughts?', '2026-07-01T14:00:00Z');
  insertCNote.run(uuid(), committeeIds['Food Planning & Cooking'], ids['Guest 1'], 'Brisket sounds great, I can help smoke it Saturday morning.', '2026-07-02T09:30:00Z');
  insertCNote.run(uuid(), committeeIds['Games & Activities'], ids['Guest 5'], "Bringing my own ping pong paddles, don't trust the house ones.", '2026-07-05T20:15:00Z');
  insertCNote.run(uuid(), committeeIds['Vibes & Extras'], ids['Rachel'], 'Grabbing some fall garland for the mantel!', '2026-07-10T11:00:00Z');

  // games bring list
  const insertGame = db.prepare('INSERT INTO games_bring (id, name, claimedBy) VALUES (?,?,?)');
  const games = [
    ['Codenames', 'Guest 4'],
    ['Uno', 'Guest 5'],
    ['Spades', null],
    ['Dominoes', 'Lance'],
    ['Heads Up', null],
    ['Werewolf', 'Guest 6'],
    ['Sonos Karaoke Night', null],
  ];
  for (const g of games) insertGame.run(uuid(), ...g);

  // tournament (single elim, 8 players pre-generated bracket, unplayed)
  const players = ['Brandon', 'Rachel', 'Lance', 'Guest 1', 'Guest 2', 'Guest 3', 'Guest 4', 'Guest 5'];
  const round1 = [];
  for (let i = 0; i < players.length; i += 2) {
    round1.push({ id: uuid(), p1: players[i], p2: players[i + 1], winner: null });
  }
  const bracket = { rounds: [round1] };
  db.prepare('INSERT INTO tournaments (id, name, playersJson, bracketJson, championId) VALUES (?,?,?,?,?)')
    .run(uuid(), 'Ping Pong Championship 🏓', JSON.stringify(players), JSON.stringify(bracket), null);

  // birthday checklist
  const insertBday = db.prepare('INSERT INTO birthday_checklist (id, forPerson, item, done) VALUES (?,?,?,?)');
  for (const person of ['Brandon', 'Rachel', 'Lance']) {
    insertBday.run(uuid(), person, `Cake/dessert plan for ${person}`, 0);
    insertBday.run(uuid(), person, `Candles for ${person}`, 0);
    insertBday.run(uuid(), person, `Group toast for ${person}`, 0);
    insertBday.run(uuid(), person, `"Hot seat" appreciation round for ${person}`, 0);
    insertBday.run(uuid(), person, `Photo moment for ${person}`, 0);
  }

  const insertSurprise = db.prepare('INSERT INTO surprise_ideas (id, personId, text, ts) VALUES (?,?,?,?)');
  insertSurprise.run(uuid(), ids['Guest 2'], 'Sneak in a photo slideshow of old birthday memories to play during dessert!', '2026-07-08T18:00:00Z');
  insertSurprise.run(uuid(), ids['Guest 6'], 'Everyone write one appreciation note, we read them aloud during the hot seat round.', '2026-07-09T12:00:00Z');

  // rooms
  const insertRoom = db.prepare('INSERT INTO rooms (id, name, bed, capacity, details) VALUES (?,?,?,?,?)');
  const rooms = [
    { id: uuid(), name: 'Top-floor Primary', bed: 'King', capacity: 2, details: 'Private balcony, gas fireplace, en suite jetted tub + steam shower' },
    { id: uuid(), name: 'Main-floor Guest 1', bed: 'Queen', capacity: 2, details: 'Private en suite' },
    { id: uuid(), name: 'Main-floor Guest 2', bed: 'Queen', capacity: 2, details: 'Near hall bath' },
    { id: uuid(), name: 'Basement Bunk Room', bed: 'Two bunks + one twin', capacity: 5, details: 'Private en suite' },
  ];
  for (const r of rooms) insertRoom.run(r.id, r.name, r.bed, r.capacity, r.details);

  const insertRoomAssign = db.prepare('INSERT INTO room_assignments (roomId, personId, note) VALUES (?,?,?)');
  insertRoomAssign.run(rooms[0].id, ids['Brandon'], 'Birthday suite!');
  insertRoomAssign.run(rooms[0].id, ids['Rachel'], '');
  insertRoomAssign.run(rooms[1].id, ids['Lance'], '');
  insertRoomAssign.run(rooms[2].id, ids['Guest 1'], '');
  insertRoomAssign.run(rooms[2].id, ids['Guest 2'], 'Light sleeper, prefers window cracked');
  insertRoomAssign.run(rooms[3].id, ids['Guest 3'], '');
  insertRoomAssign.run(rooms[3].id, ids['Guest 4'], '');
  insertRoomAssign.run(rooms[3].id, ids['Guest 5'], '');

  // carpools
  const insertCarpool = db.prepare('INSERT INTO carpools (id, driverId, seats, departure, eta) VALUES (?,?,?,?,?)');
  const cp1 = uuid(); insertCarpool.run(cp1, ids['Brandon'], 3, 'Denver, 12:00 PM Fri', '3:00 PM Fri');
  const cp2 = uuid(); insertCarpool.run(cp2, ids['Guest 3'], 2, 'Boulder, 1:00 PM Fri', '3:45 PM Fri');
  const insertRider = db.prepare('INSERT INTO carpool_riders (carpoolId, personId) VALUES (?,?)');
  insertRider.run(cp1, ids['Rachel']);
  insertRider.run(cp2, ids['Guest 4']);

  // packing list
  const insertPacking = db.prepare('INSERT INTO packing_items (id, item, isShared) VALUES (?,?,?)');
  const packItems = ['Warm jacket', 'Hat & gloves', 'Swimsuit (hot tub!)', 'Hiking shoes', 'Sunscreen', 'Water bottle', 'Layers (it gets cold at altitude)', 'Flip flops for hot tub', 'Cozy socks', 'Phone charger', 'Camera', 'Chapstick'];
  for (const item of packItems) insertPacking.run(uuid(), item, 1);

  // cleanup tasks
  const insertCleanup = db.prepare('INSERT INTO cleanup_tasks (id, task, claimedBy, done) VALUES (?,?,?,?)');
  insertCleanup.run(uuid(), 'Take out all trash & recycling', 'Guest 4', 0);
  insertCleanup.run(uuid(), 'Load & run dishwasher', 'Guest 1', 0);
  insertCleanup.run(uuid(), 'Strip all beds', null, 0);
  insertCleanup.run(uuid(), 'Cover hot tub', 'Lance', 0);
  insertCleanup.run(uuid(), 'Final walkthrough sweep', 'Brandon', 0);
  insertCleanup.run(uuid(), 'Wipe down kitchen counters', null, 0);

  console.log('Database seeded.');
}
