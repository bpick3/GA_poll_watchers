import express from 'express';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID as uuid } from 'node:crypto';
import { db } from './db.js';
import { seedIfEmpty } from './seed.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
seedIfEmpty();

const app = express();
app.use(express.json());

// identity middleware — reads person id from header, attaches person row
app.use((req, res, next) => {
  const pid = req.header('x-person-id');
  if (pid) {
    req.person = db.prepare('SELECT * FROM people WHERE id=?').get(pid) || null;
  }
  next();
});

function requireOrganizer(req, res, next) {
  if (!req.person || !req.person.isOrganizer) return res.status(403).json({ error: 'Organizer only' });
  next();
}

const api = express.Router();

// ---------- people / settings ----------
api.get('/people', (req, res) => {
  res.json(db.prepare('SELECT * FROM people ORDER BY isGuestOfHonor DESC, name ASC').all());
});
api.post('/people', (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'name required' });
  const id = uuid();
  db.prepare('INSERT INTO people (id, name, status, isOrganizer, isGuestOfHonor) VALUES (?,?,?,0,0)').run(id, name.trim(), 'confirmed');
  res.json(db.prepare('SELECT * FROM people WHERE id=?').get(id));
});
api.patch('/people/:id', (req, res) => {
  const p = db.prepare('SELECT * FROM people WHERE id=?').get(req.params.id);
  if (!p) return res.status(404).json({ error: 'not found' });
  const name = req.body.name ?? p.name;
  const status = req.body.status ?? p.status;
  const isOrganizer = req.body.isOrganizer !== undefined ? (req.body.isOrganizer ? 1 : 0) : p.isOrganizer;
  const isGuestOfHonor = req.body.isGuestOfHonor !== undefined ? (req.body.isGuestOfHonor ? 1 : 0) : p.isGuestOfHonor;
  db.prepare('UPDATE people SET name=?, status=?, isOrganizer=?, isGuestOfHonor=? WHERE id=?').run(name, status, isOrganizer, isGuestOfHonor, req.params.id);
  res.json(db.prepare('SELECT * FROM people WHERE id=?').get(req.params.id));
});
api.delete('/people/:id', (req, res) => {
  db.prepare('DELETE FROM people WHERE id=?').run(req.params.id);
  res.json({ ok: true });
});

api.get('/settings', (req, res) => {
  const rows = db.prepare('SELECT * FROM settings').all();
  const obj = {};
  for (const r of rows) obj[r.key] = r.value;
  res.json(obj);
});
api.patch('/settings', (req, res) => {
  const stmt = db.prepare('INSERT INTO settings (key, value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value');
  for (const [k, v] of Object.entries(req.body)) stmt.run(k, String(v));
  const rows = db.prepare('SELECT * FROM settings').all();
  const obj = {};
  for (const r of rows) obj[r.key] = r.value;
  res.json(obj);
});

// ---------- setup wizard ----------
function dateRange(start, end) {
  const out = [];
  const d = new Date(start + 'T00:00:00');
  const last = new Date(end + 'T00:00:00');
  while (d <= last) {
    out.push(d.toISOString().slice(0, 10));
    d.setDate(d.getDate() + 1);
  }
  return out;
}

api.post('/setup', (req, res) => {
  const b = req.body || {};
  if (!b.tripName || !b.tripStart || !b.tripEnd) {
    return res.status(400).json({ error: 'tripName, tripStart, and tripEnd are required' });
  }

  const run = db.transaction(() => {
    const setSetting = db.prepare('INSERT INTO settings (key, value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value');
    const settingFields = [
      'tripName', 'tripStart', 'tripEnd', 'address', 'rentalName', 'rentalLink',
      'costPerPerson', 'lodgingCost', 'foodCost', 'checkIn', 'checkOut', 'wifiPassword',
      'houseRules', 'quietHours', 'altitudeTips', 'emergencyInfo', 'houseDescription',
    ];
    for (const key of settingFields) {
      if (b[key] !== undefined) setSetting.run(key, String(b[key]));
    }

    // roster
    const insertPerson = db.prepare('INSERT INTO people (id, name, status, isOrganizer, isGuestOfHonor) VALUES (?,?,?,?,?)');
    const personIds = [];
    for (const p of (b.roster || [])) {
      if (!p.name || !p.name.trim()) continue;
      const id = uuid();
      insertPerson.run(id, p.name.trim(), p.status === 'maybe' ? 'maybe' : 'confirmed', p.isOrganizer ? 1 : 0, p.isGuestOfHonor ? 1 : 0);
      personIds.push({ id, ...p });
    }

    // days + blank meal rows, spanning tripStart..tripEnd
    const insertDay = db.prepare('INSERT INTO days (id, date, label, theme, sortOrder) VALUES (?,?,?,?,?)');
    const insertMeal = db.prepare('INSERT INTO meals (id, dayId, mealType, plan, cooks, cleanup) VALUES (?,?,?,?,\'\',\'\')');
    const dates = dateRange(b.tripStart, b.tripEnd);
    dates.forEach((date, i) => {
      const dayId = uuid();
      const weekday = new Date(date + 'T00:00:00').toLocaleDateString('en-US', { weekday: 'short' });
      const label = `${weekday} ${date.slice(5).replace('-', '/')}`;
      insertDay.run(dayId, date, label, 'Day plans TBD 🍂', i);
      for (const mt of ['Breakfast', 'Lunch', 'Dinner', 'Snacks']) {
        insertMeal.run(uuid(), dayId, mt, '');
      }
    });

    // rooms
    const insertRoom = db.prepare('INSERT INTO rooms (id, name, bed, capacity, details) VALUES (?,?,?,?,?)');
    for (const r of (b.rooms || [])) {
      if (!r.name || !r.name.trim()) continue;
      insertRoom.run(uuid(), r.name.trim(), r.bed || '', Number(r.capacity) || 1, r.details || '');
    }

    // payment schedule — two installments applied to every confirmed/maybe person
    const insertPayment = db.prepare('INSERT INTO payments (id, personId, dueLabel, dueDate, amount, status) VALUES (?,?,?,?,?,\'not_sent\')');
    const installments = [
      { label: b.payment1Label, date: b.payment1Date, amount: b.payment1Amount },
      { label: b.payment2Label, date: b.payment2Date, amount: b.payment2Amount },
    ].filter(p => p.label && p.date && p.amount);
    for (const person of personIds) {
      for (const inst of installments) {
        insertPayment.run(uuid(), person.id, inst.label, inst.date, Number(inst.amount));
      }
    }

    // default birthday checklist for anyone flagged as a guest of honor
    const insertBday = db.prepare('INSERT INTO birthday_checklist (id, forPerson, item, done) VALUES (?,?,?,0)');
    for (const person of personIds.filter(p => p.isGuestOfHonor)) {
      for (const item of ['Cake/dessert plan', 'Candles', 'Group toast', '"Hot seat" appreciation round', 'Photo moment']) {
        insertBday.run(uuid(), person.name.trim(), `${item} for ${person.name.trim()}`);
      }
    }

    setSetting.run('setupComplete', '1');
  });

  try {
    run();
  } catch (e) {
    return res.status(400).json({ error: e.message });
  }
  res.json({ ok: true });
});

// ---------- schedule ----------
api.get('/days', (req, res) => {
  const days = db.prepare('SELECT * FROM days ORDER BY sortOrder').all();
  const blocks = db.prepare('SELECT * FROM blocks ORDER BY sortOrder').all();
  const rsvps = db.prepare('SELECT * FROM block_rsvps').all();
  const noms = db.prepare('SELECT * FROM movie_nominations').all();
  const votes = db.prepare('SELECT * FROM movie_votes').all();
  const slots = db.prepare('SELECT * FROM hottub_slots ORDER BY rowid').all();
  const signups = db.prepare('SELECT * FROM hottub_signups').all();

  const out = days.map(d => ({
    ...d,
    blocks: blocks.filter(b => b.dayId === d.id).map(b => ({
      ...b,
      rsvps: rsvps.filter(r => r.blockId === b.id),
    })),
    movieNominations: noms.filter(n => n.dayId === d.id).map(n => ({
      ...n,
      votes: votes.filter(v => v.nominationId === n.id).map(v => v.personId),
    })),
    hottubSlots: slots.filter(s => s.dayId === d.id).map(s => ({
      ...s,
      signups: signups.filter(su => su.slotId === s.id).map(su => su.personId),
    })),
  }));
  res.json(out);
});

api.patch('/days/:id', (req, res) => {
  const d = db.prepare('SELECT * FROM days WHERE id=?').get(req.params.id);
  if (!d) return res.status(404).json({ error: 'not found' });
  const theme = req.body.theme ?? d.theme;
  db.prepare('UPDATE days SET theme=? WHERE id=?').run(theme, req.params.id);
  res.json(db.prepare('SELECT * FROM days WHERE id=?').get(req.params.id));
});

api.post('/blocks', (req, res) => {
  const { dayId, time, title, type = 'custom', location = '', ownerId = null, ownerText = '', notes = '' } = req.body;
  if (!dayId || !title) return res.status(400).json({ error: 'dayId and title required' });
  const id = uuid();
  const maxOrder = db.prepare('SELECT COALESCE(MAX(sortOrder),0) m FROM blocks WHERE dayId=?').get(dayId).m;
  db.prepare('INSERT INTO blocks (id, dayId, time, title, type, location, ownerId, ownerText, notes, sortOrder) VALUES (?,?,?,?,?,?,?,?,?,?)')
    .run(id, dayId, time, title, type, location, ownerId, ownerText, notes, maxOrder + 1);
  res.json(db.prepare('SELECT * FROM blocks WHERE id=?').get(id));
});
api.patch('/blocks/:id', (req, res) => {
  const b = db.prepare('SELECT * FROM blocks WHERE id=?').get(req.params.id);
  if (!b) return res.status(404).json({ error: 'not found' });
  const merged = { ...b, ...req.body };
  db.prepare('UPDATE blocks SET time=?, title=?, location=?, ownerId=?, ownerText=?, notes=? WHERE id=?')
    .run(merged.time, merged.title, merged.location, merged.ownerId, merged.ownerText, merged.notes, req.params.id);
  res.json(db.prepare('SELECT * FROM blocks WHERE id=?').get(req.params.id));
});
api.delete('/blocks/:id', (req, res) => {
  db.prepare('DELETE FROM blocks WHERE id=?').run(req.params.id);
  db.prepare('DELETE FROM block_rsvps WHERE blockId=?').run(req.params.id);
  res.json({ ok: true });
});
api.post('/blocks/:id/rsvp', (req, res) => {
  const personId = req.body.personId || req.person?.id;
  const status = req.body.status;
  if (!personId || !['in', 'out'].includes(status)) return res.status(400).json({ error: 'bad request' });
  db.prepare('INSERT INTO block_rsvps (blockId, personId, status) VALUES (?,?,?) ON CONFLICT(blockId, personId) DO UPDATE SET status=excluded.status')
    .run(req.params.id, personId, status);
  res.json({ ok: true });
});

api.post('/movie-nominations', (req, res) => {
  const { dayId, title, nominatedBy } = req.body;
  const id = uuid();
  db.prepare('INSERT INTO movie_nominations (id, dayId, title, nominatedBy) VALUES (?,?,?,?)').run(id, dayId, title, nominatedBy || null);
  res.json({ id });
});
api.post('/movie-nominations/:id/vote', (req, res) => {
  const personId = req.body.personId || req.person?.id;
  if (!personId) return res.status(400).json({ error: 'personId required' });
  const nom = db.prepare('SELECT * FROM movie_nominations WHERE id=?').get(req.params.id);
  if (!nom) return res.status(404).json({ error: 'not found' });
  const myVotesThisDay = db.prepare(`
    SELECT COUNT(*) c FROM movie_votes v JOIN movie_nominations n ON v.nominationId = n.id
    WHERE n.dayId = ? AND v.personId = ?`).get(nom.dayId, personId).c;
  const already = db.prepare('SELECT 1 FROM movie_votes WHERE nominationId=? AND personId=?').get(req.params.id, personId);
  if (already) {
    db.prepare('DELETE FROM movie_votes WHERE nominationId=? AND personId=?').run(req.params.id, personId);
  } else {
    if (myVotesThisDay >= 2) return res.status(400).json({ error: 'Only 2 votes per night' });
    db.prepare('INSERT INTO movie_votes (nominationId, personId) VALUES (?,?)').run(req.params.id, personId);
  }
  res.json({ ok: true });
});

api.post('/hottub-slots/:id/toggle', (req, res) => {
  const personId = req.body.personId || req.person?.id;
  if (!personId) return res.status(400).json({ error: 'personId required' });
  const existing = db.prepare('SELECT 1 FROM hottub_signups WHERE slotId=? AND personId=?').get(req.params.id, personId);
  if (existing) {
    db.prepare('DELETE FROM hottub_signups WHERE slotId=? AND personId=?').run(req.params.id, personId);
  } else {
    const count = db.prepare('SELECT COUNT(*) c FROM hottub_signups WHERE slotId=?').get(req.params.id).c;
    if (count >= 4) return res.status(400).json({ error: 'Slot full (max 4)' });
    db.prepare('INSERT INTO hottub_signups (slotId, personId) VALUES (?,?)').run(req.params.id, personId);
  }
  res.json({ ok: true });
});

// ---------- food ----------
api.get('/meals', (req, res) => res.json(db.prepare('SELECT * FROM meals').all()));
api.patch('/meals/:id', (req, res) => {
  const m = db.prepare('SELECT * FROM meals WHERE id=?').get(req.params.id);
  if (!m) return res.status(404).json({ error: 'not found' });
  const merged = { ...m, ...req.body };
  db.prepare('UPDATE meals SET plan=?, cooks=?, cleanup=? WHERE id=?').run(merged.plan, merged.cooks, merged.cleanup, req.params.id);
  res.json(db.prepare('SELECT * FROM meals WHERE id=?').get(req.params.id));
});

api.get('/dietary-notes', (req, res) => res.json(db.prepare('SELECT * FROM dietary_notes').all()));
api.put('/dietary-notes/:personId', (req, res) => {
  db.prepare('INSERT INTO dietary_notes (personId, note) VALUES (?,?) ON CONFLICT(personId) DO UPDATE SET note=excluded.note')
    .run(req.params.personId, req.body.note || '');
  res.json({ ok: true });
});

api.get('/groceries', (req, res) => res.json(db.prepare('SELECT * FROM grocery_items ORDER BY category, item').all()));
api.post('/groceries', (req, res) => {
  const { category, item, qty } = req.body;
  if (!category || !item) return res.status(400).json({ error: 'category and item required' });
  const id = uuid();
  db.prepare('INSERT INTO grocery_items (id, category, item, qty, claimedBy, checked) VALUES (?,?,?,?,NULL,0)').run(id, category, item, qty || '');
  res.json(db.prepare('SELECT * FROM grocery_items WHERE id=?').get(id));
});
api.patch('/groceries/:id', (req, res) => {
  const g = db.prepare('SELECT * FROM grocery_items WHERE id=?').get(req.params.id);
  if (!g) return res.status(404).json({ error: 'not found' });
  const merged = { ...g, ...req.body };
  db.prepare('UPDATE grocery_items SET claimedBy=?, checked=? WHERE id=?').run(merged.claimedBy, merged.checked ? 1 : 0, req.params.id);
  res.json(db.prepare('SELECT * FROM grocery_items WHERE id=?').get(req.params.id));
});
api.delete('/groceries/:id', (req, res) => {
  db.prepare('DELETE FROM grocery_items WHERE id=?').run(req.params.id);
  res.json({ ok: true });
});

api.get('/drinks-snacks', (req, res) => res.json(db.prepare('SELECT * FROM drinks_snacks').all()));
api.post('/drinks-snacks', (req, res) => {
  const id = uuid();
  db.prepare('INSERT INTO drinks_snacks (id, item, claimedBy) VALUES (?,?,NULL)').run(id, req.body.item);
  res.json(db.prepare('SELECT * FROM drinks_snacks WHERE id=?').get(id));
});
api.patch('/drinks-snacks/:id', (req, res) => {
  db.prepare('UPDATE drinks_snacks SET claimedBy=? WHERE id=?').run(req.body.claimedBy, req.params.id);
  res.json(db.prepare('SELECT * FROM drinks_snacks WHERE id=?').get(req.params.id));
});

// ---------- money ----------
api.get('/payments', (req, res) => res.json(db.prepare('SELECT * FROM payments').all()));
api.patch('/payments/:id', (req, res) => {
  const payment = db.prepare('SELECT * FROM payments WHERE id=?').get(req.params.id);
  if (!payment) return res.status(404).json({ error: 'not found' });
  const { status } = req.body;
  if (status === 'confirmed' && !(req.person && req.person.isOrganizer)) {
    return res.status(403).json({ error: 'Only the organizer can confirm payments' });
  }
  if (status === 'sent' && req.person && req.person.id !== payment.personId && !req.person.isOrganizer) {
    return res.status(403).json({ error: 'You can only mark your own payment sent' });
  }
  db.prepare('UPDATE payments SET status=? WHERE id=?').run(status, req.params.id);
  res.json(db.prepare('SELECT * FROM payments WHERE id=?').get(req.params.id));
});

api.get('/money-summary', (req, res) => {
  const people = db.prepare('SELECT * FROM people').all();
  const confirmedIds = new Set(people.filter(p => p.status === 'confirmed').map(p => p.id));
  const payments = db.prepare('SELECT * FROM payments').all();
  const confirmedPayments = payments.filter(p => confirmedIds.has(p.personId));
  const collected = confirmedPayments.filter(p => p.status === 'confirmed').reduce((s, p) => s + p.amount, 0);
  const totalNeeded = confirmedPayments.reduce((s, p) => s + p.amount, 0);
  const headcount = people.filter(p => p.status === 'confirmed' || p.status === 'maybe').length;
  const confirmedHeadcount = confirmedIds.size;
  const costPerPerson = Number(db.prepare("SELECT value FROM settings WHERE key='costPerPerson'").get()?.value) || 0;
  res.json({ collected, totalNeeded, confirmedHeadcount, headcount, costPerPerson });
});

api.get('/expenses', (req, res) => res.json(db.prepare('SELECT * FROM expenses ORDER BY date DESC').all()));
api.post('/expenses', (req, res) => {
  const { description, amount, paidBy, date } = req.body;
  if (!description || !amount || !paidBy) return res.status(400).json({ error: 'missing fields' });
  const id = uuid();
  db.prepare('INSERT INTO expenses (id, description, amount, paidBy, date) VALUES (?,?,?,?,?)')
    .run(id, description, amount, paidBy, date || new Date().toISOString().slice(0, 10));
  res.json(db.prepare('SELECT * FROM expenses WHERE id=?').get(id));
});
api.delete('/expenses/:id', (req, res) => {
  db.prepare('DELETE FROM expenses WHERE id=?').run(req.params.id);
  res.json({ ok: true });
});
api.get('/settle-up', (req, res) => {
  const people = db.prepare("SELECT * FROM people WHERE status='confirmed'").all();
  const expenses = db.prepare('SELECT * FROM expenses').all();
  const total = expenses.reduce((s, e) => s + e.amount, 0);
  const share = people.length ? total / people.length : 0;
  const paidByPerson = {};
  for (const p of people) paidByPerson[p.id] = 0;
  for (const e of expenses) paidByPerson[e.paidBy] = (paidByPerson[e.paidBy] || 0) + e.amount;
  const balances = people.map(p => ({ personId: p.id, name: p.name, paid: paidByPerson[p.id] || 0, share, net: (paidByPerson[p.id] || 0) - share }));
  // simple settle-up: debtors pay creditors
  const debtors = balances.filter(b => b.net < -0.01).map(b => ({ ...b })).sort((a, b) => a.net - b.net);
  const creditors = balances.filter(b => b.net > 0.01).map(b => ({ ...b })).sort((a, b) => b.net - a.net);
  const transactions = [];
  let di = 0, ci = 0;
  while (di < debtors.length && ci < creditors.length) {
    const d = debtors[di], c = creditors[ci];
    const amt = Math.min(-d.net, c.net);
    transactions.push({ from: d.name, to: c.name, amount: Math.round(amt * 100) / 100 });
    d.net += amt; c.net -= amt;
    if (Math.abs(d.net) < 0.01) di++;
    if (Math.abs(c.net) < 0.01) ci++;
  }
  res.json({ total, share, balances, transactions });
});

// ---------- committees ----------
api.get('/committees', (req, res) => {
  const committees = db.prepare('SELECT * FROM committees').all();
  const members = db.prepare('SELECT * FROM committee_members').all();
  const tasks = db.prepare('SELECT * FROM committee_tasks').all();
  const notes = db.prepare('SELECT * FROM committee_notes ORDER BY ts').all();
  res.json(committees.map(c => ({
    ...c,
    members: members.filter(m => m.committeeId === c.id).map(m => m.personId),
    tasks: tasks.filter(t => t.committeeId === c.id),
    notes: notes.filter(n => n.committeeId === c.id),
  })));
});
api.post('/committees/:id/join', (req, res) => {
  const personId = req.body.personId || req.person?.id;
  db.prepare('INSERT OR IGNORE INTO committee_members (committeeId, personId) VALUES (?,?)').run(req.params.id, personId);
  res.json({ ok: true });
});
api.post('/committees/:id/leave', (req, res) => {
  const personId = req.body.personId || req.person?.id;
  db.prepare('DELETE FROM committee_members WHERE committeeId=? AND personId=?').run(req.params.id, personId);
  res.json({ ok: true });
});
api.post('/committees/:id/tasks', (req, res) => {
  const { task, assigneeId, dueDate } = req.body;
  const id = uuid();
  db.prepare('INSERT INTO committee_tasks (id, committeeId, task, assigneeId, dueDate, done) VALUES (?,?,?,?,?,0)')
    .run(id, req.params.id, task, assigneeId || null, dueDate || null);
  res.json(db.prepare('SELECT * FROM committee_tasks WHERE id=?').get(id));
});
api.patch('/committee-tasks/:id', (req, res) => {
  const t = db.prepare('SELECT * FROM committee_tasks WHERE id=?').get(req.params.id);
  if (!t) return res.status(404).json({ error: 'not found' });
  const merged = { ...t, ...req.body };
  db.prepare('UPDATE committee_tasks SET task=?, assigneeId=?, dueDate=?, done=? WHERE id=?')
    .run(merged.task, merged.assigneeId, merged.dueDate, merged.done ? 1 : 0, req.params.id);
  res.json(db.prepare('SELECT * FROM committee_tasks WHERE id=?').get(req.params.id));
});
api.post('/committees/:id/notes', (req, res) => {
  const personId = req.body.personId || req.person?.id;
  const id = uuid();
  db.prepare('INSERT INTO committee_notes (id, committeeId, personId, text, ts) VALUES (?,?,?,?,?)')
    .run(id, req.params.id, personId || null, req.body.text, new Date().toISOString());
  res.json(db.prepare('SELECT * FROM committee_notes WHERE id=?').get(id));
});

// ---------- games ----------
api.get('/games-bring', (req, res) => res.json(db.prepare('SELECT * FROM games_bring').all()));
api.post('/games-bring', (req, res) => {
  const id = uuid();
  db.prepare('INSERT INTO games_bring (id, name, claimedBy) VALUES (?,?,NULL)').run(id, req.body.name);
  res.json(db.prepare('SELECT * FROM games_bring WHERE id=?').get(id));
});
api.patch('/games-bring/:id', (req, res) => {
  db.prepare('UPDATE games_bring SET claimedBy=? WHERE id=?').run(req.body.claimedBy, req.params.id);
  res.json(db.prepare('SELECT * FROM games_bring WHERE id=?').get(req.params.id));
});

api.get('/tournaments', (req, res) => {
  const rows = db.prepare('SELECT * FROM tournaments').all();
  res.json(rows.map(r => ({ ...r, players: JSON.parse(r.playersJson), bracket: JSON.parse(r.bracketJson) })));
});
api.post('/tournaments', (req, res) => {
  const { name, players } = req.body;
  const shuffled = [...players];
  const round1 = [];
  for (let i = 0; i < shuffled.length; i += 2) {
    round1.push({ id: uuid(), p1: shuffled[i], p2: shuffled[i + 1] || null, winner: shuffled[i + 1] ? null : shuffled[i] });
  }
  const id = uuid();
  db.prepare('INSERT INTO tournaments (id, name, playersJson, bracketJson, championId) VALUES (?,?,?,?,NULL)')
    .run(id, name, JSON.stringify(players), JSON.stringify({ rounds: [round1] }));
  res.json(db.prepare('SELECT * FROM tournaments WHERE id=?').get(id));
});
api.post('/tournaments/:id/advance', (req, res) => {
  const t = db.prepare('SELECT * FROM tournaments WHERE id=?').get(req.params.id);
  if (!t) return res.status(404).json({ error: 'not found' });
  const bracket = JSON.parse(t.bracketJson);
  const { matchId, winner } = req.body;
  const roundIdx = bracket.rounds.findIndex(r => r.some(m => m.id === matchId));
  const match = bracket.rounds[roundIdx].find(m => m.id === matchId);
  match.winner = winner;
  const currentRound = bracket.rounds[roundIdx];
  const allDone = currentRound.every(m => m.winner);
  let championId = t.championId;
  if (allDone) {
    if (currentRound.length === 1) {
      championId = winner;
    } else if (!bracket.rounds[roundIdx + 1]) {
      const nextRound = [];
      for (let i = 0; i < currentRound.length; i += 2) {
        nextRound.push({ id: uuid(), p1: currentRound[i].winner, p2: currentRound[i + 1]?.winner || null, winner: currentRound[i + 1] ? null : currentRound[i].winner });
      }
      bracket.rounds.push(nextRound);
    }
  }
  db.prepare('UPDATE tournaments SET bracketJson=?, championId=? WHERE id=?').run(JSON.stringify(bracket), championId, req.params.id);
  res.json({ ...db.prepare('SELECT * FROM tournaments WHERE id=?').get(req.params.id), bracket: JSON.parse(bracket ? JSON.stringify(bracket) : '{}') });
});

api.get('/birthday-checklist', (req, res) => res.json(db.prepare('SELECT * FROM birthday_checklist').all()));
api.patch('/birthday-checklist/:id', (req, res) => {
  db.prepare('UPDATE birthday_checklist SET done=? WHERE id=?').run(req.body.done ? 1 : 0, req.params.id);
  res.json({ ok: true });
});

api.get('/surprise-ideas', (req, res) => res.json(db.prepare('SELECT * FROM surprise_ideas ORDER BY ts').all()));
api.post('/surprise-ideas', (req, res) => {
  const personId = req.body.personId || req.person?.id;
  const id = uuid();
  db.prepare('INSERT INTO surprise_ideas (id, personId, text, ts) VALUES (?,?,?,?)').run(id, personId || null, req.body.text, new Date().toISOString());
  res.json(db.prepare('SELECT * FROM surprise_ideas WHERE id=?').get(id));
});

// ---------- logistics ----------
api.get('/rooms', (req, res) => {
  const rooms = db.prepare('SELECT * FROM rooms').all();
  const assigns = db.prepare('SELECT * FROM room_assignments').all();
  res.json(rooms.map(r => ({ ...r, occupants: assigns.filter(a => a.roomId === r.id) })));
});
api.post('/rooms', (req, res) => {
  const { name, bed, capacity, details } = req.body;
  if (!name) return res.status(400).json({ error: 'name required' });
  const id = uuid();
  db.prepare('INSERT INTO rooms (id, name, bed, capacity, details) VALUES (?,?,?,?,?)')
    .run(id, name, bed || '', Number(capacity) || 1, details || '');
  res.json(db.prepare('SELECT * FROM rooms WHERE id=?').get(id));
});
api.patch('/rooms/:id', (req, res) => {
  const r = db.prepare('SELECT * FROM rooms WHERE id=?').get(req.params.id);
  if (!r) return res.status(404).json({ error: 'not found' });
  const merged = { ...r, ...req.body };
  db.prepare('UPDATE rooms SET name=?, bed=?, capacity=?, details=? WHERE id=?')
    .run(merged.name, merged.bed, Number(merged.capacity) || 1, merged.details, req.params.id);
  res.json(db.prepare('SELECT * FROM rooms WHERE id=?').get(req.params.id));
});
api.delete('/rooms/:id', (req, res) => {
  db.prepare('DELETE FROM rooms WHERE id=?').run(req.params.id);
  db.prepare('DELETE FROM room_assignments WHERE roomId=?').run(req.params.id);
  res.json({ ok: true });
});
api.post('/rooms/:id/assign', (req, res) => {
  const { personId, note } = req.body;
  const room = db.prepare('SELECT * FROM rooms WHERE id=?').get(req.params.id);
  const count = db.prepare('SELECT COUNT(*) c FROM room_assignments WHERE roomId=?').get(req.params.id).c;
  if (count >= room.capacity) return res.status(400).json({ error: 'Room full' });
  db.prepare('DELETE FROM room_assignments WHERE personId=?').run(personId); // one room per person
  db.prepare('INSERT INTO room_assignments (roomId, personId, note) VALUES (?,?,?)').run(req.params.id, personId, note || '');
  res.json({ ok: true });
});
api.post('/rooms/:id/unassign', (req, res) => {
  db.prepare('DELETE FROM room_assignments WHERE roomId=? AND personId=?').run(req.params.id, req.body.personId);
  res.json({ ok: true });
});
api.patch('/room-assignments/:personId/note', (req, res) => {
  db.prepare('UPDATE room_assignments SET note=? WHERE personId=?').run(req.body.note || '', req.params.personId);
  res.json({ ok: true });
});

api.get('/carpools', (req, res) => {
  const carpools = db.prepare('SELECT * FROM carpools').all();
  const riders = db.prepare('SELECT * FROM carpool_riders').all();
  res.json(carpools.map(c => ({ ...c, riders: riders.filter(r => r.carpoolId === c.id).map(r => r.personId) })));
});
api.post('/carpools', (req, res) => {
  const { driverId, seats, departure, eta } = req.body;
  const id = uuid();
  db.prepare('INSERT INTO carpools (id, driverId, seats, departure, eta) VALUES (?,?,?,?,?)').run(id, driverId, seats, departure || '', eta || '');
  res.json(db.prepare('SELECT * FROM carpools WHERE id=?').get(id));
});
api.post('/carpools/:id/claim', (req, res) => {
  const personId = req.body.personId || req.person?.id;
  const carpool = db.prepare('SELECT * FROM carpools WHERE id=?').get(req.params.id);
  const count = db.prepare('SELECT COUNT(*) c FROM carpool_riders WHERE carpoolId=?').get(req.params.id).c;
  const already = db.prepare('SELECT 1 FROM carpool_riders WHERE carpoolId=? AND personId=?').get(req.params.id, personId);
  if (already) {
    db.prepare('DELETE FROM carpool_riders WHERE carpoolId=? AND personId=?').run(req.params.id, personId);
  } else {
    if (count >= carpool.seats) return res.status(400).json({ error: 'No seats left' });
    db.prepare('INSERT INTO carpool_riders (carpoolId, personId) VALUES (?,?)').run(req.params.id, personId);
  }
  res.json({ ok: true });
});

api.get('/packing', (req, res) => {
  const items = db.prepare('SELECT * FROM packing_items').all();
  const checks = db.prepare('SELECT * FROM packing_checks').all();
  res.json(items.map(i => ({ ...i, checks: checks.filter(c => c.itemId === i.id) })));
});
api.post('/packing', (req, res) => {
  const id = uuid();
  db.prepare('INSERT INTO packing_items (id, item, isShared) VALUES (?,?,1)').run(id, req.body.item);
  res.json(db.prepare('SELECT * FROM packing_items WHERE id=?').get(id));
});
api.post('/packing/:id/toggle', (req, res) => {
  const personId = req.body.personId || req.person?.id;
  if (!personId) return res.status(400).json({ error: 'personId required' });
  const existing = db.prepare('SELECT * FROM packing_checks WHERE itemId=? AND personId=?').get(req.params.id, personId);
  if (existing) {
    db.prepare('UPDATE packing_checks SET checked=? WHERE itemId=? AND personId=?').run(existing.checked ? 0 : 1, req.params.id, personId);
  } else {
    db.prepare('INSERT INTO packing_checks (itemId, personId, checked) VALUES (?,?,1)').run(req.params.id, personId);
  }
  res.json({ ok: true });
});

api.get('/cleanup-tasks', (req, res) => res.json(db.prepare('SELECT * FROM cleanup_tasks').all()));
api.patch('/cleanup-tasks/:id', (req, res) => {
  const t = db.prepare('SELECT * FROM cleanup_tasks WHERE id=?').get(req.params.id);
  const merged = { ...t, ...req.body };
  db.prepare('UPDATE cleanup_tasks SET claimedBy=?, done=? WHERE id=?').run(merged.claimedBy, merged.done ? 1 : 0, req.params.id);
  res.json(db.prepare('SELECT * FROM cleanup_tasks WHERE id=?').get(req.params.id));
});

app.use('/api', api);

// serve built frontend in production
const distPath = path.join(__dirname, '..', 'dist');
app.use(express.static(distPath));
app.get('*', (req, res, next) => {
  if (req.path.startsWith('/api')) return next();
  res.sendFile(path.join(distPath, 'index.html'), err => {
    if (err) res.status(404).send('Run `npm run build` first.');
  });
});

const PORT = process.env.PORT || 4001;
app.listen(PORT, () => console.log(`Cabin Trip server listening on :${PORT}`));
