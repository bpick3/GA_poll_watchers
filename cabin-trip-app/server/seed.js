import { db } from './db.js';
import { randomUUID as uuid } from 'node:crypto';

// First run only: app-level defaults that aren't trip-specific data.
// Trip name, dates, location, roster, rooms, cost, etc. all come from the
// setup wizard (see POST /api/setup in index.js) — nothing trip-specific
// is baked in here.
export function seedIfEmpty() {
  const count = db.prepare('SELECT COUNT(*) c FROM settings').get().c;
  if (count > 0) return;

  const setSetting = db.prepare('INSERT INTO settings (key, value) VALUES (?,?)');
  setSetting.run('tripName', 'Cabin Fever 2026');
  setSetting.run('setupComplete', '0');
  setSetting.run('altThemes', JSON.stringify([
    'Y2K Night 💿',
    'Cabin Casino 🎰',
    'Murder Mystery Dinner 🔪',
    'White Lies Party 🤥',
    'Decades Night (choose an era) 🕺',
  ]));

  // Committees are a generic app feature (not trip-specific personal data) —
  // always available so the group has somewhere to organize from day one.
  const insertCommittee = db.prepare('INSERT INTO committees (id, name, emoji) VALUES (?,?,?)');
  const committees = [
    { name: 'Food Planning & Cooking', emoji: '🍽️' },
    { name: 'Supplies & Shopping', emoji: '🛒' },
    { name: 'Games & Activities', emoji: '🎲' },
    { name: 'Vibes & Extras', emoji: '✨' },
  ];
  for (const c of committees) insertCommittee.run(uuid(), c.name, c.emoji);

  console.log('Database initialized (empty — run the setup wizard to configure your trip).');
}
