import Database from 'better-sqlite3';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const dbPath = path.join(__dirname, 'data.db');
const isNew = !fs.existsSync(dbPath);

export const db = new Database(dbPath);
db.pragma('journal_mode = WAL');

// --- schema ---
db.exec(`
CREATE TABLE IF NOT EXISTS people (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'confirmed', -- confirmed | maybe
  isOrganizer INTEGER NOT NULL DEFAULT 0,
  isGuestOfHonor INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS settings (
  key TEXT PRIMARY KEY,
  value TEXT
);

CREATE TABLE IF NOT EXISTS days (
  id TEXT PRIMARY KEY,
  date TEXT NOT NULL,
  label TEXT NOT NULL,
  theme TEXT NOT NULL,
  sortOrder INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS blocks (
  id TEXT PRIMARY KEY,
  dayId TEXT NOT NULL,
  time TEXT,
  title TEXT NOT NULL,
  type TEXT NOT NULL DEFAULT 'custom', -- yoga|hike|movie|hottub|custom
  location TEXT,
  ownerId TEXT,
  ownerText TEXT,
  notes TEXT,
  sortOrder INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS block_rsvps (
  blockId TEXT NOT NULL,
  personId TEXT NOT NULL,
  status TEXT NOT NULL, -- in | out
  PRIMARY KEY (blockId, personId)
);

CREATE TABLE IF NOT EXISTS movie_nominations (
  id TEXT PRIMARY KEY,
  dayId TEXT NOT NULL,
  title TEXT NOT NULL,
  nominatedBy TEXT
);

CREATE TABLE IF NOT EXISTS movie_votes (
  nominationId TEXT NOT NULL,
  personId TEXT NOT NULL,
  PRIMARY KEY (nominationId, personId)
);

CREATE TABLE IF NOT EXISTS hottub_slots (
  id TEXT PRIMARY KEY,
  dayId TEXT NOT NULL,
  startTime TEXT NOT NULL,
  label TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS hottub_signups (
  slotId TEXT NOT NULL,
  personId TEXT NOT NULL,
  PRIMARY KEY (slotId, personId)
);

CREATE TABLE IF NOT EXISTS meals (
  id TEXT PRIMARY KEY,
  dayId TEXT NOT NULL,
  mealType TEXT NOT NULL, -- Breakfast|Lunch|Dinner|Snacks
  plan TEXT,
  cooks TEXT, -- csv of person ids
  cleanup TEXT -- csv of person ids
);

CREATE TABLE IF NOT EXISTS dietary_notes (
  personId TEXT PRIMARY KEY,
  note TEXT
);

CREATE TABLE IF NOT EXISTS grocery_items (
  id TEXT PRIMARY KEY,
  category TEXT NOT NULL,
  item TEXT NOT NULL,
  qty TEXT,
  claimedBy TEXT,
  checked INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS drinks_snacks (
  id TEXT PRIMARY KEY,
  item TEXT NOT NULL,
  claimedBy TEXT
);

CREATE TABLE IF NOT EXISTS payments (
  id TEXT PRIMARY KEY,
  personId TEXT NOT NULL,
  dueLabel TEXT NOT NULL,
  dueDate TEXT NOT NULL,
  amount REAL NOT NULL,
  status TEXT NOT NULL DEFAULT 'not_sent' -- not_sent|sent|confirmed
);

CREATE TABLE IF NOT EXISTS expenses (
  id TEXT PRIMARY KEY,
  description TEXT NOT NULL,
  amount REAL NOT NULL,
  paidBy TEXT NOT NULL,
  date TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS committees (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  emoji TEXT
);

CREATE TABLE IF NOT EXISTS committee_members (
  committeeId TEXT NOT NULL,
  personId TEXT NOT NULL,
  PRIMARY KEY (committeeId, personId)
);

CREATE TABLE IF NOT EXISTS committee_tasks (
  id TEXT PRIMARY KEY,
  committeeId TEXT NOT NULL,
  task TEXT NOT NULL,
  assigneeId TEXT,
  dueDate TEXT,
  done INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS committee_notes (
  id TEXT PRIMARY KEY,
  committeeId TEXT NOT NULL,
  personId TEXT,
  text TEXT NOT NULL,
  ts TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS games_bring (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  claimedBy TEXT
);

CREATE TABLE IF NOT EXISTS tournaments (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  playersJson TEXT NOT NULL,
  bracketJson TEXT NOT NULL,
  championId TEXT
);

CREATE TABLE IF NOT EXISTS birthday_checklist (
  id TEXT PRIMARY KEY,
  forPerson TEXT NOT NULL,
  item TEXT NOT NULL,
  done INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS surprise_ideas (
  id TEXT PRIMARY KEY,
  personId TEXT,
  text TEXT NOT NULL,
  ts TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS rooms (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  bed TEXT NOT NULL,
  capacity INTEGER NOT NULL,
  details TEXT
);

CREATE TABLE IF NOT EXISTS room_assignments (
  roomId TEXT NOT NULL,
  personId TEXT NOT NULL,
  note TEXT,
  PRIMARY KEY (roomId, personId)
);

CREATE TABLE IF NOT EXISTS carpools (
  id TEXT PRIMARY KEY,
  driverId TEXT NOT NULL,
  seats INTEGER NOT NULL,
  departure TEXT,
  eta TEXT
);

CREATE TABLE IF NOT EXISTS carpool_riders (
  carpoolId TEXT NOT NULL,
  personId TEXT NOT NULL,
  PRIMARY KEY (carpoolId, personId)
);

CREATE TABLE IF NOT EXISTS packing_items (
  id TEXT PRIMARY KEY,
  item TEXT NOT NULL,
  isShared INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS packing_checks (
  itemId TEXT NOT NULL,
  personId TEXT NOT NULL,
  checked INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (itemId, personId)
);

CREATE TABLE IF NOT EXISTS cleanup_tasks (
  id TEXT PRIMARY KEY,
  task TEXT NOT NULL,
  claimedBy TEXT,
  done INTEGER NOT NULL DEFAULT 0
);
`);

export { isNew };
