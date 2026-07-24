# Cabin Fever 2026 🍂🏔️🔥

A lightweight, mobile-first collaborative trip planner for a group cabin weekend. Trip name,
dates, location, cost, roster, and rooms are all set up through an in-app **setup wizard** on
first launch — nothing trip-specific is hardcoded, so this works for any trip, not just one.

No accounts. Everyone taps their name once (after setup) and the app remembers who they are
(stored in `localStorage`) for RSVPs, claims, and payments.

## Stack

- **Frontend**: Vite + React (plain JS), polling the API every ~10s for near-real-time updates.
- **Backend**: Express + better-sqlite3, single process, single SQLite file (`server/data.db`).
- The Express server serves the built frontend as static files in production — one process, one port.

## Run locally

```bash
cd cabin-trip-app
npm install
npm run dev
```

This runs the API on `:4001` and Vite's dev server (with `/api` proxied to it) on `:5173`.
Open `http://localhost:5173`. The database is created automatically on first run
(`server/data.db`) — the app will walk you through the setup wizard (trip basics, house info,
roster, rooms) before showing the roster picker.

## Build + run as one process (what you'd deploy)

```bash
npm install
npm run build   # builds the React app into dist/
npm start        # starts Express, which serves dist/ AND the API on one port
```

By default the server listens on `PORT` env var or `4001`.

## Deploy to Railway / Render / Fly

All three just need: install, `npm run build`, then `npm start`, with a single exposed port.

- **Railway / Render**: point the service at the `cabin-trip-app/` directory (or set root directory
  to `cabin-trip-app` if deploying from this monorepo). Build command: `npm install && npm run build`.
  Start command: `npm start`. They auto-set `PORT`; the server already reads `process.env.PORT`.
- **Fly.io**: same idea — a minimal `Dockerfile` (`FROM node:20-alpine`, copy the app, `npm ci && npm run build`,
  `CMD ["npm", "start"]`) works, or use `fly launch` with a Node buildpack.

Because everything lives in one SQLite file, there's no separate database service to provision.
For persistence across deploys/restarts on these platforms, mount a small persistent volume at
`cabin-trip-app/server/` (or wherever `data.db` lands) — otherwise a redeploy resets to an
empty, unsetup database.

## First run: the setup wizard

The very first time the app starts with an empty database, it shows a setup wizard instead of
the roster picker:

1. **Trip Basics** — name, dates, address, rental listing, cost per person, two payment
   installments (label/date/amount each).
2. **House Info** — check-in/out, wifi password, house rules, quiet hours, altitude/local tips,
   emergency info, a free-text house description.
3. **Roster** — add each person, mark who's an organizer (can confirm payments) and who's a
   guest of honor 🎂 (gets a birthday checklist + a surprise-ideas board hidden from them).
4. **Rooms** — optional; bedrooms with bed type and capacity for the sleeping-arrangement
   assignments. Can also be added later.

Once submitted, the schedule's day tabs and blank meal grid are generated automatically from the
trip's date range, and the roster picker appears for everyone to tap their name.

## Share the link with the group

Once deployed and set up, just send the group the URL. Every visit after setup shows the roster
picker — everyone taps their own name and they're in. No invites, no passwords.

## Edit anything after setup

Everything from the wizard is editable later without touching code:

- **More → ⚙️ Settings**: trip name/dates/location/cost, house info, and the full roster
  (add, rename, change status, toggle organizer/guest-of-honor, remove).
- **More → 🧭 Logistics → Rooms**: add or remove rooms any time.
- **More → 🎲 Games → Tournament**: create a new bracket any time one isn't running.
- Schedule day themes, blocks, meals, groceries, committees, and packing lists are all editable
  in place from their respective tabs.

## Reset the database

Stop the server, delete `server/data.db` (and the `-shm`/`-wal` files next to it if present), then
restart (`npm start` or `npm run dev:server`). The app comes back up empty and shows the setup
wizard again.

## Notes on design decisions

- **Organizer** is a per-person flag set in the setup wizard (or toggled later in Settings) —
  there can be more than one. Only organizers can mark a payment row "Confirmed" (anyone can mark
  their own row "Sent"). This is a soft client/server check tied to identity, not real auth — fine
  for a private group trip app.
- **Surprise ideas** panel (Games → 🎂 Birthday) is hidden from whoever is currently identified as
  a guest of honor — also a soft, client-identity-based gate, not real security.
- The database starts empty aside from generic app-level defaults (a starter committee structure,
  a small alternate-theme library) — everything trip-specific comes from the setup wizard, so
  there's no empty-state wasteland once it's filled in, and nothing to hand-edit in code to reuse
  this for a different trip.
