# Cabin Trip 2026 🍂🏔️🔥

A lightweight, mobile-first collaborative trip planner for the birthday cabin weekend
(Brandon 🎂, Rachel 🎂, Lance 🎂 — Oct 2–5, 2026, East Keystone, CO).

No accounts. Everyone taps their name once and the app remembers who they are
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
Open `http://localhost:5173`. The database is created and seeded automatically on first run
(`server/data.db`) — nothing to configure.

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
`cabin-trip-app/server/` (or wherever `data.db` lands) — otherwise a redeploy resets to seed data.

## Share the link with the group

Once deployed, just send the group the URL. First visit shows a roster picker — everyone taps
their own name and they're in. No invites, no passwords.

## Edit the roster

Any user can go to **More → ⚙️ Settings** to add a guest, rename someone, toggle their status
between `confirmed`/`maybe`, or remove them — no code changes or redeploy needed.

## Reset the database

Stop the server, delete `server/data.db` (and the `-shm`/`-wal` files next to it if present), then
restart (`npm start` or `npm run dev:server`). The app re-seeds automatically from scratch.

## Notes on design decisions

- **Organizer** = Brandon. Only he can mark a payment row "Confirmed" (anyone can mark their own
  row "Sent"). This is a soft client/server check tied to identity, not real auth — fine for a
  private group trip app.
- **Surprise ideas** panel (Games → 🎂 Birthday) is hidden from whoever is currently identified as
  Brandon, Rachel, or Lance — also a soft, client-name-based gate, not real security.
- Seed data fully populates every module (schedule, food, money, committees, games, logistics) so
  the app is immediately usable with zero empty states.
