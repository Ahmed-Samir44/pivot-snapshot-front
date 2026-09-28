# Pivot Snapshot Builder — Frontend

React 19 + Vite single-page app for the [Pivot Snapshot Builder](https://github.com/Ahmed-Samir44/pivot-snapshot-back)
backend. An interactive, Excel-PivotTable-style report builder against a live SQL Server Analysis
Services cube, with saved-snapshot history and side-by-side version comparison.

**Full project history, architecture decisions, and current status live in the backend repo's
[DECISIONS.md](https://github.com/Ahmed-Samir44/pivot-snapshot-back/blob/master/DECISIONS.md)** —
it covers both repos, not just the backend (see its own section 0). Read
`## 3. الحالة الحالية للتنفيذ` there for the latest "what's done / what's left" summary.

Three pages (`react-router-dom`, `HashRouter`):

1. **Pivot Builder** (`/`) — build and run a live pivot, save it as a snapshot.
2. **Snapshot History** (`/history`) — browse saved pivot definitions and their versions.
3. **Compare Snapshots** (`/compare`) — pick 2+ saved versions with the same structure and see
   what changed between them.

## Running locally

```bash
npm install
npm run dev -- --port 5180
```

Needs the backend running at `http://localhost:5191` (or set `VITE_API_BASE_URL`) and
`VITE_DATAVERSE_ENVIRONMENT_URL` for the Dataverse-backed snapshot storage — see
`src/services/dataverseAuth.js` for the sign-in flow (a hand-rolled PKCE flow, not MSAL.js — see
its own comment for why).

## Build / lint

```bash
npm run build
npm run lint
```
