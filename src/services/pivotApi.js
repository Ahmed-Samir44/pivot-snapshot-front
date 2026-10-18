import { getDataverseAccessToken } from "./dataverseAuth";
import { reviveResultCells } from "../utils/errorCells";

// Falls back to the local dev API port when VITE_API_BASE_URL isn't set. Production builds must
// set it explicitly via .env.production (not .env.local, which should never affect a build).
const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

// `signal` (an AbortController's, passed by the caller) lets a NEWER query cancel an OLDER one
// still in flight instead of just ignoring its eventual response — the backend's own
// CubeQueryService now actually calls AdomdCommand.Cancel() when its request is aborted (see its
// own WHY comment), so an aborted fetch here stops real work on the shared cube server, not just
// discards a response client-side (requested live, 2026-10-05, "عايز تزود cancellation token").
export async function queryPivot(request, { signal } = {}) {
  // The cube itself has no per-user access control (shared service login) — this token is the
  // only thing gating who can query it at all, even though the cube query path never touches
  // Dataverse. See DECISIONS.md (2026-09-22) and dataverseAuth.js.
  const accessToken = await getDataverseAccessToken();
  const response = await fetch(`${API_BASE_URL}/api/pivot/query`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${accessToken}` },
    body: JSON.stringify(request),
    signal,
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(data.error || "An unexpected error occurred.");
  }

  // Error cells arrive as "NaN" (see utils/errorCells.js).
  return reviveResultCells(data);
}

// Excel's "Show Details" on one value cell — mirrors queryPivot exactly (same auth/error handling),
// hitting the new MDX DRILLTHROUGH endpoint instead. See PivotGrid.jsx's context menu for the caller.
export async function drillthroughPivot(request, { signal } = {}) {
  const accessToken = await getDataverseAccessToken();
  const response = await fetch(`${API_BASE_URL}/api/pivot/drillthrough`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${accessToken}` },
    body: JSON.stringify(request),
    signal,
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(data.error || "An unexpected error occurred.");
  }

  return data;
}
