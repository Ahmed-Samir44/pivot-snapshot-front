import { getDataverseAccessToken } from "./dataverseAuth";
import { apiErrorFrom } from "./apiError";
import { reviveResultCells } from "../utils/errorCells";
import { createResponseCache } from "./responseCache";

// Falls back to the local dev API port when VITE_API_BASE_URL isn't set. Production builds must
// set it explicitly via .env.production (not .env.local, which should never affect a build).
const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

// `signal` (an AbortController's, passed by the caller) lets a NEWER query cancel an OLDER one
// still in flight instead of just ignoring its eventual response — the backend's own
// CubeQueryService now actually calls AdomdCommand.Cancel() when its request is aborted (see its
// own WHY comment), so an aborted fetch here stops real work on the shared cube server, not just
// discards a response client-side (requested live, 2026-10-05, "عايز تزود cancellation token").
//
// Recent answers are remembered for 10 minutes (responseCache.js), so Undo or going back to an
// earlier layout shows at once. Only finished answers are kept — a query in flight isn't shared, so
// aborting it still stops it on the cube. `fresh` (Excel's Refresh) skips the memory, here and on
// the server ("Cache-Control: no-cache"), and keeps the new answer.
const recentResults = createResponseCache({ ttlMs: 10 * 60 * 1000, maxEntries: 20 });

export async function queryPivot(request, { signal, fresh = false } = {}) {
  const body = JSON.stringify(request);
  const remembered = fresh ? undefined : recentResults.get(body);
  if (remembered) return remembered;

  // The cube itself has no per-user access control (shared service login) — this token is the
  // only thing gating who can query it at all, even though the cube query path never touches
  // Dataverse. See DECISIONS.md (2026-09-22) and dataverseAuth.js.
  const accessToken = await getDataverseAccessToken();
  const headers = { "Content-Type": "application/json", Authorization: `Bearer ${accessToken}` };
  if (fresh) headers["Cache-Control"] = "no-cache";
  const response = await fetch(`${API_BASE_URL}/api/pivot/query`, {
    method: "POST",
    headers,
    body,
    signal,
  });

  if (!response.ok) {
    throw await apiErrorFrom(response, "An unexpected error occurred.");
  }
  const data = await response.json();

  // Error cells arrive as "NaN" (see utils/errorCells.js).
  const result = reviveResultCells(data);
  recentResults.set(body, result);
  return result;
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

  if (!response.ok) {
    throw await apiErrorFrom(response, "An unexpected error occurred.");
  }
  const data = await response.json();

  return data;
}
