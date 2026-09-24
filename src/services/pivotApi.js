import { getDataverseAccessToken } from "./dataverseAuth";

// Falls back to the local dev API port when VITE_API_BASE_URL isn't set. Production builds must
// set it explicitly via .env.production (not .env.local, which should never affect a build).
const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

export async function queryPivot(request) {
  // The cube itself has no per-user access control (shared service login) — this token is the
  // only thing gating who can query it at all, even though the cube query path never touches
  // Dataverse. See DECISIONS.md (2026-09-22) and dataverseAuth.js.
  const accessToken = await getDataverseAccessToken();
  const response = await fetch(`${API_BASE_URL}/api/pivot/query`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${accessToken}` },
    body: JSON.stringify(request),
  });

  const data = await response.json();

  if (!response.ok) {
    const error = new Error(data.error || "An unexpected error occurred.");
    // Structured fields (type/field/memberCount/canOverride) from CardinalityGuard's
    // HighCardinalityFieldException — carried on the Error object (not just its message) so the
    // caller can offer a specific "run anyway" retry instead of only showing text.
    error.details = data;
    throw error;
  }

  return data;
}
