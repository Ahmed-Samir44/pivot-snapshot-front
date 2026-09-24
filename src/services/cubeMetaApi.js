import { getDataverseAccessToken } from "./dataverseAuth";

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

// Same access-gate reasoning as pivotApi.js: the cube has no per-user access control of its own,
// so this token (not any Dataverse permission it happens to also carry) is what stops an
// unauthenticated caller from listing the cube's fields/members at all.
async function getJson(path) {
  const accessToken = await getDataverseAccessToken();
  const response = await fetch(`${API_BASE_URL}${path}`, { headers: { Authorization: `Bearer ${accessToken}` } });
  if (!response.ok) {
    throw new Error(`Request to ${path} failed with status ${response.status}`);
  }
  return response.json();
}

export function getDimensions() {
  return getJson("/api/cube/dimensions");
}

export function getMeasures() {
  return getJson("/api/cube/measures");
}

// `search`, when given, matches against the field's FULL member set on the server, not just the
// first page — a field past the server's cap (see CubeMetadataService.MaxMembers) can otherwise
// only ever offer whichever members land in the cube's own default order, unrelated to what the
// user actually typed (confirmed live 2026-09-23: a real, active doctor wasn't in the first 500
// at all, so no amount of typing their name found them).
export function getMembers(field, search) {
  const query = new URLSearchParams({ field });
  if (search) query.set("search", search);
  return getJson(`/api/cube/members?${query.toString()}`);
}
