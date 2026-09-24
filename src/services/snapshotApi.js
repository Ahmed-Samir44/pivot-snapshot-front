import { getDataverseAccessToken } from "./dataverseAuth";

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

// Every snapshot endpoint writes/reads Dataverse AS THE SIGNED-IN USER (see dataverseAuth.js) —
// unlike pivotApi.js/cubeMetaApi.js, which talk to the cube and never touch Dataverse, so they
// need no token at all. getDataverseAccessToken() silently reuses/refreshes the cached sign-in;
// SignInGate guarantees one already exists before any of this module's calls can fire.
async function request(path, options = {}) {
  const accessToken = await getDataverseAccessToken();
  const response = await fetch(`${API_BASE_URL}${path}`, {
    ...options,
    headers: {
      ...(options.headers ?? {}),
      Authorization: `Bearer ${accessToken}`,
    },
  });
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(data?.error || `Request to ${path} failed with status ${response.status}`);
  }
  return data;
}

export function saveSnapshot(name, pivotRequest, pivotResult, tableStyle = "Default", includeChart = false, chartType = "Bar", reportLayout = "Tabular") {
  return request("/api/snapshots", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, request: pivotRequest, result: pivotResult, tableStyle, includeChart, chartType, reportLayout }),
  });
}

export function listDefinitions() {
  return request("/api/snapshots");
}

export function getDeactivatedDefinitions() {
  return request("/api/snapshots/deactivated");
}

// Also deactivates every active version under this pivot — see
// DataverseSnapshotStorageService.DeactivateDefinitionAsync for why. Returns
// { deactivatedVersionCount } so the caller can tell the user how many versions were affected.
export function deactivateDefinition(definitionId) {
  return request(`/api/snapshots/${definitionId}/deactivate`, { method: "POST" });
}

export function reactivateDefinition(definitionId) {
  return request(`/api/snapshots/${definitionId}/reactivate`, { method: "POST" });
}

// The real, permanent delete — refused by the backend unless the pivot is already deactivated,
// and wipes every version under it (active or deactivated) along with the pivot itself.
export function deleteDefinition(definitionId) {
  return request(`/api/snapshots/${definitionId}`, { method: "DELETE" });
}

export function getVersions(definitionId) {
  return request(`/api/snapshots/${definitionId}/versions`);
}

export function getDeactivatedVersions(definitionId) {
  return request(`/api/snapshots/${definitionId}/versions/deactivated`);
}

export function getVersion(versionId) {
  return request(`/api/snapshots/versions/${versionId}`);
}

export function compareVersions(versionA, versionB) {
  return request(`/api/snapshots/compare?versionA=${versionA}&versionB=${versionB}`);
}

// Deactivates (not deletes) a version — see DataverseSnapshotStorageService.DeactivateVersionAsync
// for why: recoverable via Dataverse's own admin tools instead of gone forever from one click here.
export function deactivateVersion(versionId) {
  return request(`/api/snapshots/versions/${versionId}/deactivate`, { method: "POST" });
}

export function reactivateVersion(versionId) {
  return request(`/api/snapshots/versions/${versionId}/reactivate`, { method: "POST" });
}

// Reactivates every currently-deactivated version under a pivot in one call — the common gap
// after deactivating a pivot (which cascades to its versions) and then reactivating just the
// pivot itself (which deliberately doesn't cascade back — see
// DataverseSnapshotStorageService.ReactivateDefinitionAsync).
export function reactivateAllVersions(definitionId) {
  return request(`/api/snapshots/${definitionId}/versions/reactivate-all`, { method: "POST" });
}

// The real, permanent delete — the backend itself refuses this unless the version was already
// deactivated first (see DataverseSnapshotStorageService.DeleteVersionAsync), so this only ever
// makes sense called from the Deactivated section of the UI.
export function deleteVersion(versionId) {
  return request(`/api/snapshots/versions/${versionId}`, { method: "DELETE" });
}
