const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

async function request(path, options) {
  const response = await fetch(`${API_BASE_URL}${path}`, options);
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(data?.error || `Request to ${path} failed with status ${response.status}`);
  }
  return data;
}

export function saveSnapshot(name, pivotRequest, pivotResult) {
  return request("/api/snapshots", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, request: pivotRequest, result: pivotResult }),
  });
}

export function listDefinitions() {
  return request("/api/snapshots");
}

export function getVersions(definitionId) {
  return request(`/api/snapshots/${definitionId}/versions`);
}

export function getVersion(versionId) {
  return request(`/api/snapshots/versions/${versionId}`);
}

export function compareVersions(versionA, versionB) {
  return request(`/api/snapshots/compare?versionA=${versionA}&versionB=${versionB}`);
}
