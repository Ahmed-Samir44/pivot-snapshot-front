const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

async function getJson(path) {
  const response = await fetch(`${API_BASE_URL}${path}`);
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

export function getMembers(field) {
  return getJson(`/api/cube/members?field=${encodeURIComponent(field)}`);
}
