import { getDataverseAccessToken } from "./dataverseAuth";

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

// A read-only proxy over Dataverse's "stf_strategykpis" table (see
// DataverseSnapshotStorageService.GetKpisAsync's own WHY comment) — same delegated-token auth model
// as every snapshotApi.js call, since the backend forwards this same token straight to Dataverse.
export async function getKpis() {
  const accessToken = await getDataverseAccessToken();
  const response = await fetch(`${API_BASE_URL}/api/kpis`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(data?.error || `Request to /api/kpis failed with status ${response.status}`);
  }
  return data;
}
