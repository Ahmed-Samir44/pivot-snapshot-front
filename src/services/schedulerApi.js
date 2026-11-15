import { getDataverseAccessToken, getRefreshTokenForScheduling } from "./dataverseAuth";

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

async function request(path, options = {}) {
  const accessToken = await getDataverseAccessToken();
  const response = await fetch(`${API_BASE_URL}${path}`, {
    ...options,
    headers: {
      ...(options.headers ?? {}),
      Authorization: `Bearer ${accessToken}`,
    },
  });
  if (response.status === 204) return null;
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(data?.error || `Request to ${path} failed with status ${response.status}`);
  }
  return data;
}

// Whether a scheduled-run credential is currently stored server-side — see
// ScheduledRunTokenStore's own WHY comment.
export function getSchedulerStatus() {
  return request("/api/scheduler/status");
}

// Hands the CURRENT user's own refresh token to the backend so it can keep running scheduled
// pivots unattended as "this user" — the bridge the user explicitly asked for ("اعمله بحسابي
// دلوقتي") until a real service identity exists for this tenant.
export function connectScheduler() {
  const refreshToken = getRefreshTokenForScheduling();
  if (!refreshToken) {
    throw new Error("You're not signed in with a refresh token to hand over — try signing out and back in first.");
  }
  return request("/api/scheduler/connect", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ refreshToken }),
  });
}

export function disconnectScheduler() {
  return request("/api/scheduler/disconnect", { method: "POST" });
}
