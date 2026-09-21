// Falls back to the local dev API port when VITE_API_BASE_URL isn't set. Production builds must
// set it explicitly via .env.production (not .env.local, which should never affect a build).
const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

export async function queryPivot(request) {
  const response = await fetch(`${API_BASE_URL}/api/pivot/query`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(data.error || "حدث خطأ غير متوقع");
  }

  return data;
}
