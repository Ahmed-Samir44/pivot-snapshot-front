// An API failure as an Error carrying the server's own message and, for the data server's
// temporary states, its code — "cube_busy" (out of memory) or "cube_unavailable" (down/refusing) —
// which the app answers with a calm message and an automatic retry instead of a red error
// (2026-10-20, "i need the user to be pleased to use the system").
export class ApiError extends Error {
  constructor(message, { status = 0, code = null } = {}) {
    super(message);
    this.name = "ApiError";
    this.status = status;
    this.code = code;
  }
}

export const isCubeTemporarilyDown = (err) => err?.code === "cube_busy" || err?.code === "cube_unavailable";

// Reads a failed response's JSON body ({ error, code }) when there is one.
export async function apiErrorFrom(response, fallback) {
  let body = null;
  try {
    body = await response.json();
  } catch {
    // Not JSON (a proxy error page, an empty body) — the fallback text is used.
  }
  return new ApiError(body?.error || fallback, { status: response.status, code: body?.code ?? null });
}
