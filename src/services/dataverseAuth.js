// Signs the user in with their OWN Microsoft account and hands out a Dataverse access token —
// see DECISIONS.md (2026-09-22/23) for why this exists: the backend has no service identity of
// its own for Dataverse (no App Registration could be issued in this tenant), so every save/read
// against Dataverse runs AS THE SIGNED-IN USER instead. Every Dataverse write in this project
// happens because THIS user did it, not because a shared service account did.
//
// CLIENT_ID is Microsoft's own "Microsoft Azure CLI" first-party public client — pre-consented in
// effectively every Azure AD tenant already, so no App Registration of our own is needed (the
// exact same registration problem that blocked a custom one blocks nothing here). This is a
// widely used, Microsoft-documented pattern for talking to Dataverse without registering your own
// app — NOT a security workaround; it's how Microsoft's own tooling (e.g. XrmToolBox) does it too.
//
// HAND-ROLLED PKCE, not MSAL.js — MSAL was tried first (both loginPopup and loginRedirect) and
// both ultimately fail at the SAME step: the token exchange. This well-known client is registered
// in Azure AD as a native/CLI app, not a Single-Page Application, so its token endpoint never
// returns CORS headers to a browser caller — confirmed live via a plain
// "Access to fetch ... has been blocked by CORS policy" in the console, regardless of popup vs
// redirect. The fix isn't a different MSAL flow, it's moving the ONE call that needs CORS
// (the POST to /oauth2/v2.0/token) server-side, where CORS doesn't apply — see the backend's
// AuthController. Everything before that (navigating to /authorize, getting a `code` back) is a
// plain browser navigation and was never the problem.
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const AUTHORITY = "https://login.microsoftonline.com/organizations";
const DATAVERSE_ENVIRONMENT_URL = (import.meta.env.VITE_DATAVERSE_ENVIRONMENT_URL || "https://org319b4ea9.crm4.dynamics.com").replace(/\/$/, "");
const DATAVERSE_SCOPE = `${DATAVERSE_ENVIRONMENT_URL}/.default`;
const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:5191";

const STORAGE_KEY = "pivotSnapshot.dataverseTokens"; // { accessToken, expiresAt, refreshToken }
const PENDING_KEY = "pivotSnapshot.pendingPkce"; // sessionStorage: { verifier, state } — only needed across the redirect itself

// A token this close to expiry gets refreshed proactively rather than handed out and failing
// mid-request a few seconds later.
const EXPIRY_BUFFER_MS = 60_000;

function base64UrlEncode(bytes) {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function randomString(length) {
  const bytes = crypto.getRandomValues(new Uint8Array(length));
  return base64UrlEncode(bytes);
}

async function sha256Base64Url(value) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(value));
  return base64UrlEncode(new Uint8Array(digest));
}

function readTokens() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function writeTokens(tokenResponse) {
  const tokens = {
    accessToken: tokenResponse.access_token,
    expiresAt: Date.now() + tokenResponse.expires_in * 1000,
    // Keep the previous refresh token if this response didn't include a new one (a silent refresh
    // sometimes omits it, and losing it would force a full interactive sign-in next time).
    refreshToken: tokenResponse.refresh_token ?? readTokens()?.refreshToken ?? null,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(tokens));
  return tokens;
}

function clearTokens() {
  localStorage.removeItem(STORAGE_KEY);
}

async function exchangeWithBackend(body) {
  const response = await fetch(`${API_BASE_URL}/api/auth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const json = await response.json();
  if (!response.ok) {
    throw new Error(json.error_description || json.error || "Token exchange failed.");
  }
  return json;
}

// Navigates the WHOLE tab to Microsoft's login page — this call does not "return" in the normal
// sense (the page unloads before the promise would resolve); the sign-in result is picked back up
// by handlePendingRedirect() on the next page load instead, once the browser lands back here.
export async function signIn() {
  const verifier = randomString(32);
  const challenge = await sha256Base64Url(verifier);
  const state = randomString(16);

  sessionStorage.setItem(PENDING_KEY, JSON.stringify({ verifier, state }));

  const redirectUri = window.location.origin + window.location.pathname;
  const authorizeUrl = new URL(`${AUTHORITY}/oauth2/v2.0/authorize`);
  authorizeUrl.search = new URLSearchParams({
    client_id: CLIENT_ID,
    response_type: "code",
    redirect_uri: redirectUri,
    response_mode: "query",
    scope: DATAVERSE_SCOPE,
    state,
    code_challenge: challenge,
    code_challenge_method: "S256",
  }).toString();

  window.location.assign(authorizeUrl.toString());
}

// Runs once per page load: if this load IS the browser landing back from Microsoft (a `code` and
// `state` sit in the query string), completes the PKCE exchange via our backend and stores the
// resulting tokens. Every caller awaits this SAME promise, so a pending redirect only ever gets
// consumed once no matter how many places call into this module.
let initPromise = null;
export function ensureInitialized() {
  if (!initPromise) {
    initPromise = handlePendingRedirect();
  }
  return initPromise;
}

async function handlePendingRedirect() {
  const params = new URLSearchParams(window.location.search);
  const code = params.get("code");
  const state = params.get("state");
  const errorDescription = params.get("error_description");

  if (!code && !errorDescription) {
    return;
  }

  // Clean the code/state/error out of the visible URL either way, so a page refresh doesn't
  // replay a used-up (and by then invalid) authorization code.
  const cleanUrl = window.location.origin + window.location.pathname + window.location.hash;
  window.history.replaceState({}, document.title, cleanUrl);

  if (errorDescription) {
    throw new Error(errorDescription);
  }

  const pendingRaw = sessionStorage.getItem(PENDING_KEY);
  sessionStorage.removeItem(PENDING_KEY);
  if (!pendingRaw) {
    throw new Error("No pending sign-in found for this callback.");
  }
  const pending = JSON.parse(pendingRaw);
  if (pending.state !== state) {
    throw new Error("Sign-in state mismatch — please try again.");
  }

  const redirectUri = window.location.origin + window.location.pathname;
  const tokenResponse = await exchangeWithBackend({
    grantType: "authorization_code",
    code,
    codeVerifier: pending.verifier,
    redirectUri,
  });
  writeTokens(tokenResponse);
}

export function isSignedIn() {
  const tokens = readTokens();
  return Boolean(tokens?.refreshToken || (tokens?.accessToken && tokens.expiresAt > Date.now()));
}

export function signOut() {
  clearTokens();
}

// Shared by every concurrent caller currently refreshing — see getDataverseAccessToken's own WHY
// comment on the race this closes.
let refreshPromise = null;

// Returns a valid Dataverse access token for the already-signed-in account — refreshes silently
// via the backend when the cached one is expired or about to be, only throwing when there's
// nothing left to refresh from (caller should fall back to signIn()).
export async function getDataverseAccessToken() {
  await ensureInitialized();

  const tokens = readTokens();
  if (!tokens) {
    throw new Error("Not signed in.");
  }

  if (tokens.accessToken && tokens.expiresAt > Date.now() + EXPIRY_BUFFER_MS) {
    return tokens.accessToken;
  }

  if (!tokens.refreshToken) {
    throw new Error("Not signed in.");
  }

  // Concurrent callers (e.g. Promise.all([getDimensions(), getMeasures()]) in PivotBuilder,
  // both needing a token at once) must share ONE in-flight refresh, not each fire their own —
  // two independent refresh_token exchanges racing against the SAME stored refresh token can
  // both read it before either writes back, and if the token endpoint rotates the refresh token
  // (routine for Azure AD), whichever exchange lands second is using an already-invalidated
  // refresh token and fails outright. Caught live, 2026-09-27: dimensions/measures silently
  // never loaded after a long session had let the cached token expire, with no visible error
  // anywhere — exactly this race.
  if (!refreshPromise) {
    refreshPromise = exchangeWithBackend({
      grantType: "refresh_token",
      refreshToken: tokens.refreshToken,
    })
      .then((tokenResponse) => writeTokens(tokenResponse))
      .finally(() => {
        refreshPromise = null;
      });
  }
  const refreshed = await refreshPromise;
  return refreshed.accessToken;
}
