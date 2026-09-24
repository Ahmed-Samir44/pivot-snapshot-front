import { useEffect, useState } from "react";
import { ensureInitialized, isSignedIn, signIn } from "../../services/dataverseAuth";

// Nothing under this component renders (so no API call can fire) until the user has actually
// signed in — see dataverseAuth.js for why: this is a full-tab redirect to Microsoft and back
// (not a popup), so "ready" here means "we've checked whether this page load IS that redirect
// coming back and, if so, finished the token exchange" before deciding whether to show the app
// or the sign-in button.
export default function SignInGate({ children }) {
  const [ready, setReady] = useState(false);
  const [signedIn, setSignedIn] = useState(false);
  const [signingIn, setSigningIn] = useState(false);
  const [error, setError] = useState(null);

  useEffect(() => {
    ensureInitialized()
      .then(() => setSignedIn(isSignedIn()))
      .catch((err) => setError(err.message || "Sign-in failed."))
      .finally(() => setReady(true));
  }, []);

  const handleSignIn = async () => {
    setSigningIn(true);
    setError(null);
    try {
      // Navigates the whole tab away — nothing after this line runs until the browser lands
      // back here on the next page load.
      await signIn();
    } catch (err) {
      setError(err.message || "Sign-in failed.");
      setSigningIn(false);
    }
  };

  if (!ready) {
    return null;
  }

  if (!signedIn) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-gradient-to-br from-slate-50 via-blue-50 to-slate-100 px-4">
        <div className="card w-full max-w-sm text-center">
          <h1 className="mb-2 text-xl font-bold text-ink">Pivot Snapshot Builder</h1>
          <p className="mb-4 text-sm text-muted">
            This tool queries cube data and saves to Dataverse as you — sign in with your Microsoft account to continue.
          </p>
          <button type="button" onClick={handleSignIn} disabled={signingIn} className="btn-primary w-full disabled:cursor-not-allowed disabled:opacity-50">
            {signingIn ? "Signing in…" : "Sign in with Microsoft"}
          </button>
          {error && <p className="mt-3 text-sm text-red-700">{error}</p>}
        </div>
      </div>
    );
  }

  return children;
}
