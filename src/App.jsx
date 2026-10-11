import { HashRouter, Routes, Route } from "react-router-dom";
import PivotBuilder from "./pages/PivotBuilder";
import SnapshotHistory from "./pages/SnapshotHistory";
import SignInGate from "./components/layout/SignInGate";
import ErrorBoundary from "./components/layout/ErrorBoundary";

// HashRouter (not BrowserRouter): this is served as a static bundle with no server-side routing
// configured, so a path-based router would 404 on refresh/direct navigation.
//
// SignInGate wraps EVERYTHING — every page queries the cube and/or Dataverse, both of which now
// require a signed-in Microsoft account (see DECISIONS.md, 2026-09-23) — nothing underneath it
// renders, so no page can fire an API call, until the user has actually signed in.
//
// ErrorBoundary wraps the routed pages (not SignInGate itself, which has nothing to render yet
// anyway before auth completes) — see its own WHY comment for the exact silent-blank-page crash
// this exists to catch instead.
function App() {
  return (
    <SignInGate>
      <HashRouter>
        <ErrorBoundary>
          <Routes>
            <Route path="/" element={<PivotBuilder />} />
            <Route path="/history" element={<SnapshotHistory />} />
          </Routes>
        </ErrorBoundary>
      </HashRouter>
    </SignInGate>
  );
}

export default App;
