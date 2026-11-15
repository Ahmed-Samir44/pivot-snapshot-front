import { lazy, Suspense } from "react";
import { HashRouter, Routes, Route } from "react-router-dom";
import SignInGate from "./components/layout/SignInGate";
import ErrorBoundary from "./components/layout/ErrorBoundary";

// Lazy, not a static import — each page pulls in its own large, page-specific dependency tree
// (FieldPicker.jsx/PivotGrid.jsx for the Builder; nothing nearly as heavy for History) that the
// OTHER page never uses. A static import put both in the one main bundle regardless of which page
// you actually opened first — confirmed as the single biggest chunk of that bundle's size (perf
// investigation, 2026-11-09: "عايز احسن واسرع السيستم كله").
const PivotBuilder = lazy(() => import("./pages/PivotBuilder"));
const SnapshotHistory = lazy(() => import("./pages/SnapshotHistory"));
const KpiLinks = lazy(() => import("./pages/KpiLinks"));

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
          <Suspense fallback={null}>
            <Routes>
              <Route path="/" element={<PivotBuilder />} />
              <Route path="/history" element={<SnapshotHistory />} />
              <Route path="/kpi-links" element={<KpiLinks />} />
            </Routes>
          </Suspense>
        </ErrorBoundary>
      </HashRouter>
    </SignInGate>
  );
}

export default App;
