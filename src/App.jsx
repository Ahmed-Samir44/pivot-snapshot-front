import { HashRouter, Routes, Route } from "react-router-dom";
import PivotBuilder from "./pages/PivotBuilder";
import SnapshotHistory from "./pages/SnapshotHistory";
import CompareSnapshots from "./pages/CompareSnapshots";
import SignInGate from "./components/layout/SignInGate";

// HashRouter (not BrowserRouter): this is served as a static bundle with no server-side routing
// configured, so a path-based router would 404 on refresh/direct navigation.
//
// SignInGate wraps EVERYTHING — every page queries the cube and/or Dataverse, both of which now
// require a signed-in Microsoft account (see DECISIONS.md, 2026-09-23) — nothing underneath it
// renders, so no page can fire an API call, until the user has actually signed in.
function App() {
  return (
    <SignInGate>
      <HashRouter>
        <Routes>
          <Route path="/" element={<PivotBuilder />} />
          <Route path="/history" element={<SnapshotHistory />} />
          <Route path="/compare" element={<CompareSnapshots />} />
        </Routes>
      </HashRouter>
    </SignInGate>
  );
}

export default App;
