import { HashRouter, Routes, Route } from "react-router-dom";
import PivotBuilder from "./pages/PivotBuilder";
import SnapshotHistory from "./pages/SnapshotHistory";
import CompareSnapshots from "./pages/CompareSnapshots";

// HashRouter (not BrowserRouter): this is served as a static bundle with no server-side routing
// configured, so a path-based router would 404 on refresh/direct navigation.
function App() {
  return (
    <HashRouter>
      <Routes>
        <Route path="/" element={<PivotBuilder />} />
        <Route path="/history" element={<SnapshotHistory />} />
        <Route path="/compare" element={<CompareSnapshots />} />
      </Routes>
    </HashRouter>
  );
}

export default App;
