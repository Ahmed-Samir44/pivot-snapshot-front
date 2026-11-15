import { useEffect, useState } from "react";
import { Link } from "react-router-dom";
import { CalendarClock, Link2, Loader2, Search } from "lucide-react";
import AppShell from "../components/layout/AppShell";
import { listDefinitions } from "../services/snapshotApi";

// Mirrors SnapshotHistory.jsx's own schedule-hours parsing (PivotSchedulerBackgroundService.
// ParseHours's frontend counterpart) — kept as its own small copy rather than a shared import
// since this page only needs the DISPLAY label, not the editing logic that lives with the button.
function formatSchedule(interval) {
  if (!interval || interval === "Off") return null;
  if (interval === "Hourly") return "Hourly";
  if (interval === "Daily") return "Daily";
  if (interval === "Weekly") return "Weekly";
  const hours = Number(interval);
  if (!Number.isFinite(hours) || hours <= 0) return null;
  if (hours === 1) return "Hourly";
  if (hours === 24) return "Daily";
  if (hours === 168) return "Weekly";
  return `Every ${Number(hours.toFixed(2))}h`;
}

// A plain "which KPI is linked to which pivot" list (2026-11-15, "محتاج صفحة يبقى فيها اي kpi
// مربوط ب pivot") — read-only, no new data: every field here already comes back on
// listDefinitions() (LinkedKpiName/ScheduleInterval, added for the badge/button in
// SnapshotHistory.jsx). This just filters that same list down to the ones actually linked, instead
// of someone having to scroll the full Saved Pivots list hunting for the 🔗 badge.
export default function KpiLinks() {
  const [definitions, setDefinitions] = useState(null);
  const [error, setError] = useState(null);
  const [searchTerm, setSearchTerm] = useState("");

  useEffect(() => {
    listDefinitions()
      .then(setDefinitions)
      .catch((err) => setError(err.message));
  }, []);

  const linked = (definitions ?? []).filter((d) => d.linkedKpiName);
  const query = searchTerm.trim().toLowerCase();
  const visible = query
    ? linked.filter((d) => d.linkedKpiName.toLowerCase().includes(query) || d.name.toLowerCase().includes(query))
    : linked;

  return (
    <AppShell>
      <header className="mb-6">
        <h1 className="text-3xl font-bold text-ink">KPI Links</h1>
        <p className="mt-1 text-muted">Every saved pivot currently linked to a Strategy KPI.</p>
      </header>

      <div className="relative mb-6 max-w-md">
        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted" aria-hidden="true" />
        <input
          type="text"
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          placeholder="Search by KPI or pivot name…"
          className="w-full rounded-lg border border-slate-200 bg-white py-2 pl-9 pr-4 text-sm shadow-sm transition-colors focus:border-primary focus:outline-none focus:ring-2 focus:ring-primary/20"
        />
      </div>

      {error && <p className="mb-6 rounded-xl border border-red-300 bg-red-50 px-4 py-3 text-red-800">{error}</p>}

      <div className="rounded-2xl border border-slate-200/70 bg-white shadow-ds-lg">
        {definitions === null && !error && (
          <p className="flex items-center gap-2 px-5 py-8 text-sm text-muted">
            <Loader2 className="h-4 w-4 animate-spin" /> Loading…
          </p>
        )}
        {definitions !== null && linked.length === 0 && (
          <div className="flex flex-col items-center gap-2 px-5 py-14 text-center">
            <Link2 className="h-7 w-7 text-slate-300" aria-hidden="true" />
            <p className="text-sm font-medium text-ink">No pivots linked to a KPI yet</p>
            <p className="max-w-sm text-xs text-muted">Link a saved pivot to a KPI from Snapshot History (the 🔗 icon on any pivot) to see it here.</p>
          </div>
        )}
        {definitions !== null && linked.length > 0 && visible.length === 0 && (
          <p className="px-5 py-8 text-center text-sm text-muted">No matches.</p>
        )}
        {visible.length > 0 && (
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-slate-100 text-left text-xs font-bold uppercase tracking-wide text-muted">
                <th className="px-5 py-3">KPI</th>
                <th className="px-5 py-3">Pivot</th>
                <th className="px-5 py-3">Schedule</th>
                <th className="px-5 py-3">Created</th>
                <th className="px-5 py-3"></th>
              </tr>
            </thead>
            <tbody>
              {visible.map((d) => {
                const schedule = formatSchedule(d.scheduleInterval);
                return (
                  <tr key={d.id} className="border-b border-slate-50 last:border-0 hover:bg-slate-50">
                    <td className="px-5 py-3 font-medium text-ink">{d.linkedKpiName}</td>
                    <td className="px-5 py-3 text-ink">{d.name}</td>
                    <td className="px-5 py-3">
                      {schedule ? (
                        <span className="inline-flex items-center gap-1 text-xs font-medium text-primary-strong">
                          <CalendarClock className="h-3.5 w-3.5" />
                          {schedule}
                        </span>
                      ) : (
                        <span className="text-xs text-muted">Off</span>
                      )}
                    </td>
                    <td className="px-5 py-3 text-xs text-muted">{new Date(d.createdAtUtc).toLocaleDateString()}</td>
                    <td className="px-5 py-3 text-right">
                      <Link to={`/history?definitionId=${d.id}`} className="text-xs font-semibold text-primary-strong hover:underline">
                        View versions
                      </Link>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </AppShell>
  );
}
