// Adapted from the Segmentation project's DashboardLayout — same sidebar shell and gold/ink
// branding, but without react-router-dom: this app has a single page for now. Nav items below
// are placeholders for Phase C (Save Snapshot / History / Compare will become real routes once
// they exist — see DECISIONS.md) rather than working links.
import { Activity, LayoutDashboard, History, GitCompare } from "lucide-react";

const NAV_ITEMS = [
  { label: "Pivot Builder", icon: LayoutDashboard, active: true },
  { label: "Snapshot History", icon: History, active: false },
  { label: "Compare Snapshots", icon: GitCompare, active: false },
];

export default function AppShell({ children }) {
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-blue-50 to-slate-100">
      <aside className="fixed left-0 top-0 z-50 h-full w-64 glass-panel border-r border-white/50">
        <div className="border-b border-slate-200 p-6">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-gradient-to-br from-gold to-gold-strong">
              <Activity className="h-6 w-6 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-ink">Pivot Snapshot</h1>
              <p className="text-xs text-muted">Builder</p>
            </div>
          </div>
        </div>

        <nav className="space-y-2 p-4">
          {NAV_ITEMS.map(({ label, icon: Icon, active }) => (
            <div
              key={label}
              aria-disabled={!active}
              title={active ? undefined : "Coming in a later phase"}
              className={`flex items-center gap-3 rounded-xl px-4 py-3 transition-all duration-200 ${
                active ? "bg-gradient-to-r from-gold to-gold-strong text-white shadow-lg" : "text-muted opacity-60"
              }`}
            >
              <Icon className="h-5 w-5" />
              <span className="font-semibold">{label}</span>
            </div>
          ))}
        </nav>

        <div className="absolute bottom-0 left-0 right-0 border-t border-slate-200 p-4">
          <p className="text-center text-xs text-muted">© {new Date().getFullYear()} Andalusia Group</p>
        </div>
      </aside>

      <main className="ml-64 p-8">{children}</main>
    </div>
  );
}
