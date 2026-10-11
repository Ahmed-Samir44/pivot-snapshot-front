// Adapted from the Segmentation project's DashboardLayout — same sidebar shell and gold/ink
// branding, now with real routes (Pivot Builder / Snapshot History).
import { useState } from "react";
import { Link, useLocation } from "react-router-dom";
import { Activity, LayoutDashboard, History, ChevronsLeft, ChevronsRight, LogOut } from "lucide-react";
import { signOut } from "../../services/dataverseAuth";

const NAV_ITEMS = [
  { path: "/", label: "Pivot Builder", icon: LayoutDashboard },
  { path: "/history", label: "Snapshot History", icon: History },
];

const COLLAPSED_STORAGE_KEY = "pivotSnapshot.sidebarCollapsed";

// Exported so Modal.jsx can both read the INITIAL persisted sidebar state and subscribe to LIVE
// toggles — its backdrop needs to know the sidebar's current width (w-16 collapsed vs w-64
// expanded) to avoid covering only part of it, same reason its own left offset exists at all (see
// Modal.jsx's own WHY comment).
export const SIDEBAR_COLLAPSED_EVENT = "pivotSnapshot:sidebarCollapsedChanged";

export function readCollapsed() {
  try {
    return localStorage.getItem(COLLAPSED_STORAGE_KEY) === "true";
  } catch {
    // Per-browser convenience only — nothing breaks if this can't be read (private window,
    // storage disabled); the sidebar just starts expanded.
    return false;
  }
}

export default function AppShell({ children }) {
  const location = useLocation();
  const [collapsed, setCollapsed] = useState(readCollapsed);

  // A plain reload (not a client-side route change) so SignInGate re-runs its isSignedIn() check
  // from scratch and shows the sign-in screen — same effect the password-change auto-recovery in
  // dataverseAuth.js's getDataverseAccessToken relies on, just user-initiated here instead of
  // triggered by a failed token refresh. There was no way to switch accounts/sign out at all
  // before this (caught live, 2026-10-06, after hitting the auto-recovery path twice from a
  // password reset and asking for a manual way to do the same thing).
  const handleSignOut = () => {
    signOut();
    window.location.reload();
  };

  const toggleCollapsed = () => {
    // Computed outside the setCollapsed updater (not read from `collapsed` directly either, to
    // stay correct under React's automatic batching) — dispatching the event FROM INSIDE the
    // updater ran it during React's render/commit of THIS component, which is exactly the "update
    // a different component (Modal) while rendering a different component (AppShell)" misuse React
    // warns about (caught live 2026-10-07, right after adding this event). A plain side effect
    // after the state update call, in the event handler itself, is the correct place for it.
    const next = !collapsed;
    setCollapsed(next);
    try {
      localStorage.setItem(COLLAPSED_STORAGE_KEY, String(next));
    } catch {
      // Same no-op-on-failure as readCollapsed above.
    }
    // The native "storage" event only fires in OTHER tabs/windows, never the one that made the
    // change — Modal.jsx needs to react to a toggle in THIS same tab too (e.g. the sidebar
    // collapsed/expanded while a Modal was already open), so a plain custom event on `window`
    // covers the same-tab case the storage event can't (see Modal.jsx's own WHY comment).
    window.dispatchEvent(new CustomEvent(SIDEBAR_COLLAPSED_EVENT, { detail: next }));
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-blue-50 to-slate-100">
      <aside
        className={`fixed left-0 top-0 z-50 h-full glass-panel border-r border-white/50 transition-[width] duration-200 ${
          collapsed ? "w-16" : "w-64"
        }`}
      >
        <div className={`flex items-center border-b border-slate-200 p-4 ${collapsed ? "justify-center" : "justify-between gap-3"}`}>
          <div className={`flex items-center gap-3 ${collapsed ? "" : "min-w-0"}`}>
            <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-gold to-gold-strong">
              <Activity className="h-6 w-6 text-white" />
            </div>
            {!collapsed && (
              <div className="min-w-0">
                <h1 className="truncate text-xl font-bold text-ink">Pivot Snapshot</h1>
                <p className="text-xs text-muted">Builder</p>
              </div>
            )}
          </div>
          {!collapsed && (
            <button
              type="button"
              onClick={toggleCollapsed}
              className="shrink-0 rounded-lg p-1.5 text-muted hover:bg-white/50 hover:text-ink"
              aria-label="Collapse sidebar"
              title="Collapse sidebar"
            >
              <ChevronsLeft className="h-5 w-5" />
            </button>
          )}
        </div>

        {collapsed && (
          <button
            type="button"
            onClick={toggleCollapsed}
            className="mx-auto mt-2 flex items-center justify-center rounded-lg p-1.5 text-muted hover:bg-white/50 hover:text-ink"
            aria-label="Expand sidebar"
            title="Expand sidebar"
          >
            <ChevronsRight className="h-5 w-5" />
          </button>
        )}

        <nav className={`space-y-2 p-4 ${collapsed ? "px-2" : ""}`}>
          {NAV_ITEMS.map(({ path, label, icon: Icon }) => {
            const active = location.pathname === path;
            return (
              <Link
                key={path}
                to={path}
                title={collapsed ? label : undefined}
                className={`flex items-center gap-3 rounded-xl py-3 transition-all duration-200 ${collapsed ? "justify-center px-0" : "px-4"} ${
                  active ? "bg-gradient-to-r from-gold to-gold-strong text-white shadow-lg" : "text-ink hover:bg-white/50"
                }`}
              >
                <Icon className="h-5 w-5 shrink-0" />
                {!collapsed && <span className="font-semibold">{label}</span>}
              </Link>
            );
          })}
          <button
            type="button"
            onClick={handleSignOut}
            title={collapsed ? "Sign Out" : undefined}
            className={`flex w-full items-center gap-3 rounded-xl py-3 text-ink transition-all duration-200 hover:bg-white/50 ${collapsed ? "justify-center px-0" : "px-4"}`}
          >
            <LogOut className="h-5 w-5 shrink-0" />
            {!collapsed && <span className="font-semibold">Sign Out</span>}
          </button>
        </nav>

        {!collapsed && (
          <div className="absolute bottom-0 left-0 right-0 border-t border-slate-200 p-4">
            <p className="text-center text-xs text-muted">© {new Date().getFullYear()} Andalusia Group</p>
          </div>
        )}
      </aside>

      <main className={`p-8 transition-[margin] duration-200 ${collapsed ? "ml-16" : "ml-64"}`}>{children}</main>
    </div>
  );
}
