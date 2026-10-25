// The app's frame. 2026-10-20 ("like excel but more modern"): Excel for the web's chrome — a slim top
// bar (app name, the two screens as tabs, page actions where Excel has Share, the account menu)
// instead of the old 256px left sidebar, so the worksheet gets the full width.
import { useEffect, useRef, useState } from "react";
import { Link, useLocation } from "react-router-dom";
import { Activity, LogOut } from "lucide-react";
import { getSignedInUser, signOut } from "../../services/dataverseAuth";

// short: the label on a phone, where both full names don't fit beside the actions (2026-10-25).
const NAV_ITEMS = [
  { path: "/", label: "Pivot Builder", short: "Builder" },
  { path: "/history", label: "Snapshot History", short: "History" },
];

function initialsOf(user) {
  const source = user.name || user.email || "";
  const parts = source.replace(/@.*/, "").split(/[\s._-]+/).filter(Boolean);
  return ((parts[0]?.[0] ?? "") + (parts[1]?.[0] ?? "")).toUpperCase() || "?";
}

// The signed-in user's initials; opens their name/email and Sign Out.
function AccountMenu() {
  const [open, setOpen] = useState(false);
  const [user] = useState(getSignedInUser);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return undefined;
    const close = (e) => {
      if (!ref.current?.contains(e.target)) setOpen(false);
    };
    const onKey = (e) => e.key === "Escape" && setOpen(false);
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", close);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  // A plain reload (not a route change) so SignInGate re-checks from scratch and shows sign-in
  // (2026-10-06: the only way to switch accounts).
  const handleSignOut = () => {
    signOut();
    window.location.reload();
  };

  return (
    <div className="relative" ref={ref}>
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        aria-haspopup="menu"
        aria-expanded={open}
        title={user.name || user.email || "Account"}
        className="flex h-8 w-8 items-center justify-center rounded-full bg-ink text-xs font-semibold text-white ring-offset-2 transition hover:ring-2 hover:ring-primary/50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary"
      >
        {initialsOf(user)}
      </button>
      {open && (
        <div role="menu" className="absolute right-0 top-10 z-[60] w-64 rounded-lg border border-slate-200 bg-white p-1 text-sm shadow-lg">
          {(user.name || user.email) && (
            <div className="border-b border-slate-100 px-3 py-2">
              {user.name && <p className="truncate font-semibold text-ink">{user.name}</p>}
              {user.email && <p className="truncate text-xs text-muted">{user.email}</p>}
            </div>
          )}
          <button type="button" role="menuitem" onClick={handleSignOut} className="mt-1 flex w-full items-center gap-2 rounded-md px-3 py-2 text-left text-slate-700 hover:bg-slate-100">
            <LogOut className="h-4 w-4" />
            Sign Out
          </button>
        </div>
      )}
    </div>
  );
}

// workspace: an app-like page that owns the whole viewport (the Pivot Builder) — from the xl
// breakpoint up, <main> is exactly the space under the top bar and lays its children out as a
// column, so the grid flexes into the remaining height and scrolls inside itself. Below xl it falls
// back to the normal scrolling document flow.
// actions: the page's own buttons at the right of the top bar (Excel's Share spot) — the Builder's
// Export and Save Snapshot.
export default function AppShell({ children, workspace = false, actions = null }) {
  const location = useLocation();

  return (
    // ds-shell-bg: the design system's shell atmosphere (index.css), on the shell only.
    <div className="ds-shell-bg min-h-screen">
      {/* Narrow screens (2026-10-25): tighter spacing, the name hidden under the logo, tabs on one line —
          the bar used to push Export/Save off a phone screen and scroll the whole page sideways. */}
      <header className="sticky top-0 z-50 flex h-12 items-center gap-2 border-b border-slate-200 bg-white px-2 sm:gap-6 sm:px-4">
        <Link to="/" className="flex shrink-0 items-center gap-2.5" title="Pivot Snapshot">
          <span className="flex h-7 w-7 items-center justify-center rounded-control bg-gradient-to-br from-primary to-primary-deep">
            <Activity className="h-4 w-4 text-white" />
          </span>
          <span className="hidden text-[15px] font-bold tracking-tight text-ink md:inline">Pivot Snapshot</span>
        </Link>

        {/* The two screens as tabs; the active one carries the gold underline, like the ribbon's. */}
        <nav className="flex h-full items-stretch gap-1" aria-label="Main">
          {NAV_ITEMS.map(({ path, label, short }) => {
            const active = location.pathname === path;
            return (
              <Link
                key={path}
                to={path}
                aria-current={active ? "page" : undefined}
                className={`relative flex items-center whitespace-nowrap px-2 text-sm transition-colors sm:px-3 ${active ? "font-semibold text-ink" : "font-medium text-slate-500 hover:text-ink"}`}
              >
                <span className="sm:hidden">{short}</span>
                <span className="hidden sm:inline">{label}</span>
                {active && <span className="absolute inset-x-3 bottom-0 h-[2px] rounded-full bg-primary" aria-hidden="true" />}
              </Link>
            );
          })}
        </nav>

        <div className="ml-auto flex shrink-0 items-center gap-1.5 sm:gap-2">
          {actions}
          <span className="mx-1 hidden h-6 w-px bg-slate-200 sm:block" aria-hidden="true" />
          <AccountMenu />
        </div>
      </header>

      {/* A workspace page: tight padding and, from lg up, exactly the height under the top bar (the
          grid fills what's left — no page scrollbar; lg since 2026-10-25). A document page (History): normal padding,
          the page scrolls. */}
      <main className={workspace ? "flex flex-col gap-2 p-2 lg:h-[calc(100vh-3rem)] lg:overflow-hidden" : "mx-auto max-w-[1600px] p-6"}>{children}</main>
    </div>
  );
}
