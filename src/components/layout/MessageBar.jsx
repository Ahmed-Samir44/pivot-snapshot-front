import { AlertTriangle, Info, XCircle } from "lucide-react";

const TONES = {
  info: { icon: Info, bar: "border-sky-200 bg-sky-50 text-sky-900", iconColor: "text-sky-600" },
  warning: { icon: AlertTriangle, bar: "border-amber-200 bg-amber-50 text-amber-900", iconColor: "text-amber-600" },
  error: { icon: XCircle, bar: "border-red-200 bg-red-50 text-red-900", iconColor: "text-red-600" },
};

// Excel's message bar (the strip under the ribbon that says e.g. "PROTECTED VIEW … Enable
// Editing"): one slim line with an icon, the message and an optional action button — replaces the
// tall padded banners (2026-10-20, "like excel but more modern").
export default function MessageBar({ tone = "info", action = null, children }) {
  const { icon: Icon, bar, iconColor } = TONES[tone] ?? TONES.info;
  return (
    <div role={tone === "error" ? "alert" : "status"} className={`flex min-h-9 shrink-0 items-center gap-2.5 rounded-md border px-3 py-1.5 text-sm ${bar}`}>
      <Icon className={`h-4 w-4 shrink-0 ${iconColor}`} />
      <p className="min-w-0 flex-1 truncate" title={typeof children === "string" ? children : undefined}>
        {children}
      </p>
      {action && (
        <button
          type="button"
          onClick={action.onClick}
          disabled={action.disabled}
          className="shrink-0 rounded border border-current/20 bg-white/70 px-2.5 py-0.5 text-xs font-semibold hover:bg-white disabled:cursor-not-allowed disabled:opacity-50"
        >
          {action.label}
        </button>
      )}
    </div>
  );
}
