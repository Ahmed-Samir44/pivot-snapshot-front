import { useEffect, useRef } from "react";
import { AlertTriangle, CheckCircle2, X, XCircle } from "lucide-react";

const TONES = {
  success: { icon: CheckCircle2, color: "text-emerald-600" },
  warning: { icon: AlertTriangle, color: "text-amber-600" },
  error: { icon: XCircle, color: "text-red-600" },
};

// A success hides itself after a few seconds; a warning or error stays until it's closed.
function Toast({ tone, text, onClose }) {
  const { icon: Icon, color } = TONES[tone] ?? TONES.success;
  // The timer belongs to the message, not to each re-render of the page (onClose is a new function
  // every render, which would keep restarting it).
  const onCloseRef = useRef(onClose);
  useEffect(() => {
    onCloseRef.current = onClose;
  });
  useEffect(() => {
    if (tone !== "success") return undefined;
    const timer = setTimeout(() => onCloseRef.current(), 5000);
    return () => clearTimeout(timer);
  }, [tone, text]);
  return (
    <div role={tone === "error" ? "alert" : "status"} className="pointer-events-auto flex w-80 items-start gap-2.5 rounded-lg border border-slate-200 bg-white px-3 py-2.5 text-sm text-slate-700 shadow-lg">
      <Icon className={`mt-0.5 h-4 w-4 shrink-0 ${color}`} />
      <p className="min-w-0 flex-1">{text}</p>
      <button type="button" onClick={onClose} className="-mr-1 shrink-0 rounded p-0.5 text-slate-400 hover:bg-slate-100 hover:text-ink" aria-label="Dismiss">
        <X className="h-3.5 w-3.5" />
      </button>
    </div>
  );
}

// Results of an action the user just took (Save Snapshot, Export) as small notices stacked at the
// bottom right, above the status bar (2026-10-20) — they used to be text lines in the save card.
export default function Toasts({ items }) {
  if (!items.length) return null;
  return (
    <div className="pointer-events-none fixed bottom-12 right-4 z-[55] flex flex-col items-end gap-2">
      {items.map((item) => (
        <Toast key={`${item.tone}:${item.text}`} {...item} />
      ))}
    </div>
  );
}
