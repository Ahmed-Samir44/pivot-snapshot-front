// A thin indeterminate sweep along a button's bottom edge while it's busy (2026-11-12, "عايز
// الزرار نفسه يبقى زي ال... progress bar" — the button itself, not just its label text, should
// read as "working"). There's no real percentage to report for a Save/Export request (just an
// awaited promise), so this is an indeterminate sweep, the same "we're working on it" shape Office
// itself uses for an in-progress command, not a determinate fill.
// The button needs `relative overflow-hidden` itself for this to clip correctly; `tone` picks a
// sweep color visible against that button's own background ("light" on a dark/primary button,
// "dark" on a white/secondary one).
export default function ButtonProgressBar({ active, tone = "light" }) {
  if (!active) return null;
  return (
    <span className="pointer-events-none absolute inset-x-0 bottom-0 h-0.5 overflow-hidden">
      <span className={`absolute inset-y-0 w-2/5 animate-progress-sweep ${tone === "light" ? "bg-white/80" : "bg-primary/70"}`} />
    </span>
  );
}
