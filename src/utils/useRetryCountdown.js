import { useEffect, useRef, useState } from "react";

// While `active`, counts down from `seconds` and then calls onRetry (once per activation) — the
// data server's busy/unreachable states retry on their own instead of waiting for the user
// (2026-10-20). Returns the seconds left, or null when inactive. A new `key` (e.g. the next failed
// attempt) restarts the count.
export function useRetryCountdown(active, seconds, onRetry, key = 0) {
  // { run, left }: `run` identifies the activation the count belongs to, so a stale count from a
  // previous one is never shown.
  const [tick, setTick] = useState({ run: null, left: seconds });
  const onRetryRef = useRef(onRetry);
  useEffect(() => {
    onRetryRef.current = onRetry;
  });
  const run = active ? `${key}` : null;
  useEffect(() => {
    if (run === null) return undefined;
    const until = Date.now() + seconds * 1000;
    const timer = setInterval(() => {
      const left = Math.max(0, Math.ceil((until - Date.now()) / 1000));
      setTick({ run, left });
      if (left === 0) {
        clearInterval(timer);
        onRetryRef.current();
      }
    }, 1000);
    return () => clearInterval(timer);
  }, [run, seconds]);
  if (run === null) return null;
  return tick.run === run ? tick.left : seconds;
}
