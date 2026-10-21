import { useEffect, useRef, useState } from "react";

// Excel's Undo / Redo for the PivotTable (2026-10-21, "the system must be as excel"): every change
// to `state` (a plain object of everything that describes the pivot) is remembered; undo() hands the
// previous version to `apply`, redo() the next. A change made BY undo/redo isn't recorded again.
// `limit` versions are kept, like Excel's own undo depth.
export function useUndoHistory(state, apply, { limit = 100 } = {}) {
  const [stacks, setStacks] = useState({ past: [], future: [] });
  const current = useRef({ key: JSON.stringify(state), value: state });
  const applying = useRef(false);
  const key = JSON.stringify(state);

  useEffect(() => {
    if (key === current.current.key) return;
    const previous = current.current.value;
    current.current = { key, value: state };
    if (applying.current) {
      applying.current = false;
      return;
    }
    setStacks((s) => ({ past: [...s.past, previous].slice(-limit), future: [] }));
    // `state` is represented by `key`; re-running on the object identity alone would record nothing new.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key, limit]);

  const undo = () => {
    if (stacks.past.length === 0) return;
    const previous = stacks.past[stacks.past.length - 1];
    // Only flag the next change as ours when there will BE one (else the flag would swallow the
    // user's next real edit).
    if (JSON.stringify(previous) === current.current.key) {
      setStacks((s) => ({ ...s, past: s.past.slice(0, -1) }));
      return;
    }
    setStacks((s) => ({ past: s.past.slice(0, -1), future: [current.current.value, ...s.future] }));
    applying.current = true;
    apply(previous);
  };

  const redo = () => {
    if (stacks.future.length === 0) return;
    const next = stacks.future[0];
    if (JSON.stringify(next) === current.current.key) {
      setStacks((s) => ({ ...s, future: s.future.slice(1) }));
      return;
    }
    setStacks((s) => ({ past: [...s.past, current.current.value], future: s.future.slice(1) }));
    applying.current = true;
    apply(next);
  };

  return { undo, redo, canUndo: stacks.past.length > 0, canRedo: stacks.future.length > 0 };
}
