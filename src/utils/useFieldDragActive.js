import { useEffect, useState } from "react";
import { FIELD_DRAG_MIME } from "./dragDrop";

// True while a pivot field is being dragged anywhere on the page — every area (Filters, Columns,
// Rows, Values) lights up as a target the moment a drag starts, not only once the pointer is over it
// (2026-10-20; shared by all four areas since 2026-10-21).
export function useFieldDragActive() {
  const [active, setActive] = useState(false);
  useEffect(() => {
    const start = (e) => {
      if (e.dataTransfer?.types?.includes(FIELD_DRAG_MIME)) setActive(true);
    };
    const stop = () => setActive(false);
    document.addEventListener("dragstart", start);
    document.addEventListener("dragend", stop);
    document.addEventListener("drop", stop);
    return () => {
      document.removeEventListener("dragstart", start);
      document.removeEventListener("dragend", stop);
      document.removeEventListener("drop", stop);
    };
  }, []);
  return active;
}
