// Shared native-HTML5-DnD payload (no library) for moving a field between any of the pivot
// builder's zones — Filters, Rows, Columns, Values, and the Available Fields source list.
// dataTransfer carries a plain JSON string across the browser's DOM event system, which is what
// lets completely separate React component instances (each zone owns its own state slice)
// coordinate a move without any shared component tree above FieldPicker.
export const FIELD_DRAG_MIME = "application/x-pivot-field";

export function setFieldDragPayload(event, field, sourceZone) {
  event.dataTransfer.setData(FIELD_DRAG_MIME, JSON.stringify({ field, source: sourceZone }));
  event.dataTransfer.effectAllowed = "move";
}

export function readFieldDragPayload(event) {
  const raw = event.dataTransfer.getData(FIELD_DRAG_MIME);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}
