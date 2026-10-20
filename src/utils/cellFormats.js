// Number formats set on individual PivotTable cells (Home › Number on a selection, 2026-10-20 —
// Excel formats just the selected cells; Value Field Settings › Number Format still formats a whole
// value field). A cell is identified by its row and column labels (not its position), so its format
// stays with it when the pivot re-queries or re-sorts, like Excel's "Preserve cell formatting".
export const cellFormatKey = (rowHeader, columnHeader) => JSON.stringify([rowHeader?.labels ?? [], columnHeader?.labels ?? []]);

export const hasCellFormats = (cellFormats) => Boolean(cellFormats) && Object.keys(cellFormats).length > 0;
