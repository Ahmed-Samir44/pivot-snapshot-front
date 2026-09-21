export default function PivotGrid({ result }) {
  if (!result) {
    return null;
  }

  const { rowHeaders, columnHeaders, cells } = result;
  // rowHeaders/columnHeaders are PivotHeader objects ({ labels, isTotal }), one per dimension
  // level tuple. isTotal marks a Grand Total row/column (see PivotHeader in the backend) so it
  // can be styled distinctly, the way Excel bolds/shades its Grand Total row and column.
  const rowDepth = rowHeaders[0]?.labels.length ?? 0;
  const columnDepth = columnHeaders[0]?.labels.length ?? 0;

  const totalCellClass = "bg-gold/20 font-bold";

  return (
    <div className="card overflow-x-auto">
      <table className="w-full border-collapse text-sm">
        <thead>
          {Array.from({ length: columnDepth }).map((_, level) => (
            <tr key={level}>
              {level === 0 && <th colSpan={rowDepth} rowSpan={columnDepth} className="border border-slate-200" />}
              {columnHeaders.map((col, colIndex) => (
                <th
                  key={colIndex}
                  className={`border border-slate-200 bg-gold/10 px-3 py-2 text-ink ${col.isTotal ? totalCellClass : ""}`}
                >
                  {col.labels[level]}
                </th>
              ))}
            </tr>
          ))}
        </thead>
        <tbody>
          {rowHeaders.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.labels.map((label, level) => (
                <th
                  key={level}
                  scope="row"
                  className={`border border-slate-200 bg-gold/10 px-3 py-2 text-left text-ink ${row.isTotal ? totalCellClass : ""}`}
                >
                  {label}
                </th>
              ))}
              {cells[rowIndex].map((cell, colIndex) => (
                <td
                  key={colIndex}
                  className={`border border-slate-200 px-3 py-2 text-right ${
                    row.isTotal || columnHeaders[colIndex].isTotal ? totalCellClass : ""
                  }`}
                >
                  {cell ?? ""}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
