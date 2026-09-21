export default function PivotGrid({ result }) {
  if (!result) {
    return null;
  }

  const { rowHeaders, columnHeaders, cells } = result;
  // rowHeaders/columnHeaders are PivotHeader objects ({ labels, isTotal }), one per dimension
  // level tuple (e.g. a row with two row-fields produces labels ["Cairo", "Dr. Amin"]). isTotal
  // marks a Grand Total row/column (see PivotHeader in the backend) so it can be styled distinctly,
  // the way Excel bolds/shades its Grand Total row and column.
  const rowDepth = rowHeaders[0]?.labels.length ?? 0;
  const columnDepth = columnHeaders[0]?.labels.length ?? 0;

  return (
    <table className="pivot-grid">
      <thead>
        {Array.from({ length: columnDepth }).map((_, level) => (
          <tr key={level}>
            {level === 0 && <th colSpan={rowDepth} rowSpan={columnDepth} />}
            {columnHeaders.map((col, colIndex) => (
              <th key={colIndex} className={col.isTotal ? "is-total" : undefined}>
                {col.labels[level]}
              </th>
            ))}
          </tr>
        ))}
      </thead>
      <tbody>
        {rowHeaders.map((row, rowIndex) => (
          <tr key={rowIndex} className={row.isTotal ? "is-total" : undefined}>
            {row.labels.map((label, level) => (
              <th key={level} scope="row" className={row.isTotal ? "is-total" : undefined}>
                {label}
              </th>
            ))}
            {cells[rowIndex].map((cell, colIndex) => (
              <td
                key={colIndex}
                className={row.isTotal || columnHeaders[colIndex].isTotal ? "is-total" : undefined}
              >
                {cell ?? ""}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
