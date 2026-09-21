export default function PivotGrid({ result }) {
  if (!result) {
    return null;
  }

  const { rowHeaders, columnHeaders, cells } = result;
  // rowHeaders/columnHeaders are lists of label tuples, one per dimension level (e.g. a row with
  // two row-fields produces ["Cairo", "Dr. Amin"]). Depth = how many header rows/columns to render.
  const rowDepth = rowHeaders[0]?.length ?? 0;
  const columnDepth = columnHeaders[0]?.length ?? 0;

  return (
    <table className="pivot-grid">
      <thead>
        {Array.from({ length: columnDepth }).map((_, level) => (
          <tr key={level}>
            {level === 0 && <th colSpan={rowDepth} rowSpan={columnDepth} />}
            {columnHeaders.map((col, colIndex) => (
              <th key={colIndex}>{col[level]}</th>
            ))}
          </tr>
        ))}
      </thead>
      <tbody>
        {rowHeaders.map((row, rowIndex) => (
          <tr key={rowIndex}>
            {row.map((label, level) => (
              <th key={level} scope="row">
                {label}
              </th>
            ))}
            {cells[rowIndex].map((cell, colIndex) => (
              <td key={colIndex}>{cell ?? ""}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
