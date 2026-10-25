import { formatNumber } from "../../utils/numberFormat";
import { measuresPlacementSupportsChart } from "../../utils/chartLayout";

// Live-preview mirror of the backend's PivotChartRenderer.cs (same three chart types, same
// first-value-field-only scope, same multi-series-from-a-Columns-field support added 2026-09-24)
// — actual <svg> elements, not an HTML canvas or charting library, so what the user sees here is
// what the saved snapshot will actually contain.
const WIDTH = 640;
const HEIGHT = 320;
const PADDING_LEFT = 40;
const PADDING_BOTTOM = 60;
const PADDING_TOP = 20;
// Reserved above PADDING_TOP for a legend row, only when there's more than one series — see
// PivotChartRenderer.cs's own comment on the same constant.
const LEGEND_ROW_HEIGHT = 16;
const BAR_GAP = 8;
// Value labels on bars/points only up to this many categories — past it they overlap into an
// unreadable band (caught 2026-10-18 with 30 bars). Excel shows none by default. Same constant in
// PivotChartRenderer.cs.
const MAX_VALUE_LABELS = 12;
// Must stay in sync with PivotChartRenderer.cs's identical array (saved snapshot vs. live preview).
// Synapse/HMIS teal scale alternating with its neutrals (2026-10-25, replacing the gold set), so
// neighbouring slices/series stay distinguishable.
const SLICE_COLORS = ["#126e5c", "#64748b", "#4bb69a", "#0a463c", "#82d2ba", "#94a3b8", "#278f76", "#334155", "#b7e7d7", "#cbd5e1"];

// One CATEGORY per detail row, named by its whole path ("Egypt / ER" — 2026-10-18: it used to be
// labels[0] alone, so with 2+ Rows fields every bar read "Egypt"). One SERIES per detail COLUMN — every value field and
// every Columns item (2026-10-15: it used to chart the FIRST value field only, and needed Σ Values
// innermost in Columns), named by its labels ("2024 / Revenue"); a lone series needs no name.
// Total columns are left out the same way total rows are. Mirrors PivotChartRenderer.ExtractSeries.
function extractSeries(result) {
  const categoryRows = result.rowHeaders.map((row, r) => (row.isTotal ? null : { row, r })).filter(Boolean);
  if (categoryRows.length === 0) return null;

  const categories = categoryRows.map(({ row }) => row.labels.filter((l) => l !== "").join(" / "));

  // Detail columns only — every total column (Grand Total, and since 2026-10-15 the "2024 Total"
  // column subtotals) is left out, same as PivotChartRenderer.ExtractSeries.
  const detailColumns = result.columnHeaders.map((h, c) => (h.isTotal ? null : c)).filter((c) => c !== null);
  if (detailColumns.length === 0) return null;
  const series = detailColumns.map((colIndex) => ({
    label: detailColumns.length === 1 ? "" : result.columnHeaders[colIndex].labels.filter((l) => l !== "").join(" / "),
    // An error cell (NaN) plots as 0, like Excel's charts do for #DIV/0! (mirrors PivotChartRenderer).
    values: categoryRows.map(({ r }) => (Number.isNaN(result.cells[r][colIndex]) ? 0 : (result.cells[r][colIndex] ?? 0))),
  }));

  return { categories, series };
}

function CategoryLabel({ x, label }) {
  return (
    <text x={x} y={HEIGHT - PADDING_BOTTOM + 14} textAnchor="end" fill="#6a7380" transform={`rotate(-40 ${x} ${HEIGHT - PADDING_BOTTOM + 14})`}>
      {label}
    </text>
  );
}

function ValueLabel({ x, y, value, format }) {
  return (
    <text x={x} y={y - 6} textAnchor="middle" fill="#1f2430">
      {formatNumber(value, format)}
    </text>
  );
}

// Swatch + label per series, left to right, WRAPPING onto a new row when the next one would run
// past the right edge (2026-10-15 — every value field is a series now, so legends got long enough to
// be cut off). Mirrors PivotChartRenderer.LegendLayout exactly. Returns each entry's { x, row }.
function legendLayout(series) {
  const items = [];
  let x = PADDING_LEFT;
  let row = 0;
  series.forEach((s) => {
    const width = 13 + s.label.length * 6 + 14;
    if (x > PADDING_LEFT && x + width > WIDTH - 10) {
      x = PADDING_LEFT;
      row += 1;
    }
    items.push({ x, row });
    x += width;
  });
  return { items, rows: series.length > 1 ? row + 1 : 0 };
}

function SeriesLegend({ series }) {
  const { items } = legendLayout(series);
  return (
    <>
      {series.map((s, i) => (
        <g key={i}>
          <rect x={items[i].x} y={10 + items[i].row * LEGEND_ROW_HEIGHT - 8} width={9} height={9} fill={SLICE_COLORS[i % SLICE_COLORS.length]} />
          <text x={items[i].x + 13} y={10 + items[i].row * LEGEND_ROW_HEIGHT} fill="#1f2430">
            {s.label}
          </text>
        </g>
      ))}
    </>
  );
}

function BarChart({ data, format }) {
  const max = Math.max(...data.series.flatMap((s) => s.values));
  if (max <= 0) return null;

  const paddingTop = PADDING_TOP + legendLayout(data.series).rows * LEGEND_ROW_HEIGHT;
  const chartWidth = WIDTH - PADDING_LEFT - 10;
  const chartHeight = HEIGHT - paddingTop - PADDING_BOTTOM;
  const slotWidth = chartWidth / data.categories.length;
  const barWidth = (slotWidth - BAR_GAP) / data.series.length;

  return (
    <>
      <line x1={PADDING_LEFT} y1={HEIGHT - PADDING_BOTTOM} x2={WIDTH - 10} y2={HEIGHT - PADDING_BOTTOM} stroke="#d8d5dd" />
      {data.categories.map((label, c) => {
        const groupX = PADDING_LEFT + c * slotWidth + BAR_GAP / 2;
        return (
          <g key={c}>
            {data.series.map((s, si) => {
              const value = s.values[c];
              const barHeight = value <= 0 ? 0 : (value / max) * chartHeight;
              const x = groupX + si * barWidth;
              const y = HEIGHT - PADDING_BOTTOM - barHeight;
              return (
                <g key={si}>
                  <rect x={x} y={y} width={Math.max(barWidth, 1)} height={barHeight} fill={SLICE_COLORS[si % SLICE_COLORS.length]} />
                  {data.series.length === 1 && data.categories.length <= MAX_VALUE_LABELS && <ValueLabel x={x + barWidth / 2} y={y} value={value} format={format} />}
                </g>
              );
            })}
            <CategoryLabel x={groupX + (slotWidth - BAR_GAP) / 2} label={label} />
          </g>
        );
      })}
      {data.series.length > 1 && <SeriesLegend series={data.series} />}
    </>
  );
}

function LineChart({ data, format }) {
  const max = Math.max(...data.series.flatMap((s) => s.values));
  if (max <= 0) return null;

  const paddingTop = PADDING_TOP + legendLayout(data.series).rows * LEGEND_ROW_HEIGHT;
  const chartWidth = WIDTH - PADDING_LEFT - 10;
  const chartHeight = HEIGHT - paddingTop - PADDING_BOTTOM;
  const slotWidth = chartWidth / data.categories.length;

  return (
    <>
      <line x1={PADDING_LEFT} y1={HEIGHT - PADDING_BOTTOM} x2={WIDTH - 10} y2={HEIGHT - PADDING_BOTTOM} stroke="#d8d5dd" />
      {data.series.map((s, si) => {
        const color = SLICE_COLORS[si % SLICE_COLORS.length];
        const points = s.values.map((value, i) => ({
          x: PADDING_LEFT + i * slotWidth + slotWidth / 2,
          y: HEIGHT - PADDING_BOTTOM - (value <= 0 ? 0 : (value / max) * chartHeight),
          value,
        }));
        return (
          <g key={si}>
            <polyline points={points.map((p) => `${p.x},${p.y}`).join(" ")} fill="none" stroke={color} strokeWidth={2} />
            {points.map((p, i) => (
              <g key={i}>
                <circle cx={p.x} cy={p.y} r={3} fill={color} />
                {data.series.length === 1 && data.categories.length <= MAX_VALUE_LABELS && <ValueLabel x={p.x} y={p.y} value={p.value} format={format} />}
              </g>
            ))}
          </g>
        );
      })}
      {data.categories.map((label, c) => (
        <CategoryLabel key={c} x={PADDING_LEFT + c * slotWidth + slotWidth / 2} label={label} />
      ))}
      {data.series.length > 1 && <SeriesLegend series={data.series} />}
    </>
  );
}

// Pie stays single-series (Excel's own pivot chart behaves the same way) — uses just the FIRST
// series when a Columns field would otherwise produce several, same as before 2026-09-24.
function PieChart({ data }) {
  const bars = data.categories.map((label, c) => ({ label, value: data.series[0].values[c] })).filter((b) => b.value > 0);
  const total = bars.reduce((sum, b) => sum + b.value, 0);
  if (bars.length === 0 || total <= 0) return null;

  const cx = WIDTH / 4;
  const cy = HEIGHT / 2;
  const r = Math.min(cy, cx) - 20;
  const legendX = WIDTH / 2 + 20;
  const legendY = PADDING_TOP + 10;

  // A single slice can't be drawn as an SVG arc (start/end angle coincide) — a plain circle
  // instead, same as the backend.
  if (bars.length === 1) {
    return (
      <>
        <circle cx={cx} cy={cy} r={r} fill={SLICE_COLORS[0]} />
        <rect x={legendX} y={legendY - 9} width={10} height={10} fill={SLICE_COLORS[0]} />
        <text x={legendX + 16} y={legendY} fill="#1f2430">
          {bars[0].label} (100.0%)
        </text>
      </>
    );
  }

  // Cumulative sweep-so-far per slice, computed without a mutated accumulator variable —
  // startAngles[i] is where slice i begins, derived purely from the slices before it.
  const startAngles = bars.reduce((acc, bar) => [...acc, acc[acc.length - 1] + (bar.value / total) * 2 * Math.PI], [-Math.PI / 2]);

  const slices = bars.map((bar, i) => {
    const startAngle = startAngles[i];
    const sweep = (bar.value / total) * 2 * Math.PI;
    const endAngle = startAngle + sweep;
    const x1 = cx + r * Math.cos(startAngle);
    const y1 = cy + r * Math.sin(startAngle);
    const x2 = cx + r * Math.cos(endAngle);
    const y2 = cy + r * Math.sin(endAngle);
    const largeArcFlag = sweep > Math.PI ? 1 : 0;
    const path = `M ${cx},${cy} L ${x1},${y1} A ${r},${r} 0 ${largeArcFlag} 1 ${x2},${y2} Z`;
    return { path, color: SLICE_COLORS[i % SLICE_COLORS.length], label: bar.label, percent: (bar.value / total) * 100 };
  });

  return (
    <>
      {slices.map((s, i) => (
        <path key={i} d={s.path} fill={s.color} />
      ))}
      {slices.map((s, i) => (
        <g key={i}>
          <rect x={legendX} y={legendY + i * 18 - 9} width={10} height={10} fill={s.color} />
          <text x={legendX + 16} y={legendY + i * 18} fill="#1f2430">
            {s.label} ({s.percent.toFixed(1)}%)
          </text>
        </g>
      ))}
    </>
  );
}

export default function PivotChart({ result, valueFields, chartType = "Bar", rows }) {
  // Mirrors the backend's PivotChartRenderer.Render: categories are rows and every detail column is a
  // series, wherever Σ Values sits in Columns — with Σ Values in Rows (each row a measure) there's no chart.
  if (!result || !valueFields?.length || result.columnHeaders.length === 0 || !measuresPlacementSupportsChart(rows)) {
    return null;
  }

  const data = extractSeries(result);
  if (!data) {
    return null;
  }

  const format = valueFields[0]?.format;

  return (
    <div className="card mt-6 overflow-x-auto">
      <svg width={WIDTH} height={HEIGHT} viewBox={`0 0 ${WIDTH} ${HEIGHT}`} style={{ fontFamily: "Segoe UI,Arial,sans-serif", fontSize: 11 }}>
        {chartType === "Line" && <LineChart data={data} format={format} />}
        {chartType === "Pie" && <PieChart data={data} />}
        {chartType !== "Line" && chartType !== "Pie" && <BarChart data={data} format={format} />}
      </svg>
    </div>
  );
}
