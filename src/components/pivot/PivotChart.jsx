import { formatNumber } from "../../utils/numberFormat";
import { MAX_BARS } from "../../utils/chartLayout";

// Live-preview mirror of the backend's PivotChartRenderer.cs (same three chart types, same
// first-value-field-only scope, same MaxBars cap, same multi-series-from-a-Columns-field support
// added 2026-09-24) — actual <svg> elements, not an HTML canvas or charting library, so what the
// user sees here is what the saved snapshot will actually contain.
const WIDTH = 640;
const HEIGHT = 320;
const PADDING_LEFT = 40;
const PADDING_BOTTOM = 60;
const PADDING_TOP = 20;
// Reserved above PADDING_TOP for a legend row, only when there's more than one series — see
// PivotChartRenderer.cs's own comment on the same constant.
const LEGEND_ROW_HEIGHT = 16;
const BAR_GAP = 8;
// MAX_BARS lives in ../../utils/chartLayout.js (shared with chartUnavailableReason, used by
// PivotBuilder.jsx to explain — flagged live 2026-09-27 — why a large pivot silently shows no
// chart, instead of it looking broken) — kept out of this file so mixing component and
// non-component exports here doesn't break React Fast Refresh.
// Must stay in sync with PivotChartRenderer.cs's identical array (saved snapshot vs. live preview).
const SLICE_COLORS = ["#AE8C67", "#6a7380", "#63BE7B", "#F8696B", "#FFEB84", "#638EC6", "#9a7b57", "#1f2430", "#c9a876", "#8a94a3"];

// One CATEGORY per detail row (labels[0]). One SERIES per detail column GROUP of the FIRST value
// field only — with no Columns field, columnHeaders has exactly one column per measure, so
// there's exactly 1 series (unlabeled, no legend needed). With a Columns field, each distinct
// dim-combo (e.g. each month) becomes its own series — see PivotChartRenderer.ExtractSeries for
// the full WHY (this fixes the bug caught live 2026-09-24: charting only column 0 silently showed
// just the FIRST column's values). A Grand Total column group is excluded the same way
// GrandTotalsProcessor/ShowValuesAsProcessor exclude it from their own sums.
function extractSeries(result, valueCount) {
  const categoryRows = result.rowHeaders.map((row, r) => (row.isTotal ? null : { row, r })).filter(Boolean);
  if (categoryRows.length === 0 || categoryRows.length > MAX_BARS) return null;

  const categories = categoryRows.map(({ row }) => row.labels[0]);

  const hasGrandTotalColumns =
    result.columnHeaders.length >= valueCount && result.columnHeaders.slice(-valueCount).every((h) => h.isTotal);
  const detailColumnCount = hasGrandTotalColumns ? result.columnHeaders.length - valueCount : result.columnHeaders.length;
  const seriesCount = Math.max(Math.floor(detailColumnCount / valueCount), 1);

  const series = Array.from({ length: seriesCount }, (_, s) => {
    const colIndex = s * valueCount;
    const label = seriesCount === 1 ? "" : result.columnHeaders[colIndex].labels.slice(0, -1).join(" / ");
    const values = categoryRows.map(({ r }) => result.cells[r][colIndex] ?? 0);
    return { label, values };
  });

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

// One row, left to right, swatch + label per series — mirrors PivotChartRenderer.AppendSeriesLegend.
// Each series' x position is computed from the widths of every series before it, rather than a
// mutated accumulator variable (same "derive, don't mutate" approach PieChart's startAngles uses).
function SeriesLegend({ series }) {
  const y = 10;
  const xPositions = series.reduce(
    (acc, s) => [...acc, acc[acc.length - 1] + 13 + s.label.length * 6 + 14],
    [PADDING_LEFT],
  );

  return (
    <>
      {series.map((s, i) => (
        <g key={i}>
          <rect x={xPositions[i]} y={y - 8} width={9} height={9} fill={SLICE_COLORS[i % SLICE_COLORS.length]} />
          <text x={xPositions[i] + 13} y={y} fill="#1f2430">
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

  const paddingTop = data.series.length > 1 ? PADDING_TOP + LEGEND_ROW_HEIGHT : PADDING_TOP;
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
                  {data.series.length === 1 && <ValueLabel x={x + barWidth / 2} y={y} value={value} format={format} />}
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

  const paddingTop = data.series.length > 1 ? PADDING_TOP + LEGEND_ROW_HEIGHT : PADDING_TOP;
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
                {data.series.length === 1 && <ValueLabel x={p.x} y={p.y} value={p.value} format={format} />}
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

export default function PivotChart({ result, valueFields, chartType = "Bar", valuesPlacement = "Columns" }) {
  // v1 scope cut, mirrors the backend's PivotChartRenderer.Render exactly: extractSeries assumes
  // categories are ROWS and series are COLUMN groups — correct only when measures are on Columns.
  // With measures on Rows, consecutive rows ARE the different measures for the same dim-combo, so
  // no chart is drawn rather than a meaningless mixed one.
  if (!result || !valueFields?.length || result.columnHeaders.length === 0 || valuesPlacement === "Rows") {
    return null;
  }

  const data = extractSeries(result, valueFields.length);
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
