import { formatNumber } from "../../utils/numberFormat";

// Live-preview mirror of the backend's PivotChartRenderer.cs (same three chart types, same
// first-value-field-only scope, same MaxBars cap) — actual <svg> elements, not an HTML canvas or
// charting library, so what the user sees here is what the saved snapshot will actually contain.
const WIDTH = 640;
const HEIGHT = 320;
const PADDING_LEFT = 40;
const PADDING_BOTTOM = 60;
const PADDING_TOP = 20;
const BAR_GAP = 8;
const MAX_BARS = 50;
const SLICE_COLORS = ["#AE8C67", "#6a7380", "#63BE7B", "#F8696B", "#FFEB84", "#638EC6", "#9a7b57", "#1f2430", "#c9a876", "#8a94a3"];

function extractBars(result) {
  const bars = result.rowHeaders
    .map((row, r) => (row.isTotal ? null : { label: row.labels[0], value: result.cells[r][0] ?? 0 }))
    .filter(Boolean);
  return bars.length === 0 || bars.length > MAX_BARS ? null : bars;
}

function ValueAndLabel({ x, y, label, value, format }) {
  return (
    <>
      <text x={x} y={y - 6} textAnchor="middle" fill="#1f2430">
        {formatNumber(value, format)}
      </text>
      <text x={x} y={HEIGHT - PADDING_BOTTOM + 14} textAnchor="end" fill="#6a7380" transform={`rotate(-40 ${x} ${HEIGHT - PADDING_BOTTOM + 14})`}>
        {label}
      </text>
    </>
  );
}

function BarChart({ bars, format }) {
  const max = Math.max(...bars.map((b) => b.value));
  if (max <= 0) return null;

  const chartWidth = WIDTH - PADDING_LEFT - 10;
  const chartHeight = HEIGHT - PADDING_TOP - PADDING_BOTTOM;
  const barWidth = chartWidth / bars.length - BAR_GAP;

  return (
    <>
      <line x1={PADDING_LEFT} y1={HEIGHT - PADDING_BOTTOM} x2={WIDTH - 10} y2={HEIGHT - PADDING_BOTTOM} stroke="#d8d5dd" />
      {bars.map((bar, i) => {
        const barHeight = bar.value <= 0 ? 0 : (bar.value / max) * chartHeight;
        const x = PADDING_LEFT + i * (chartWidth / bars.length) + BAR_GAP / 2;
        const y = HEIGHT - PADDING_BOTTOM - barHeight;
        return (
          <g key={i}>
            <rect x={x} y={y} width={Math.max(barWidth, 1)} height={barHeight} fill="#AE8C67" />
            <ValueAndLabel x={x + barWidth / 2} y={y} label={bar.label} value={bar.value} format={format} />
          </g>
        );
      })}
    </>
  );
}

function LineChart({ bars, format }) {
  const max = Math.max(...bars.map((b) => b.value));
  if (max <= 0) return null;

  const chartWidth = WIDTH - PADDING_LEFT - 10;
  const chartHeight = HEIGHT - PADDING_TOP - PADDING_BOTTOM;
  const slotWidth = chartWidth / bars.length;

  const points = bars.map((b, i) => ({
    x: PADDING_LEFT + i * slotWidth + slotWidth / 2,
    y: HEIGHT - PADDING_BOTTOM - (b.value <= 0 ? 0 : (b.value / max) * chartHeight),
    label: b.label,
    value: b.value,
  }));

  return (
    <>
      <line x1={PADDING_LEFT} y1={HEIGHT - PADDING_BOTTOM} x2={WIDTH - 10} y2={HEIGHT - PADDING_BOTTOM} stroke="#d8d5dd" />
      <polyline points={points.map((p) => `${p.x},${p.y}`).join(" ")} fill="none" stroke="#AE8C67" strokeWidth={2} />
      {points.map((p, i) => (
        <g key={i}>
          <circle cx={p.x} cy={p.y} r={3} fill="#AE8C67" />
          <ValueAndLabel x={p.x} y={p.y} label={p.label} value={p.value} format={format} />
        </g>
      ))}
    </>
  );
}

function PieChart({ bars }) {
  const positive = bars.filter((b) => b.value > 0);
  const total = positive.reduce((sum, b) => sum + b.value, 0);
  if (positive.length === 0 || total <= 0) return null;

  const cx = WIDTH / 4;
  const cy = HEIGHT / 2;
  const r = Math.min(cy, cx) - 20;
  const legendX = WIDTH / 2 + 20;
  const legendY = PADDING_TOP + 10;

  // A single slice can't be drawn as an SVG arc (start/end angle coincide) — a plain circle
  // instead, same as the backend.
  if (positive.length === 1) {
    return (
      <>
        <circle cx={cx} cy={cy} r={r} fill={SLICE_COLORS[0]} />
        <rect x={legendX} y={legendY - 9} width={10} height={10} fill={SLICE_COLORS[0]} />
        <text x={legendX + 16} y={legendY} fill="#1f2430">
          {positive[0].label} (100.0%)
        </text>
      </>
    );
  }

  // Cumulative sweep-so-far per slice, computed without a mutated accumulator variable —
  // startAngles[i] is where slice i begins, derived purely from the slices before it.
  const startAngles = positive.reduce(
    (acc, bar) => [...acc, acc[acc.length - 1] + (bar.value / total) * 2 * Math.PI],
    [-Math.PI / 2],
  );

  const slices = positive.map((bar, i) => {
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

export default function PivotChart({ result, valueFields, chartType = "Bar" }) {
  if (!result || !valueFields?.length || result.columnHeaders.length === 0) {
    return null;
  }

  const bars = extractBars(result);
  if (!bars) {
    return null;
  }

  const format = valueFields[0]?.format;

  return (
    <div className="card mt-6 overflow-x-auto">
      <svg width={WIDTH} height={HEIGHT} viewBox={`0 0 ${WIDTH} ${HEIGHT}`} style={{ fontFamily: "Segoe UI,Arial,sans-serif", fontSize: 11 }}>
        {chartType === "Line" && <LineChart bars={bars} format={format} />}
        {chartType === "Pie" && <PieChart bars={bars} />}
        {chartType !== "Line" && chartType !== "Pie" && <BarChart bars={bars} format={format} />}
      </svg>
    </div>
  );
}
