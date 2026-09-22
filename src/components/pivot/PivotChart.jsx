import { formatNumber } from "../../utils/numberFormat";

// Live-preview mirror of the backend's PivotChartRenderer.cs (same bar-chart-only, first-value-
// field-only, MaxBars cap) — an actual <svg>, not an HTML canvas element, so what the user sees
// here is what the saved snapshot will actually contain (no JS-driven charting library involved
// on either side).
const WIDTH = 640;
const HEIGHT = 320;
const PADDING_LEFT = 40;
const PADDING_BOTTOM = 60;
const PADDING_TOP = 20;
const BAR_GAP = 8;
const MAX_BARS = 50;

export default function PivotChart({ result, valueFields }) {
  if (!result || !valueFields?.length || result.columnHeaders.length === 0) {
    return null;
  }

  const bars = result.rowHeaders
    .map((row, r) => (row.isTotal ? null : { label: row.labels[0], value: result.cells[r][0] ?? 0 }))
    .filter(Boolean);

  if (bars.length === 0 || bars.length > MAX_BARS) {
    return null;
  }

  const max = Math.max(...bars.map((b) => b.value));
  if (max <= 0) {
    return null;
  }

  const chartWidth = WIDTH - PADDING_LEFT - 10;
  const chartHeight = HEIGHT - PADDING_TOP - PADDING_BOTTOM;
  const barWidth = chartWidth / bars.length - BAR_GAP;
  const format = valueFields[0]?.format;

  return (
    <div className="card mt-6 overflow-x-auto">
      <svg width={WIDTH} height={HEIGHT} viewBox={`0 0 ${WIDTH} ${HEIGHT}`} style={{ fontFamily: "Segoe UI,Arial,sans-serif", fontSize: 11 }}>
        <line x1={PADDING_LEFT} y1={HEIGHT - PADDING_BOTTOM} x2={WIDTH - 10} y2={HEIGHT - PADDING_BOTTOM} stroke="#d8d5dd" />
        {bars.map((bar, i) => {
          const barHeight = bar.value <= 0 ? 0 : (bar.value / max) * chartHeight;
          const x = PADDING_LEFT + i * (chartWidth / bars.length) + BAR_GAP / 2;
          const y = HEIGHT - PADDING_BOTTOM - barHeight;
          return (
            <g key={i}>
              <rect x={x} y={y} width={Math.max(barWidth, 1)} height={barHeight} fill="#AE8C67" />
              <text x={x + barWidth / 2} y={y - 4} textAnchor="middle" fill="#1f2430">
                {formatNumber(bar.value, format)}
              </text>
              <text
                x={x + barWidth / 2}
                y={HEIGHT - PADDING_BOTTOM + 14}
                textAnchor="end"
                fill="#6a7380"
                transform={`rotate(-40 ${x + barWidth / 2} ${HEIGHT - PADDING_BOTTOM + 14})`}
              >
                {bar.label}
              </text>
            </g>
          );
        })}
      </svg>
    </div>
  );
}
