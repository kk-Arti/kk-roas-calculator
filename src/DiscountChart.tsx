import type { DiscountRow } from './calc'

const W = 420
const H = 240
const PAD = { top: 28, right: 8, bottom: 34, left: 8 }

export function DiscountChart({ rows }: { rows: DiscountRow[] }) {
  const nums = rows.map((r) => (typeof r.breakEvenRoas === 'number' ? r.breakEvenRoas : 0))
  const max = Math.max(...nums, 0.01)
  const innerW = W - PAD.left - PAD.right
  const innerH = H - PAD.top - PAD.bottom
  const slot = innerW / rows.length
  const barW = Math.min(40, slot * 0.62)
  const base = PAD.top + innerH

  const summary = rows.map((r) => `${r.discount}%: ${typeof r.breakEvenRoas === 'number' ? r.breakEvenRoas.toFixed(2) : 'never'}`).join(', ')

  return (
    <figure className="chart">
      <figcaption className="chart-title">Break-even ROAS by discount</figcaption>
      <svg viewBox={`0 0 ${W} ${H}`} role="img" aria-label={`Break-even ROAS by discount. ${summary}`}>
        <line x1={PAD.left} x2={W - PAD.right} y1={base} y2={base} className="axis" />
        {rows.map((r, i) => {
          const cx = PAD.left + slot * i + slot / 2
          const v = typeof r.breakEvenRoas === 'number' ? r.breakEvenRoas : null
          const h = v === null ? 0 : (v / max) * innerH
          return (
            <g key={r.discount}>
              {v !== null && <rect className="bar" x={cx - barW / 2} y={base - h} width={barW} height={h} rx={3} />}
              <text x={cx} y={v === null ? base - 8 : base - h - 8} textAnchor="middle" className="bar-value">
                {v === null ? 'never' : v.toFixed(2)}
              </text>
              <text x={cx} y={base + 22} textAnchor="middle" className="bar-label">
                {r.discount}%
              </text>
            </g>
          )
        })}
      </svg>
    </figure>
  )
}
