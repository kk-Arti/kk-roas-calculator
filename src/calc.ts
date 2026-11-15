// Pure calculation logic. Mirrors KK_Break-even_ROAS_Calculator.xlsx one to one.
// Percent inputs are entered as whole numbers (20 means 20%).

export const PCT_MIN = 0
export const PCT_MAX = 95
export const MAX_PRODUCTS = 20
export const DISCOUNT_STEPS = [0, 5, 10, 15, 20, 25, 30]

export type FieldKey =
  | 'name'
  | 'price'
  | 'cogs'
  | 'shipping'
  | 'fees'
  | 'other'
  | 'discount'
  | 'target'

export interface Product {
  id: string
  example?: boolean
  name: string
  price: string
  cogs: string
  shipping: string
  fees: string
  other: string
  discount: string
  target: string
}

export const PERCENT_FIELDS: FieldKey[] = ['fees', 'discount', 'target']
export const MONEY_FIELDS: FieldKey[] = ['price', 'cogs', 'shipping', 'other']

export const FIELD_LABELS: Record<FieldKey, string> = {
  name: 'Product name',
  price: 'Price ($)',
  cogs: 'COGS ($)',
  shipping: 'Shipping & pick/pack ($)',
  fees: 'Payment & platform fees (% of net price)',
  other: 'Other variable cost ($)',
  discount: 'Discount in ads (%)',
  target: 'Target contribution after ads (% of revenue, covers opex and profit)',
}

export interface Result {
  netPrice: number
  contribution: number
  margin: number
  breakEvenRoas: number | 'never'
  targetRoas: number | 'not reachable'
  maxCpa: number
  afterAdsPer100: number
  above: boolean
}

export type Evaluation =
  | { kind: 'empty' }
  | { kind: 'invalid'; errors: Partial<Record<FieldKey, string>> }
  | { kind: 'ok'; result: Result }

/** Parses a user typed number. Accepts "1,5" and "$59". Empty string gives null. */
export function parseNum(raw: string): number | null {
  const s = raw.trim().replace(/[$\s%]/g, '').replace(',', '.')
  if (s === '') return null
  if (!/^-?\d*\.?\d+$|^-?\d+\.$/.test(s)) return NaN
  return Number(s)
}

export function parseRoas(raw: string): { value: number | null; error?: string } {
  const v = parseNum(raw)
  if (v === null) return { value: null, error: 'Enter your first-order ROAS, for example 2.2' }
  if (Number.isNaN(v)) return { value: null, error: 'Use a number, for example 2.2' }
  if (v < 0) return { value: null, error: 'ROAS cannot be negative' }
  if (v > 100) return { value: null, error: 'That looks too high. Use a value up to 100' }
  return { value: v }
}

export function parseTarget(raw: string): { value: number | null; error?: string } {
  const v = parseNum(raw)
  if (v === null) return { value: null, error: 'Enter a target, for example 10 for 10%' }
  if (Number.isNaN(v)) return { value: null, error: 'Use a number, for example 10 for 10%' }
  if (v < PCT_MIN || v > PCT_MAX) return { value: null, error: `Enter a percent from ${PCT_MIN} to ${PCT_MAX}` }
  return { value: v / 100 }
}

export interface Inputs {
  price: number
  cogs: number
  shipping: number
  fees: number // fraction, 0.035
  other: number
  discount: number // fraction
  target: number // fraction
}

/** Validates one product. Empty cost fields count as 0. Price is required. */
export function validate(p: Product): { inputs?: Inputs; errors: Partial<Record<FieldKey, string>>; empty: boolean } {
  const errors: Partial<Record<FieldKey, string>> = {}
  const vals: Record<string, number> = {}
  const allBlank = [...MONEY_FIELDS, ...PERCENT_FIELDS].every((k) => p[k].trim() === '')

  for (const k of MONEY_FIELDS) {
    const v = parseNum(p[k])
    if (v === null) {
      if (k === 'price') errors[k] = 'Enter the price'
      vals[k] = 0
    } else if (Number.isNaN(v)) errors[k] = 'Use a number, for example 12.50'
    else if (v < 0) errors[k] = 'Cannot be negative'
    else if (k === 'price' && v === 0) errors[k] = 'Price must be above 0'
    else vals[k] = v
  }
  for (const k of PERCENT_FIELDS) {
    const v = parseNum(p[k])
    if (v === null) vals[k] = 0
    else if (Number.isNaN(v)) errors[k] = 'Use a number, for example 20 for 20%'
    else if (v < PCT_MIN || v > PCT_MAX) errors[k] = `Enter a percent from ${PCT_MIN} to ${PCT_MAX}. Type 20 for 20%`
    else vals[k] = v / 100
  }

  if (allBlank) return { errors: {}, empty: true }
  if (Object.keys(errors).length) return { errors, empty: false }
  return {
    errors,
    empty: false,
    inputs: {
      price: vals.price,
      cogs: vals.cogs,
      shipping: vals.shipping,
      other: vals.other,
      fees: vals.fees,
      discount: vals.discount,
      target: vals.target,
    },
  }
}

export function compute(i: Inputs, roas: number): Result {
  const netPrice = i.price * (1 - i.discount)
  const contribution = netPrice - i.cogs - i.shipping - i.other - netPrice * i.fees
  const margin = contribution / netPrice
  const breakEvenRoas = margin <= 0 ? 'never' : 1 / margin
  const targetRoas = margin - i.target <= 0 ? 'not reachable' : 1 / (margin - i.target)
  const maxCpa = Math.max(contribution, 0)
  const afterAdsPer100 = (roas * margin - 1) * 100
  return {
    netPrice,
    contribution,
    margin,
    breakEvenRoas,
    targetRoas,
    maxCpa,
    afterAdsPer100,
    above: afterAdsPer100 >= 0,
  }
}

export function evaluate(p: Product, roas: number | null): Evaluation {
  const v = validate(p)
  if (v.empty) return { kind: 'empty' }
  if (!v.inputs) return { kind: 'invalid', errors: v.errors }
  return { kind: 'ok', result: compute(v.inputs, roas ?? 0) }
}

export interface DiscountRow {
  discount: number // whole percent
  netPrice: number
  contribution: number
  margin: number | null
  breakEvenRoas: number | 'never'
  changeVsNone: number | null
}

/** "What a discount does": same product at 0..30% discount. Uses product COGS + shipping + other as variable cost. */
export function discountTable(i: Inputs): DiscountRow[] {
  const variable = i.cogs + i.shipping + i.other
  const rows = DISCOUNT_STEPS.map((d) => {
    const netPrice = i.price * (1 - d / 100)
    const contribution = netPrice - variable - netPrice * i.fees
    const margin = netPrice === 0 ? null : contribution / netPrice
    const breakEvenRoas: number | 'never' = margin === null || margin <= 0 ? 'never' : 1 / margin
    return { discount: d, netPrice, contribution, margin, breakEvenRoas, changeVsNone: null as number | null }
  })
  const base = rows[0].breakEvenRoas
  for (const r of rows) {
    r.changeVsNone = typeof r.breakEvenRoas === 'number' && typeof base === 'number' ? r.breakEvenRoas - base : null
  }
  return rows
}

// Formatting

export const fmtRoas = (v: number | string) => (typeof v === 'number' ? v.toFixed(2) : v)
export const fmtPct = (v: number) => `${(v * 100).toFixed(1)}%`
export function fmtMoney(v: number, signed = false) {
  const abs = Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  if (v < 0) return `−$${abs}` // minus sign, not a dash
  return `${signed ? '+' : ''}$${abs}`
}

// Examples

export const EXAMPLES: Omit<Product, 'id'>[] = [
  { example: true, name: 'Starter bundle', price: '59', cogs: '14', shipping: '8', fees: '3.5', other: '1.5', discount: '0', target: '10' },
  { example: true, name: 'Single pouch', price: '39', cogs: '12', shipping: '7', fees: '3.5', other: '1.0', discount: '20', target: '10' },
  { example: true, name: '3-pack bundle', price: '99', cogs: '30', shipping: '9', fees: '3.5', other: '2.0', discount: '15', target: '10' },
]

// CSV

function csvCell(v: string | number) {
  const s = String(v)
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
}

export function toCsv(products: Product[], roasRaw: string): string {
  const roas = parseRoas(roasRaw).value
  const header = [
    'Product',
    'Price ($)',
    'COGS ($)',
    'Shipping & pick/pack ($)',
    'Payment & platform fees (%)',
    'Other variable cost ($)',
    'Discount in ads (%)',
    'Target contribution after ads (%)',
    'Net price ($)',
    'Contribution per order ($)',
    'Contribution margin (%)',
    'Break-even ROAS',
    'Target ROAS',
    'Max CPA at break-even ($)',
    'Contribution after ads per $100 ($)',
    'Status',
  ]
  const lines = [[`First-order ROAS`, roasRaw.trim()].map(csvCell).join(','), '', header.map(csvCell).join(',')]
  for (const p of products) {
    const e = evaluate(p, roas)
    const base = [p.name, p.price, p.cogs, p.shipping, p.fees, p.other, p.discount, p.target]
    let res: (string | number)[]
    if (e.kind === 'ok') {
      const r = e.result
      res = [
        r.netPrice.toFixed(2),
        r.contribution.toFixed(2),
        (r.margin * 100).toFixed(1),
        fmtRoas(r.breakEvenRoas),
        fmtRoas(r.targetRoas),
        r.maxCpa.toFixed(2),
        roas === null ? '' : r.afterAdsPer100.toFixed(2),
        roas === null ? '' : r.above ? 'Above break-even' : 'Below break-even',
      ]
    } else if (e.kind === 'invalid') res = ['', '', '', '', '', '', '', 'Check inputs']
    else res = ['', '', '', '', '', '', '', '']
    lines.push([...base, ...res].map(csvCell).join(','))
  }
  return lines.join('\r\n') + '\r\n'
}
