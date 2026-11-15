import { describe, expect, it } from 'vitest'
import { EXAMPLES, compute, discountTable, evaluate, parseRoas, toCsv, validate, type Product } from './calc'

const ex = EXAMPLES.map((e, i) => ({ ...e, id: String(i) })) as Product[]
const ok = (p: Product, roas = 2.2) => {
  const e = evaluate(p, roas)
  if (e.kind !== 'ok') throw new Error('expected ok')
  return e.result
}

// Values copied from KK_Break-even_ROAS_Calculator.xlsx (Calculator, rows 9 to 11, ROAS 2.2)
const XLSX = [
  { net: 59, contrib: 33.435, margin: 0.566694915254237, be: 1.76461791535816, target: 2.14272743780643, cpa: 33.435, after: 24.6728813559322, above: true },
  { net: 31.2, contrib: 10.108, margin: 0.323974358974359, be: 3.08666402849228, target: 4.46479679450487, cpa: 10.108, after: -28.725641025641, above: false },
  { net: 84.15, contrib: 40.20475, margin: 0.477774806892454, be: 2.09303627058992, target: 2.64707964045015, cpa: 40.20475, after: 5.11045751633987, above: true },
]

describe('examples at first-order ROAS 2.2 (brief)', () => {
  it('margin 56.7% / 32.4% / 47.8%', () => {
    expect(ex.map((p) => (ok(p).margin * 100).toFixed(1))).toEqual(['56.7', '32.4', '47.8'])
  })
  it('break-even 1.76 / 3.09 / 2.09', () => {
    expect(ex.map((p) => (ok(p).breakEvenRoas as number).toFixed(2))).toEqual(['1.76', '3.09', '2.09'])
  })
  it('after ads +$24.67 / -$28.73 / +$5.11', () => {
    expect(ex.map((p) => ok(p).afterAdsPer100.toFixed(2))).toEqual(['24.67', '-28.73', '5.11'])
  })
})

describe('matches xlsx cell by cell', () => {
  ex.forEach((p, i) => {
    it(p.name, () => {
      const r = ok(p)
      const x = XLSX[i]
      expect(r.netPrice).toBeCloseTo(x.net, 9)
      expect(r.contribution).toBeCloseTo(x.contrib, 9)
      expect(r.margin).toBeCloseTo(x.margin, 12)
      expect(r.breakEvenRoas as number).toBeCloseTo(x.be, 12)
      expect(r.targetRoas as number).toBeCloseTo(x.target, 12)
      expect(r.maxCpa).toBeCloseTo(x.cpa, 9)
      expect(r.afterAdsPer100).toBeCloseTo(x.after, 10)
      expect(r.above).toBe(x.above)
    })
  })
})

describe('discount table matches xlsx "Discount impact" sheet', () => {
  it('starter bundle 0..30%', () => {
    const v = validate(ex[0])
    const rows = discountTable(v.inputs!)
    const be = [1.76461791535816, 1.83240296519088, 1.91409981435755, 2.01448096486207, 2.14078374455733, 2.30453746500879, 2.52529884741203]
    const margin = [0.566694915254237, 0.545731489741302, 0.52243879472693, 0.496405782652044, 0.467118644067797, 0.433926553672316, 0.395992736077482]
    expect(rows.map((r) => r.discount)).toEqual([0, 5, 10, 15, 20, 25, 30])
    rows.forEach((r, i) => {
      expect(r.breakEvenRoas as number).toBeCloseTo(be[i], 12)
      expect(r.margin!).toBeCloseTo(margin[i], 12)
    })
    expect(rows[6].changeVsNone!).toBeCloseTo(0.76068093205387, 12)
  })
})

describe('edge cases', () => {
  const base: Product = { id: 'x', name: 'X', price: '20', cogs: '15', shipping: '5', fees: '3', other: '0', discount: '0', target: '10' }
  it('margin <= 0 gives never and not reachable, max CPA 0', () => {
    const r = ok(base)
    expect(r.breakEvenRoas).toBe('never')
    expect(r.targetRoas).toBe('not reachable')
    expect(r.maxCpa).toBe(0)
    expect(r.above).toBe(false)
  })
  it('target above margin gives not reachable', () => {
    const r = ok({ ...base, cogs: '5', shipping: '0', target: '80' })
    expect(typeof r.breakEvenRoas).toBe('number')
    expect(r.targetRoas).toBe('not reachable')
  })
  it('percent outside 0..95 is an error and nothing is computed', () => {
    for (const bad of ['96', '-1', '120', 'abc']) {
      const e = evaluate({ ...base, discount: bad }, 2)
      expect(e.kind).toBe('invalid')
      if (e.kind === 'invalid') expect(e.errors.discount).toBeTruthy()
    }
    expect(evaluate({ ...base, fees: '95' }, 2).kind).toBe('ok')
    expect(evaluate({ ...base, fees: '0' }, 2).kind).toBe('ok')
  })
  it('missing price is an error, blank row is empty', () => {
    expect(evaluate({ ...base, price: '' }, 2).kind).toBe('invalid')
    expect(evaluate({ ...base, price: '0' }, 2).kind).toBe('invalid')
    const blank = { id: 'b', name: '', price: '', cogs: '', shipping: '', fees: '', other: '', discount: '', target: '' }
    expect(evaluate(blank, 2).kind).toBe('empty')
  })
  it('accepts comma decimals', () => {
    expect(ok({ ...base, cogs: '5,5', shipping: '0' }).contribution).toBeCloseTo(20 - 5.5 - 0.6, 10)
  })
  it('roas parsing', () => {
    expect(parseRoas('2.2').value).toBe(2.2)
    expect(parseRoas('').error).toBeTruthy()
    expect(parseRoas('-1').error).toBeTruthy()
  })
  it('compute is the plain formula', () => {
    const r = compute({ price: 100, cogs: 20, shipping: 10, other: 0, fees: 0.1, discount: 0.5, target: 0 }, 3)
    expect(r.netPrice).toBe(50)
    expect(r.contribution).toBeCloseTo(15, 10)
    expect(r.margin).toBeCloseTo(0.3, 10)
  })
})

describe('csv', () => {
  it('has header, all products and quotes commas', () => {
    const csv = toCsv([...ex, { ...ex[0], id: 'q', name: 'Bundle, large' }], '2.2')
    const lines = csv.trim().split('\r\n')
    expect(lines[0]).toBe('First-order ROAS,2.2')
    expect(lines[2].startsWith('Product,Price ($)')).toBe(true)
    expect(lines).toHaveLength(7)
    expect(lines[3]).toContain('Starter bundle,59,14,8,3.5,1.5,0,10,59.00,33.44,56.7,1.76,2.14,33.44,24.67,Above break-even')
    expect(lines[4]).toContain('Below break-even')
    expect(lines[6].startsWith('"Bundle, large"')).toBe(true)
  })
})
