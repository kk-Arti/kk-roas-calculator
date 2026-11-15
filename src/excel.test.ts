import { describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { writeFileSync } from 'node:fs'
import { buildWorkbook } from './excel'
import { EXAMPLES, type Product } from './calc'

describe('excel export', () => {
  it('writes the visitor products with formulas and the expected cached results', async () => {
    const products: Product[] = EXAMPLES.map((e, i) => ({ ...e, id: `x${i}` }))
    const blob = await buildWorkbook(products, '2.2', '10')
    const buf = Buffer.from(await blob.arrayBuffer())
    if (process.env.XLSX_OUT) writeFileSync(process.env.XLSX_OUT, buf)
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(buf)
    const ws = wb.getWorksheet('Calculator')!
    expect(ws.getCell('C6').value).toBe(2.2)
    expect(ws.getCell('C7').value).toBeCloseTo(0.1)
    expect(ws.getCell('B10').value).toBe('Starter bundle')
    const r = (a: string) => (ws.getCell(a).value as { result: number }).result
    expect(r('L10')).toBeCloseTo(1.76, 2)
    expect(r('L11')).toBeCloseTo(3.09, 2)
    expect(r('L12')).toBeCloseTo(2.09, 2)
    expect(r('O11')).toBeCloseTo(-28.73, 2)
    expect((ws.getCell('M10').value as { formula: string }).formula).toContain('$C$7')
  })
})
