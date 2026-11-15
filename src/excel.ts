// Branded Excel export: a live copy of the calculator with the visitor's products and working formulas.
// Formulas mirror calc.ts and the original KK_Break-even_ROAS_Calculator.xlsx one to one.
import type { Workbook, Worksheet, Cell } from 'exceljs'
import { LOGO_PNG_BASE64 } from './assets/logo64'
import { evaluate, parseRoas, parseTarget, validate, MAX_PRODUCTS, type Product } from './calc'

const MESSAGE_URL =
  'https://www.linkedin.com/messaging/compose/?profileUrn=urn%3Ali%3Afsd_profile%3AACoAADxgieABwE7747g9IZ4BfTrBChanNiZJwr4'
const TOOL_URL = 'https://kk-arti.github.io/kk-roas-calculator/'

const C = {
  dark: 'FF1C1C1C',
  dark2: 'FF242424',
  green: 'FF179158',
  greenDeep: 'FF072B1A',
  greenBright: 'FF00D370',
  greenText: 'FF0F7A49',
  inputFill: 'FFEAF6EF',
  outFill: 'FFF5F5F5',
  grey: 'FF6B6B6B',
  line: 'FFD9D9D9',
  red: 'FFC0392B',
  blue: 'FF0000FF',
  white: 'FFFFFFFF',
}
const FONT = 'Arial'
const fill = (argb: string) => ({ type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb } })
const thin = { style: 'thin' as const, color: { argb: C.line } }
const box = { top: thin, left: thin, bottom: thin, right: thin }

const ROAS_CELL = 'C6'
const TARGET_CELL = 'C7'
const HEAD_ROW = 9
const FIRST = HEAD_ROW + 1
const LAST = HEAD_ROW + MAX_PRODUCTS

const HEADERS = [
  'Product',
  'Price ($)',
  'COGS ($)',
  'Shipping & pick/pack ($)',
  'Payment & platform fees (% of net price)',
  'Other variable cost ($)',
  'Discount in ads (%)',
  'Net price ($)',
  'Contribution per order ($)',
  'Contribution margin (%)',
  'Break-even ROAS',
  'Target ROAS',
  'Max CPA at break-even ($)',
  'After ads per $100 of ads ($)',
  'Status at your ROAS',
]
const WIDTHS = [3, 26, 10, 10, 13, 15, 13, 12, 12, 14, 14, 12, 14, 14, 15, 18, 4]

function band(ws: Worksheet, row: number, argb: string) {
  for (let c = 1; c <= 16; c++) ws.getCell(row, c).fill = fill(argb)
}

function formula(cell: Cell, f: string, result: number | string | undefined) {
  cell.value = result === undefined ? { formula: f } : { formula: f, result }
}

export async function buildWorkbook(products: Product[], roasRaw: string, targetRaw: string): Promise<Blob> {
  const ExcelJS = (await import('exceljs')).default
  const wb: Workbook = new ExcelJS.Workbook()
  wb.creator = 'K&K Media'
  wb.title = 'Break-even ROAS Calculator'
  const ws = wb.addWorksheet('Calculator', {
    views: [{ state: 'frozen', ySplit: HEAD_ROW, xSplit: 2, showGridLines: false }],
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  })
  WIDTHS.forEach((w, i) => (ws.getColumn(i + 1).width = w))

  // Brand header
  for (const r of [1, 2, 3, 4]) band(ws, r, C.dark)
  ws.getRow(1).height = 12
  ws.getRow(2).height = 30
  ws.getRow(3).height = 20
  ws.getRow(4).height = 12
  const logo = wb.addImage({ base64: LOGO_PNG_BASE64, extension: 'png' })
  ws.addImage(logo, { tl: { col: 1.1, row: 1.05 }, ext: { width: 42, height: 42 } })
  const title = ws.getCell('C2')
  title.value = 'Break-even ROAS Calculator'
  title.font = { name: FONT, size: 18, bold: true, color: { argb: C.white } }
  const sub = ws.getCell('C3')
  sub.value = 'K&K Media, built by Artem Kalachev'
  sub.font = { name: FONT, size: 10, color: { argb: 'FFCDCDCD' } }
  const msg = ws.getCell('M6')
  msg.value = { text: 'Message me on LinkedIn', hyperlink: MESSAGE_URL }
  msg.font = { name: FONT, size: 11, bold: true, color: { argb: C.greenText }, underline: true }
  const web = ws.getCell('M7')
  web.value = { text: 'Open the online calculator', hyperlink: TOOL_URL }
  web.font = { name: FONT, size: 10, color: { argb: C.grey }, underline: true }

  // ROAS input
  ws.getCell('B6').value = 'Your first-order ROAS'
  ws.getCell('B6').font = { name: FONT, size: 12, bold: true }
  const roas = parseRoas(roasRaw).value
  const rc = ws.getCell(ROAS_CELL)
  rc.value = roas ?? 2.2
  rc.numFmt = '0.00'
  rc.font = { name: FONT, size: 14, bold: true, color: { argb: C.blue } }
  rc.fill = fill(C.inputFill)
  rc.border = { top: { style: 'medium', color: { argb: C.green } }, left: { style: 'medium', color: { argb: C.green } }, bottom: { style: 'medium', color: { argb: C.green } }, right: { style: 'medium', color: { argb: C.green } } }
  rc.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(6).height = 24
  const hint = ws.getCell('E6')
  hint.value = 'Use ROAS on first orders from new customers. Blended ROAS includes repeat buyers and makes the first order look better than it is.'
  hint.font = { name: FONT, size: 9, italic: true, color: { argb: C.grey } }
  ws.getCell('B7').value = 'Target margin after ads'
  ws.getCell('B7').font = { name: FONT, size: 12, bold: true }
  const target = parseTarget(targetRaw).value
  const tc = ws.getCell(TARGET_CELL)
  tc.value = target ?? 0.1
  tc.numFmt = '0.0%'
  tc.font = rc.font
  tc.fill = fill(C.inputFill)
  tc.border = rc.border
  tc.alignment = rc.alignment
  tc.dataValidation = { type: 'decimal', operator: 'between', formulae: [0, 0.95], allowBlank: false, showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Enter a percentage', error: 'Type 10% (or 0.1), not 10. Allowed range 0% to 95%.' }
  ws.getRow(7).height = 24
  ws.getCell('E7').value = 'What each order should leave after ad spend to pay opex and profit. Drives Target ROAS.'
  ws.getCell('E7').font = { name: FONT, size: 9, italic: true, color: { argb: C.grey } }
  ws.getCell('B8').value = 'Green cells with blue text are inputs. Everything else calculates. Type percents as 20%.'
  ws.getCell('B8').font = { name: FONT, size: 9, color: { argb: C.grey } }

  // Table header
  const hr = ws.getRow(HEAD_ROW)
  hr.height = 48
  HEADERS.forEach((h, i) => {
    const c = hr.getCell(i + 2)
    c.value = h
    c.font = { name: FONT, size: 9, bold: true, color: { argb: C.white } }
    c.fill = fill(i < 7 ? C.green : C.dark2)
    c.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' }
    c.border = box
  })

  // Product rows: visitor's products first, the rest blank and ready to fill
  const filled = products.filter((p) => !validate(p).empty)
  for (let k = 0; k < MAX_PRODUCTS; k++) {
    const r = FIRST + k
    const row = ws.getRow(r)
    row.height = 18
    const p = filled[k]
    const v = p ? validate(p) : undefined
    const inp = v?.inputs
    const ev = p && roas !== null ? evaluate({ ...p, target: target === null ? '' : targetRaw }, roas) : undefined
    const res = ev && ev.kind === 'ok' ? ev.result : undefined

    const inputs: [string, unknown, string][] = [
      ['B', p ? p.name || `Product ${k + 1}` : null, '@'],
      ['C', inp?.price ?? null, '$#,##0.00'],
      ['D', inp?.cogs ?? null, '$#,##0.00'],
      ['E', inp?.shipping ?? null, '$#,##0.00'],
      ['F', inp?.fees ?? null, '0.0%'],
      ['G', inp?.other ?? null, '$#,##0.00'],
      ['H', inp?.discount ?? null, '0.0%'],
    ]
    for (const [col, val, fmt] of inputs) {
      const c = ws.getCell(`${col}${r}`)
      if (val !== null && val !== undefined) c.value = val as string | number
      c.numFmt = fmt
      c.fill = fill(C.inputFill)
      c.font = { name: FONT, size: 10, color: { argb: C.blue } }
      c.border = box
      if ('FH'.includes(col)) {
        c.dataValidation = {
          type: 'decimal',
          operator: 'between',
          formulae: [0, 0.95],
          allowBlank: true,
          showErrorMessage: true,
          errorStyle: 'stop',
          errorTitle: 'Enter a percentage',
          error: 'Type 20% (or 0.2), not 20. Allowed range 0% to 95%.',
        }
      }
    }

    const blank = res === undefined && !p ? '' : undefined
    const T = `$${TARGET_CELL[0]}$${TARGET_CELL.slice(1)}`
    const RO = `$${ROAS_CELL[0]}$${ROAS_CELL.slice(1)}`
    formula(ws.getCell(`I${r}`), `IF(C${r}="","",IF(OR(F${r}>=1,H${r}>=1),"check % input",C${r}*(1-H${r})))`, res?.netPrice ?? blank)
    formula(ws.getCell(`J${r}`), `IF(OR(C${r}="",ISTEXT(I${r})),"",I${r}-D${r}-E${r}-G${r}-I${r}*F${r})`, res?.contribution ?? blank)
    formula(ws.getCell(`K${r}`), `IF(OR(C${r}="",ISTEXT(I${r}),I${r}=0),"",J${r}/I${r})`, res?.margin ?? blank)
    formula(ws.getCell(`L${r}`), `IF(K${r}="","",IF(K${r}<=0,"never",1/K${r}))`, res?.breakEvenRoas ?? blank)
    formula(ws.getCell(`M${r}`), `IF(K${r}="","",IF(K${r}-${T}<=0,"not reachable",1/(K${r}-${T})))`, res?.targetRoas ?? blank)
    formula(ws.getCell(`N${r}`), `IF(J${r}="","",MAX(J${r},0))`, res?.maxCpa ?? blank)
    formula(ws.getCell(`O${r}`), `IF(K${r}="","",(${RO}*K${r}-1)*100)`, res?.afterAdsPer100 ?? blank)
    formula(
      ws.getCell(`P${r}`),
      `IF(K${r}="","",IF(O${r}>=0,"Above break-even","Below break-even"))`,
      res ? (res.above ? 'Above break-even' : 'Below break-even') : blank,
    )
    const outFmt: Record<string, string> = { I: '$#,##0.00', J: '$#,##0.00;-$#,##0.00', K: '0.0%', L: '0.00', M: '0.00', N: '$#,##0.00', O: '+$#,##0.00;-$#,##0.00', P: '@' }
    for (const [col, fmt] of Object.entries(outFmt)) {
      const c = ws.getCell(`${col}${r}`)
      c.numFmt = fmt
      c.fill = fill(C.outFill)
      c.border = box
      c.font = { name: FONT, size: 10, bold: 'LP'.includes(col) }
      c.alignment = { horizontal: col === 'P' ? 'left' : 'right' }
    }
  }
  ws.addConditionalFormatting({
    ref: `P${FIRST}:P${LAST}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['"Below break-even"'], priority: 1, style: { font: { color: { argb: C.red }, bold: true } } },
      { type: 'cellIs', operator: 'equal', formulae: ['"Above break-even"'], priority: 2, style: { font: { color: { argb: C.greenText }, bold: true } } },
    ],
  })
  ws.addConditionalFormatting({
    ref: `O${FIRST}:O${LAST}`,
    rules: [{ type: 'cellIs', operator: 'lessThan', formulae: [0], priority: 3, style: { font: { color: { argb: C.red } } } }],
  })

  // How it works
  let r = LAST + 2
  ws.getCell(`B${r}`).value = 'How the numbers work'
  ws.getCell(`B${r}`).font = { name: FONT, size: 12, bold: true }
  const notes = [
    'Contribution per order = net price minus COGS, shipping, other variable cost and fees.',
    'Contribution margin = contribution ÷ net price.',
    'Break-even ROAS = 1 ÷ contribution margin. Below it, every first order loses money.',
    'Target ROAS = 1 ÷ (contribution margin minus target contribution %).',
    'Max CPA at break-even = contribution per order: the most you can pay for one first order.',
    'After ads per $100 = (first-order ROAS × contribution margin minus 1) × 100. Before opex, so it is not net profit.',
    'First order only. Subscription and repeat purchases change the picture: judge them by cohort payback days.',
  ]
  for (const n of notes) {
    r++
    ws.getCell(`B${r}`).value = n
    ws.getCell(`B${r}`).font = { name: FONT, size: 10, color: { argb: 'FF3A3A3A' } }
  }

  // Closing CTA
  r += 2
  for (const rr of [r, r + 1, r + 2]) band(ws, rr, C.inputFill)
  ws.getRow(r).height = 10
  ws.getRow(r + 2).height = 10
  const q = ws.getCell(`B${r + 1}`)
  q.value = 'Need one ROAS target for the whole account? It depends on how your ad spend splits across products.'
  q.font = { name: FONT, size: 11, bold: true, color: { argb: C.greenDeep } }
  ws.getRow(r + 1).height = 22
  const cta = ws.getCell(`M${r + 1}`)
  cta.value = { text: 'Send me your mix', hyperlink: MESSAGE_URL }
  cta.font = { name: FONT, size: 11, bold: true, color: { argb: C.greenText }, underline: true }

  ws.headerFooter.oddFooter = '&L&8K&K Media, Break-even ROAS Calculator&R&8kkmedia.io'

  const buf = await wb.xlsx.writeBuffer()
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}
