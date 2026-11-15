import { useEffect, useMemo, useState } from 'react'
import {
  FIELD_LABELS,
  MAX_PRODUCTS,
  PERCENT_FIELDS,
  discountTable,
  evaluate,
  fmtMoney,
  fmtPct,
  fmtRoas,
  parseRoas,
  parseTarget,
  validate,
  type Evaluation,
  type FieldKey,
  type Product,
} from './calc'
import { blankProduct, load, newId, save, type State } from './storage'
import { DiscountChart } from './DiscountChart'

const LINKEDIN = 'https://www.linkedin.com/in/artem-kalachev/'
// Opens a new LinkedIn message to Artem (same link as the Message button on his profile)
const LINKEDIN_MESSAGE =
  'https://www.linkedin.com/messaging/compose/?profileUrn=urn%3Ali%3Afsd_profile%3AACoAADxgieABwE7747g9IZ4BfTrBChanNiZJwr4'
const LOGO = `${import.meta.env.BASE_URL}kk-logo-96.png`

const INPUT_FIELDS: FieldKey[] = ['price', 'cogs', 'shipping', 'fees', 'other', 'discount']

const SHORT: Record<FieldKey, { title: string; unit: string }> = {
  name: { title: 'Product', unit: '' },
  price: { title: 'Price', unit: '$' },
  cogs: { title: 'COGS', unit: '$' },
  shipping: { title: 'Shipping & pick/pack', unit: '$' },
  fees: { title: 'Payment & platform fees', unit: '% of net price' },
  other: { title: 'Other variable cost', unit: '$' },
  discount: { title: 'Discount in ads', unit: '%' },
  target: { title: 'Target after ads', unit: '% of revenue' },
}

const productLabel = (p: Product, i: number) => p.name.trim() || `Product ${i + 1}`

export default function App() {
  const [state, setState] = useState<State>(load)
  useEffect(() => save(state), [state])

  const roasParsed = parseRoas(state.roas)
  const roas = roasParsed.value
  const targetParsed = parseTarget(state.target)
  const withTarget = useMemo(
    () => state.products.map((p) => ({ ...p, target: targetParsed.value === null ? '' : state.target })),
    [state.products, state.target, targetParsed.value],
  )
  const evals = useMemo(() => withTarget.map((p) => evaluate(p, roas)), [withTarget, roas])
  const hasExamples = state.products.some((p) => p.example)

  const setRoas = (roas: string) => setState((s) => ({ ...s, roas }))
  const setTarget = (target: string) => setState((s) => ({ ...s, target }))
  const update = (id: string, key: FieldKey, value: string) =>
    setState((s) => ({
      ...s,
      products: s.products.map((p) => (p.id === id ? { ...p, [key]: value, example: false } : p)),
    }))
  const add = () =>
    setState((s) => (s.products.length >= MAX_PRODUCTS ? s : { ...s, products: [...s.products, blankProduct()] }))
  const duplicate = (id: string) =>
    setState((s) => {
      if (s.products.length >= MAX_PRODUCTS) return s
      const i = s.products.findIndex((p) => p.id === id)
      const src = s.products[i]
      const copy: Product = { ...src, id: newId(), example: false, name: src.name ? `${src.name} copy` : '' }
      const products = [...s.products]
      products.splice(i + 1, 0, copy)
      return { ...s, products }
    })
  const remove = (id: string) =>
    setState((s) => {
      const products = s.products.filter((p) => p.id !== id)
      return { ...s, products: products.length ? products : [blankProduct()] }
    })
  const clearExamples = () =>
    setState((s) => {
      const products = s.products.filter((p) => !p.example)
      return { ...s, products: products.length ? products : [blankProduct()], selectedId: null }
    })

  const [xlsxBusy, setXlsxBusy] = useState(false)
  const downloadExcel = async () => {
    setXlsxBusy(true)
    try {
      const { buildWorkbook } = await import('./excel')
      const blob = await buildWorkbook(state.products, state.roas, state.target)
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'KK Break-even ROAS Calculator.xlsx'
      document.body.appendChild(a)
      a.click()
      a.remove()
      setTimeout(() => URL.revokeObjectURL(url), 1000)
    } finally {
      setXlsxBusy(false)
    }
  }

  const full = state.products.length >= MAX_PRODUCTS

  return (
    <>
      <div className="screen">
        <header className="site-header">
          <div className="wrap header-row">
            <div className="brand">
              <img src={LOGO} alt="" width="36" height="36" />
              <span>K&amp;K Media</span>
            </div>
            <a className="header-link" href={LINKEDIN_MESSAGE} target="_blank" rel="noopener noreferrer">
              Message me
            </a>
          </div>
        </header>

        <main className="wrap">
          <section className="hero">
            <p className="eyebrow">Free tool for DTC brands</p>
            <h1>Break-even ROAS calculator</h1>
            <p className="lead">
              See the ROAS where each product stops losing money on the first order, and the ROAS you need to hit your
              profit target.
            </p>
          </section>

          <section className="card roas-card" aria-labelledby="roas-label">
            <div className="roas-main">
            <label id="roas-label" htmlFor="roas" className="roas-label">
              Your first-order ROAS
            </label>
            <div className="roas-row">
              <input
                id="roas"
                className={`roas-input${roasParsed.error ? ' invalid' : ''}`}
                inputMode="decimal"
                autoComplete="off"
                value={state.roas}
                onChange={(e) => setRoas(e.target.value)}
                aria-invalid={!!roasParsed.error}
                aria-describedby="roas-hint roas-err"
              />
              <span className="roas-x" aria-hidden="true">
                x
              </span>
            </div>
            {roasParsed.error && (
              <p id="roas-err" className="field-error" role="alert">
                {roasParsed.error}
              </p>
            )}
            </div>
            <div className="roas-main">
              <label htmlFor="target" className="roas-label">
                Target margin after ads
              </label>
              <div className="roas-row">
                <input
                  id="target"
                  className={`roas-input${targetParsed.error ? ' invalid' : ''}`}
                  inputMode="decimal"
                  autoComplete="off"
                  value={state.target}
                  onChange={(e) => setTarget(e.target.value)}
                  aria-invalid={!!targetParsed.error}
                  aria-describedby="target-hint target-err"
                />
                <span className="roas-x" aria-hidden="true">
                  %
                </span>
              </div>
              {targetParsed.error ? (
                <p id="target-err" className="field-error" role="alert">
                  {targetParsed.error}
                </p>
              ) : (
                <p id="target-hint" className="target-hint">
                  What each order should leave after ad spend to pay opex and profit. Drives Target ROAS.
                </p>
              )}
            </div>
            <p id="roas-hint" className="hint">
              Use ROAS on first orders from new customers. Blended ROAS includes repeat buyers and makes the first order
              look better than it is.
            </p>
          </section>

          <section aria-labelledby="products-h">
            <div className="section-head">
              <h2 id="products-h">Your products</h2>
              <div className="toolbar">
                {hasExamples && (
                  <button type="button" className="btn ghost" onClick={clearExamples}>
                    Clear examples
                  </button>
                )}
                <button type="button" className="btn ghost" onClick={downloadExcel} disabled={xlsxBusy}>
                  {xlsxBusy ? 'Preparing file…' : 'Download Excel'}
                </button>
                <button type="button" className="btn ghost" onClick={() => window.print()}>
                  Print / Save as PDF
                </button>
              </div>
            </div>
            <p className="help-note">
              Stuck on a number or not sure where to find it?{' '}
              <a href={LINKEDIN_MESSAGE} target="_blank" rel="noopener noreferrer">
                Message me
              </a>{' '}
              and I'll help.
            </p>

            <div className="cards">
              {state.products.map((p, i) => (
                <ProductCard
                  key={p.id}
                  p={p}
                  index={i}
                  ev={evals[i]}
                  roasOk={roas !== null}
                  canAdd={!full}
                  onChange={update}
                  onDuplicate={duplicate}
                  onRemove={remove}
                />
              ))}
            </div>

            <ProductTable
              products={state.products}
              evals={evals}
              roasOk={roas !== null}
              canAdd={!full}
              onChange={update}
              onDuplicate={duplicate}
              onRemove={remove}
            />

            <div className="add-row">
              <button type="button" className="btn primary" onClick={add} disabled={full}>
                + Add product
              </button>
              <span className="muted small">
                {state.products.length} of {MAX_PRODUCTS} products
              </span>
            </div>

            <MixCta evals={evals} roas={roas} />
          </section>

          <DiscountBlock
            products={state.products}
            selectedId={state.selectedId}
            onSelect={(selectedId) => setState((s) => ({ ...s, selectedId }))}
          />

          <Notes />
        </main>

        <Footer />
      </div>

      <PrintView state={state} evals={evals} roas={roas} />
    </>
  )
}

// Under the products: a status count, not a total. Break-even ROAS does not add up across
// products; the account-wide number depends on how ad spend splits, which we invite people to ask about.
function MixCta({ evals, roas }: { evals: Evaluation[]; roas: number | null }) {
  const ok = evals.filter((e): e is Extract<Evaluation, { kind: 'ok' }> => e.kind === 'ok')
  const below = ok.filter((e) => !e.result.above).length
  const showCount = roas !== null && ok.length > 0
  return (
    <div className="mix-cta">
      {showCount && (
        <p className="mix-count">
          At {roas.toFixed(2)}x,{' '}
          {below === 0 ? (
            <strong className="pos">
              {ok.length === 1 ? 'your product covers its ad spend' : `all ${ok.length} products cover their ad spend`}
            </strong>
          ) : (
            <>
              <strong className="neg">
                {below} of {ok.length} products
              </strong>{' '}
              {below === 1 ? 'loses' : 'lose'} money
            </>
          )}{' '}
          on the first order.
        </p>
      )}
      <div className="mix-body">
        <div>
          <p className="mix-title">Need one ROAS target for the whole account?</p>
          <p className="muted">
            There is no simple total here. It depends on how your ad spend splits across these products. Send me your
            mix and I will work it out with you.
          </p>
        </div>
        <a className="btn primary" href={LINKEDIN_MESSAGE} target="_blank" rel="noopener noreferrer">
          Message me
        </a>
      </div>
    </div>
  )
}

// Product card (phone)

interface RowProps {
  p: Product
  index: number
  ev: Evaluation
  roasOk: boolean
  canAdd: boolean
  onChange: (id: string, key: FieldKey, value: string) => void
  onDuplicate: (id: string) => void
  onRemove: (id: string) => void
}

function NumInput({
  p,
  k,
  error,
  onChange,
  labelledBy,
  compact,
}: {
  p: Product
  k: FieldKey
  error?: string
  onChange: RowProps['onChange']
  labelledBy?: string
  compact?: boolean
}) {
  const id = `${p.id}-${k}${compact ? '-t' : ''}`
  const isPct = PERCENT_FIELDS.includes(k)
  return (
    <div className={`num${error ? ' invalid' : ''}`}>
      {!isPct && <span className="adorn pre">$</span>}
      <input
        id={id}
        inputMode="decimal"
        autoComplete="off"
        value={p[k]}
        placeholder="0"
        onChange={(e) => onChange(p.id, k, e.target.value)}
        aria-invalid={!!error}
        aria-labelledby={labelledBy}
        aria-label={labelledBy ? undefined : FIELD_LABELS[k]}
        aria-describedby={error ? `${id}-err` : undefined}
        className={isPct ? 'has-post' : 'has-pre'}
      />
      {isPct && <span className="adorn post">%</span>}
      {error && (
        <span id={`${id}-err`} className="field-error" role="alert">
          {error}
        </span>
      )}
    </div>
  )
}

function RowActions({ p, index, canAdd, onDuplicate, onRemove }: Pick<RowProps, 'p' | 'index' | 'canAdd' | 'onDuplicate' | 'onRemove'>) {
  const name = productLabel(p, index)
  return (
    <div className="row-actions">
      <button
        type="button"
        className="icon-btn"
        onClick={() => onDuplicate(p.id)}
        disabled={!canAdd}
        aria-label={`Duplicate ${name}`}
        title="Duplicate"
      >
        <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
          <rect x="8" y="8" width="12" height="12" rx="2" fill="none" stroke="currentColor" strokeWidth="2" />
          <path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2" fill="none" stroke="currentColor" strokeWidth="2" />
        </svg>
      </button>
      <button type="button" className="icon-btn" onClick={() => onRemove(p.id)} aria-label={`Delete ${name}`} title="Delete">
        <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
          <path d="M5 7h14M10 7V5h4v2M7 7l1 12h8l1-12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinejoin="round" />
        </svg>
      </button>
    </div>
  )
}

function ProductCard({ p, index, ev, roasOk, canAdd, onChange, onDuplicate, onRemove }: RowProps) {
  const errors = ev.kind === 'invalid' ? ev.errors : {}
  const nameId = `${p.id}-name`
  return (
    <article className="card product-card" aria-label={productLabel(p, index)}>
      <div className="pc-head">
        <div className="pc-name">
          <label htmlFor={nameId} className="field-label">
            Product name {p.example && <span className="badge">Example</span>}
          </label>
          <input
            id={nameId}
            className="text-input"
            value={p.name}
            placeholder={`Product ${index + 1}`}
            onChange={(e) => onChange(p.id, 'name', e.target.value)}
            autoComplete="off"
          />
        </div>
        <RowActions p={p} index={index} canAdd={canAdd} onDuplicate={onDuplicate} onRemove={onRemove} />
      </div>

      <div className="pc-fields">
        {INPUT_FIELDS.map((k) => (
          <div key={k} className={`pc-field${k === 'target' ? ' wide' : ''}`}>
            <span id={`${p.id}-${k}-lbl`} className="field-label">
              {FIELD_LABELS[k]}
            </span>
            <NumInput p={p} k={k} error={errors[k]} onChange={onChange} labelledBy={`${p.id}-${k}-lbl`} />
          </div>
        ))}
      </div>

      <CardResult ev={ev} roasOk={roasOk} />
    </article>
  )
}

function StatusPill({ above }: { above: boolean }) {
  return <span className={`pill ${above ? 'up' : 'down'}`}>{above ? 'Above break-even' : 'Below break-even'}</span>
}

function CardResult({ ev, roasOk }: { ev: Evaluation; roasOk: boolean }) {
  if (ev.kind === 'empty') return <p className="pc-empty muted">Enter a price and costs to see results.</p>
  if (ev.kind === 'invalid') return <p className="pc-empty warn">Fix the highlighted fields to see results.</p>
  const r = ev.result
  return (
    <div className="pc-result" aria-live="polite">
      <div className="big-pair">
        <div>
          <span className="metric-label">Break-even ROAS</span>
          <span className="big">{fmtRoas(r.breakEvenRoas)}</span>
        </div>
        <div>
          <span className="metric-label">Target ROAS</span>
          <span className="big secondary">{fmtRoas(r.targetRoas)}</span>
        </div>
      </div>
      {roasOk && (
        <div className={`after-ads ${r.above ? 'up' : 'down'}`}>
          <StatusPill above={r.above} />
          <span>
            <strong>{fmtMoney(r.afterAdsPer100, true)}</strong> per $100 of ads after ad spend
          </span>
        </div>
      )}
      <dl className="mini">
        <div>
          <dt>Net price</dt>
          <dd>{fmtMoney(r.netPrice)}</dd>
        </div>
        <div>
          <dt>Contribution per order</dt>
          <dd>{fmtMoney(r.contribution)}</dd>
        </div>
        <div>
          <dt>Contribution margin</dt>
          <dd>{fmtPct(r.margin)}</dd>
        </div>
        <div>
          <dt>Max CPA at break-even</dt>
          <dd>{fmtMoney(r.maxCpa)}</dd>
        </div>
      </dl>
    </div>
  )
}

// Product table (desktop)

function ProductTable({
  products,
  evals,
  roasOk,
  canAdd,
  onChange,
  onDuplicate,
  onRemove,
}: {
  products: Product[]
  evals: Evaluation[]
  roasOk: boolean
  canAdd: boolean
  onChange: RowProps['onChange']
  onDuplicate: RowProps['onDuplicate']
  onRemove: RowProps['onRemove']
}) {
  return (
    <div className="table-wrap card">
      <table className="ptable">
        <thead>
          <tr className="group-row">
            <th colSpan={1 + INPUT_FIELDS.length} scope="colgroup">
              Your inputs
            </th>
            <th colSpan={5} scope="colgroup" className="res-start">
              Results
            </th>
            <th aria-label="Actions" />
          </tr>
          <tr>
            {(['name', ...INPUT_FIELDS] as FieldKey[]).map((k) => (
              <th key={k} id={`th-${k}`} scope="col" title={FIELD_LABELS[k]}>
                {SHORT[k].title}
                {SHORT[k].unit && <span className="unit">{SHORT[k].unit}</span>}
              </th>
            ))}
            <th scope="col" className="res-start">
              Margin
            </th>
            <th scope="col">Break-even ROAS</th>
            <th scope="col">
              Target ROAS<span className="unit">at your target margin</span>
            </th>
            <th scope="col">
              Max CPA<span className="unit">at break-even</span>
            </th>
            <th scope="col">
              After ads<span className="unit">per $100 of ads, status</span>
            </th>
            <th aria-hidden="true" />
          </tr>
        </thead>
        <tbody>
          {products.map((p, i) => {
            const ev = evals[i]
            const errors = ev.kind === 'invalid' ? ev.errors : {}
            return (
              <tr key={p.id}>
                <td className="name-cell">
                  {p.example && <span className="badge">Example</span>}
                  <input
                    className="text-input"
                    value={p.name}
                    placeholder={`Product ${i + 1}`}
                    aria-label={`Product name, row ${i + 1}`}
                    onChange={(e) => onChange(p.id, 'name', e.target.value)}
                    autoComplete="off"
                  />
                </td>
                {INPUT_FIELDS.map((k) => (
                  <td key={k}>
                    <NumInput p={p} k={k} error={errors[k]} onChange={onChange} compact />
                  </td>
                ))}
                <TableResult ev={ev} roasOk={roasOk} />
                <td>
                  <RowActions p={p} index={i} canAdd={canAdd} onDuplicate={onDuplicate} onRemove={onRemove} />
                </td>
              </tr>
            )
          })}
        </tbody>
      </table>
    </div>
  )
}

function TableResult({ ev, roasOk }: { ev: Evaluation; roasOk: boolean }) {
  if (ev.kind !== 'ok')
    return (
      <td colSpan={5} className={`res-start ${ev.kind === 'invalid' ? 'warn' : 'muted'}`}>
        {ev.kind === 'invalid' ? 'Fix the highlighted fields to see results' : 'Enter a price to see results'}
      </td>
    )
  const r = ev.result
  return (
    <>
      <td className="res-start num-cell">{fmtPct(r.margin)}</td>
      <td className="num-cell strong">{fmtRoas(r.breakEvenRoas)}</td>
      <td className="num-cell">{fmtRoas(r.targetRoas)}</td>
      <td className="num-cell">{fmtMoney(r.maxCpa)}</td>
      <td className={`num-cell after-cell ${roasOk ? (r.above ? 'pos' : 'neg') : ''}`}>
        {roasOk && (
          <>
            {fmtMoney(r.afterAdsPer100, true)}
            <StatusPill above={r.above} />
          </>
        )}
      </td>
    </>
  )
}

// Discount block

function DiscountBlock({
  products,
  selectedId,
  onSelect,
}: {
  products: Product[]
  selectedId: string | null
  onSelect: (id: string) => void
}) {
  const valid = products
    .map((p, i) => ({ p, i, v: validate(p) }))
    .filter((x) => x.v.inputs && x.v.inputs.price > 0)
  const current = valid.find((x) => x.p.id === selectedId) ?? valid[0]

  return (
    <section className="card discount" aria-labelledby="disc-h">
      <h2 id="disc-h">What a discount does</h2>
      <p className="muted">
        Same product, same ads. A bigger discount in your ads means you need a higher ROAS just to break even.
      </p>
      {!current ? (
        <p className="muted">Add a product with a valid price and costs to see this.</p>
      ) : (
        <DiscountBody valid={valid} current={current} onSelect={onSelect} />
      )}
    </section>
  )
}

function DiscountBody({
  valid,
  current,
  onSelect,
}: {
  valid: { p: Product; i: number; v: ReturnType<typeof validate> }[]
  current: { p: Product; i: number; v: ReturnType<typeof validate> }
  onSelect: (id: string) => void
}) {
  const rows = discountTable(current.v.inputs!)
  return (
    <>
      <div className="select-row">
        <label htmlFor="disc-product" className="field-label">
          Product
        </label>
        <select id="disc-product" value={current.p.id} onChange={(e) => onSelect(e.target.value)}>
          {valid.map((x) => (
            <option key={x.p.id} value={x.p.id}>
              {productLabel(x.p, x.i)}
            </option>
          ))}
        </select>
        <span className="muted small">Uses its price, costs and fees. The discount column replaces its own.</span>
      </div>
      <div className="disc-grid">
        <table className="dtable">
          <caption className="sr-only">Margin and break-even ROAS by discount for {productLabel(current.p, current.i)}</caption>
          <thead>
            <tr>
              <th scope="col">Discount</th>
              <th scope="col">Net price</th>
              <th scope="col">Margin</th>
              <th scope="col">Break-even ROAS</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.discount}>
                <th scope="row">{r.discount}%</th>
                <td>{fmtMoney(r.netPrice)}</td>
                <td className={r.margin !== null && r.margin <= 0 ? 'neg' : ''}>{r.margin === null ? '' : fmtPct(r.margin)}</td>
                <td className="strong">
                  {fmtRoas(r.breakEvenRoas)}
                  {r.changeVsNone !== null && r.discount > 0 && (
                    <span className="delta"> +{r.changeVsNone.toFixed(2)}</span>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <DiscountChart rows={rows} />
      </div>
    </>
  )
}

// Notes and footer

const NOTES: [string, string][] = [
  ['Contribution per order', 'Net price minus COGS, shipping, other variable cost and fees. Net price is price after the discount in your ads.'],
  ['Contribution margin', 'Contribution divided by net price.'],
  ['Break-even ROAS', '1 divided by contribution margin. Below it, every first order loses money.'],
  ['Target ROAS', '1 divided by (contribution margin minus your target contribution). The ROAS you need to also cover opex and profit.'],
  ['Max CPA at break-even', 'Equal to contribution per order: the most you can pay for one first order.'],
  ['After ads per $100', '(Your first-order ROAS × contribution margin minus 1) × 100. Before opex, so it is not net profit.'],
]

function Notes() {
  return (
    <section className="card notes" aria-labelledby="notes-h">
      <h2 id="notes-h">How the numbers work</h2>
      <dl>
        {NOTES.map(([t, d]) => (
          <div key={t}>
            <dt>{t}</dt>
            <dd>{d}</dd>
          </div>
        ))}
      </dl>
      <p className="callout">
        First order only, before opex. Subscription and repeat purchases change the picture: judge them by cohort payback
        days.
      </p>
    </section>
  )
}

function Footer() {
  return (
    <footer className="site-footer">
      <div className="wrap footer-inner">
        <div className="footer-cta">
          <p className="footer-title">Want a second pair of eyes on your numbers?</p>
          <a className="btn primary big" href={LINKEDIN_MESSAGE} target="_blank" rel="noopener noreferrer">
            Message me on LinkedIn
          </a>
        </div>
        <div className="footer-meta">
          <img src={LOGO} alt="" width="28" height="28" />
          <span>
            Built by{' '}
            <a href={LINKEDIN} target="_blank" rel="noopener noreferrer">
              Artem Kalachev
            </a>
            , K&amp;K Media
          </span>
        </div>
        <p className="privacy small muted">No tracking, no cookies. Your numbers stay in this browser.</p>
      </div>
    </footer>
  )
}

// Print view

function PrintView({ state, evals, roas }: { state: State; evals: Evaluation[]; roas: number | null }) {
  const date = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' })
  return (
    <div className="print-only" aria-hidden="true">
      <div className="pv-head">
        <img src={LOGO} alt="" width="40" height="40" />
        <div>
          <h1>Break-even ROAS calculator</h1>
          <p>
            K&amp;K Media · {date} · First-order ROAS: <strong>{state.roas || 'not set'}</strong>
          </p>
        </div>
      </div>
      <table className="pv-table">
        <thead>
          <tr>
            <th>Product</th>
            <th>Price</th>
            <th>COGS</th>
            <th>Shipping</th>
            <th>Fees</th>
            <th>Other</th>
            <th>Discount</th>
            <th>Target</th>
            <th>Margin</th>
            <th>Break-even ROAS</th>
            <th>Target ROAS</th>
            <th>Max CPA</th>
            <th>After ads / $100</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          {state.products.map((p, i) => {
            const ev = evals[i]
            const r = ev.kind === 'ok' ? ev.result : null
            const pct = (v: string) => (v.trim() ? `${v}%` : '')
            const usd = (v: string) => (v.trim() ? `$${v}` : '')
            return (
              <tr key={p.id}>
                <td>{productLabel(p, i)}</td>
                <td>{usd(p.price)}</td>
                <td>{usd(p.cogs)}</td>
                <td>{usd(p.shipping)}</td>
                <td>{pct(p.fees)}</td>
                <td>{usd(p.other)}</td>
                <td>{pct(p.discount)}</td>
                <td>{pct(state.target)}</td>
                {r ? (
                  <>
                    <td>{fmtPct(r.margin)}</td>
                    <td>
                      <strong>{fmtRoas(r.breakEvenRoas)}</strong>
                    </td>
                    <td>{fmtRoas(r.targetRoas)}</td>
                    <td>{fmtMoney(r.maxCpa)}</td>
                    <td>{roas !== null ? fmtMoney(r.afterAdsPer100, true) : ''}</td>
                    <td>{roas !== null ? (r.above ? 'Above break-even' : 'Below break-even') : ''}</td>
                  </>
                ) : (
                  <td colSpan={6}>{ev.kind === 'invalid' ? 'Check inputs' : ''}</td>
                )}
              </tr>
            )
          })}
        </tbody>
      </table>
      <h2>How the numbers work</h2>
      <ul>
        {NOTES.map(([t, d]) => (
          <li key={t}>
            <strong>{t}:</strong> {d}
          </li>
        ))}
        <li>
          <strong>First order only, before opex.</strong> Subscription and repeat purchases change the picture: judge them
          by cohort payback days.
        </li>
      </ul>
      <p className="pv-foot">Built by Artem Kalachev, K&amp;K Media. linkedin.com/in/artem-kalachev</p>
    </div>
  )
}
