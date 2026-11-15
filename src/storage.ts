import { EXAMPLES, type Product } from './calc'

const KEY = 'kk-roas-calculator:v1'

export interface State {
  roas: string
  products: Product[]
  selectedId: string | null
}

let seq = 0
export const newId = () => `p${Date.now().toString(36)}${(seq++).toString(36)}`

export const blankProduct = (): Product => ({
  id: newId(),
  name: '',
  price: '',
  cogs: '',
  shipping: '',
  fees: '',
  other: '',
  discount: '',
  target: '',
})

export const exampleState = (): State => ({
  roas: '2.2',
  products: EXAMPLES.map((e) => ({ ...e, id: newId() })),
  selectedId: null,
})

export function load(): State {
  try {
    const raw = window.localStorage.getItem(KEY)
    if (!raw) return exampleState()
    const s = JSON.parse(raw) as State
    if (typeof s.roas !== 'string' || !Array.isArray(s.products)) return exampleState()
    return s
  } catch {
    return exampleState()
  }
}

export function save(s: State) {
  try {
    window.localStorage.setItem(KEY, JSON.stringify(s))
  } catch {
    // storage blocked: the calculator still works for this visit
  }
}
