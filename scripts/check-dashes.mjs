// Fails if any interface text contains an en dash (U+2013) or em dash (U+2014).
import { readFileSync, readdirSync, statSync, existsSync } from 'node:fs'
import { join } from 'node:path'

const roots = ['src', 'index.html', 'dist']
const exts = /\.(tsx?|css|html|js|svg|json)$/
const bad = /[–—]/
const hits = []
const walk = (p) => {
  if (!existsSync(p)) return
  if (statSync(p).isDirectory()) return readdirSync(p).forEach((f) => walk(join(p, f)))
  if (!exts.test(p)) return
  readFileSync(p, 'utf8').split('\n').forEach((line, i) => {
    if (bad.test(line)) hits.push(`${p}:${i + 1}: ${line.trim().slice(0, 120)}`)
  })
}
roots.forEach(walk)
if (hits.length) {
  console.error(`Found ${hits.length} line(s) with en or em dashes:\n` + hits.join('\n'))
  process.exit(1)
}
console.log('No en or em dashes found.')
