/// <reference types="vitest" />
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Absolute URL for Open Graph (LinkedIn needs a full URL). Overridable in CI.
const SITE_URL = process.env.SITE_URL ?? 'https://kk-arti.github.io/kk-roas-calculator/'

// Strict CSP for the built page: nothing leaves the browser, nothing loads from other hosts.
const CSP =
  "default-src 'self'; img-src 'self' data:; style-src 'self' 'unsafe-inline'; font-src 'self'; script-src 'self'; connect-src 'none'; object-src 'none'; base-uri 'self'; form-action 'none'"

export default defineConfig(({ command }) => ({
  base: './',
  plugins: [
    react(),
    {
      name: 'og-urls-and-csp',
      transformIndexHtml: (html: string) =>
        html
          .replaceAll('%SITE_URL%', SITE_URL)
          .replaceAll('%OG_IMAGE%', `${SITE_URL}og-image.png`)
          .replace('<!--CSP-->', command === 'build' ? `<meta http-equiv="Content-Security-Policy" content="${CSP}" />` : ''),
    },
  ],
  server: { host: '127.0.0.1', port: 5180 },
  test: { environment: 'node' },
}))
