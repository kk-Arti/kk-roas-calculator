# Break-even ROAS Calculator | K&K Media

Static web calculator: per product break-even ROAS, target ROAS, max CPA and contribution after ads on the first order.
Logic matches `KK_Break-even_ROAS_Calculator.xlsx` (checked by `npm test`).

- No backend, no analytics, no cookies, no third party scripts or fonts. Inputs live only in the browser's localStorage.
- Stack: Vite + React + TypeScript. Font: Montserrat via @fontsource (bundled).
- Deploy: GitHub Pages via `.github/workflows/deploy.yml` on push to `main`.

## Structure

- `src/calc.ts` pure logic (formulas, validation, discount table, CSV)
- `src/calc.test.ts` vitest, examples from the xlsx
- `src/App.tsx` UI: ROAS input, product cards (mobile) / table (desktop), discount block, notes, footer
- `public/` logo, favicon, Open Graph image 1200x627
- `scripts/check-dashes.mjs` fails if any UI text has an en or em dash

## Commands

    npm install
    npm run dev
    npm test
    npm run check:dashes
    npm run build
