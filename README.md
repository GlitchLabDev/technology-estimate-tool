# Technology Estimate Tool

A React + Vite + TypeScript app that reproduces the Excel estimate logic from the two provided workbooks in the browser.

## How it works

- `public/workbook.json` contains the combined contents of `Technology_Estimate_Template-71626.xlsx` and `High_Level_costs_for_Network_kit_(1).xlsx`.
- `src/App.tsx` loads the HyperFormula spreadsheet engine, maps the `Test Fit` inputs to the workbook, and reads the computed `Tech Estimate Summary` and `Estimate Sheet` outputs.
- The grouped input form lets users set square footage, seats, room counts, and security quantities.
- The summary table shows Cap, Non-Cap, and Project totals, plus detailed line items that can be exported to CSV.

## Regenerate `workbook.json`

If the Excel files change, run:

```bash
python3 convert.py
```

This script parses both workbooks, trims empty rows/columns, removes external workbook link markers, and patches an obvious typo (`'Test Fit'!E341` -> `'Test Fit'!E34`).

## Local development

```bash
cd /home/ubuntu/estimator-app
npm install
npm run dev
```

## Production build

```bash
npm run build
npm run preview
```

## Lovable / GitHub import

This is a standard Vite + React + TypeScript project with `package.json` at the root, so it can be imported into Lovable or deployed from GitHub to most static hosts. The build output is in `dist/`.

## License note

HyperFormula is used under the `gpl-v3` license key. For commercial production use, replace it with a valid proprietary HyperFormula license.
