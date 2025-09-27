# Mutual Fund Dashboard (Vite + React + TypeScript)

Zero-backend. Import your `MFPortfolio.xlsx`, press **Refresh NAVs**, see KPIs, table and charts, then **Export Excel**.

## Deploy on Vercel (no installs)
1. Create this repo on GitHub (files at root).
2. Vercel → Add New Project → Import this repo.
   - Framework: **Vite**
   - Build: `npm run build`
   - Output: `dist`

## Excel requirements
- Sheet name: `MFPortfolio`
- Columns (in order): see app header list in code or exported Excel.
