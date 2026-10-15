# Sales Reporting

Mobile form: pick a store, enter VIP / SKY / ARI / VIP Total / VIP MTD / AT / KT / AT Total / AT MTD / Safari / Safari MTD, hit Submit — values land in that store's row in the Google Sheet.

## One-time setup
1. Open the sheet → Extensions → Apps Script → paste `Code.gs` → Deploy → New deployment → Web app (Execute as *Me*, access *Anyone*).
2. Copy the `/exec` URL into `SCRIPT_URL` in `index.html`, commit, push.
3. Enable GitHub Pages (main branch, root).
