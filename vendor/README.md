# Vendored libraries

These two files are **not** in the repo — download them once and commit them
here. They are loaded from `./vendor/` rather than a public CDN so the app
works offline on the shop floor and so the exact bytes are pinned.

## 1. `xlsx.full.min.js` — SheetJS **0.19.3**

Download: `https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js`

Do **not** use the cdnjs/npm copy. The newest version published to npm is
0.18.5, which is vulnerable to CVE-2023-30533 (prototype pollution when
*reading* a crafted file). This app calls `XLSX.read()` on operator-supplied
files during bulk import, so that path is directly exposed. The fix ships only
from SheetJS's own CDN.

## 2. `chart.umd.min.js` — Chart.js **4.4.1** (or any pinned 4.x)

Download: `https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js`

Previously loaded as an unpinned `chart.js` URL, which resolves to whatever is
latest — a major release could have broken the charts without a single line of
this repo changing.
