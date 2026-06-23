# Preventa ASICS S1-27 — Dashboard

Password-protected, public static dashboard of the ASICS S1-27 preline footwear order.
Pie charts break the order down by **Género**, **Categoría** (Running / CPS / Sportstyle) and
**Modelo**, with buttons to switch the quantity view between **Compra**, **Extra** and **Cancel**.

The published `index.html` is **AES-256-GCM encrypted**: without the password it is unreadable
gibberish, so it's safe in a public repo. Decryption happens in the browser (Web Crypto API).

## Requirements
- Python 3 with `openpyxl` and `cryptography` (already installed on the build machine).

## Build / rebuild

```bash
python site/build.py --password "YOUR-PASSWORD"
```

This reads `PRELINE ASICS S1-27 footwear.xlsx` (kept locally, git-ignored), aggregates the
`Preline order` sheet, and regenerates `index.html`. The script aborts if the unit totals don't
match the expected figures (Compra 76,764 / Extra 2,844 / Cancel 888).

**To change the password** or **update the data** (drop in a new workbook of the same shape):
just rerun the command and commit the new `index.html`.

## Test locally
`crypto.subtle` only works on `https://` or `localhost`, so open via a local server (not file://):

```bash
python -m http.server 8000
# then open http://localhost:8000/  and enter the password
```

## Publish to GitHub Pages
1. Create a new **public** repo on github.com (e.g. `preventa-asics`). Don't add any files.
2. From this folder:
   ```bash
   git init
   git add .
   git commit -m "Preventa ASICS dashboard"
   git branch -M main
   git remote add origin https://github.com/<you>/preventa-asics.git
   git push -u origin main
   ```
3. Repo **Settings → Pages → Source: Deploy from a branch → `main` / `/ (root)` → Save**.
4. After ~1 minute the site is live at `https://<you>.github.io/preventa-asics/`.

`.gitignore` keeps the raw Excel file and the plaintext `site/app.html` out of the repo — only the
encrypted `index.html` is published.
