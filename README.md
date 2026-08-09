# eLS — Enhanced Link Support Portal

Internal operations portal for the Invalid Claims & Repay Management team.  
Covers customer directories, portals, reason codes, job aids, templates, and the eLS Compass decision wizard.

---

## Deploying to GitHub Pages

### One-time setup

**1. Create the repository on github.com**

Go to → https://github.com/new

```
Repository name:  els-portal       (or any name you prefer)
Visibility:       Private          ← keeps it off the public internet
```

Click **Create repository**.

---

**2. Push this folder from PowerShell**

```powershell
cd "C:\Users\eLNunez\.bob\playground\eLS\Operations\Reports\els-app"

git init
git add .
git commit -m "Initial deploy"
git remote add origin https://github.com/<YOUR-USERNAME>/els-portal.git
git push -u origin main
```

> Replace `<YOUR-USERNAME>` with your GitHub username.

---

**3. Enable GitHub Pages**

```
github.com/<YOUR-USERNAME>/els-portal
  → Settings
    → Pages (left sidebar)
      → Source: GitHub Actions    ← select this
```

The workflow in `.github/workflows/deploy.yml` runs automatically on every push.  
Wait ~1 minute, then refresh Pages settings — your URL will appear.

---

**4. Your site is live at**

```
https://<YOUR-USERNAME>.github.io/els-portal/
```

---

## Updating Content

Edit `src/main.js` (or any source file), then push. **No local build step needed** —  
GitHub Actions inlines everything and deploys automatically.

```powershell
# 1. Edit src/main.js  (data), public/compass.html, or index.html (layout)

# 2. Commit and push — site redeploys in ~1 minute
git add .
git commit -m "Update portal content"
git push
```

| What to update | File |
|---|---|
| Customer directory data | `src/main.js` → `DIRECT`, `ACOSTA`, `NATDIST` arrays |
| Portal links | `src/main.js` → `PORTALS` array |
| Reason codes | `src/main.js` → `RC` array |
| Repay letter links | `src/main.js` → `REPAY_LETTERS` object |
| BPS / Customer SOP | `src/main.js` → `BPS_DATA` array |
| eLS Compass wizard | `public/compass.html` |
| Styling | `src/style.css` |
| Page layout / HTML | `index.html` |

---

## How the build works

On every push to `main`, GitHub Actions runs `.github/workflows/deploy.yml`:

1. Checks out the repo
2. Runs an inline Python script that:
   - Reads `src/main.js` and inlines it into `index.html`
   - Reads `public/compass.html` and inlines it as an `srcdoc` iframe
   - Writes the fully self-contained `index.html`
3. Uploads and deploys to GitHub Pages

The built `index.html` also works as a **local file** — double-click to open in a browser, no server needed.

---

## Project Structure

```
els-app/
├── index.html              ← page layout + CSS (JS/compass are inlined at build time)
├── src/
│   ├── main.js             ← all data + application logic
│   └── style.css           ← all styles
├── public/
│   └── compass.html        ← eLS Compass decision wizard (source)
├── .github/
│   └── workflows/
│       └── deploy.yml      ← GitHub Actions: builds index.html, deploys to Pages
├── .nojekyll               ← required for GitHub Pages to serve files correctly
├── build-inline.ps1        ← optional: run locally to rebuild index.html manually
└── README.md
```

---

## Notes

- The portal contains **no backend, no database, no login system**.  
  All data lives in `src/main.js` as JavaScript arrays.
- Since data includes internal contacts and emails, keep the repo **Private**  
  and share the Pages URL only with your team.
- `index.html` is committed with a `src=` reference to `main.js` (not inlined).  
  The inline build happens in CI. To build locally, run `.\build-inline.ps1`.
