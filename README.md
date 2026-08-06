# eLS — Enhanced Link Support Portal

Internal operations portal for the Invalid Claims & Repay Management team.  
Covers customer directories, portals, reason codes, job aids, templates, and the eLS Compass decision wizard.

---

## Deploying to GitHub Pages (Personal — github.com)

### Prerequisites
- A **github.com** account
- Git installed on your machine

### Steps

**1. Create the repository on github.com**

Go to → https://github.com/new

```
Repository name:  els-portal       (or any name you prefer)
Visibility:       Private          ← keeps it off the public internet
                                     (GitHub Pages still works on private repos)
```
Click **Create repository**.

---

**2. Push this folder from PowerShell**

Open PowerShell, navigate to this folder, then run:

```powershell
cd "C:\Users\eLNunez\.bob\playground\eLS\Operations\Reports\els-app"

git init
git add .
git commit -m "Initial deploy"
git remote add origin https://github.com/<YOUR-USERNAME>/els-portal.git
git push -u origin main
```

> Replace `<YOUR-USERNAME>` with your GitHub username.  
> GitHub will prompt for your credentials the first time.

---

**3. Enable GitHub Pages**

```
github.com/<YOUR-USERNAME>/els-portal
  → Settings
    → Pages (left sidebar)
      → Source: GitHub Actions    ← select this
```

The workflow in `.github/workflows/deploy.yml` triggers automatically.  
Wait ~1 minute, then refresh the Pages settings — your URL will appear.

---

**4. Your site is live at**

```
https://<YOUR-USERNAME>.github.io/els-portal/
```

Share this link with your team. Anyone with the link can access it  
(no login required — it is URL-access only since GitHub Pages on private  
repos is still publicly reachable by URL).

> **Want stricter access?** Upgrade to GitHub Pro ($4/mo) to enable  
> "Private Pages" — which requires a GitHub login to view.  
> For a free alternative, see the section below.

---

## Updating Content

After editing any source file, rebuild `index.html` then push:

```powershell
# 1. Edit src/main.js, src/style.css, or public/compass.html

# 2. Rebuild the self-contained index.html
.\build-inline.ps1

# 3. Commit and push — site redeploys automatically in ~1 min
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

## Project Structure

```
els-app/
├── index.html              ← self-contained app (CSS + JS + compass inlined)
├── src/
│   ├── main.js             ← all data + application logic
│   └── style.css           ← all styles
├── public/
│   └── compass.html        ← eLS Compass decision wizard (source)
├── .github/
│   └── workflows/
│       └── deploy.yml      ← GitHub Actions auto-deploy on push to main
├── .nojekyll               ← required for GitHub Pages to serve files correctly
├── build-inline.ps1        ← PowerShell script: rebuild index.html from sources
└── README.md
```

---

## Notes

- The portal contains **no backend, no database, no login system**.  
  All data lives in `src/main.js` as JavaScript arrays.
- `index.html` is fully self-contained — works as a **local file** too  
  (double-click to open in a browser, no server needed).
- Since data includes internal contacts and emails, keep the repo **Private**  
  and share the Pages URL only with your team.
