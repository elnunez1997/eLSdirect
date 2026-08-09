# Operations Reports — eLS Tools

Internal tools portal for the Invalid Claims & Repay Management team.

---

## Apps

| File | Description |
|---|---|
| `Footprint_App.html` | eLS Enhanced Link Support portal (main app) |
| `Footprint_Directory.html` | Customer directory tool |
| `els-compass-enhanced-link-support.html` | eLS Compass — Enhanced Link Support wizard |
| `els-compass-invalids-dia-path-wizard.html` | eLS Compass — Invalids DIA Path wizard |
| `5-digit-upc-extractor.html` | 5-digit UPC extractor tool |

---

## Deploying to GitHub Pages

### One-time setup

**1. Create the repository on github.com**

Go to → https://github.com/new

```
Repository name:  operations-reports   (or any name you prefer)
Visibility:       Private              ← keeps it off the public internet
```

Click **Create repository**.

---

**2. Push this folder from PowerShell**

```powershell
cd "C:\Users\eLNunez\.bob\playground\Operations\Reports"

git remote add origin https://github.com/<YOUR-USERNAME>/operations-reports.git
git push -u origin main
```

> Replace `<YOUR-USERNAME>` with your GitHub username.

---

**3. Enable GitHub Pages**

```
github.com/<YOUR-USERNAME>/operations-reports
  → Settings
    → Pages (left sidebar)
      → Source: GitHub Actions    ← select this
```

---

**4. Your apps will be live at**

```
https://<YOUR-USERNAME>.github.io/operations-reports/Footprint_App.html
https://<YOUR-USERNAME>.github.io/operations-reports/Footprint_Directory.html
https://<YOUR-USERNAME>.github.io/operations-reports/els-compass-enhanced-link-support.html
```

---

## Updating Content

Edit any `.html` file, then push. GitHub Actions deploys automatically in ~1 minute.

```powershell
git add .
git commit -m "Update content"
git push
```

---

## Notes

- All apps are **fully self-contained HTML files** — no backend, no database, no login.
- Excel workbooks (`.xlsx`, `.xlsm`) are excluded from the repo via `.gitignore`.
- Keep the repository **Private** since files may contain internal contacts and data.
