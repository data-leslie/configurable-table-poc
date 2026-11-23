# Configurable Table — GitHub Pages PoC

Proves that the same table, hosted as a **network-enabled** Tableau extension
(not a sandboxed data app), can do the two things the sandbox blocked:

- **Real file downloads** (CSV + PDF)
- **Per-user saved layouts** (localStorage)

…and that a **Viewer-licence** user can use all of it inside Tableau.

Uses a small **sample dataset** so the PoC validates the *environment* with no
data-access dependency. Real data wiring + the Azure move come after it's proven.

---

## Step 1 — Put these files on GitHub Pages

1. Create a **new GitHub repo** (public is fine for the PoC), e.g. `configurable-table-poc`.
2. Upload the **contents of this folder** to the repo root:
   - `index.html`
   - `ConfigurableTablePoC.trex`
   - `src/` (the whole folder: `app.js`, `styles.css`, `tableau.extensions.1.latest.min.js`, `vendor/`)
   (Drag-and-drop in the GitHub web UI → “Add file → Upload files”, or `git push`.)
3. Repo **Settings → Pages →** Source: **Deploy from a branch**, Branch: **main / (root)** → Save.
4. Wait ~1 minute. GitHub shows your site URL, e.g.
   `https://your-username.github.io/configurable-table-poc/`

**Quick check:** open
`https://your-username.github.io/configurable-table-poc/index.html`
directly in your browser — you should see the table render with sample data,
and **Download → CSV** should download a file. (This already proves hosting +
downloads work; the next steps put it inside Tableau.)

## Step 2 — Point the manifest at your URL

Edit **`ConfigurableTablePoC.trex`** → `<source-location><url>` → replace the
placeholder with your real page URL ending in `/index.html`, e.g.:

```
<url>https://your-username.github.io/configurable-table-poc/index.html</url>
```

Commit it. (Keep a copy of this `.trex` handy — you select it in Tableau.)

## Step 3 — Allow the extension in Tableau (admin)

Tableau Cloud → **Settings → Extensions**:
1. Under **Dashboard & Viz Extensions**, ensure extensions are enabled.
2. **Add to the safe list** your origin — `https://your-username.github.io` — set
   **Full Data access = Allowed**, and allow user prompts as needed.

## Step 4 — Add it to a dashboard

1. New (or existing) **Dashboard**.
2. From **Objects**, drag **Extension** onto the dashboard.
3. In the dialog choose **Access Local Extensions** → select your
   `ConfigurableTablePoC.trex` file.
4. Accept the permissions prompt (Full Data). The table renders.
5. **Publish** the workbook (to a normal project, not Personal Space, so you can
   share it to a test Viewer).

## Step 5 — Test — especially as a Viewer

As yourself first, then **signed in as a Viewer-licence user**:

- Table renders, filters/sort/currency work.
- **Columns ▾** → toggle columns (client-side show/hide).
- **Download → CSV** and **→ PDF** → **real files download**.
- Rearrange columns → **Save view** → reload the dashboard → layout **persists**.
- **⚙ Diagnostics** → confirms `localStorage: WORKS` and a test download fires.

If a Viewer can download and save a layout → the self-hosted approach is proven.
Next: wire real data (read the worksheet’s data or the VizQL Data Service) and
move hosting to your **Azure Static Web App** for production.

---

### Notes
- Everything is client-side; the sample data never leaves the browser.
- For production in Azure, the same files + `.trex` are used — only the URL
  changes, and you can add a CSP that blocks all egress except Tableau.
