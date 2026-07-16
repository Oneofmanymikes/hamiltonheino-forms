# Naomi Davison — Campaign Forms System

Four embeddable web forms (**Commit to Vote · Request a Sign · Commit to Donate · Volunteer**) for
**Naomi Davison, Ontario Liberal candidate for York–Simcoe**.

Every submission → a row in a Google Sheet (one tab per form, each with an `In Liberalist?` checkbox) →
a notification email to **info@votenaomi.ca**. Forms are served by a single Google Apps Script web app,
embedded as iframes on the GitHub Pages "Get Involved" page and on any partner site. Supporters are pushed
to **ontarioliberal.ca/yorksimcoe** for more information.

## Files

| File | What it is |
|---|---|
| `apps-script/Code.gs` | Backend: config, form definitions, submit handler, Sheet writer, email notifier |
| `apps-script/Form.html` | Templated form UI (renders all four forms) |
| `apps-script/appsscript.json` | Manifest: OAuth scopes + web-app deployment settings |
| `index.html` | GitHub Pages "Get Involved" page (embeds all four forms) |
| `embed-snippets.html` | Single-form iframe snippets to hand to partner websites |

---

## Deploy runbook

Legend: **🧑 YOU** = a physical login / browser / account action only you can do · **🤖 CLAUDE** = I can run it once you've done the prerequisite.

### 1. 🧑 Create the Google Sheet
Signed in as **info@votenaomi.ca**, create a blank Sheet named **"Vote Naomi — Form Submissions"**.
Copy its ID from the URL (`/spreadsheets/d/`**`THIS_PART`**`/edit`) and send it to me. Tabs auto-create on first submit.

### 2. 🤖 CLAUDE — set the Sheet ID
I paste your ID into `SHEET_ID` in `apps-script/Code.gs`.

### 3. 🧑 clasp login
In a terminal (as the Windows user **Micha**, signed into Google as **info@votenaomi.ca** in the browser that opens):
```
clasp login
```
This stores credentials at `~/.clasprc.json`. Because I run as the same Windows user, **once you've done this I can run the remaining clasp steps.**

### 4. 🤖 CLAUDE — create + push the Apps Script project
From `apps-script/`:
```
clasp create --type standalone --title "Naomi Davison — Campaign Forms"
clasp push -f
```
(Creates `.clasp.json` locally — gitignored. `-f` pushes the manifest.)

### 5. 🧑 Authorize scopes (one-time browser consent)
The web app runs **as info@votenaomi.ca** and needs Sheets + Send-Email permission. Open the editor:
```
clasp open-script
```
Run any function once (e.g. `doGet`) → **Review permissions** → choose the info@votenaomi.ca account →
*Advanced ▸ Go to project (unsafe)* (safe — you own it) → **Allow**.

### 6. 🤖 CLAUDE — deploy the web app
```
clasp deploy --description "v1 — go-live"
clasp list-deployments        # grab the /exec URL
```
I copy the `/exec` URL into `index.html` (`BASE_URL`) and `embed-snippets.html`.

> The four form URLs are that base + `?form=vote | sign | donate | volunteer`.

### 7. 🧑 Publish GitHub Pages
Create a repo (e.g. `votenaomi-forms`), push `index.html` (and `embed-snippets.html` if you like), then
**Settings ▸ Pages ▸ Deploy from branch ▸ main /(root)**. Your Get Involved page goes live at
`https://<user>.github.io/votenaomi-forms/` (or your custom domain).

### 8. 🧑🤖 Test (checklist below)

---

## Updating later (important)

Editing `Code.gs`/`Form.html` and re-pushing does **NOT** change the live site. You must publish a **new
version to the existing deployment** — the `/exec` URL stays the same:

```
clasp push -f
clasp list-deployments                     # note the deploymentId (AKfy...)
clasp redeploy <deploymentId> -d "what changed"
```

⚠️ Do **not** `clasp deploy` a brand-new deployment for updates — that mints a *new* `/exec` URL and leaves
every embedded iframe pointing at the old code. Always `redeploy` the same deploymentId.

---

## Testing checklist

- [ ] `SHEET_ID` set in `Code.gs`; web app deployed (execute as info@votenaomi.ca / access Anyone).
- [ ] Open each `…/exec?form=vote|sign|donate|volunteer` directly — form renders and is styled.
- [ ] Submit one test per form → new row on the correct tab with a working `In Liberalist?` checkbox.
- [ ] Notification email arrives at info@votenaomi.ca with all fields + row number; Reply-To = submitter.
- [ ] Honeypot: a submission with the hidden `_hp` field filled creates **no** row.
- [ ] GitHub Pages page shows all four iframes and each submits successfully.
- [ ] "Learn more" links point to ontarioliberal.ca/yorksimcoe.
- [ ] (Optional) Conditional formatting highlights un-entered rows (`In Liberalist?` unchecked → amber).

---

## Compliance notes

- **Donations are a pledge only** — no card/payment data collected. Real contributions run through the
  official Elections-compliant page. `DONATE_URL` currently points to ontarioliberal.ca/yorksimcoe; swap it
  for the exact donation URL if there is a dedicated one.
- **CASL consent** — every form has a required consent checkbox stored as Yes/No + timestamp + source.
- **PII** — the Sheet holds personal info. Restrict sharing to campaign staff; never commit real
  submissions to the repo.
