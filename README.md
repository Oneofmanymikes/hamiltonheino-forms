# Heino Doessing — Campaign Forms System

Four embeddable web forms (**Commit to Vote · Request a Sign · Commit to Donate · Volunteer**) for
**Heino Doessing, Ontario Liberal candidate for Hamilton East–Stoney Creek**.

Every submission → a row in a Google Sheet (one tab per form, each with an `In Liberalist?` checkbox) →
a notification email to **hello@hamiltonheino.ca**. Forms are served by a single Google Apps Script web app,
embedded as iframes on the GitHub Pages "Get Involved" page and on any partner site. Supporters are pushed
to **ontarioliberal.ca** for more information.

## Files

| File | What it is |
|---|---|
| `apps-script/Code.gs` | Backend: config, form definitions, submit handler, Sheet writer, email notifier |
| `apps-script/Form.html` | Templated form UI (renders all four forms) |
| `apps-script/appsscript.json` | Manifest: OAuth scopes + web-app deployment settings |
| `index.html` | Landing page (hero + About + priorities + Get Involved CTA) |
| `style.css` | Shared styles (OLP red theme) |
| `CNAME` | Custom domain for GitHub Pages (`hamiltonheino.ca`) |
| `getinvolved/index.html` | The hub — four action cards |
| `getinvolved/{commit-to-vote,lawn-sign,donate,volunteer}/index.html` | One page per form (iframe embed) |
| `embed-snippets.html` | Single-form iframe snippets to hand to partner websites |
| `assets/heino-doessing.jpg` | About photo (landscape ~1400×821) |

---

## Deploy runbook

Legend: **🧑 YOU** = a physical login / browser / account action only you can do · **🤖 CLAUDE** = I can run it once you've done the prerequisite.

### 1. 🧑 Create the Google Sheet
Signed in as **hello@hamiltonheino.ca**, create a blank Sheet named **"Hamilton Heino — Form Submissions"**.
Copy its ID from the URL (`/spreadsheets/d/`**`THIS_PART`**`/edit`) and send it to me. Tabs auto-create on first submit.

### 2. 🤖 CLAUDE — set the Sheet ID
I paste your ID into `SHEET_ID` in `apps-script/Code.gs`.

### 3–6. 🧑 Deploy the Apps Script backend

**Recommended path: manual paste-deploy** (no clasp — clasp *writes* are blocked by the build machine's
sandbox classifier when deploying a public web app):

1. Go to **script.google.com** signed in as the data-owning account → **New project**.
2. Paste `apps-script/Code.gs` (with `SHEET_ID` filled in). Add an HTML file named **`Form`** and paste
   `apps-script/Form.html`. Project Settings → *Show "appsscript.json"* → paste `apps-script/appsscript.json`.
3. **Deploy → New deployment → Web app**: Execute as **Me**, Who has access **Anyone**. Authorize the
   scopes (Sheets + send email) at the consent screen.
4. Copy the **`/exec` URL** and send it back — it replaces `PASTE_HEINO_EXEC_ID_HERE` in the four form
   pages and `embed-snippets.html`.

<details><summary>Alternative: clasp (only if you run it yourself)</summary>

⚠️ clasp 3.3.0 has **no `--user` flag** (single global `~/.clasprc.json`) — **back up `~/.clasprc.json`
before `clasp login`** so you don't clobber another project's credentials.

### 3. 🧑 clasp login
In a terminal (as the Windows user **Micha**, signed into Google as **hello@hamiltonheino.ca** in the browser that opens):
```
clasp login
```
This stores credentials at `~/.clasprc.json`. Because I run as the same Windows user, **once you've done this I can run the remaining clasp steps.**

### 4. 🤖 CLAUDE — create + push the Apps Script project
From `apps-script/`:
```
clasp create --type standalone --title "Heino Doessing — Campaign Forms"
clasp push -f
```
(Creates `.clasp.json` locally — gitignored. `-f` pushes the manifest.)

### 5. 🧑 Authorize scopes (one-time browser consent)
The web app runs **as hello@hamiltonheino.ca** and needs Sheets + Send-Email permission. Open the editor:
```
clasp open-script
```
Run any function once (e.g. `doGet`) → **Review permissions** → choose the hello@hamiltonheino.ca account →
*Advanced ▸ Go to project (unsafe)* (safe — you own it) → **Allow**.

### 6. 🤖 CLAUDE — deploy the web app
```
clasp deploy --description "v1 — go-live"
clasp list-deployments        # grab the /exec URL
```
I copy the `/exec` URL into the four form pages and `embed-snippets.html`.

> The four form URLs are that base + `?form=vote | sign | donate | volunteer`.

</details>

### 7. 🧑 Publish GitHub Pages
Create the repo **`hamiltonheino-forms`** on github.com (API repo-creation is blocked from this shell),
push, then **Settings ▸ Pages ▸ Deploy from branch ▸ main /(root)** and set the custom domain to
**`hamiltonheino.ca`** (check the spelling — GitHub commits that string straight into `CNAME`).

DNS at the registrar:
- Apex `@` → **A** → `185.199.108.153`, `185.199.109.153`, `185.199.110.153`, `185.199.111.153`
- Apex `@` → **AAAA** → `2606:50c0:8000::153`, `2606:50c0:8001::153`, `2606:50c0:8002::153`, `2606:50c0:8003::153`
- `www` → **CNAME** → the apex (or `<account>.github.io`)
- **Keep the existing MX/SPF/DKIM records** so `hello@hamiltonheino.ca` keeps receiving.

Then tick **Enforce HTTPS** once the cert provisions.

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

- [ ] `SHEET_ID` set in `Code.gs`; web app deployed (execute as hello@hamiltonheino.ca / access Anyone).
- [ ] Open each `…/exec?form=vote|sign|donate|volunteer` directly — form renders and is styled.
- [ ] Submit one test per form → new row on the correct tab with a working `In Liberalist?` checkbox.
- [ ] Notification email arrives at hello@hamiltonheino.ca with all fields + row number; Reply-To = submitter.
- [ ] Honeypot: a submission with the hidden `_hp` field filled creates **no** row.
- [ ] GitHub Pages page shows all four iframes and each submits successfully.
- [ ] "Learn more" links point to ontarioliberal.ca.
- [ ] (Optional) Conditional formatting highlights un-entered rows (`In Liberalist?` unchecked → amber).

---

## Compliance notes

- **Donations are a pledge only** — no card/payment data collected. Real contributions run through the
  official Elections-compliant page. `DONATE_URL` points to
  `ontarioliberal.ca/donate/?pla=37` — PLA **37** is Hamilton East–Stoney Creek, so credit flows to the
  right riding association.
- **CASL consent** — every form has a required consent checkbox stored as Yes/No + timestamp + source.
- **PII** — the Sheet holds personal info. Restrict sharing to campaign staff; never commit real
  submissions to the repo.
