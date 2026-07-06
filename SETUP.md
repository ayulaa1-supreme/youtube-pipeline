# Setup Guide — from zero to your own summary library in ~45 minutes

This walkthrough sets up the full pipeline: Google Sheet → Apps Script → Supadata → Claude API → Netlify → Gmail. No server, no local code. Cost: ~₪0.55 (~$0.15) per video, everything else on free tiers.

---

## Step 1 — Collect your three keys (~10 min)

| Key | Where | Notes |
|---|---|---|
| **Claude API key** | [console.anthropic.com](https://console.anthropic.com) → API Keys → Create Key | Starts with `sk-ant-`. You'll need a payment method; usage is pay-per-use. |
| **Supadata key** | [supadata.ai](https://supadata.ai) → sign up → Dashboard → API Key | Free tier: 100 videos/month. |
| **Netlify token** | [app.netlify.com/user/applications](https://app.netlify.com/user/applications) → New access token | This is a *personal access token*, not a site key. |

Copy each key somewhere safe for Step 4. Also decide on an **admin code** — a long random string (20+ characters, not a word). It authorizes video deletion from your library.

## Step 2 — Create the Google Sheet (~5 min)

1. Create a new Google Sheet.
2. In the first tab (default name `גיליון1` — or rename and set the `SHEET_NAME` property later), add headers in row 1:

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| url | Title | Topic | Status | Date Processed | Drive Link |

Columns D-F are filled automatically. A `Subscribers` tab is created automatically the first time someone subscribes.

## Step 3 — Paste the script (~2 min)

1. In the Sheet: **Extensions → Apps Script**.
2. Delete the default code, paste the full contents of [`YouTube_to_Doc.gs`](YouTube_to_Doc.gs), save.

The file contains **no secrets** — all your values go into Script Properties next.

## Step 4 — Script Properties (~5 min)

In the Apps Script editor: **Project Settings (⚙ in the left rail) → Script Properties → Add script property**, one row per value:

| Property | Required | Value |
|---|---|---|
| `CLAUDE_API_KEY` | ✅ | from Step 1 |
| `NETLIFY_TOKEN` | ✅ | from Step 1 |
| `NETLIFY_SITE` | ✅ | bare site name only, e.g. `my-yt-summaries` — **no** `https://`, **no** `.netlify.app` |
| `EMAIL_TO` | ✅ | your email (owner notifications) |
| `SUPADATA_KEY` | ✅ | from Step 1 |
| `ADMIN_CODE` | ✅ | your long random string from Step 1 |
| `WEBAPP_URL` | after Step 5 | the `/exec` URL |
| `DRIVE_FOLDER_ID` | optional | Drive folder id for HTML backups |
| `STATIC_PAGES_FOLDER_ID` | optional | Drive folder with extra pages (about.html etc.) |
| `SHEET_NAME` | optional | only if you renamed the tab (default `גיליון1`) |

> Pick a site name that isn't taken: the first run creates `<NETLIFY_SITE>.netlify.app` automatically.

## Step 5 — Deploy the web app (~5 min)

This enables the subscribe form and admin deletion on your library page.

1. **Deploy → New deployment → Web app.** Execute as: **Me**. Who has access: **Anyone**.
2. Authorize the permissions Google asks for.
3. Copy the resulting `/exec` URL → add it as the `WEBAPP_URL` script property.

## Step 6 — Trigger (~2 min)

Run `setupDailyTrigger()` once from the editor (▶ Run). It installs a daily 08:00 trigger.

Prefer instant processing on every sheet edit instead? Run this once in the editor instead of the daily trigger:

```javascript
ScriptApp.newTrigger("processNewVideos")
  .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
  .onChange().create();
```

Both are safe: the script takes a lock, so overlapping trigger firings can't double-process a video.

## Step 7 — First video (~3 min of your time)

1. Paste a YouTube URL in column A, a title in column B, a topic in column C.
2. Run `processNewVideos` manually (or wait for the trigger).
3. Watch column D: `Processing...` → `Done` (2-4 minutes), then check your email and the link in column F.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `Missing Script Properties: ...` | A required property is empty | Add it in Project Settings → Script Properties (exact spelling) |
| `NETLIFY_SITE must be the bare site name` | You put a URL in `NETLIFY_SITE` | Use the name only, e.g. `my-yt-summaries` |
| `Transcript fetch failed: 401` | Bad/missing Supadata key | Re-paste the key — no spaces or line breaks |
| `Netlify ... 401 Access Denied` | Bad/expired Netlify token | Generate a new token, update the property |
| Row stuck on `Processing...` >6 min | Execution died (Apps Script 6-min limit) or still running | Check **Executions** in the editor; if dead, clear the Status cell to retry |
| `Error: No transcript found` | Video has no captions | Not fixable — pick a captioned video |
| Two sites appeared on Netlify | `NETLIFY_SITE` was wrong at some point | Fix the property, delete the junk site in Netlify |

## Security notes

- Secrets live only in Script Properties — never in code, never in git.
- Deletion requires `ADMIN_CODE` server-side; the public page embeds only a one-way SHA-256 verifier.
- If you ever suspect the admin code leaked, just change the `ADMIN_CODE` property and let the next video (or `redeployIndex()`) refresh the page.
