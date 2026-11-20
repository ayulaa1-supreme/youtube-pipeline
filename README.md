# 📺 YouTube → HTML Summary Pipeline

**Automated pipeline: YouTube URL → Hebrew HTML summary → Netlify library → Gmail notification**

Built by [Ayelet Shachak Shul](https://ayelet-yt-summaries.netlify.app/about.html) · 2026

---

## What it does

Add a YouTube URL to a Google Sheet → the pipeline fetches the transcript, generates a Hebrew summary page with Claude, deploys it to Netlify, and emails you (and any subscribers) when it's live.

**Live library:** [ayelet-yt-summaries.netlify.app](https://ayelet-yt-summaries.netlify.app)
**Case study:** [Full technical breakdown](https://ayelet-yt-summaries.netlify.app/youtube-pipeline-casestudy.html)

---

## How it works

```
Google Sheets (URL input)
       ↓  daily time trigger (or run processNewVideos manually)
Apps Script
       ↓  Supadata API
YouTube transcript
       ↓  Claude API (claude-sonnet-4-6)
Hebrew HTML summary
       ↓  Netlify ZIP Deploy API
Live page + updated library index
       ↓  Gmail / MailApp
Owner notification + subscriber emails
```

---

## Stack

| Layer | Tool | Notes |
|---|---|---|
| Input | Google Sheets | URL, title, topic per row |
| Orchestration | Google Apps Script | Daily time trigger |
| Transcript | [Supadata.ai](https://supadata.ai) | YouTube transcript API |
| Summarization | Claude API (`claude-sonnet-4-6`) | `max_tokens: 16000` |
| Hosting | Netlify | ZIP Deploy API, free tier |
| Notification | Gmail / MailApp | Apps Script built-in |
| Backup | Google Drive | Auto-saves each HTML file |

**Cost per video:** ~₪0.55 (Claude API)
**Processing time:** ~2-3 minutes end-to-end

---

## Output format

Each summary page includes:
- **Level 1:** Tools and options mentioned in the video
- **Level 2:** Methods and approaches
- **Level 3:** Deep dive — definitions, examples, action steps

**Iron rule:** only content explicitly stated in the transcript. No inventions, no additions.

---

## Setup

### Prerequisites
- Google account (Sheets + Apps Script)
- [Supadata API key](https://supadata.ai)
- [Anthropic API key](https://console.anthropic.com)
- [Netlify personal access token](https://app.netlify.com/user/applications)

### Configuration

All user-specific values live in **Script Properties**, not in the code — so updating the code can never wipe your keys, and there is nothing secret to accidentally commit.

Apps Script editor → **Project Settings (⚙) → Script Properties**: add `CLAUDE_API_KEY`, `NETLIFY_TOKEN`, `NETLIFY_SITE` (bare site name only), `EMAIL_TO`, `SUPADATA_KEY`, `ADMIN_CODE`, `WEBAPP_URL` (after deploy), and optionally `DRIVE_FOLDER_ID`, `STATIC_PAGES_FOLDER_ID`, `SHEET_NAME`. See `.env.example` for the full list.

📖 **Step-by-step walkthrough: [SETUP.md](SETUP.md)** — from zero to a live library in ~45 minutes, including troubleshooting.

### Google Sheet structure

The active tab (default name: `גיליון1`) must have these columns:

| A: URL | B: Title | C: Topic | D: Status | E: Date | F: Link |
|---|---|---|---|---|---|
| `https://youtube.com/watch?v=...` | Optional title | `AI Tools` | (auto-filled) | (auto-filled) | (auto-filled) |

A second tab named `Subscribers` is created automatically the first time someone subscribes.

### Deploy

1. Paste `YouTube_to_Doc.gs` into the Apps Script editor of your sheet
2. Add your values as Script Properties (see Configuration above / [SETUP.md](SETUP.md))
3. Run `setupDailyTrigger()` once — installs a daily 08:00 trigger that calls `processNewVideos`
4. **Deploy → New deployment → Web app** (execute as Me, access Anyone) — save the resulting `/exec` URL as the `WEBAPP_URL` script property

---

## Bug fixes vs. earlier versions

This revision (`v2026-06-30`) addresses three real issues found in code review:

| Fix | What changed | Why it matters |
|---|---|---|
| **Double email** | `sendToAllSubscribers` now skips `CONFIG.EMAIL_TO` | Previously, if the owner email was also in the Subscribers tab, every new video sent two emails (owner + subscriber). |
| **Delete endpoint auth** | `doGet` now requires `?code=<ADMIN_CODE>` on `action=delete` | Previously, anyone who knew the `WEBAPP_URL` could call `?action=delete&vid=X` and delete any video. |
| **Broken typewriter** | Removed dead `setTimeout(type,...)` block referencing missing DOM elements `#tw` / `#cur` | Was throwing a silent JS error on every library page load. |

`testSubscribe()` was also removed — running it once was the cause of the double-email bug.

### Revision `v2026-07-06`

| Fix | What changed | Why it matters |
|---|---|---|
| **Duplicate processing** | `processNewVideos` now takes a `LockService` lock; overlapping trigger firings skip instead of racing | An onChange trigger fires once per sheet edit — entering one video row is several edits. Without the lock, one new video could be processed (and emailed) up to 7 times. |
| **Secrets → Script Properties** | The `CONFIG` block reads everything via `PropertiesService`; the code file contains no user values | Pasting a code update used to risk overwriting real keys with placeholders. Now code and config are fully separated. |
| **Admin code exposure** | The index page embedded `base64(ADMIN_CODE)` — reversible by anyone via view-source. Replaced with a one-way SHA-256 verifier | The delete-endpoint protection was effectively bypassable. After updating, **rotate your `ADMIN_CODE`** — the old one should be treated as public. |
| **Fail-fast config check** | `assertConfig_()` validates required properties (and that `NETLIFY_SITE` is a bare name, not a URL) before any run | A wrong `NETLIFY_SITE` used to silently create a junk Netlify site; missing keys surfaced as cryptic 401s deep in the pipeline. |

### Revision `v2026-11-20b`: short summaries + magazine design

| Change | What changed | Why it matters |
|---|---|---|
| **Short summaries** | New prompt: ~300-380 words (hard ceiling 520 in the quality gate), 1-3 parts shaped by the video instead of a fixed 3-level template, one bold phrase per paragraph, 3-5 action steps, no dates, no emoji | Summaries were 770-1,240 words with 9-17 steps and 15-37 bold phrases. The fixed template forced every video into the same long shape. |
| **Claude writes content only** | Claude returns an `<article class="summary">` fragment; `buildSummaryPage_` builds the page around it | No CSS in the prompt (fewer tokens, smaller stored HTML), every page gets the same frame. |
| **Shared stylesheet** | `summary.css` is generated by `summaryCss_()` and deployed on every deploy; pages link to it | One design change restyles every page on the next deploy. |
| **Old pages restyled at deploy** | `prepareSummaryForDeploy_` strips the old `<style>` from legacy pages, links `summary.css`, adds a topic label and removes emoji | Existing pages match the new design right away, before they are regenerated. |
| **Magazine index** | New `buildIndexPage`: newest summary as cover story, pull quote, numbered table of contents with topic filter and search, key figure per summary in its topic color | Replaces the carousels. Admin (`#admin`) and subscribe keep working. |
| **Topic colors** | AI Tools pink, Productivity teal, Entrepreneurship blue (`TOPIC_COLORS_`) | The reader sees the topic before reading it. |
| **Fonts** | IBM Plex Sans Hebrew + Space Mono, loaded from Google Fonts | Segoe UI exists only on Windows; phones and Macs fell back to Arial. |
| **Fact check before publishing** | Two agents: the writer, then one checker that gets only the summary and the transcript, removes invented statements and corrects distorted ones, and returns the clean summary with a change list (logged). Wording and framing are left alone; the checker may not add content | The "only what was said" rule was a prompt instruction with no check behind it. |
| **Regeneration without emails** | `startRegeneration()` rewrites every published summary in the short format, 1 per run on a 5-minute trigger, each fact-checked, then redeploys once and emails the owner a report | Clearing the status column would have emailed every subscriber once per video. |

After updating the code: run `redeployIndex()` (new design, old content). To rewrite the old summaries in the short format, run `startRegeneration()` once; `regenerationStatus()` shows progress.

### Revision `v2026-11-20`

| Fix | What changed | Why it matters |
|---|---|---|
| **No admin UI in public HTML** | Banner, delete buttons and lock icon are no longer in the page. Open `<site>/#admin`, enter the code, and they are created client-side | Crawlers, link previews and reader modes ignore CSS, so they showed the admin banner and delete buttons to everyone. |
| **Delete via POST only** | New `doPost` handles `action=delete`; the GET route is gone | The admin code no longer appears in URLs or Apps Script request logs. |
| **Real numbers in HTML** | Video and topic counters are rendered server-side; cards stay visible without JS; topic count = topics that have videos | Without JS (previews, crawlers) the page showed 0 videos and invisible cards. |
| **Share previews** | `og:` / description tags on the index and on every summary page (injected at deploy) | WhatsApp / LinkedIn showed a bare URL. |
| **Honest subscribe status** | Subscribe reads the JSON response and shows an error on failure | The form used to show success even when the request failed. |
| **Static folder** | Images (`og-*.png`) are deployed too; summary backups (`<title>_<videoId>.html`) in the same folder are skipped | Backups were being published a second time under odd URLs. |
| **Hebrew topic labels** | `TOPIC_LABELS` maps sheet topics to display names | Consistent Hebrew UI; sheet values unchanged. |

After updating: **Deploy → Manage deployments → edit → New version** (keeps the same `/exec` URL), access **Anyone**, then run `redeployIndex()`. To delete a video, open `<site>/#admin`.

---

## Sync with clasp (optional)

```bash
npm install -g @google/clasp
clasp login
clasp clone YOUR_SCRIPT_ID   # Apps Script → Settings → Script ID
clasp push                   # push local changes
```

`.clasp.json` and `.clasprc.json` are gitignored.

---

## Files

```
YouTube_to_Doc.gs    # Main pipeline (Apps Script)
appsscript.json      # Apps Script manifest (timezone, scopes, webapp config)
SETUP.md             # Step-by-step setup walkthrough + troubleshooting
static-pages/        # Extra site pages (case study) — deployed from Drive, versioned here
.env.example         # Required Script Properties (placeholders only)
.gitignore           # Standard ignores + clasp secrets
README.md            # This file
LICENSE              # MIT
```

---

## License

MIT. Fork, adapt, share — just don't push your API keys.
