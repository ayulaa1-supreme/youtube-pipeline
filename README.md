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

In `YouTube_to_Doc.gs`, fill in the `CONFIG` object. See `.env.example` for the full list of required values.

```javascript
const CONFIG = {
  CLAUDE_API_KEY:         "sk-ant-...",
  NETLIFY_TOKEN:          "nfp_...",
  NETLIFY_SITE:           "my-yt-summaries",
  EMAIL_TO:               "you@example.com",
  SUPADATA_KEY:           "...",
  DRIVE_FOLDER_ID:        "...",            // optional
  STATIC_PAGES_FOLDER_ID: "...",            // optional
  WEBAPP_URL:             "https://script.google.com/macros/s/.../exec",
  ADMIN_CODE:             "long-random-string",
  // ...
};
```

> ⚠️ **Never commit this file with real API keys.**
> For production, prefer [PropertiesService](https://developers.google.com/apps-script/guides/properties) to store secrets.

### Google Sheet structure

The active tab (default name: `גיליון1`) must have these columns:

| A: URL | B: Title | C: Topic | D: Status | E: Date | F: Link |
|---|---|---|---|---|---|
| `https://youtube.com/watch?v=...` | Optional title | `AI Tools` | (auto-filled) | (auto-filled) | (auto-filled) |

A second tab named `Subscribers` is created automatically the first time someone subscribes.

### Deploy

1. Paste `YouTube_to_Doc.gs` into the Apps Script editor of your sheet
2. Fill the `CONFIG` block
3. Run `setupDailyTrigger()` once — installs a daily 08:00 trigger that calls `processNewVideos`
4. **Deploy → New deployment → Web app** (execute as Me, access Anyone) — copy the resulting `/exec` URL into `CONFIG.WEBAPP_URL` and redeploy once more so the URL is baked into the index page

---

## Bug fixes vs. earlier versions

This revision (`v2026-06-30`) addresses three real issues found in code review:

| Fix | What changed | Why it matters |
|---|---|---|
| **Double email** | `sendToAllSubscribers` now skips `CONFIG.EMAIL_TO` | Previously, if the owner email was also in the Subscribers tab, every new video sent two emails (owner + subscriber). |
| **Delete endpoint auth** | `doGet` now requires `?code=<ADMIN_CODE>` on `action=delete` | Previously, anyone who knew the `WEBAPP_URL` could call `?action=delete&vid=X` and delete any video. |
| **Broken typewriter** | Removed dead `setTimeout(type,...)` block referencing missing DOM elements `#tw` / `#cur` | Was throwing a silent JS error on every library page load. |

`testSubscribe()` was also removed — running it once was the cause of the double-email bug.

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
.env.example         # Required config values (placeholders only)
.gitignore           # Standard ignores + clasp secrets
README.md            # This file
LICENSE              # MIT
```

---

## License

MIT. Fork, adapt, share — just don't push your API keys.
