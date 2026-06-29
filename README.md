# 📺 YouTube → HTML Summary Pipeline

**Automated pipeline: YouTube URL → Hebrew HTML summary → Netlify → Gmail**

Built by [Ayelet Shachak Shul](https://ayelet-yt-summaries.netlify.app/about.html) · 2026

---

## What it does

Add a YouTube URL to a Google Sheet → the pipeline fires **automatically**, and within minutes a branded Hebrew summary page is live on Netlify with a Gmail notification. Zero manual work, zero button clicks.

> 🔁 **Fully automatic:** an `onChange` installable trigger watches the Sheet at all times. The moment a new row is added, the entire pipeline runs — no cron job, no manual trigger, nothing to remember.

**Live library:** [ayelet-yt-summaries.netlify.app](https://ayelet-yt-summaries.netlify.app)  
**Case study:** [Full technical breakdown](https://ayelet-yt-summaries.netlify.app/youtube-pipeline-casestudy.html)

---

## Goals

- **Capture knowledge:** every video watched becomes a documented, searchable asset
- **Zero manual work:** add one URL, everything else happens automatically
- **Consistent format:** every summary follows the same 3-level structure with action steps
- **Instant access:** an online library you can search, filter, and share
- **Near-zero cost:** under ₪1 per video, no server infrastructure needed

---

## How it works

```
Google Sheets (URL input)
       ↓  onChange trigger — fires instantly on every new row, automatically
Apps Script
       ↓  Supadata API
YouTube transcript
       ↓  Claude API (claude-sonnet-4-6)
Hebrew HTML summary
       ↓  Netlify ZIP Deploy API
Live page + updated library index
       ↓  Gmail API
Email notification
```

---

## Stack

| Layer | Tool | Notes |
|---|---|---|
| Input | Google Sheets | onChange installable trigger |
| Orchestration | Google Apps Script | No separate server needed |
| Transcript | [Supadata.ai](https://supadata.ai) | YouTube transcript API |
| Summarization | Claude API (`claude-sonnet-4-6`) | max_tokens: 16,000 |
| Hosting | Netlify | ZIP Deploy API, free tier |
| Notification | Gmail | Apps Script MailApp |
| Backup | Google Drive | Auto-saves each HTML file |

**Cost per video:** < ₪1 (Claude API only)  
**Processing time:** ~2 minutes end-to-end

---

## Output format

Each summary page includes:
- **Level 1:** Tools and options mentioned in the video
- **Level 2:** Methods and approaches
- **Level 3:** Deep dive — definitions, examples, action steps

Rule: only content explicitly stated in the transcript. No inventions, no additions.

---

## Topics

- `AI Tools`
- `Productivity`
- `Entrepreneurship`

---

## Setup

### Prerequisites
- Google account (Sheets + Apps Script)
- [Supadata API key](https://supadata.ai)
- [Anthropic API key](https://console.anthropic.com)
- Netlify account + site ID + personal access token

### Configuration

In `YouTube_to_Doc.gs`, update the `CONFIG` object:

```javascript
var CONFIG = {
  ANTHROPIC_API_KEY: "sk-ant-...",
  SUPADATA_API_KEY:  "...",
  NETLIFY_SITE_ID:   "...",
  NETLIFY_TOKEN:     "...",
  TOPICS:            ["AI Tools", "Productivity", "Entrepreneurship"],
  INDEX_URL:         "https://your-site.netlify.app",
  DRIVE_FOLDER_ID:   ""  // optional
};
```

> ⚠️ **Never share this file publicly with API keys inside.**  
> Use [Google Apps Script Properties](https://developers.google.com/apps-script/guides/properties) to store secrets safely in production.

### Google Sheet structure

| Column A | Column B |
|---|---|
| YouTube URL | Topic |
| https://youtube.com/watch?v=... | AI Tools |

### Install the trigger (one-time setup)

In the Apps Script editor:  
**Triggers → Add Trigger → `processNewVideos` → From spreadsheet → On change**

This is a one-time step. After setup, the trigger runs forever — every time a new URL is added to the Sheet, the full pipeline fires automatically with no manual action required.

---

## Sync with clasp (optional)

To connect this repo to your live Apps Script project:

```bash
npm install -g @google/clasp
clasp login
clasp clone YOUR_SCRIPT_ID   # Apps Script > Settings > Script ID
```

Then push changes:
```bash
clasp push
```

---

## Files

```
YouTube_to_Doc.gs                # Main pipeline (Apps Script)
library-preview.html             # Local preview of the library index
about.html                       # About page
brand-guidelines.html            # Design system reference
youtube-pipeline-casestudy.html  # Full case study
tools-comparison.html            # Stack comparison vs Vercel / Supabase / GitHub
README.md                        # This file
```

---

## License

Personal project. Feel free to fork and adapt — just don't share API keys. 🔑
