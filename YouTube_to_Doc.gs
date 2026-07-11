// YouTube to HTML Doc - Google Apps Script
// ===========================================
// SETTINGS
// ===========================================
// All user-specific values live in Script Properties, NOT in this file —
// so pasting a code update can never wipe your keys.
// Apps Script editor → Project Settings (⚙) → Script Properties → add:
//   CLAUDE_API_KEY, NETLIFY_TOKEN, NETLIFY_SITE, EMAIL_TO, SUPADATA_KEY,
//   ADMIN_CODE, WEBAPP_URL, DRIVE_FOLDER_ID (optional), STATIC_PAGES_FOLDER_ID (optional)
// Full walkthrough: SETUP.md in the repo.

const PROPS_ = PropertiesService.getScriptProperties();
function prop_(key, fallback) { return PROPS_.getProperty(key) || fallback || ""; }

const CONFIG = {
  CLAUDE_API_KEY:         prop_("CLAUDE_API_KEY"),          // https://console.anthropic.com
  NETLIFY_TOKEN:          prop_("NETLIFY_TOKEN"),           // https://app.netlify.com/user/applications
  NETLIFY_SITE:           prop_("NETLIFY_SITE"),            // site NAME only, e.g. "my-yt-summaries" — no https://, no .netlify.app
  EMAIL_TO:               prop_("EMAIL_TO"),                // where you (the owner) receive notifications
  SHEET_NAME:             prop_("SHEET_NAME", "גיליון1"),    // active sheet/tab name
  COL_URL:                1,
  COL_TITLE:              2,
  COL_TOPIC:              3,
  COL_STATUS:             4,
  COL_DATE:               5,
  COL_LINK:               6,
  SUPADATA_KEY:           prop_("SUPADATA_KEY"),            // https://supadata.ai
  DRIVE_FOLDER_ID:        prop_("DRIVE_FOLDER_ID"),         // optional backup folder
  STATIC_PAGES_FOLDER_ID: prop_("STATIC_PAGES_FOLDER_ID"),  // optional: Drive folder with about.html, casestudy, etc.
  START_ROW:              2,
  TOPICS:                 ["AI Tools", "Productivity", "Entrepreneurship"],
  WEBAPP_URL:             prop_("WEBAPP_URL"),              // fill after Deploy as Web App (the /exec URL)
  ADMIN_CODE:             prop_("ADMIN_CODE"),              // long random string — authorizes deletions
  SUBSCRIBERS_SHEET:      "Subscribers"
};

// Fail fast with a clear message instead of a cryptic 401 somewhere downstream.
function assertConfig_() {
  const missing = ["CLAUDE_API_KEY", "NETLIFY_TOKEN", "NETLIFY_SITE", "EMAIL_TO", "SUPADATA_KEY"]
    .filter(function(k) { return !CONFIG[k]; });
  if (missing.length) {
    throw new Error("Missing Script Properties: " + missing.join(", ") + " — add them under Project Settings → Script Properties (see SETUP.md)");
  }
  if (/[:\/.]/.test(CONFIG.NETLIFY_SITE)) {
    throw new Error('NETLIFY_SITE must be the bare site name, e.g. "my-yt-summaries" — not a URL');
  }
}

// ===========================================
// MAIN
// ===========================================
function processNewVideos() {
  // The onChange trigger fires once per sheet edit, and entering one video row is
  // several edits (URL, title, topic). Without a lock, those firings overlap: each
  // sees the row as pending, so one new video gets processed — and emailed — N times.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log("Another run is already processing — skipping this trigger firing.");
    return;
  }
  try {
    // Re-scan until a pass finds nothing, so rows added while a long run was busy
    // still get processed before the lock is released.
    let rounds = 0;
    while (processPendingVideos() > 0 && ++rounds < 5);
  } finally {
    lock.releaseLock();
  }
}

function processPendingVideos() {
  assertConfig_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { Logger.log("Sheet not found: " + CONFIG.SHEET_NAME); return 0; }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) { Logger.log("No rows to process"); return 0; }

  const data = sheet.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 6).getValues();
  let processed = 0;

  data.forEach((row, i) => {
    const url    = (row[CONFIG.COL_URL - 1] || "").toString().trim();
    const status = (row[CONFIG.COL_STATUS - 1] || "").toString().trim();
    if (!url || status !== "") return;

    const sheetRow = CONFIG.START_ROW + i;
    sheet.getRange(sheetRow, CONFIG.COL_STATUS).setValue("Processing...");
    SpreadsheetApp.flush(); // make the claim visible immediately, not on lazy write-out

    try {
      const videoId = extractVideoId(url);
      const title   = (row[CONFIG.COL_TITLE - 1] || "Video " + videoId).toString().trim();
      const topic   = (row[CONFIG.COL_TOPIC - 1] || "General").toString().trim();

      const transcript = getYouTubeTranscript(url);
      if (!transcript) throw new Error("No transcript found");

      const html = generateValidatedHtml_(transcript, title, url);

      // Store HTML in ScriptProperties by videoId — reliable lookup for future deploys
      PropertiesService.getScriptProperties().setProperty("vid_html_" + videoId, html);

      // Extract Hebrew title from generated HTML and save separately
      const heTitle = extractHebrewTitle(html) || title;
      PropertiesService.getScriptProperties().setProperty("vid_title_he_" + videoId, heTitle);

      saveToDrive(html, heTitle, videoId);

      const netlifyUrl = deployAllToNetlify(sheet, data, html, videoId, heTitle, topic, i);

      sendEmailNotification(heTitle, url, netlifyUrl, topic);
      sendToAllSubscribers(heTitle, url, netlifyUrl, topic);

      sheet.getRange(sheetRow, CONFIG.COL_STATUS).setValue("Done");
      sheet.getRange(sheetRow, CONFIG.COL_DATE).setValue(new Date().toLocaleDateString("he-IL"));
      sheet.getRange(sheetRow, CONFIG.COL_LINK).setValue(netlifyUrl);
      data[i][CONFIG.COL_STATUS - 1] = "Done"; // sync in-memory so next video sees this as Done
      processed++;

    } catch (err) {
      Logger.log("Error: " + err.message);
      sheet.getRange(sheetRow, CONFIG.COL_STATUS).setValue("Error: " + err.message);
    }
  });

  Logger.log("Processed: " + processed);
  return processed;
}

// ===========================================
// NETLIFY DEPLOY
// ===========================================
// Static pages (about.html, casestudy, etc.) from the Drive folder.
// A Netlify ZIP deploy replaces the WHOLE site, so any deploy path that
// skips these files silently deletes them from production. Every deploy
// path must call this.
function addStaticPageBlobs_(blobs) {
  if (!CONFIG.STATIC_PAGES_FOLDER_ID) return;
  try {
    const folder = DriveApp.getFolderById(CONFIG.STATIC_PAGES_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.HTML);
    while (files.hasNext()) {
      const f = files.next();
      blobs.push(Utilities.newBlob(f.getBlob().getDataAsString(), "text/html", f.getName()));
      Logger.log("Added static page: " + f.getName());
    }
  } catch(e) { Logger.log("Static pages folder error: " + e.message); }
}

function deployAllToNetlify(sheet, data, newHtml, newVideoId, newTitle, newTopic, newRowIndex) {
  const videos = [];
  data.forEach((row, i) => {
    const url    = (row[CONFIG.COL_URL - 1] || "").toString().trim();
    const title  = (row[CONFIG.COL_TITLE - 1] || "").toString().trim();
    const topic  = (row[CONFIG.COL_TOPIC - 1] || "General").toString().trim();
    const status = (row[CONFIG.COL_STATUS - 1] || "").toString().trim();
    const date   = (row[CONFIG.COL_DATE - 1] || "").toString().trim();
    if (!url) return;
    const vid = extractVideoId(url);
    if (!vid) return;
    if (i === newRowIndex) {
      videos.push({ videoId: newVideoId, title: newTitle || title, topic: newTopic, date: new Date().toLocaleDateString("he-IL") });
    } else if (status.includes("Done")) {
      const storedHeTitle = PropertiesService.getScriptProperties().getProperty("vid_title_he_" + vid) || title;
      videos.push({ videoId: vid, title: storedHeTitle, topic: topic, date: date });
    }
  });

  const siteId = getOrCreateNetlifySite(CONFIG.NETLIFY_SITE);
  const blobs = [];

  blobs.push(Utilities.newBlob(buildIndexPage(videos), "text/html", "index.html"));
  blobs.push(Utilities.newBlob(newHtml, "text/html", newVideoId + ".html"));

  videos.forEach(function(v) {
    if (v.videoId === newVideoId) return;
    // 1. Try ScriptProperties (videos processed after the fix)
    const storedHtml = PropertiesService.getScriptProperties().getProperty("vid_html_" + v.videoId);
    if (storedHtml) {
      blobs.push(Utilities.newBlob(storedHtml, "text/html", v.videoId + ".html"));
      return;
    }
    // 2. Search Drive by videoId anywhere in filename — no title dependency
    try {
      const query = 'title contains "' + v.videoId + '" and mimeType = "text/html" and trashed = false';
      const files = DriveApp.searchFiles(query);
      if (files.hasNext()) {
        const html = files.next().getBlob().getDataAsString();
        blobs.push(Utilities.newBlob(html, "text/html", v.videoId + ".html"));
        // Cache in ScriptProperties so next deploy doesn't need Drive
        PropertiesService.getScriptProperties().setProperty("vid_html_" + v.videoId, html);
        Logger.log("Found in Drive: " + v.videoId);
      } else {
        Logger.log("WARNING: no HTML found for video " + v.videoId);
      }
    } catch(e) { Logger.log("Drive search failed for: " + v.videoId + " — " + e.message); }
  });

  addStaticPageBlobs_(blobs); // without this, every new-video deploy wipes the static pages

  const zipBlob = Utilities.zip(blobs, "deploy.zip");
  const deployResp = UrlFetchApp.fetch(
    "https://api.netlify.com/api/v1/sites/" + siteId + "/deploys",
    {
      method: "post",
      muteHttpExceptions: true,
      headers: {
        "Authorization": "Bearer " + CONFIG.NETLIFY_TOKEN,
        "Content-Type": "application/zip"
      },
      payload: zipBlob.getBytes()
    }
  );

  const deployData = JSON.parse(deployResp.getContentText());
  if (deployData.error) throw new Error("Netlify: " + deployData.error);

  waitForDeploy(deployData.id);
  return "https://" + CONFIG.NETLIFY_SITE + ".netlify.app/" + newVideoId + ".html";
}

function getOrCreateNetlifySite(name) {
  const listResp = UrlFetchApp.fetch("https://api.netlify.com/api/v1/sites?filter=owner", {
    headers: { "Authorization": "Bearer " + CONFIG.NETLIFY_TOKEN },
    muteHttpExceptions: true
  });
  const sites = JSON.parse(listResp.getContentText());
  if (Array.isArray(sites)) {
    const existing = sites.find(function(s) { return s.name === name; });
    if (existing) return existing.id;
  }
  const createResp = UrlFetchApp.fetch("https://api.netlify.com/api/v1/sites", {
    method: "post",
    muteHttpExceptions: true,
    headers: {
      "Authorization": "Bearer " + CONFIG.NETLIFY_TOKEN,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({ name: name })
  });
  const site = JSON.parse(createResp.getContentText());
  if (!site.id) throw new Error("Netlify site creation failed: " + createResp.getContentText());
  return site.id;
}

function waitForDeploy(deployId) {
  for (var i = 0; i < 10; i++) {
    Utilities.sleep(3000);
    const resp = UrlFetchApp.fetch("https://api.netlify.com/api/v1/deploys/" + deployId, {
      headers: { "Authorization": "Bearer " + CONFIG.NETLIFY_TOKEN }
    });
    if (JSON.parse(resp.getContentText()).state === "ready") return;
  }
}

// ===========================================
// INDEX PAGE
// ===========================================
// SHA-256 hex digest — used to embed a one-way verifier in the index page.
// Never embed the admin code itself (or any reversible encoding of it) in public HTML.
function sha256Hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function(b) { return ("0" + (b & 0xFF).toString(16)).slice(-2); })
    .join("");
}

function buildIndexPage(videos) {
  const byTopic = {};
  CONFIG.TOPICS.forEach(function(t) { byTopic[t] = []; });
  videos.forEach(function(v) {
    const t = v.topic || "General";
    if (!byTopic[t]) byTopic[t] = [];
    byTopic[t].push(v);
  });

  const totalVideos = videos.length;
  const totalTopics = CONFIG.TOPICS.length;

  var svgIcons = {
    "AI Tools":    '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="url(#grd-lib)" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="6" height="6" rx="1"/><path d="M3 9h2M3 15h2M19 9h2M19 15h2M9 3v2M15 3v2M9 19v2M15 19v2"/><rect x="3" y="3" width="18" height="18" rx="3"/></svg>',
    "Productivity":'<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="url(#grd-lib)" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><polyline points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>',
    "Entrepreneurship": '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="url(#grd-lib)" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>',
    "General":     '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="url(#grd-lib)" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21l-7-5-7 5V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2z"/></svg>'
  };

  var sections = "";
  Object.keys(byTopic).forEach(function(topic) {
    const vids = byTopic[topic];
    if (vids.length === 0) return;
    const iconSvg = svgIcons[topic] || svgIcons["General"];
    var cards = "";
    vids.forEach(function(v) {
      cards += '<a href="' + v.videoId + '.html" class="video-card anim-ready" data-vid="' + v.videoId + '">' +
        '<button class="delete-btn" title="מחיקה" aria-label="מחיקת סרטון">✕</button>' +
        '<div class="card-topic">' + topic + '</div>' +
        '<div class="card-title">' + (v.title || v.videoId) + '</div>' +
        '<div class="card-date">' + (v.date || "") + '</div>' +
        '</a>';
    });
    sections += '<div class="topic-section anim-ready">' +
      '<div class="topic-header">' +
      '<span class="topic-icon" aria-hidden="true">' + iconSvg + '</span>' +
      '<span class="topic-name">' + topic + '</span>' +
      '<span class="topic-count" aria-label="' + vids.length + ' סיכומים">' + vids.length + '</span>' +
      '<div class="car-nav">' +
      '<button class="car-btn car-prev" aria-label="הסיכומים הקודמים">›</button>' +
      '<button class="car-btn car-next" aria-label="הסיכומים הבאים">‹</button>' +
      '</div>' +
      '</div>' +
      '<div class="video-track" role="group" aria-label="סיכומים בנושא ' + topic + '">' + cards + '</div>' +
      '</div>';
  });

  if (!sections) {
    sections = '<div class="empty">No videos yet - add a URL to the Google Sheet</div>';
  }

  return '<!DOCTYPE html>' +
'<html lang="he" dir="rtl">' +
'<head>' +
'<meta charset="UTF-8">' +
'<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
'<title>ספריית סיכומים | אילת שחק שול</title>' +
'<link rel="preconnect" href="https://fonts.googleapis.com">' +
'<link href="https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;700;800;900&display=swap" rel="stylesheet">' +
'<style>' +
'@keyframes fadeUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}' +
'.anim-ready{opacity:0}.anim-fadeup{animation:fadeUp 0.7s cubic-bezier(0.22,1,0.36,1) both}' +
'.topic-icon{width:28px;height:28px;flex-shrink:0;display:flex;align-items:center;justify-content:center}' +
'*{box-sizing:border-box;margin:0;padding:0}' +
'body{font-family:Heebo,Arial,sans-serif;direction:rtl;' +
'background:#fff;' +
'background-image:linear-gradient(rgba(0,0,0,0.04) 1px,transparent 1px),linear-gradient(90deg,rgba(0,0,0,0.04) 1px,transparent 1px);' +
'background-size:28px 28px;color:#1f2937;line-height:1.7}' +
'.container{max-width:900px;margin:0 auto;padding:0 1.2rem 4rem}' +
/* Hero — compact bar */
'.hero{background:#fff;border-bottom:1px solid #f0f0f0;padding:1.6rem 2rem 1.4rem;' +
'position:relative;display:flex;align-items:center;justify-content:space-between;gap:1rem}' +
'.hero::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;' +
'background:linear-gradient(90deg,#0d9488,#ec4899,#0d9488)}' +
'.hero-inner{animation:fadeUp .5s cubic-bezier(.22,1,.36,1) both}' +
'.footer-credit{text-align:center;padding:2rem 1rem;font-size:0.8rem;color:#9ca3af;border-top:1px solid #f0f0f0}' +
'.footer-credit b{color:#ec4899;font-weight:600}' +
'h1{font-size:1.5rem;font-weight:800;color:#111827;margin-bottom:.15rem;line-height:1.3}' +
'h1 span{color:#ec4899}' +
'.sub{font-size:.82rem;color:#6b7280}' +
'.stats{display:flex;align-items:center;gap:1.5rem;flex-shrink:0}' +
'.stat{text-align:center}' +
'.stat-num{display:block;font-size:1.4rem;font-weight:800;color:#ec4899;line-height:1}' +
'.stat-num.teal{color:#0d9488}' +
'.stat-label{font-size:.7rem;color:#6b7280;font-weight:500;margin-top:.1rem}' +
/* Topic sections */
'.topic-section{background:white;border:1px solid #f3f4f6;border-radius:14px;' +
'padding:1.4rem;margin-bottom:1.2rem;box-shadow:0 2px 12px rgba(0,0,0,0.04);' +
'border-right:4px solid #0d9488}' +
'.topic-header{display:flex;align-items:center;gap:0.75rem;margin-bottom:1.1rem;' +
'padding-bottom:0.9rem;border-bottom:1px solid #f9fafb}' +
'.topic-icon{font-size:1.3rem;flex-shrink:0}' +
'.topic-name{font-size:1rem;font-weight:700;color:#111827;flex:1}' +
'.topic-count{background:#fce7f3;color:#ec4899;border:1px solid #fbcfe8;' +
'border-radius:50%;width:28px;height:28px;display:flex;align-items:center;' +
'justify-content:center;font-size:0.78rem;font-weight:700;flex-shrink:0}' +
/* Cards */
/* Carousel track — horizontal scroll-snap keeps each section one row tall as the library grows */
'.video-track{display:flex;gap:0.8rem;overflow-x:auto;scroll-snap-type:x mandatory;' +
'padding:0.2rem 2px 0.6rem;-webkit-overflow-scrolling:touch;scrollbar-width:thin;scrollbar-color:#e5e7eb transparent}' +
'.video-track::-webkit-scrollbar{height:6px}' +
'.video-track::-webkit-scrollbar-thumb{background:#e5e7eb;border-radius:3px}' +
'.video-track .video-card{flex:0 0 230px;scroll-snap-align:start}' +
'.car-nav{display:flex;gap:0.4rem;flex-shrink:0}' +
'.car-btn{background:#fff;border:1px solid #e5e7eb;color:#0d9488;border-radius:50%;width:28px;height:28px;' +
'display:none;align-items:center;justify-content:center;cursor:pointer;font-size:1rem;line-height:1;padding:0;' +
'transition:background .15s,color .15s}' +
'.car-btn:hover{background:#0d9488;color:#fff;border-color:#0d9488}' +
'.video-card{background:#fafafa;border:1px solid #f3f4f6;border-radius:10px;' +
'padding:1.1rem;text-decoration:none;color:inherit;display:block;' +
'border-top:3px solid #0d9488;' +
'transition:border-color 0.18s,transform 0.18s,box-shadow 0.18s}' +
'.video-card:hover{border-top-color:#ec4899;transform:translateY(-3px);box-shadow:0 6px 20px rgba(236,72,153,0.12)}' +
'.card-topic{font-size:0.7rem;color:#0d9488;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:0.4rem}' +
'.card-title{font-size:0.88rem;font-weight:600;color:#111827;line-height:1.45;margin-bottom:0.6rem}' +
'.card-date{font-size:0.72rem;color:#6b7280}' +
'.empty{text-align:center;padding:3rem;color:#9ca3af}' +
/* Search bar */
'.search-wrap{background:#fff;border-bottom:1px solid #f0f0f0;padding:0.7rem 2rem}' +
'.search-inner{position:relative;max-width:900px;margin:0 auto}' +
'.search-icon{position:absolute;top:50%;transform:translateY(-50%);left:0.85rem;font-size:0.9rem;pointer-events:none;line-height:1}' +
'.search-input{width:100%;padding:0.6rem 1rem 0.6rem 2.5rem;border:1px solid #e5e7eb;border-radius:10px;' +
'font-family:Heebo,Arial,sans-serif;font-size:0.9rem;color:#111827;outline:none;direction:rtl;' +
'transition:border-color 0.15s,box-shadow 0.15s;background:#fafafa}' +
'.search-input:focus{border-color:#0d9488;box-shadow:0 0 0 3px rgba(13,148,136,0.1);background:#fff}' +
'.no-results{display:none;text-align:center;padding:3rem 1rem;color:#9ca3af;font-size:0.9rem}' +
/* Admin delete button — hidden unless admin mode active */
'.delete-btn{display:none;position:absolute;top:.45rem;left:.45rem;' +
'background:#fee2e2;border:1px solid #fca5a5;color:#991b1b;' +
'border-radius:50%;width:22px;height:22px;font-size:.68rem;font-weight:800;' +
'cursor:pointer;align-items:center;justify-content:center;z-index:10;padding:0;line-height:1;' +
'transition:background .15s,color .15s}' +
'.delete-btn:hover{background:#dc2626;color:#fff;border-color:#dc2626}' +
'.admin-mode .delete-btn{display:flex}' +
'.admin-mode .video-card{position:relative}' +
'.admin-banner{display:none;background:#111827;color:#f9a8d4;text-align:center;' +
'padding:.5rem 1rem;font-size:.78rem;font-weight:600;letter-spacing:.04em}' +
'.admin-mode .admin-banner{display:block}' +
/* Lock button in hero */
'.lock-btn{background:none;border:none;cursor:pointer;font-size:1.1rem;opacity:0.35;' +
'transition:opacity .2s;padding:.2rem;line-height:1;flex-shrink:0}' +
'.lock-btn:hover{opacity:1}' +
/* Subscribe card */
'.subscribe-card{background:#fff;border:1px solid #f0f0f0;border-radius:14px;' +
'padding:1.4rem 1.6rem;margin-bottom:1.2rem;' +
'box-shadow:0 1px 8px rgba(0,0,0,0.04);display:flex;align-items:center;' +
'gap:1.2rem;flex-wrap:wrap}' +
'.subscribe-card .sub-label{font-size:.85rem;font-weight:700;color:#111827;white-space:nowrap}' +
'.subscribe-card .sub-label span{color:#ec4899}' +
'.sub-form{display:flex;gap:.5rem;flex:1;min-width:220px}' +
'.sub-input{flex:1;padding:.5rem .9rem;border:1px solid #e5e7eb;border-radius:8px;' +
'font-family:Heebo,Arial,sans-serif;font-size:.85rem;outline:none;direction:rtl;' +
'transition:border-color .15s}' +
'.sub-input:focus{border-color:#0d9488;box-shadow:0 0 0 2px rgba(13,148,136,0.1)}' +
'.sub-btn{background:#ec4899;color:#fff;border:none;border-radius:8px;' +
'padding:.5rem 1rem;font-family:Heebo,Arial,sans-serif;font-size:.82rem;' +
'font-weight:700;cursor:pointer;white-space:nowrap;transition:background .15s}' +
'.sub-btn:hover{background:#db2777}' +
'.sub-msg{font-size:.78rem;color:#0d9488;margin-top:.4rem;min-height:1rem;width:100%}' +
'.sub-msg.err{color:#dc2626}' +
/* Keyboard focus — visible outline on everything interactive */
'.video-card:focus-visible,.car-btn:focus-visible,.lock-btn:focus-visible,.sub-btn:focus-visible,' +
'.delete-btn:focus-visible{outline:2px solid #0d9488;outline-offset:2px}' +
/* Mobile */
'@media(max-width:640px){' +
'.hero{flex-direction:column;align-items:flex-start;gap:0.8rem;padding:1.2rem 1.2rem 1rem}' +
'.stats{width:100%;justify-content:flex-start;gap:1.2rem}' +
'.search-wrap{padding:0.7rem 1.2rem}' +
'.subscribe-card{padding:1.1rem 1.2rem}' +
'.topic-section{padding:1.1rem}' +
'.video-track .video-card{flex-basis:78%}' + /* card peeks — invites swipe */
'}' +
'</style></head><body>' +
'<div class="admin-banner">⚙️ מצב ניהול — כפתורי מחיקה פעילים</div>' +
'<svg width="0" height="0" style="position:absolute;overflow:hidden;"><defs>' +
'<linearGradient id="grd-lib" x1="0%" y1="0%" x2="100%" y2="100%">' +
'<stop offset="0%" stop-color="#0d9488"/><stop offset="100%" stop-color="#ec4899"/>' +
'</linearGradient></defs></svg>' +
'<div class="hero">' +
'<div class="hero-inner">' +
'<h1>ספריית <span>הסיכומים</span> שלי</h1>' +
'<p class="sub">סיכומים מדויקים - רק מה שנאמר, עם צעדי פעולה</p>' +
'</div>' +
'<div class="stats">' +
'<div class="stat"><span class="stat-num" id="cv">0</span><span class="stat-label">סרטונים</span></div>' +
'<div class="stat"><span class="stat-num teal" id="ct">0</span><span class="stat-label">נושאים</span></div>' +
'<button class="lock-btn" id="lock-btn" onclick="adminLogin()" title="כניסת מנהלת" aria-label="כניסת מנהלת">🔒</button>' +
'</div></div>' +
'<div class="search-wrap"><div class="search-inner">' +
'<span class="search-icon" aria-hidden="true">🔍</span>' +
'<input class="search-input" id="srch" type="search" placeholder="חיפוש לפי שם סרטון..." aria-label="חיפוש לפי שם סרטון" autocomplete="off">' +
'</div></div>' +
'<div class="container">' +
'<div class="subscribe-card">' +
'<div class="sub-label">📬 סיכומים חדשים <span>ישירות למייל</span></div>' +
'<div class="sub-form">' +
'<input type="email" id="sub-email" class="sub-input" placeholder="המייל שלך..." aria-label="כתובת מייל להרשמה לעדכונים" autocomplete="email">' +
'<button class="sub-btn" onclick="doSubscribe()">הרשמה</button>' +
'</div>' +
'<div id="sub-msg" class="sub-msg" role="status" aria-live="polite"></div>' +
'</div>' +
sections +
'<div class="no-results" id="no-res">לא נמצאו תוצאות לחיפוש זה 🔍</div>' +
'<div style="text-align:center;padding:2.5rem 1rem 1rem;display:flex;flex-wrap:wrap;justify-content:center;gap:.75rem">' +
'<a href="/youtube-pipeline-casestudy.html" style="display:inline-flex;align-items:center;background:#0d9488;color:#fff;padding:.7rem 1.6rem;border-radius:10px;text-decoration:none;font-weight:700;font-size:.9rem;letter-spacing:.01em;transition:opacity .2s" onmouseover="this.style.opacity=\'.85\'" onmouseout="this.style.opacity=\'1\'">Case Study - איך זה נבנה</a>' +
'<a href="/about.html" style="display:inline-flex;align-items:center;background:#ec4899;color:#fff;padding:.7rem 1.6rem;border-radius:10px;text-decoration:none;font-weight:700;font-size:.9rem;letter-spacing:.01em;transition:opacity .2s" onmouseover="this.style.opacity=\'.85\'" onmouseout="this.style.opacity=\'1\'">אודות הפרויקט</a>' +
'</div>' +
'<div class="footer-credit">נבנה עם 💛 <span>Claude</span> · אילת שחק שול · 2026</div>' +
'</div>' +
'<script>' +
'var obs=new IntersectionObserver(function(e){e.forEach(function(x){if(x.isIntersecting){x.target.classList.add("anim-fadeup");obs.unobserve(x.target);}});},{threshold:0.1,rootMargin:"0px 0px -30px 0px"});' +
'document.querySelectorAll(".anim-ready").forEach(function(el){obs.observe(el);});' +
'function countUp(el,target,dur){var s=performance.now();(function step(now){var p=Math.min((now-s)/dur,1);el.textContent=Math.round((1-Math.pow(1-p,3))*target);if(p<1)requestAnimationFrame(step);})(performance.now());}' +
'setTimeout(function(){countUp(document.getElementById("cv"),' + totalVideos + ',1200);setTimeout(function(){countUp(document.getElementById("ct"),' + totalTopics + ',800);},200);},400);' +
'var WEBAPP_URL="' + CONFIG.WEBAPP_URL + '";' +
'var ADMIN_HASH="' + sha256Hex_(CONFIG.ADMIN_CODE) + '";' + /* SHA-256, one-way — the code itself is NOT recoverable from the page */
'var adminCode=null;' + /* set after successful login; sent with delete requests for server-side auth */
/* Admin login via lock button */
'function adminLogin(){' +
'if(document.body.classList.contains("admin-mode")){' +
'document.body.classList.remove("admin-mode");' +
'adminCode=null;' +
'document.getElementById("lock-btn").textContent="🔒";return;}' +
'var code=prompt("קוד גישה:");' +
'if(!code)return;' +
'sha256Hex(code).then(function(h){' +
'if(h===ADMIN_HASH){' +
'document.body.classList.add("admin-mode");' +
'adminCode=code;' +
'document.getElementById("lock-btn").textContent="🔓";' +
'attachDeleteHandlers();' +
'}else{alert("קוד שגוי");}});}' +
'function sha256Hex(s){return crypto.subtle.digest("SHA-256",new TextEncoder().encode(s)).then(function(buf){return Array.prototype.map.call(new Uint8Array(buf),function(b){return("0"+b.toString(16)).slice(-2)}).join("")});}' +
/* Attach delete buttons */
'function attachDeleteHandlers(){' +
'document.querySelectorAll(".delete-btn").forEach(function(btn){' +
'if(btn.dataset.bound)return;btn.dataset.bound="1";' +
'btn.addEventListener("click",function(e){' +
'e.preventDefault();e.stopPropagation();' +
'var card=btn.closest(".video-card");' +
'var vid=card.getAttribute("data-vid");' +
'var titleText=card.querySelector(".card-title").textContent;' +
'if(!confirm("למחוק: "+titleText+"?"))return;' +
'btn.textContent="...";btn.style.pointerEvents="none";' +
'if(!WEBAPP_URL){alert("WEBAPP_URL ריק — הוסיפי אותו ב-Script Properties");return;}' +
'if(!adminCode){alert("חסר קוד אדמין");return;}' +
'fetch(WEBAPP_URL+"?action=delete&vid="+encodeURIComponent(vid)+"&code="+encodeURIComponent(adminCode),{mode:"no-cors"})' +
'.then(function(){card.style.opacity="0.35";card.style.pointerEvents="none";' +
'card.querySelector(".card-title").textContent="נמחק — מתעדכן...";})' +
'.catch(function(){alert("שגיאה — בדקי את WEBAPP_URL");});' +
'});});}' +
/* Subscribe */
'function doSubscribe(){' +
'var email=document.getElementById("sub-email").value.trim();' +
'var msg=document.getElementById("sub-msg");' +
'if(!email||!email.includes("@")){msg.className="sub-msg err";msg.textContent="מייל לא תקין";return;}' +
'if(!WEBAPP_URL){msg.className="sub-msg err";msg.textContent="שירות ההרשמה לא מוגדר עדיין";return;}' +
'msg.className="sub-msg";msg.textContent="שולח...";' +
'var img=new Image();' +
'img.onload=img.onerror=function(){msg.textContent="✓ נרשמת! בדקי את תיבת הדואר לאישור";' +
'document.getElementById("sub-email").value="";};' +
'img.src=WEBAPP_URL+"?action=subscribe&email="+encodeURIComponent(email)+"&t="+Date.now();} ' +
'document.getElementById("srch").addEventListener("input",function(){' +
'var q=this.value.trim().toLowerCase();var vis=0;' +
'document.querySelectorAll(".topic-section").forEach(function(sec){' +
'var cv=0;sec.querySelectorAll(".video-card").forEach(function(c){' +
'var m=!q||c.textContent.toLowerCase().indexOf(q)>-1;' +
'c.style.display=m?"":"none";if(m)cv++;});' +
'sec.style.display=cv?"":"none";if(cv)vis++;});' +
'var nr=document.getElementById("no-res");' +
'if(nr)nr.style.display=q&&!vis?"block":"none";' +
'window.dispatchEvent(new Event("resize"));});' + /* re-check carousel arrows after filtering */
/* Carousels: arrows appear only when a section actually overflows */
'document.querySelectorAll(".topic-section").forEach(function(sec){' +
'var track=sec.querySelector(".video-track");' +
'var prev=sec.querySelector(".car-prev"),next=sec.querySelector(".car-next");' +
'if(!track||!prev||!next)return;' +
'function upd(){var over=track.scrollWidth>track.clientWidth+4;' +
'prev.style.display=over?"flex":"none";next.style.display=over?"flex":"none";}' +
'function step(){return Math.max(track.clientWidth*0.8,240);}' +
'next.addEventListener("click",function(){track.scrollBy({left:-step(),behavior:"smooth"});});' + /* RTL: forward = scroll left */
'prev.addEventListener("click",function(){track.scrollBy({left:step(),behavior:"smooth"});});' +
'window.addEventListener("resize",upd);upd();});' +
'</script></body></html>';
}

function getIconCode(topic) {
  var codes = { "AI Tools": "129302", "Productivity": "128200", "Entrepreneurship": "128640", "General": "128204" };
  return codes[topic] || "128204";
}

// ===========================================
// YOUTUBE TRANSCRIPT (via Supadata.ai)
// ===========================================
function getYouTubeTranscript(url) {
  const resp = UrlFetchApp.fetch(
    "https://api.supadata.ai/v1/youtube/transcript?url=" + encodeURIComponent(url),
    {
      muteHttpExceptions: true,
      headers: { "x-api-key": CONFIG.SUPADATA_KEY }
    }
  );

  if (resp.getResponseCode() !== 200) throw new Error("Transcript fetch failed: " + resp.getResponseCode());

  const data = JSON.parse(resp.getContentText());
  if (!data || !data.content) throw new Error("No transcript content");

  return Array.isArray(data.content)
    ? data.content.map(function(s) { return s.text || ""; }).join(" ").replace(/\s+/g, " ").trim()
    : data.content.toString().trim();
}

function extractVideoId(url) {
  const match = url.match(/(?:v=|youtu\.be\/)([a-zA-Z0-9_-]{11})/);
  return match ? match[1] : null;
}

// ===========================================
// CLAUDE API
// ===========================================
function generateHTMLWithClaude(transcript, title, sourceUrl) {
  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    muteHttpExceptions: true,
    headers: {
      "x-api-key": CONFIG.CLAUDE_API_KEY,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json"
    },
    payload: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 16000,
      messages: [{ role: "user", content: buildPrompt(transcript, title, sourceUrl) }]
    })
  });

  const result = JSON.parse(response.getContentText());
  if (result.error) throw new Error("Claude API: " + result.error.message);

  const text = result.content[0].text;
  const htmlMatch = text.match(/```html\n?([\s\S]*?)```/) ||
                    text.match(/(<!DOCTYPE[\s\S]*<\/html>)/i);
  return htmlMatch ? htmlMatch[1].trim() : text.trim();
}

// ===========================================
// QUALITY GATE
// ===========================================
// A summary page goes public (Netlify) and lands in inboxes the moment it is
// generated, so it must be checked BEFORE deploy, not after. The gate verifies
// the structure the prompt demands; content quality stays on the prompt's
// IRON RULE. Claude output varies between calls, so one retry is worth it;
// a second failure marks the row Error and alerts the owner instead of
// publishing a broken page.
function validateSummaryHtml_(html) {
  const problems = [];
  if (!html || html.length < 3000) {
    problems.push("page too short (" + (html ? html.length : 0) + " chars)");
    return problems;
  }
  if (!/<\/html>\s*$/i.test(html)) problems.push("truncated - no closing </html>");
  if (html.indexOf("```") !== -1) problems.push("markdown fences left in output");
  if (!/<style/i.test(html)) problems.push("missing <style> block");
  if (!/dir=["']?rtl/i.test(html)) problems.push("no RTL direction");
  // Content sections only — cosmetic parts (page-footer) are not worth
  // blocking a publish over: 7 of the first 12 production pages lack the
  // footer and are perfectly usable.
  ["hero", "central-idea", "level-badge", "final-message"].forEach(function(cls) {
    if (html.indexOf(cls) === -1) problems.push("missing section: ." + cls);
  });
  const h1 = extractHebrewTitle(html);
  if (!h1 || !/[א-ת]/.test(h1)) problems.push("h1 missing or not Hebrew");
  return problems;
}

function generateValidatedHtml_(transcript, title, sourceUrl) {
  let problems = [];
  for (let attempt = 1; attempt <= 2; attempt++) {
    const html = generateHTMLWithClaude(transcript, title, sourceUrl);
    problems = validateSummaryHtml_(html);
    if (problems.length === 0) return html;
    Logger.log("Quality gate attempt " + attempt + " failed: " + problems.join("; "));
  }
  sendQualityAlert_(title, sourceUrl, problems);
  throw new Error("Quality gate: " + problems.join("; "));
}

function sendQualityAlert_(title, sourceUrl, problems) {
  try {
    GmailApp.sendEmail(CONFIG.EMAIL_TO, "Summary failed quality gate: " + title, "", {
      htmlBody: '<div dir="rtl" style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
        '<h2 style="color:#ec4899;">הסיכום לא עמד בתקן ולא פורסם</h2>' +
        '<p><strong>סרטון:</strong> <a href="' + sourceUrl + '">' + title + '</a></p>' +
        '<p><strong>מה נמצא (אחרי שני ניסיונות):</strong></p>' +
        '<ul>' + problems.map(function(p) { return '<li>' + p + '</li>'; }).join('') + '</ul>' +
        '<p>הסרטון מסומן Error בגיליון. מחיקת הסטטוס תפעיל ניסיון חדש.</p></div>',
      name: "YouTube Doc Bot"
    });
  } catch (e) { Logger.log("Quality alert email failed: " + e.message); }
}

function buildPrompt(transcript, title, sourceUrl) {
  var css = "@keyframes fadeUp{from{opacity:0;transform:translateY(22px)}to{opacity:1;transform:translateY(0)}}" +
    "@keyframes slideRight{from{opacity:0;transform:translateX(-12px)}to{opacity:1;transform:translateX(0)}}" +
    ".anim-ready{opacity:0}.anim-fadeup{animation:fadeUp 0.75s cubic-bezier(0.22,1,0.36,1) both}" +
    "*{box-sizing:border-box;margin:0;padding:0}" +
    "body{font-family:Segoe UI,Arial,sans-serif;background:#fff;" +
    "background-image:linear-gradient(#e5e7eb 1px,transparent 1px),linear-gradient(90deg,#e5e7eb 1px,transparent 1px);" +
    "background-size:28px 28px;color:#1f2937;line-height:1.7;padding:2rem 1rem}" +
    ".container{max-width:820px;margin:0 auto}" +
    "a.back{display:block;text-align:right;font-size:0.85rem;color:#6b7280;text-decoration:none;margin-bottom:1.5rem}" +
    "a.back:hover{color:#ec4899}" +
    ".hero{text-align:center;padding:2.5rem 1.5rem;background:white;border-radius:16px;" +
    "margin-bottom:2rem;border:2px solid #ec489933;box-shadow:0 4px 24px #ec489911;animation:fadeUp 0.5s ease}" +
    ".hero .tag{display:inline-block;background:#fce7f3;color:#ec4899;border:1px solid #f9a8d4;" +
    "border-radius:20px;padding:0.3rem 1rem;font-size:0.8rem;margin-bottom:1rem;font-weight:600}" +
    ".hero h1{font-size:1.9rem;font-weight:800;color:#1f2937;margin-bottom:1rem;line-height:1.3}" +
    ".hero .subtitle{color:#6b7280;font-size:1rem;max-width:540px;margin:0 auto 1.2rem}" +
    ".hero a.yt{display:inline-block;background:#ec4899;color:#fff;text-decoration:none;" +
    "padding:0.55rem 1.4rem;border-radius:8px;font-size:0.9rem;font-weight:600}" +
    ".central-idea{background:#fce7f3;border:1px solid #f9a8d4;border-right:4px solid #ec4899;" +
    "border-radius:12px;padding:1.2rem 1.5rem;margin-bottom:2rem;font-size:1rem;color:#1f2937}" +
    ".central-idea strong{color:#ec4899}" +
    ".level{background:white;border:1px solid #e5e7eb;border-radius:14px;padding:1.6rem;" +
    "margin-bottom:1.5rem;box-shadow:0 2px 12px #0000000a;transition:box-shadow 0.2s,transform 0.2s}" +
    ".level:hover{box-shadow:0 6px 24px #0000001a;transform:translateY(-2px)}" +
    ".level-header{display:flex;align-items:center;gap:0.8rem;margin-bottom:1rem}" +
    ".level-badge{background:linear-gradient(135deg,#ec4899,#374151);color:#fff;font-size:0.75rem;" +
    "font-weight:700;padding:0.25rem 0.75rem;border-radius:20px}" +
    ".level h2{font-size:1.2rem;font-weight:700;color:#1f2937}" +
    ".tagline{color:#6b7280;font-size:0.9rem;margin-bottom:1rem;padding-bottom:1rem;border-bottom:1px solid #f3f4f6}" +
    ".tagline strong{color:#ec4899}" +
    ".tool-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:0.75rem;margin-bottom:1rem}" +
    ".tool-card{background:#fafafa;border:1px solid #e5e7eb;border-radius:10px;padding:0.9rem 1rem;" +
    "transition:border-color 0.2s,transform 0.2s}" +
    ".tool-card:hover{border-color:#f9a8d4;transform:translateY(-2px)}" +
    ".tool-name{font-weight:700;color:#ec4899;margin-bottom:0.3rem;font-size:0.95rem}" +
    ".tool-card p{font-size:0.85rem;color:#6b7280}" +
    ".context-ways{display:flex;flex-direction:column;gap:0.6rem;margin-bottom:1rem}" +
    ".way{background:#fafafa;border:1px solid #e5e7eb;border-radius:8px;padding:0.75rem 1rem;" +
    "display:flex;gap:0.75rem;align-items:flex-start}" +
    ".way-num{background:#fce7f3;color:#ec4899;border:1px solid #f9a8d4;border-radius:50%;" +
    "width:24px;height:24px;display:flex;align-items:center;justify-content:center;" +
    "font-size:0.75rem;font-weight:700;flex-shrink:0;margin-top:0.1rem}" +
    ".way p{font-size:0.88rem;color:#374151}" +
    ".way p strong{color:#1f2937}" +
    ".definition-box{background:#111827;border-radius:10px;padding:1rem 1.2rem;margin-bottom:1rem}" +
    ".def-label{font-size:0.72rem;font-weight:700;color:#0d9488;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:0.4rem}" +
    ".definition-box p{font-size:1rem;font-weight:700;color:#f9fafb;line-height:1.5}" +
    ".def-sub{font-size:0.82rem;font-weight:400;color:#9ca3af;margin-top:0.4rem}" +
    ".system-grid{display:flex;flex-direction:column;gap:0.6rem;margin-bottom:1rem}" +
    ".system-option{background:#fafafa;border:1px solid #e5e7eb;border-radius:8px;padding:0.75rem 1rem;" +
    "display:flex;justify-content:space-between;align-items:center;gap:1rem}" +
    ".opt-name{font-weight:700;color:#ec4899;font-size:0.9rem;min-width:130px}" +
    ".opt-desc{font-size:0.85rem;color:#6b7280;flex:1}" +
    ".difficulty{font-size:0.75rem;padding:0.2rem 0.6rem;border-radius:20px;white-space:nowrap;font-weight:600}" +
    ".easy{background:#f3f4f6;color:#374151;border:1px solid #d1d5db}" +
    ".medium{background:#fef3c7;color:#92400e;border:1px solid #fcd34d}" +
    ".advanced{background:#fee2e2;color:#991b1b;border:1px solid #fca5a5}" +
    ".example-box{background:#f0fdfa;border:1px solid #99f6e4;border-right:3px solid #0d9488;" +
    "border-radius:10px;padding:1rem 1.2rem;margin-bottom:1rem}" +
    ".example-title{font-size:0.78rem;font-weight:700;color:#0d9488;margin-bottom:0.7rem;text-transform:uppercase}" +
    ".example-steps{display:flex;flex-direction:column;gap:0.5rem}" +
    ".example-step{display:flex;gap:0.75rem;align-items:flex-start}" +
    ".step-icon{font-size:1rem;flex-shrink:0;margin-top:0.05rem}" +
    ".example-step p{font-size:0.85rem;color:#374151}" +
    ".example-step p strong{color:#1f2937}" +
    ".actions-box{background:#f0fdfa;border:1px solid #99f6e4;border-right:3px solid #0d9488;" +
    "border-radius:10px;padding:1rem 1.2rem;margin-top:0.5rem}" +
    ".actions-title{font-size:0.8rem;font-weight:700;color:#0d9488;margin-bottom:0.6rem}" +
    ".actions-box ol{padding-right:1.2rem;display:flex;flex-direction:column;gap:0.4rem}" +
    ".actions-box li{font-size:0.88rem;color:#374151}" +
    ".actions-box li strong{color:#0f766e}" +
    ".section-label{font-size:0.85rem;font-weight:700;color:#ec4899;margin-bottom:0.6rem;display:block}" +
    ".final-message{background:white;border:1px solid #e5e7eb;border-radius:14px;padding:1.5rem;" +
    "margin-top:2rem;text-align:center;box-shadow:0 2px 12px #0000000a}" +
    ".final-message p{color:#6b7280;font-size:0.95rem;max-width:580px;margin:0 auto 1rem}" +
    ".final-message p strong{color:#ec4899}" +
    ".source{font-size:0.78rem;color:#9ca3af;margin-top:1rem}" +
    ".source a{color:#ec4899;text-decoration:none}" +
    ".page-footer{margin-top:3rem;padding:1.5rem;text-align:center;border-top:1px solid #f0f0f0}" +
    ".page-footer .back-link{display:inline-flex;align-items:center;gap:0.5rem;background:#f0fdf9;" +
    "border:1px solid #99f6e4;border-radius:20px;padding:0.45rem 1.1rem;color:#0d9488;" +
    "text-decoration:none;font-size:0.88rem;font-weight:600;transition:all 0.15s}" +
    ".page-footer .back-link:hover{background:#0d9488;color:#fff;border-color:#0d9488}" +
    ".page-footer .credit{font-size:0.75rem;color:#9ca3af;margin-top:0.75rem}";

  return "You generate a complete RTL Hebrew HTML summary page from a YouTube transcript.\n" +
    "IRON RULE: ONLY content from the transcript. Zero inventions. Zero additions.\n" +
    "If something is NOT explicitly said in the transcript — do NOT write it. Not as an example, not as a tip, not as an explanation. If the transcript has only 2 tools, write 2 — not 3. If a name or concept is not mentioned — do not add it. Every sentence must be traceable to the transcript.\n\n" +
    "Title: " + title + "\n" +
    "Source: " + sourceUrl + "\n\n" +
    "Use EXACTLY this CSS inside <style> tags:\n" + css + "\n\n" +
    "Build this EXACT HTML structure:\n\n" +
    "1. <a href='/' class='back'>חזרה לכל הסרטונים</a>\n\n" +
    "2. .hero: .tag (speaker+date), h1 (Hebrew title), .subtitle (one sentence), <a class='yt' href='SOURCE_URL'>▶ לסרטון המקורי</a>\n\n" +
    "3. .central-idea: core message of the video\n\n" +
    "4. .level (רמה 1): .level-badge 'רמה 1', h2, .tagline, .tool-grid with .tool-card items (each option/tool mentioned: .tool-name + p description), .actions-box with ol>li actionable steps\n\n" +
    "5. .level (רמה 2): .level-badge 'רמה 2', h2, .tagline, .context-ways with .way items (.way-num + p for each method/approach), .actions-box\n\n" +
    "6. .level (רמה 3): .level-badge 'רמה 3', h2, .tagline, .definition-box (.def-label + exact quote in p + .def-sub explaining speaker's own system), .context-ways (2 capabilities of the system), .example-box (.example-title + .example-steps with .step-icon + p for each step of the example), .section-label + .system-grid (.system-option with .opt-name + .opt-desc + .difficulty easy/medium/advanced), .actions-box\n\n" +
    "7. .final-message: closing thought + .source with links\n\n" +
    "8. <script> IntersectionObserver that adds class 'anim-fadeup' to '.anim-ready' elements when visible\n\n" +
    "9. .page-footer: <a href='/' class='back-link'>&#x2190; ספריית הסיכומים</a> then <p class='credit'>נבנה עם &#x1F49B; Claude &middot; אילת שחק שול &middot; 2026</p>\n\n" +
    "Add class='anim-ready' to each .level and .final-message.\n" +
    "NO dark backgrounds except .definition-box.\n" +
    "Mixed Hebrew/English titles: always add dir='rtl' to the h1 tag, and wrap each English word/phrase inside <bdi> tags so bidi rendering is correct — e.g. <h1 dir='rtl'>שליטה בסוכני AI של <bdi>Claude</bdi> – 5 דוגמאות</h1>.\n" +
    "Quotes and definitions in .definition-box: translate to Hebrew, with the original English in parentheses after — e.g. '\"פרויקט הוא בית קבוע לזרמי עבודה חוזרים\" (A project is a permanent home for recurring work streams)'.\n" +
    "Return complete HTML only. No markdown fences.\n\n" +
    "TRANSCRIPT:\n" + transcript.substring(0, 12000);
}

// ===========================================
// EXTRACT HEBREW TITLE FROM GENERATED HTML
// ===========================================
function extractHebrewTitle(html) {
  if (!html) return null;
  // Match the <h1> tag (may have attributes like dir='rtl')
  var match = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  if (!match) return null;
  // Strip inner HTML tags (<bdi>, <span>, etc.) to get plain text
  return match[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim() || null;
}

// ===========================================
// WEB APP HANDLER — Deploy as Web App to enable delete/subscribe
// Execute as: Me | Who has access: Anyone
// ===========================================
function doGet(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    if (!e || !e.parameter || !e.parameter.action) {
      output.setContent(JSON.stringify({status:"error",message:"No action"}));
      return output;
    }
    var action = e.parameter.action;
    var vid    = (e.parameter.vid || "").toString().trim();
    var email  = (e.parameter.email || "").toString().trim().toLowerCase();
    var token  = (e.parameter.token || "").toString().trim();
    var code   = (e.parameter.code || "").toString();

    if (action === "delete" && vid) {
      // Server-side auth: ADMIN_CODE must be set AND match the code from the request.
      // Without this, anyone who knows the WEBAPP_URL could delete any video.
      if (!CONFIG.ADMIN_CODE || code !== CONFIG.ADMIN_CODE) {
        output.setContent(JSON.stringify({status:"error",message:"Unauthorized"}));
        return output;
      }
      deleteVideoFromLibrary(vid);
      output.setContent(JSON.stringify({status:"ok",deleted:vid}));
    } else if (action === "subscribe" && email) {
      var result = addSubscriber(email);
      output.setContent(JSON.stringify({status:"ok",message:result}));
    } else if (action === "unsubscribe" && token) {
      // Show confirmation page — don't remove yet
      var libUrl = "https://" + CONFIG.NETLIFY_SITE + ".netlify.app";
      var confirmUrl = ScriptApp.getService().getUrl() + "?action=confirm_unsubscribe&token=" + encodeURIComponent(token);
      return HtmlService.createHtmlOutput(
        '<html dir="rtl"><body style="font-family:Arial,sans-serif;text-align:center;padding:4rem 2rem;color:#111827;background:#fafafa">' +
        '<div style="max-width:420px;margin:0 auto;background:#fff;border-radius:16px;padding:2.5rem;box-shadow:0 2px 20px rgba(0,0,0,0.07)">' +
        '<div style="height:3px;background:linear-gradient(90deg,#0d9488,#ec4899);border-radius:3px;margin-bottom:2rem"></div>' +
        '<div style="font-size:2rem;margin-bottom:1rem">📬</div>' +
        '<h2 style="margin-bottom:.75rem;font-size:1.2rem">הסרה מרשימת התפוצה</h2>' +
        '<p style="color:#6b7280;font-size:.9rem;margin-bottom:2rem">האם לבטל את ההרשמה לסיכומים של אילת שחק שול?</p>' +
        '<div style="display:flex;gap:.75rem;justify-content:center;flex-wrap:wrap">' +
        '<a href="' + confirmUrl + '" style="background:#dc2626;color:#fff;padding:.65rem 1.4rem;border-radius:8px;text-decoration:none;font-weight:700;font-size:.9rem">כן, הסירי אותי</a>' +
        '<a href="' + libUrl + '" style="background:#f0fdf9;color:#0d9488;border:1px solid #99f6e4;padding:.65rem 1.4rem;border-radius:8px;text-decoration:none;font-weight:700;font-size:.9rem">לא, המשיכי לשלוח</a>' +
        '</div></div></body></html>');
    } else if (action === "confirm_unsubscribe" && token) {
      removeSubscriberByToken(token);
      var libUrl2 = "https://" + CONFIG.NETLIFY_SITE + ".netlify.app";
      return HtmlService.createHtmlOutput(
        '<html dir="rtl"><head><meta charset="UTF-8"></head><body style="font-family:Arial,sans-serif;text-align:center;padding:4rem 2rem;color:#111827;background:#fafafa">' +
        '<div style="max-width:420px;margin:0 auto;background:#fff;border-radius:16px;padding:2.5rem;box-shadow:0 2px 20px rgba(0,0,0,0.07)">' +
        '<div style="height:3px;background:linear-gradient(90deg,#0d9488,#ec4899);border-radius:3px;margin-bottom:2rem"></div>' +
        '<div style="font-size:2.5rem;margin-bottom:1rem;color:#0d9488">✓</div>' +
        '<h2 style="color:#0d9488;margin-bottom:.75rem;font-size:1.3rem">הוסרת בהצלחה</h2>' +
        '<p style="color:#6b7280;font-size:.9rem;margin-bottom:2rem">לא תקבלי יותר עדכונים מהספרייה.<br>תמיד אפשר להירשם מחדש.</p>' +
        '<a href="' + libUrl2 + '" style="display:inline-block;background:#ec4899;color:#fff;padding:.65rem 1.4rem;border-radius:8px;text-decoration:none;font-weight:700;font-size:.9rem">חזרה לספרייה</a>' +
        '</div></body></html>');
    } else {
      output.setContent(JSON.stringify({status:"error",message:"Unknown action"}));
    }
  } catch(err) {
    output.setContent(JSON.stringify({status:"error",message:err.message}));
  }
  return output;
}

function deleteVideoFromLibrary(videoId) {
  // 1. Mark as Deleted in sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 6).getValues();
  data.forEach(function(row, i) {
    const url = (row[CONFIG.COL_URL-1]||"").toString().trim();
    if (!url) return;
    const vid = extractVideoId(url);
    if (vid === videoId) {
      sheet.getRange(CONFIG.START_ROW + i, CONFIG.COL_STATUS).setValue("Deleted");
    }
  });

  // 2. Remove from ScriptProperties
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("vid_html_"      + videoId);
  props.deleteProperty("vid_title_he_"  + videoId);

  // 3. Redeploy index without this video
  redeployIndex();
  Logger.log("Deleted and redeployed: " + videoId);
}

// ===========================================
// SUBSCRIBERS
// Sheet "Subscribers": A=Email, B=Token, C=SignupDate, D=Status, E=RemovedAt
// ===========================================
function getSubscribersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.SUBSCRIBERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SUBSCRIBERS_SHEET);
    sh.getRange("A1:E1").setValues([["Email","Token","SignupDate","Status","RemovedAt"]]);
    sh.getRange("A1:E1").setFontWeight("bold");
  }
  return sh;
}

function addSubscriber(email) {
  var sh = getSubscribersSheet();
  var data = sh.getDataRange().getValues();
  // Check if already subscribed
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0]||"").toString().toLowerCase() === email && data[i][3] === "Active") {
      return "already_subscribed";
    }
    if ((data[i][0]||"").toString().toLowerCase() === email && data[i][3] === "Unsubscribed") {
      sh.getRange(i+1, 4).setValue("Active"); // resubscribe
      sendSubscribeConfirmation(email, data[i][1]);
      return "resubscribed";
    }
  }
  var token = Utilities.getUuid();
  sh.appendRow([email, token, new Date().toLocaleDateString("he-IL"), "Active"]);
  sendSubscribeConfirmation(email, token);
  return "subscribed";
}

function removeSubscriberByToken(token) {
  var sh = getSubscribersSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][1]||"").toString() === token) {
      var ts = new Date().toLocaleString("he-IL", {timeZone:"Asia/Jerusalem"});
      sh.getRange(i+1, 4).setValue("Unsubscribed");
      sh.getRange(i+1, 5).setValue(ts);
      return;
    }
  }
}

function sendSubscribeConfirmation(email, token) {
  var unsubUrl = CONFIG.WEBAPP_URL + "?action=unsubscribe&token=" + token;
  var libUrl   = "https://" + CONFIG.NETLIFY_SITE + ".netlify.app";
  MailApp.sendEmail({
    to: email,
    subject: "✓ נרשמת לספריית הסיכומים של אילת",
    htmlBody:
      '<div dir="rtl" style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;color:#111827">' +
      '<div style="height:3px;background:linear-gradient(90deg,#0d9488,#ec4899)"></div>' +
      '<div style="padding:2rem">' +
      '<h2 style="color:#ec4899;margin-bottom:.5rem">נרשמת!</h2>' +
      '<p style="color:#374151">תקבלי התראה בכל פעם שיתווסף סיכום חדש לספרייה.</p>' +
      '<p style="margin-top:1.5rem"><a href="' + libUrl + '" style="background:#ec4899;color:#fff;padding:.6rem 1.2rem;border-radius:8px;text-decoration:none;font-weight:700">לספרייה →</a></p>' +
      '<p style="margin-top:2rem;font-size:.78rem;color:#9ca3af">לא רוצה יותר? <a href="' + unsubUrl + '" style="color:#6b7280">הסירי אותי מהרשימה</a></p>' +
      '</div></div>'
  });
}

function sendToAllSubscribers(title, youtubeUrl, netlifyUrl, topic) {
  var sh = getSubscribersSheet();
  var data = sh.getDataRange().getValues();
  var libUrl = "https://" + CONFIG.NETLIFY_SITE + ".netlify.app";
  var ownerEmail = (CONFIG.EMAIL_TO || "").toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var email  = (data[i][0]||"").toString().trim();
    var token  = (data[i][1]||"").toString();
    var status = (data[i][3]||"").toString();
    if (!email || status !== "Active") continue;
    // Skip the owner — they already get the dedicated owner notification via sendEmailNotification.
    // Without this filter, the owner receives two emails per video (owner + subscriber).
    if (email.toLowerCase() === ownerEmail) continue;
    var unsubUrl = CONFIG.WEBAPP_URL + "?action=unsubscribe&token=" + token;
    try {
      MailApp.sendEmail({
        to: email,
        subject: "📖 סיכום חדש: " + title,
        htmlBody:
          '<div dir="rtl" style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;color:#111827">' +
          '<div style="height:3px;background:linear-gradient(90deg,#0d9488,#ec4899)"></div>' +
          '<div style="padding:2rem">' +
          '<div style="font-size:.75rem;font-weight:700;color:#0d9488;text-transform:uppercase;margin-bottom:.5rem">' + topic + '</div>' +
          '<h2 style="margin:0 0 1rem;line-height:1.3">' + title + '</h2>' +
          '<div style="display:flex;gap:.75rem;margin-top:1.5rem">' +
          '<a href="' + netlifyUrl + '" style="background:#ec4899;color:#fff;padding:.55rem 1.2rem;border-radius:8px;text-decoration:none;font-weight:700;font-size:.9rem">לסיכום המלא →</a>' +
          '<a href="' + youtubeUrl + '" style="background:#f3f4f6;color:#374151;padding:.55rem 1.2rem;border-radius:8px;text-decoration:none;font-weight:700;font-size:.9rem">▶ לסרטון</a>' +
          '<a href="' + libUrl + '" style="background:#f0fdf9;color:#0d9488;border:1px solid #99f6e4;padding:.55rem 1.2rem;border-radius:8px;text-decoration:none;font-weight:700;font-size:.9rem">כל הספרייה</a>' +
          '</div>' +
          '<p style="margin-top:2rem;font-size:.75rem;color:#9ca3af">הגעת למייל זה כי נרשמת לספריית הסיכומים של אילת שחק שול. <a href="' + unsubUrl + '" style="color:#6b7280">הסרה מהרשימה</a></p>' +
          '</div></div>'
      });
    } catch(err) {
      Logger.log("Failed to send to " + email + ": " + err.message);
    }
  }
}

// ===========================================
// REDEPLOY INDEX ONLY (no new video)
// ===========================================
function redeployIndex() {
  assertConfig_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 6).getValues();
  const props = PropertiesService.getScriptProperties();
  const videos = [];

  data.forEach(function(row) {
    const url    = (row[CONFIG.COL_URL    - 1] || "").toString().trim();
    const status = (row[CONFIG.COL_STATUS - 1] || "").toString().trim();
    const topic  = (row[CONFIG.COL_TOPIC  - 1] || "General").toString().trim();
    const date   = (row[CONFIG.COL_DATE   - 1] || "").toString().trim();
    const rawTitle = (row[CONFIG.COL_TITLE - 1] || "").toString().trim();
    if (!url || !status.includes("Done")) return;
    const vid = extractVideoId(url);
    if (!vid) return;
    const heTitle = props.getProperty("vid_title_he_" + vid) || rawTitle;
    videos.push({ videoId: vid, title: heTitle, topic: topic, date: date });
  });

  const siteId = getOrCreateNetlifySite(CONFIG.NETLIFY_SITE);
  const blobs = [Utilities.newBlob(buildIndexPage(videos), "text/html", "index.html")];

  videos.forEach(function(v) {
    const html = props.getProperty("vid_html_" + v.videoId);
    if (html) blobs.push(Utilities.newBlob(html, "text/html", v.videoId + ".html"));
  });

  addStaticPageBlobs_(blobs);

  const zipBlob = Utilities.zip(blobs, "deploy.zip");
  const resp = UrlFetchApp.fetch(
    "https://api.netlify.com/api/v1/sites/" + siteId + "/deploys",
    {
      method: "post",
      muteHttpExceptions: true,
      headers: {
        "Authorization": "Bearer " + CONFIG.NETLIFY_TOKEN,
        "Content-Type": "application/zip"
      },
      payload: zipBlob.getBytes()
    }
  );
  const deployData = JSON.parse(resp.getContentText());
  if (deployData.error) throw new Error("Netlify: " + deployData.error);
  waitForDeploy(deployData.id);
  Logger.log("Index redeployed.");
}

// ===========================================
// GOOGLE DRIVE BACKUP
// ===========================================
function saveToDrive(html, title, videoId) {
  const fileName = sanitizeFilename(title) + "_" + videoId + ".html";
  const blob = Utilities.newBlob(html, "text/html", fileName);
  const folder = CONFIG.DRIVE_FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID)
    : DriveApp.getRootFolder();
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function sanitizeFilename(name) {
  return name.replace(/[^a-zA-Z0-9\s_-]/g, "").replace(/\s+/g, "_").substring(0, 50);
}

// ===========================================
// OWNER EMAIL NOTIFICATION
// ===========================================
function sendEmailNotification(title, youtubeUrl, netlifyUrl, topic) {
  const indexUrl = "https://" + CONFIG.NETLIFY_SITE + ".netlify.app";
  GmailApp.sendEmail(CONFIG.EMAIL_TO, "New summary ready: " + title, "", {
    htmlBody: '<div dir="rtl" style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
      '<h2 style="color:#ec4899;">New summary ready!</h2>' +
      '<p><strong>Video:</strong> <a href="' + youtubeUrl + '">' + title + '</a></p>' +
      '<p><strong>Topic:</strong> ' + topic + '</p>' +
      '<p style="margin-top:1rem;">' +
      '<a href="' + netlifyUrl + '" style="background:#ec4899;color:white;padding:0.6rem 1.2rem;border-radius:8px;text-decoration:none;font-weight:600;">Open summary</a>' +
      '&nbsp;&nbsp;' +
      '<a href="' + indexUrl + '" style="background:#0d9488;color:white;padding:0.6rem 1.2rem;border-radius:8px;text-decoration:none;font-weight:600;">All videos</a>' +
      '</p></div>',
    name: "YouTube Doc Bot"
  });
}

// ===========================================
// DAILY TRIGGER - run once to set up
// ===========================================
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("processNewVideos").timeBased().atHour(8).everyDays(1).create();
  Logger.log("Daily trigger created - runs every day at 08:00");
}

// ===========================================
// DEBUG - inspect what's stored in Drive and ScriptProperties
// ===========================================
function debugDriveFiles() {
  Logger.log("=== Searching Drive for HTML files ===");
  try {
    const allHtml = DriveApp.searchFiles('mimeType = "text/html" and trashed = false');
    let count = 0;
    while (allHtml.hasNext()) {
      const f = allHtml.next();
      Logger.log("FILE: " + f.getName() + " | ID: " + f.getId());
      count++;
    }
    Logger.log("Total HTML files found: " + count);
  } catch(e) { Logger.log("Search error: " + e.message); }

  Logger.log("=== ScriptProperties ===");
  const props = PropertiesService.getScriptProperties().getProperties();
  Object.keys(props).forEach(function(k) { Logger.log(k + ": " + props[k].substring(0, 80)); });
}
