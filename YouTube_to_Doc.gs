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
  SUBSCRIBERS_SHEET:      "Subscribers",
  // Hebrew display names for topics. Sheet values stay in English.
  TOPIC_LABELS:           { "AI Tools": "כלי AI", "Productivity": "פרודוקטיביות", "Entrepreneurship": "יזמות", "General": "כללי" }
};

// Display label for a topic (falls back to the raw sheet value)
function topicLabel_(t) { return CONFIG.TOPIC_LABELS[t] || t; }

// Public base URL of the library site
function siteUrl_() { return "https://" + CONFIG.NETLIFY_SITE + ".netlify.app"; }

// Escape text for safe use inside HTML attributes / text
function escHtml_(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Share preview (WhatsApp / LinkedIn) tags for a summary page.
// Summary HTML comes from Claude without them, so they are injected at deploy time.
function addShareMeta_(html, title, videoId) {
  if (!html || /property=["']og:title/i.test(html)) return html;
  var t = escHtml_(title || "סיכום סרטון");
  var tags =
    '<meta name="description" content="' + t + ' | סיכום בעברית: רק מה שנאמר, עם צעדי פעולה">' +
    '<meta property="og:type" content="article"><meta property="og:locale" content="he_IL">' +
    '<meta property="og:site_name" content="ספריית הסיכומים | אילת שחק שול">' +
    '<meta property="og:title" content="' + t + '">' +
    '<meta property="og:description" content="סיכום בעברית: רק מה שנאמר, עם צעדי פעולה">' +
    '<meta property="og:url" content="' + siteUrl_() + '/' + videoId + '.html">' +
    '<meta property="og:image" content="' + siteUrl_() + '/og-library.png">' +
    '<meta name="twitter:card" content="summary_large_image">';
  return /<\/head>/i.test(html) ? html.replace(/<\/head>/i, tags + "</head>") : html;
}

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

      const html = generateValidatedHtml_(transcript, title, url, topic);

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
// Copies the HTML pages in the repo's static-pages/ folder (main branch) into
// the Drive static folder, then redeploys. GitHub is the source of truth for
// these pages; the Drive folder is what the deploy reads.
const STATIC_PAGES_REPO_ = "https://raw.githubusercontent.com/ayulaa1-supreme/youtube-pipeline/main/static-pages/";
const STATIC_PAGES_SYNCED_ = ["youtube-pipeline-casestudy.html", "about.html"];

function syncStaticPagesFromGitHub() {
  assertConfig_();
  if (!CONFIG.STATIC_PAGES_FOLDER_ID) throw new Error("STATIC_PAGES_FOLDER_ID is not set");
  const folder = DriveApp.getFolderById(CONFIG.STATIC_PAGES_FOLDER_ID);
  STATIC_PAGES_SYNCED_.forEach(function(name) {
    const res = UrlFetchApp.fetch(STATIC_PAGES_REPO_ + name + "?t=" + Date.now(), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) { Logger.log("Skipped " + name + ": HTTP " + res.getResponseCode()); return; }
    const html = res.getContentText("UTF-8");
    if (!/^<!DOCTYPE html>/i.test(html.trim())) { Logger.log("Skipped " + name + ": not an HTML page"); return; }
    const files = folder.getFilesByName(name);
    if (files.hasNext()) files.next().setContent(html);
    else folder.createFile(Utilities.newBlob(html, "text/html", name));
    Logger.log("Synced " + name + " (" + html.length + " chars)");
  });
  redeployIndex();
}

function addStaticPageBlobs_(blobs) {
  if (!CONFIG.STATIC_PAGES_FOLDER_ID) return;
  try {
    const folder = DriveApp.getFolderById(CONFIG.STATIC_PAGES_FOLDER_ID);
    const allowed = ["text/html", "image/png", "image/jpeg", "image/svg+xml", "image/webp"];
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName();
      const mime = f.getMimeType();
      if (allowed.indexOf(mime) === -1) continue;
      // Summary backups (<title>_<videoId>.html) may live in the same folder.
      // They are already deployed as <videoId>.html, so skip the duplicates.
      if (/_[A-Za-z0-9_-]{11}\.html$/.test(name)) continue;
      if (mime === "text/html") {
        blobs.push(Utilities.newBlob(f.getBlob().getDataAsString(), "text/html", name));
      } else {
        blobs.push(f.getBlob().setName(name));
      }
      Logger.log("Added static file: " + name);
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

  const htmlById = {};
  htmlById[newVideoId] = newHtml;
  blobs.push(Utilities.newBlob(buildIndexPage(videos, htmlById), "text/html", "index.html"));
  blobs.push(Utilities.newBlob(summaryCss_(), "text/css", "summary.css"));
  blobs.push(Utilities.newBlob(prepareSummaryForDeploy_(newHtml, { title: newTitle, videoId: newVideoId, topic: newTopic }), "text/html", newVideoId + ".html"));

  videos.forEach(function(v) {
    if (v.videoId === newVideoId) return;
    // 1. Try ScriptProperties (videos processed after the fix)
    const storedHtml = PropertiesService.getScriptProperties().getProperty("vid_html_" + v.videoId);
    if (storedHtml) {
      blobs.push(Utilities.newBlob(prepareSummaryForDeploy_(storedHtml, v), "text/html", v.videoId + ".html"));
      return;
    }
    // 2. Search Drive by videoId anywhere in filename — no title dependency
    try {
      const query = 'title contains "' + v.videoId + '" and mimeType = "text/html" and trashed = false';
      const files = DriveApp.searchFiles(query);
      if (files.hasNext()) {
        const html = files.next().getBlob().getDataAsString();
        blobs.push(Utilities.newBlob(prepareSummaryForDeploy_(html, v), "text/html", v.videoId + ".html"));
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

// videos: sheet order (oldest first). htmlById: stored summary pages, used for
// the one-line description on each card.
function buildIndexPage(videos, htmlById) {
  htmlById = htmlById || {};
  const props = PropertiesService.getScriptProperties();
  const items = videos.map(function(v) {
    const html = htmlById[v.videoId] || props.getProperty("vid_html_" + v.videoId) || "";
    const info = summaryCardInfo_(html, v.title);
    return { v: v, info: info, topic: v.topic || "General" };
  }).reverse(); // newest first

  const total = items.length;
  const topicsUsed = [];
  items.forEach(function(it) { if (topicsUsed.indexOf(it.topic) === -1) topicsUsed.push(it.topic); });
  const topicOrder = CONFIG.TOPICS.concat(["General"]).filter(function(t) { return topicsUsed.indexOf(t) !== -1; })
    .concat(topicsUsed.filter(function(t) { return CONFIG.TOPICS.indexOf(t) === -1 && t !== "General"; }));

  // Escapes text and isolates each English run in <bdi>, so mixed
  // Hebrew/English titles keep their word order (e.g. "ה-Workflow שלי ב-Claude Code").
  function esc(s) {
    return stripEmoji_(s).replace(/[A-Za-z][A-Za-z0-9.+#'’ -]*[A-Za-z0-9]|[A-Za-z]|[^A-Za-z]+/g, function(t) {
      return /^[A-Za-z]/.test(t) ? "<bdi>" + escHtml_(t) + "</bdi>" : escHtml_(t);
    });
  }

  // ---- library list ----
  // Hooks the page script and the admin delete rely on — keep them:
  // #toc, #srch, #no-res, .toc-item[data-t], a.video-card[data-vid], .card-title, .tab[data-f]
  // The subscribe form was removed from the page (Nov 2026). doSubscribe() and the backend stay, so it can return.
  let rows = "";
  items.forEach(function(it, i) {
    const v = it.v;
    rows += '<li class="toc-item" data-topic="' + topicKey_(it.topic) + '" data-t="' + escHtml_(it.topic) + '">' +
      '<a href="' + v.videoId + '.html" class="video-card" data-vid="' + v.videoId + '">' +
      '<span class="toc-cat">' + esc(topicLabel_(it.topic)) + (i === 0 ? '<span class="toc-new">חדש</span>' : '') + '</span>' +
      '<h3 class="card-title">' + esc(v.title || v.videoId) + '</h3>' +
      (it.info.dek ? '<p class="toc-dek">' + esc(it.info.dek) + '</p>' : '') +
      '</a></li>';
  });

  let tabs = '<button type="button" class="tab" aria-pressed="true" data-f="all">הכול <span class="cnt">' + total + '</span></button>';
  topicOrder.forEach(function(t) {
    const n = items.filter(function(it) { return it.topic === t; }).length;
    tabs += '<button type="button" class="tab" aria-pressed="false" data-f="' + escHtml_(t) + '" data-topic="' + topicKey_(t) + '">' +
      esc(topicLabel_(t)) + ' <span class="cnt">' + n + '</span></button>';
  });

  // Colour on this page: only the pink brand dot. Everything else is ink and thin rules,
  // in the same type and spacing as the summary pages (summary.css).
  const css =
    ":root{--pink:#EC4899}" +
    ".mast-top{display:flex;justify-content:space-between;gap:12px;padding:24px 0 12px;border-bottom:1px solid var(--ink);font-family:'Space Mono',monospace;font-size:12px;letter-spacing:.5px;color:var(--muted2)}" +
    ".mast-top a{text-decoration:none;color:var(--ink)}" +
    ".hero{display:grid;grid-template-columns:minmax(0,1fr);gap:24px;padding:28px 0 20px;border-bottom:1px solid var(--rule)}" +
    ".hero-txt{display:flex;flex-direction:column;gap:16px}" +
    ".hero h1{margin:0;font-weight:200;font-size:clamp(40px,11vw,64px);line-height:1.05;letter-spacing:-1px}" +
    ".hero h1 span{color:var(--pink)}" +
    ".hero-dek{margin:0;max-width:520px;font-weight:300;font-size:19px;line-height:1.6;color:var(--ink2)}" +
    ".steps{list-style:none;margin:0;padding:0;counter-reset:st;border-bottom:1px solid var(--rule)}" +
    ".steps li{counter-increment:st;display:grid;grid-template-columns:32px minmax(0,1fr);gap:2px 12px;padding:12px 0;border-top:1px solid var(--rule)}" +
    ".steps li::before{content:counter(st,decimal-leading-zero);grid-row:span 2;font-family:'Space Mono',monospace;font-size:12px;color:var(--muted2);padding-top:4px}" +
    ".steps b{font-weight:600;font-size:16px;line-height:1.5}" +
    ".steps span{font-size:14px;line-height:1.55;color:var(--muted)}" +
    ".hero-links{display:flex;margin-top:2px}" +
    ".hero-links a{display:inline-flex;align-items:center;min-height:44px;font-family:'Space Mono',monospace;font-size:12px;letter-spacing:.5px;text-decoration:none}" +
    ".lib{padding:28px 0 8px;display:flex;flex-direction:column;gap:20px}" +
    ".lib-head{display:flex;align-items:flex-end;justify-content:space-between;gap:20px}" +
    ".lib-head h2{margin:0;font-weight:300;font-size:28px;line-height:1.2}" +
    ".srch{flex:0 1 300px;min-width:0}" +
    ".srch input{width:100%;height:44px;padding:0 2px;border:none;border-bottom:1px solid var(--ink);background:transparent;font-family:inherit;font-size:16px;color:var(--ink);border-radius:0}" +
    ".tabs{display:flex;flex-wrap:wrap;gap:0 22px;border-bottom:1px solid var(--rule)}" +
    ".tab{min-height:44px;padding:0;margin-bottom:-1px;background:transparent;border:none;border-bottom:2px solid transparent;font-family:inherit;font-size:15px;color:var(--muted);cursor:pointer;display:inline-flex;align-items:center;gap:6px}" +
    ".tab[aria-pressed=true]{border-bottom-color:var(--ink);color:var(--ink);font-weight:600}" +
    ".cnt{font-family:'Space Mono',monospace;font-size:11px;color:var(--muted2)}" +
    ".toc{list-style:none;margin:0;padding:0;display:grid;grid-template-columns:minmax(0,1fr)}" +
    ".video-card{position:relative;display:flex;flex-direction:column;gap:6px;padding:22px 0;border-bottom:1px solid var(--rule);text-decoration:none;color:var(--ink)}" +
    ".toc-cat{display:flex;align-items:center;gap:10px;font-family:'Space Mono',monospace;font-size:11px;letter-spacing:1px;color:var(--muted2)}" +
    ".toc-new{font-weight:700;color:var(--ink)}" +
    ".toc-new::before{content:'';display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--pink);margin-left:6px;vertical-align:1px}" +
    ".card-title{margin:0;font-weight:500;font-size:19px;line-height:1.45}" +
    ".video-card:hover .card-title{text-decoration:underline;text-decoration-color:var(--pink);text-underline-offset:5px}" +
    ".toc-dek{margin:0;font-size:15px;line-height:1.65;color:var(--muted);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}" +
    ".no-results{display:none;margin:0;padding:32px 0;text-align:center;font-size:17px;color:var(--muted2)}" +
    ".foot{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:12px 16px;margin-top:24px;padding:8px 0 40px;font-size:13px;color:var(--muted2)}" +
    ".foot nav{display:flex;flex-wrap:wrap;gap:0 24px}" +
    ".foot nav a{display:inline-flex;align-items:center;min-height:44px}" +
    ".foot .dot{display:inline-block;width:7px;height:7px;border-radius:50%;background:var(--pink);margin-right:6px}" +
    ".delete-btn{display:none;position:absolute;top:20px;left:0;background:#fee2e2;border:1px solid #fca5a5;color:#991b1b;border-radius:50%;width:26px;height:26px;font-size:.75rem;font-weight:700;cursor:pointer;align-items:center;justify-content:center;z-index:2;padding:0}" +
    ".admin-mode .delete-btn{display:flex}" +
    ".admin-mode .video-card{padding-left:40px}" +
    ".admin-banner{background:#111827;color:#f9a8d4;text-align:center;padding:.5rem 1rem;font-size:.8rem;font-weight:600}" +
    "@media(min-width:880px){" +
    ".hero{gap:40px;padding:64px 0 48px}" +
    ".hero-dek{font-size:21px}.steps b{font-size:17px}.steps span{font-size:15px}.steps li{padding:16px 0}" +
    ".steps{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));column-gap:40px;border-bottom:none}.steps li{border-top:1px solid var(--ink)}" +
    ".lib{padding-top:56px}" +
    ".toc{grid-template-columns:repeat(2,minmax(0,1fr));column-gap:56px}" +
    "}";

  return '<!DOCTYPE html>' +
'<html lang="he" dir="rtl">' +
'<head>' +
'<meta charset="UTF-8">' +
'<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
'<title>ספריית הסיכומים | אילת שחק שול</title>' +
'<meta name="description" content="סיכומים מדויקים של סרטוני YouTube בעברית, רק מה שנאמר, עם צעדי פעולה. נבנה אוטומטית עם Apps Script ו-Claude.">' +
'<link rel="canonical" href="' + siteUrl_() + '/">' +
'<meta property="og:type" content="website"><meta property="og:locale" content="he_IL">' +
'<meta property="og:site_name" content="אילת שחק שול">' +
'<meta property="og:title" content="ספריית הסיכומים | אילת שחק שול">' +
'<meta property="og:description" content="סיכומים מדויקים של סרטוני YouTube בעברית, רק מה שנאמר, עם צעדי פעולה.">' +
'<meta property="og:url" content="' + siteUrl_() + '/">' +
'<meta property="og:image" content="' + siteUrl_() + '/og-library.png">' +
'<meta property="og:image:width" content="1200"><meta property="og:image:height" content="630">' +
'<meta name="twitter:card" content="summary_large_image">' +
FONT_LINK_ +
'<link rel="stylesheet" href="/summary.css">' +
'<style>' + css + '</style>' +
'</head><body>' +
'<div class="wrap">' +
'<header>' +
'<div class="mast-top"><span>אילת שחק שול</span><span>' + total + ' סיכומים</span></div>' +
'<div class="hero">' +
'<div class="hero-txt">' +
'<h1>ספריית הסיכומים<span>.</span></h1>' +
'<p class="hero-dek">כלי שהופך סרטון <bdi>YouTube</bdi> לסיכום קצר בעברית. קוראים בכמה דקות במקום לצפות בכל הסרטון.</p>' +
'</div>' +
'<div>' +
'<ol class="steps" aria-label="איך זה עובד">' +
'<li><b>מוסיפים קישור לסרטון</b><span>זה כל מה שצריך לעשות.</span></li>' +
'<li><b>הסיכום מגיע במייל</b><span>רק מה שנאמר בסרטון, עם צעדים לביצוע.</span></li>' +
'<li><b>ונשמר כאן, בספרייה</b><span>כל הסיכומים במקום אחד, לפי נושא.</span></li>' +
'</ol>' +
'<div class="hero-links"><a href="/youtube-pipeline-casestudy.html">איך זה נבנה ←</a></div>' +
'</div>' +
'</div>' +
'</header>' +
'<main>' +
(total ?
'<section class="lib" aria-labelledby="lib-title">' +
'<div class="lib-head"><h2 id="lib-title">הסיכומים</h2>' +
'<div class="srch"><input id="srch" type="search" placeholder="חיפוש" aria-label="חיפוש סיכום" autocomplete="off"></div></div>' +
'<div class="tabs" role="group" aria-label="סינון לפי נושא">' + tabs + '</div>' +
'<ol class="toc" id="toc" data-mode="all">' + rows + '</ol>' +
'<p class="no-results" id="no-res">לא נמצאו סיכומים לחיפוש הזה.</p>' +
'</section>' : '<p class="no-results" style="display:block">עדיין אין סיכומים</p>') +
'</main>' +
'<footer class="foot"><nav><a href="/youtube-pipeline-casestudy.html">איך זה נבנה</a><a href="/about.html">אודות הפרויקט</a></nav>' +
'<span>נבנה עם <bdi>Claude</bdi> · אילת שחק שול · 2026<span class="dot" aria-hidden="true"></span></span></footer>' +
'</div>' +
'<script>' +
'var WEBAPP_URL="' + CONFIG.WEBAPP_URL + '";' +
'var ADMIN_HASH="' + sha256Hex_(CONFIG.ADMIN_CODE) + '";' + /* SHA-256, one-way — the code itself is NOT recoverable from the page */
'var adminCode=null;' +
'function sha256Hex(s){return crypto.subtle.digest("SHA-256",new TextEncoder().encode(s)).then(function(buf){return Array.prototype.map.call(new Uint8Array(buf),function(b){return("0"+b.toString(16)).slice(-2)}).join("")});}' +
/* Admin: no admin UI in the public page. Open <site>/#admin to log in; the banner and delete buttons are created only after a correct code. */
'function adminLogin(){' +
'var code=prompt("קוד גישה:");if(!code)return;' +
'sha256Hex(code).then(function(h){' +
'if(h!==ADMIN_HASH){alert("קוד שגוי");return;}' +
'adminCode=code;document.body.classList.add("admin-mode");' +
'var bn=document.createElement("div");bn.className="admin-banner";bn.textContent="מצב ניהול: כפתורי מחיקה פעילים";document.body.insertBefore(bn,document.body.firstChild);' +
'document.querySelectorAll(".video-card").forEach(function(card){' +
'var btn=document.createElement("button");btn.className="delete-btn";btn.type="button";btn.title="מחיקה";btn.setAttribute("aria-label","מחיקת סרטון");btn.textContent="✕";' +
'btn.addEventListener("click",function(e){e.preventDefault();e.stopPropagation();deleteCard(card,btn);});' +
'card.insertBefore(btn,card.firstChild);});});}' +
'function deleteCard(card,btn){' +
'var vid=card.getAttribute("data-vid");var titleText=card.querySelector(".card-title").textContent;' +
'if(!confirm("למחוק: "+titleText+"?"))return;' +
'if(!WEBAPP_URL){alert("WEBAPP_URL ריק. צריך להוסיף אותו ב-Script Properties");return;}' +
'btn.textContent="...";btn.disabled=true;' +
/* POST keeps the admin code out of URLs and server logs */
'fetch(WEBAPP_URL,{method:"POST",body:new URLSearchParams({action:"delete",vid:vid,code:adminCode})})' +
'.then(function(r){return r.json();})' +
'.then(function(d){if(d.status!=="ok")throw new Error(d.message||"error");' +
'card.style.opacity="0.35";card.style.pointerEvents="none";card.querySelector(".card-title").textContent="נמחק, הדף מתעדכן...";})' +
'.catch(function(err){btn.textContent="✕";btn.disabled=false;alert("המחיקה נכשלה: "+err.message);});}' +
'if(location.hash==="#admin"){history.replaceState(null,"",location.pathname);setTimeout(adminLogin,300);}' +
/* Subscribe */
'function doSubscribe(){' +
'var email=document.getElementById("sub-email").value.trim();' +
'var msg=document.getElementById("sub-msg");' +
'if(!email||!email.includes("@")){msg.className="sub-msg err";msg.textContent="מייל לא תקין";return;}' +
'if(!WEBAPP_URL){msg.className="sub-msg err";msg.textContent="שירות ההרשמה לא מוגדר עדיין";return;}' +
'msg.className="sub-msg";msg.textContent="שולח...";' +
'fetch(WEBAPP_URL+"?action=subscribe&email="+encodeURIComponent(email))' +
'.then(function(r){return r.json();})' +
'.then(function(d){if(d.status!=="ok")throw new Error();' +
'msg.className="sub-msg";msg.textContent="נרשמת. בדקי את תיבת הדואר לאישור";document.getElementById("sub-email").value="";})' +
'.catch(function(){msg.className="sub-msg err";msg.textContent="ההרשמה לא עברה. אפשר לנסות שוב מאוחר יותר";});}' +
/* Filter tabs + search */
'var toc=document.getElementById("toc"),srch=document.getElementById("srch"),cur="all";' +
'function applyFilter(){if(!toc)return;var q=srch.value.trim().toLowerCase();var vis=0;' +
'toc.setAttribute("data-mode",(cur==="all"&&!q)?"all":"filtered");' +
'toc.querySelectorAll(".toc-item").forEach(function(li){' +
'var ok=(cur==="all"||li.getAttribute("data-t")===cur)&&(!q||li.textContent.toLowerCase().indexOf(q)>-1);' +
'li.style.display=ok?"":"none";if(ok&&!(li.classList.contains("is-featured")&&cur==="all"&&!q))vis++;});' +
'document.getElementById("no-res").style.display=vis?"none":"block";}' +
'document.querySelectorAll(".tab").forEach(function(b){b.addEventListener("click",function(){' +
'cur=b.getAttribute("data-f");document.querySelectorAll(".tab").forEach(function(x){x.setAttribute("aria-pressed",x===b?"true":"false");});applyFilter();});});' +
'if(srch)srch.addEventListener("input",applyFilter);' +
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
// Returns the <article class="summary"> fragment. The page around it
// (head, stylesheet, navigation, footer) is built by buildSummaryPage_.
function callClaude_(prompt, maxTokens) {
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
      max_tokens: maxTokens,
      messages: [{ role: "user", content: prompt }]
    })
  });
  const result = JSON.parse(response.getContentText());
  if (result.error) throw new Error("Claude API: " + result.error.message);
  return result.content[0].text;
}

function extractArticle_(text) {
  const art = String(text || "").match(/<article[\s\S]*<\/article>/i);
  return art ? stripEmoji_(art[0].trim()) : String(text || "").trim();
}

function generateHTMLWithClaude(transcript, title, sourceUrl) {
  return extractArticle_(callClaude_(buildPrompt(transcript, title, sourceUrl), 6000));
}

// ===========================================
// FACT CHECK (second agent)
// ===========================================
// Two agents: the writer (generateHTMLWithClaude) and one checker. The IRON RULE
// in the writer's prompt is an instruction, not a guarantee, so before a summary
// is published the checker compares it with the transcript, fixes or removes
// every statement the transcript does not support, and returns the clean summary
// with a list of what it changed. It sees only the summary and the transcript,
// not the writer's instructions.
function factCheckAndFix_(fragment, transcript) {
  const prompt =
    "You are a strict fact checker. Below is a TRANSCRIPT of a video and a Hebrew SUMMARY of it (an HTML fragment).\n" +
    "Check every statement in the SUMMARY against the TRANSCRIPT and fix only real problems:\n" +
    "- invented: a fact, number, price, name, tool, feature or claim that the transcript does not contain at all -> remove it.\n" +
    "- distorted: the transcript says something materially different (wrong number, wrong tool, reversed meaning, overstated certainty) -> correct it to what the transcript says.\n" +
    "Do NOT change wording, emphasis, framing, grouping or labels that mean the same thing. An action step is fine if it is based on something " +
    "the speaker recommended, suggested, demonstrated or gave as a tip. Never add new content.\n" +
    "If removing a statement leaves an empty list item or paragraph, remove that element. Keep the same HTML structure, classes and attributes.\n\n" +
    "Output exactly two blocks and nothing else:\n" +
    "1. The full corrected <article>...</article> fragment (unchanged if nothing was wrong).\n" +
    "2. <changes>one line per change: 'removed: <text> | <reason in English>' or 'fixed: <old> -> <new> | <reason>'. Empty if nothing changed.</changes>\n\n" +
    "SUMMARY:\n" + fragment + "\n\n" +
    "TRANSCRIPT:\n" + String(transcript || "").substring(0, MAX_TRANSCRIPT_CHARS_);
  const text = callClaude_(prompt, 6000);
  const fixed = extractArticle_(text.replace(/<changes>[\s\S]*?<\/changes>/i, ""));
  const ch = text.match(/<changes>([\s\S]*?)<\/changes>/i);
  const changes = ch ? ch[1].split("\n").map(function(l) { return l.trim(); }).filter(function(l) { return l; }) : [];
  return { fragment: fixed, changes: changes };
}

// ===========================================
// QUALITY GATE
// ===========================================
// A summary page goes public (Netlify) and lands in inboxes the moment it is
// generated, so it must be checked BEFORE deploy, not after. The gate verifies
// the structure and the length the prompt demands; the fact check below verifies
// that every statement comes from the transcript. Claude output varies between calls, so one retry is worth
// it; a second failure marks the row Error and alerts the owner instead of
// publishing a broken page.
const MAX_SUMMARY_WORDS_ = 520; // prompt asks for ~350; this is the hard ceiling

function countWords_(html) {
  const text = String(html || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  return text ? text.split(" ").length : 0;
}

function validateSummaryHtml_(fragment, transcript) {
  const problems = [];
  if (!fragment || !/^<article[^>]*class=["']summary["']/i.test(fragment)) {
    problems.push("no <article class=\"summary\"> fragment");
    return problems;
  }
  if (!/<\/article>\s*$/i.test(fragment)) problems.push("truncated - no closing </article>");
  if (fragment.indexOf("```") !== -1) problems.push("markdown fences left in output");
  if (/<(style|script|html|head|body)\b/i.test(fragment)) problems.push("page-level tags inside the fragment");
  ["s-tldr", "s-part", "s-actions"].forEach(function(cls) {
    if (fragment.indexOf(cls) === -1) problems.push("missing section: ." + cls);
  });
  const bad = tagNestingError_(fragment);
  if (bad) problems.push("broken HTML structure: " + bad);
  const h1 = extractHebrewTitle(fragment);
  if (!h1 || !/[א-ת]/.test(h1)) problems.push("h1 missing or not Hebrew");
  const words = countWords_(fragment);
  if (words > MAX_SUMMARY_WORDS_) problems.push("too long: " + words + " words (max " + MAX_SUMMARY_WORDS_ + ")");
  const actions = (fragment.match(/<section class=["']s-actions["'][\s\S]*?<\/section>/i) || [""])[0];
  const nActions = (actions.match(/<li\b/gi) || []).length;
  if (nActions > 5) problems.push("too many action steps: " + nActions + " (max 5)");
  return problems;
}

// Checks that block tags open and close in the right order. A swapped pair
// (e.g. </section></div>) makes the browser end the article early and the rest
// of the page falls outside the layout. Returns "" when the nesting is valid.
function tagNestingError_(html) {
  const tracked = /^(article|header|section|div|ol|ul|li|dl|dt|dd|p|h1|h2|h3|figure|blockquote|figcaption)$/;
  const stack = [];
  const re = /<(\/?)([a-z0-9]+)\b[^>]*?(\/?)>/gi;
  let m;
  while ((m = re.exec(String(html || "")))) {
    const tag = m[2].toLowerCase();
    if (!tracked.test(tag) || m[3]) continue;
    if (!m[1]) { stack.push(tag); continue; }
    if (!stack.length) return "</" + tag + "> without an opening tag";
    const open = stack.pop();
    if (open !== tag) return "<" + open + "> closed by </" + tag + ">";
  }
  return stack.length ? "<" + stack[stack.length - 1] + "> never closed" : "";
}

// The key figure on the index card must come from the video, not from Claude.
// A number that is not in the transcript is dropped (not a reason to fail).
function checkFigure_(fragment, transcript) {
  const m = fragment.match(/data-figure=["']([^"']*)["']/i);
  if (!m) return fragment;
  const fig = m[1].trim();
  const digits = fig.replace(/[^0-9]/g, "");
  const inTranscript = digits ? String(transcript || "").indexOf(digits) !== -1 : fig.length > 0;
  if (inTranscript && fig.length <= 8) return fragment;
  return fragment.replace(/\sdata-figure(-label)?=["'][^"']*["']/gi, "");
}

function generateValidatedHtml_(transcript, title, sourceUrl, topic) {
  let problems = [];
  for (let attempt = 1; attempt <= 2; attempt++) {
    let fragment = generateHTMLWithClaude(transcript, title, sourceUrl);
    problems = validateSummaryHtml_(fragment, transcript);
    if (problems.length) {
      Logger.log("Quality gate attempt " + attempt + " failed: " + problems.join("; "));
      continue;
    }
    // Second agent: check against the transcript and fix.
    const checked = factCheckAndFix_(fragment, transcript);
    Logger.log("Fact check: " + (checked.changes.length ? checked.changes.length + " change(s): " + checked.changes.join(" || ") : "no changes"));
    problems = validateSummaryHtml_(checked.fragment, transcript);
    // The checker may only remove or correct, never add.
    if (!problems.length && countWords_(checked.fragment) > countWords_(fragment) + 20) problems.push("checker added content");
    if (!problems.length) {
      return buildSummaryPage_(checkFigure_(checked.fragment, transcript), { topic: topic, title: title, sourceUrl: sourceUrl });
    }
    Logger.log("Attempt " + attempt + " failed: " + problems.join("; "));
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

// The whole transcript goes to Claude. The old 12,000-character cut kept only
// the first ~12 minutes of a video, so the end of longer videos was never summarized.
const MAX_TRANSCRIPT_CHARS_ = 150000;

function buildPrompt(transcript, title, sourceUrl) {
  return "You write a SHORT Hebrew summary of a YouTube video, as an HTML fragment.\n\n" +
    "IRON RULE: ONLY content from the transcript. Zero inventions. Zero additions.\n" +
    "If something is NOT explicitly said in the transcript, do NOT write it. Not as an example, not as a tip, not as an explanation. " +
    "No dates, years or speaker names unless they are said in the transcript. Every sentence must be traceable to the transcript.\n\n" +
    "LENGTH IS THE PRODUCT. The reader wants the video in 2 minutes of reading.\n" +
    "- Total: about 300-380 Hebrew words. Never more than 450.\n" +
    "- Choose. Leave out repetitions, intros, anecdotes, sponsor reads and anything that is not a takeaway.\n" +
    "- Each paragraph: at most 2 short sentences.\n" +
    "- Bold (<strong>) at most ONE phrase per paragraph or list item: the single thing to remember. Never bold whole sentences.\n" +
    "- Plain, direct Hebrew. No filler words, no marketing tone, no emoji, no exclamation marks.\n\n" +
    "Title: " + title + "\n" +
    "Source: " + sourceUrl + "\n\n" +
    "Return ONLY this fragment, nothing before or after it, no <html>/<head>/<style>/<script>:\n\n" +
    "<article class=\"summary\" data-figure=\"F\" data-figure-label=\"L\">\n" +
    "  <header>\n" +
    "    <h1 dir=\"rtl\">Hebrew title</h1>\n" +
    "    <p class=\"s-dek\">One sentence: what the video gives the reader. Max 20 words.</p>\n" +
    "  </header>\n" +
    "  <section class=\"s-tldr\"><p>The core message in 1-2 sentences, max 35 words, with ONE <strong>.</p></section>\n" +
    "  PARTS (1 to 3 of them, see below)\n" +
    "  QUOTE (optional, at most one)\n" +
    "  <section class=\"s-actions\"><h2>מה עושים עם זה</h2><ol>\n" +
    "    <li><strong>Verb phrase, 2-4 words</strong> rest of the step.</li>  (3 to 5 steps, max 15 words each, only steps the video supports)\n" +
    "  </ol></section>\n" +
    "  <div class=\"s-meta\"><p class=\"s-tools\">הכלים בסרטון: Tool · Tool</p><p class=\"s-source\">מקור: <a href=\"SOURCE_URL\" dir=\"ltr\">Original English video title</a></p></div>\n" +
    "</article>\n\n" +
    "PARTS. Build the parts from what the video ACTUALLY contains. Do not force a structure: a video about 5 agents gets one part listing 5 agents; " +
    "a video with one argument may need a single part. Each part is:\n" +
    "<section class=\"s-part\"><div>\n" +
    "  <h2>Short Hebrew heading, max 6 words</h2>\n" +
    "  <p class=\"s-lead\">Optional: one sentence, max 20 words.</p>\n" +
    "  then EITHER 1-3 short paragraphs <p>...</p>\n" +
    "  OR a list of the video's items (tools, methods, examples, steps):\n" +
    "  <ol class=\"s-items\">\n" +
    "    <li><div><h3>Item name</h3><p>What it is or does, max 25 words.</p></div>\n" +
    "      <dl class=\"s-facts\"><div><dt>מתאים ל</dt><dd>...</dd></div><div><dt>עלות</dt><dd>...</dd></div></dl></li>\n" +
    "  </ol>\n" +
    "  Use <dl class=\"s-facts\"> ONLY for short facts the video states about that item (who it fits, cost, time, difficulty). No facts in the video, no <dl>.\n" +
    "</div></section>\n\n" +
    "QUOTE (optional): the single most memorable line of the video, translated to Hebrew:\n" +
    "<figure class=\"s-quote\"><blockquote>Hebrew translation, max 20 words</blockquote><figcaption dir=\"ltr\">Original English words</figcaption></figure>\n\n" +
    "data-figure: one number that the video itself uses and that captures it (e.g. 5 for five agents, 30 for thirty days, 90%). " +
    "data-figure-label: 1-2 Hebrew words for what the number counts (e.g. סוכנים, יום). If no number fits, omit both attributes.\n" +
    "s-tools: only tools or products named in the video. If none, omit that <p>.\n\n" +
    "Mixed Hebrew/English: wrap every English word or phrase in <bdi>, e.g. <h1 dir=\"rtl\">שליטה בסוכני AI של <bdi>Claude</bdi></h1>.\n" +
    "Address the reader in plural (אתם / בחרו / חברו).\n\n" +
    "TRANSCRIPT:\n" + transcript.substring(0, MAX_TRANSCRIPT_CHARS_);
}

// ===========================================
// SITE STYLE (shared by every summary page)
// ===========================================
// One stylesheet, deployed as /summary.css on every deploy. Summary pages link
// to it instead of carrying their own <style>, so a design change applies to
// every page on the next deploy. It styles both the current page format
// (.summary) and the legacy format (.hero/.level...) of pages generated
// before the short format, so old pages look right until they are regenerated.
const FONT_LINK_ =
  '<link rel="preconnect" href="https://fonts.googleapis.com">' +
  '<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Hebrew:wght@200;300;400;500;600&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">';

// Topic colors. Key = sheet value of the topic column.
const TOPIC_COLORS_ = {
  "AI Tools":         { c: "#BE185D", t: "#FDF2F8" },
  "Productivity":     { c: "#0F766E", t: "#F0FDFA" },
  "Entrepreneurship": { c: "#1D4ED8", t: "#EFF6FF" },
  "General":          { c: "#1B2A4A", t: "#F1F2F5" }
};
function topicColor_(t) { return TOPIC_COLORS_[t] || TOPIC_COLORS_["General"]; }
// Topic key safe for an HTML attribute / CSS selector
function topicKey_(t) { return TOPIC_COLORS_[t] ? t.replace(/\s+/g, "-").toLowerCase() : "general"; }

function topicCss_() {
  return Object.keys(TOPIC_COLORS_).map(function(t) {
    return '[data-topic="' + topicKey_(t) + '"]{--accent:' + TOPIC_COLORS_[t].c + ';--tint:' + TOPIC_COLORS_[t].t + '}';
  }).join("");
}

function summaryCss_() {
  return "" +
  ":root{--ink:#1B2A4A;--ink2:#3A4763;--muted:#4A5468;--muted2:#5B6478;--rule:#D9D6CF;--paper:#FBFAF8;--accent:#BE185D;--tint:#FDF2F8}" +
  topicCss_() +
  "*{box-sizing:border-box}" +
  "html{-webkit-text-size-adjust:100%}" +
  "body{margin:0;background:var(--paper);color:var(--ink);font-family:'IBM Plex Sans Hebrew',Arial,sans-serif;line-height:1.7;font-size:17px}" +
  "a{color:var(--ink)}" +
  "a:focus-visible,button:focus-visible,input:focus-visible{outline:2px solid var(--accent);outline-offset:3px}" +
  ".mono{font-family:'Space Mono',monospace}" +
  ".wrap{max-width:1040px;margin:0 auto;padding:0 24px}" +
  ".s-nav{display:flex;flex-wrap:wrap;justify-content:space-between;gap:8px;padding:28px 0 14px;border-bottom:1px solid var(--ink);font-family:'Space Mono',monospace;font-size:12px;letter-spacing:.5px}" +
  ".s-nav a{text-decoration:none}.s-nav span{color:var(--muted2)}" +
  ".s-kicker{display:inline-flex;align-items:center;gap:8px;font-family:'Space Mono',monospace;font-size:12px;font-weight:700;letter-spacing:1px;color:var(--accent)}" +
  ".s-kicker::before{content:'';width:8px;height:8px;border-radius:50%;background:var(--accent)}" +
  ".s-double{border-top:1px solid var(--ink);border-bottom:1px solid var(--ink);height:3px}" +
  "bdi{unicode-bidi:isolate}" +
  // ---- current format ----
  ".summary{counter-reset:part}" +
  ".summary>header{padding:56px 0 40px;display:flex;flex-direction:column;gap:18px;max-width:820px}" +
  ".summary h1{margin:0;font-weight:300;font-size:clamp(32px,6vw,56px);line-height:1.15;letter-spacing:-1px}" +
  ".s-dek{margin:0;font-weight:300;font-size:21px;line-height:1.6;color:var(--ink2)}" +
  ".s-yt{align-self:flex-start;display:inline-flex;align-items:center;min-height:44px;font-size:15px;font-weight:600;text-decoration:underline;text-decoration-color:var(--accent);text-decoration-thickness:2px;text-underline-offset:6px}" +
  ".s-tldr,.s-part,.s-actions{display:grid;grid-template-columns:160px minmax(0,1fr);gap:16px 48px;padding:52px 0;border-bottom:1px solid var(--rule)}" +
  ".s-tldr::before{content:'בשורה אחת';font-family:'Space Mono',monospace;font-size:12px;font-weight:700;letter-spacing:1px;color:var(--accent);padding-top:10px}" +
  ".s-tldr p{margin:0;font-weight:300;font-size:clamp(22px,2.8vw,30px);line-height:1.5}" +
  ".s-tldr strong{font-weight:600}" +
  ".s-part{counter-increment:part}" +
  ".s-part::before{content:counter(part,decimal-leading-zero);font-weight:200;font-size:88px;line-height:.85;color:var(--accent)}" +
  ".s-part>div{min-width:0;display:flex;flex-direction:column;gap:20px}" +
  ".s-part h2{margin:0;font-weight:300;font-size:32px;line-height:1.25}" +
  ".s-lead{margin:0;font-size:18px;line-height:1.7;color:var(--muted)}" +
  ".s-part p{margin:0;font-size:17px;line-height:1.75;color:var(--ink2)}" +
  ".s-part strong,.s-actions strong{font-weight:600;color:var(--ink)}" +
  ".s-items{list-style:none;margin:0;padding:0;border-bottom:1px solid var(--rule);counter-reset:item}" +
  ".s-items>li{counter-increment:item;display:grid;grid-template-columns:44px minmax(0,1fr) auto;gap:6px 24px;padding:22px 0;border-top:1px solid var(--rule)}" +
  ".s-items>li::before{content:counter(item,decimal-leading-zero);font-family:'Space Mono',monospace;font-size:12px;color:var(--accent);padding-top:6px}" +
  ".s-items h3{margin:0 0 4px;font-weight:600;font-size:19px;line-height:1.4;color:var(--ink)}" +
  ".s-items li>div>p{font-size:16px}" +
  ".s-facts{margin:0;display:flex;flex-direction:column;gap:8px;min-width:150px;max-width:240px}" +
  ".s-facts dt{font-family:'Space Mono',monospace;font-size:11px;letter-spacing:1px;color:var(--muted2)}" +
  ".s-facts dd{margin:2px 0 0;font-size:15px;line-height:1.5}" +
  ".s-quote{margin:0;padding:60px 0;border-bottom:1px solid var(--rule);display:flex;flex-direction:column;align-items:center;gap:14px;text-align:center}" +
  ".s-quote blockquote{margin:0;max-width:780px;font-weight:200;font-size:clamp(26px,3.8vw,40px);line-height:1.4}" +
  ".s-quote figcaption{max-width:640px;font-size:13px;color:var(--muted2)}" +
  ".s-actions{border-bottom:none}" +
  ".s-actions h2{margin:0;padding-top:8px;font-family:'Space Mono',monospace;font-size:12px;font-weight:700;letter-spacing:1px;color:var(--accent)}" +
  ".s-actions ol{list-style:none;margin:0;padding:24px 28px;background:var(--tint);counter-reset:act;display:flex;flex-direction:column;gap:14px}" +
  ".s-actions li{counter-increment:act;position:relative;padding-right:40px;font-size:18px;line-height:1.6}" +
  ".s-actions li::before{content:counter(act,decimal-leading-zero);position:absolute;right:0;top:6px;font-family:'Space Mono',monospace;font-size:12px;line-height:1;color:var(--accent)}" +
  ".s-tools,.s-source{margin:0;font-family:'Space Mono',monospace;font-size:12px;color:var(--muted2)}" +
  ".s-meta{display:flex;flex-wrap:wrap;justify-content:space-between;gap:12px;padding:22px 0;border-top:1px solid var(--rule)}" +
  ".s-foot{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:16px;padding:28px 0 48px;border-top:1px solid var(--rule);font-size:14px;color:var(--muted2)}" +
  ".s-foot .dot{display:inline-block;width:7px;height:7px;border-radius:50%;background:#EC4899;margin-right:6px}" +
  "@media(max-width:720px){" +
  ".s-tldr,.s-part,.s-actions{grid-template-columns:minmax(0,1fr);padding:40px 0}" +
  ".s-part::before{font-size:64px}" +
  ".s-items>li{grid-template-columns:32px minmax(0,1fr)}" +
  ".s-facts{grid-column:2;max-width:none}" +
  ".s-actions ol{padding:20px}" +
  ".s-actions li{font-size:17px;padding-right:34px}" +
  "}" +
  // ---- legacy format (pages generated before the short format) ----
  ".legacy .container{max-width:860px;margin:0 auto;padding:0 24px}" +
  ".legacy a.back{display:block;padding:28px 0 14px;border-bottom:1px solid var(--ink);font-family:'Space Mono',monospace;font-size:12px;text-decoration:none}" +
  ".legacy .hero{padding:52px 0 36px;border-bottom:4px double var(--ink);display:flex;flex-direction:column;gap:16px}" +
  ".legacy .hero .tag{display:none}" +
  ".legacy .hero h1{margin:0;font-weight:300;font-size:clamp(30px,5.5vw,52px);line-height:1.15;letter-spacing:-1px}" +
  ".legacy .subtitle{margin:0;font-weight:300;font-size:20px;line-height:1.6;color:var(--ink2)}" +
  ".legacy a.yt{align-self:flex-start;font-size:15px;font-weight:600;text-decoration:underline;text-decoration-color:var(--accent);text-decoration-thickness:2px;text-underline-offset:6px}" +
  ".legacy .central-idea{padding:44px 0;border-bottom:1px solid var(--rule);font-weight:300;font-size:clamp(20px,2.4vw,25px);line-height:1.6}" +
  ".legacy .central-idea strong{font-weight:600;color:var(--ink)}" +
  ".legacy .central-idea strong:first-child{display:block;font-family:'Space Mono',monospace;font-size:12px;letter-spacing:1px;color:var(--accent);margin-bottom:8px}" +
  ".legacy .anim-ready{opacity:1}" +
  ".legacy .level{padding:48px 0;border-bottom:1px solid var(--rule);display:flex;flex-direction:column;gap:18px}" +
  ".legacy .level-header{display:flex;flex-direction:column;gap:6px}" +
  ".legacy .level-badge{font-family:'Space Mono',monospace;font-size:12px;font-weight:700;letter-spacing:1px;color:var(--accent)}" +
  ".legacy .level h2{margin:0;font-weight:300;font-size:30px;line-height:1.25}" +
  ".legacy .tagline{margin:0;font-size:17px;color:var(--muted)}" +
  ".legacy .tool-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,240px),1fr));column-gap:40px}" +
  ".legacy .tool-card{padding:16px 0;border-top:1px solid var(--rule)}" +
  ".legacy .tool-name{font-weight:600;font-size:16px;margin-bottom:4px}" +
  ".legacy .tool-card p,.legacy .way p,.legacy .example-step p,.legacy .opt-desc{margin:0;font-size:15px;line-height:1.7;color:var(--muted)}" +
  ".legacy .context-ways{display:flex;flex-direction:column}" +
  ".legacy .way{display:flex;gap:14px;padding:18px 0;border-top:1px solid var(--rule)}" +
  ".legacy .way-num{flex:0 0 22px;font-family:'Space Mono',monospace;font-size:12px;color:var(--accent);padding-top:3px}" +
  ".legacy .definition-box{padding:24px 0;border-top:1px solid var(--ink);border-bottom:1px solid var(--ink)}" +
  ".legacy .def-label,.legacy .example-title,.legacy .section-label,.legacy .actions-title{display:block;font-family:'Space Mono',monospace;font-size:12px;font-weight:700;letter-spacing:1px;color:var(--accent);margin-bottom:10px}" +
  ".legacy .definition-box p{margin:0;font-weight:300;font-size:22px;line-height:1.5}" +
  ".legacy .def-sub{font-size:14px;color:var(--muted2);margin-top:8px}" +
  ".legacy .example-steps{display:flex;flex-direction:column;gap:12px;border-right:1px solid var(--rule);padding-right:18px}" +
  ".legacy .example-step{position:relative}" +
  ".legacy .example-step::before{content:'';position:absolute;right:-23px;top:10px;width:9px;height:9px;border-radius:50%;background:var(--accent)}" +
  ".legacy .step-icon{display:none}" +
  ".legacy .system-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,200px),1fr));gap:0 28px}" +
  ".legacy .system-option{padding:16px 0;border-top:1px solid var(--ink);display:flex;flex-direction:column;gap:4px}" +
  ".legacy .opt-name{font-weight:600}" +
  ".legacy .difficulty{font-family:'Space Mono',monospace;font-size:12px;color:var(--muted2)}" +
  ".legacy .actions-box{background:var(--tint);padding:24px 28px}" +
  ".legacy .actions-box ol{margin:0;padding-right:20px;display:flex;flex-direction:column;gap:10px}" +
  ".legacy .actions-box li{font-size:16px;line-height:1.65}" +
  ".legacy .actions-box li::marker{color:var(--accent);font-family:'Space Mono',monospace;font-size:13px}" +
  ".legacy strong{font-weight:600}" +
  ".legacy .final-message{padding:56px 0;text-align:center;border-bottom:1px solid var(--rule)}" +
  ".legacy .final-message p{margin:0 auto 14px;max-width:640px;font-size:16px;color:var(--muted)}" +
  ".legacy .final-message p:first-child{font-weight:300;font-size:clamp(21px,2.6vw,27px);color:var(--ink);line-height:1.5}" +
  ".legacy .source{font-family:'Space Mono',monospace;font-size:12px;color:var(--muted2)}" +
  ".legacy .page-footer{padding:28px 0 48px;text-align:center}" +
  ".legacy .page-footer .credit{font-size:13px;color:var(--muted2)}";
}

// Strip emoji (and the joiners/variation selectors that build them).
function stripEmoji_(s) {
  return String(s || "").replace(/[\p{Extended_Pictographic}\uFE0F\u200D\u20E3]/gu, "").replace(/ {2,}/g, " ");
}

// Wraps the article fragment Claude returns into a full page.
function buildSummaryPage_(fragment, info) {
  const topic = info.topic || "General";
  const label = escHtml_(topicLabel_(topic));
  const title = extractHebrewTitle(fragment) || info.title || "סיכום";
  // YouTube link goes right after the article header
  const yt = '<a class="s-yt" href="' + escHtml_(info.sourceUrl) + '">לסרטון המקורי ↗</a>';
  let body = fragment.replace(/<\/header>/i, yt + "</header>");
  body = body.replace(/<header>/i, '<header><span class="s-kicker">' + label + '</span>');
  return '<!DOCTYPE html><html lang="he" dir="rtl"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<title>' + escHtml_(title) + '</title>' + FONT_LINK_ +
    '<link rel="stylesheet" href="/summary.css">' +
    '</head><body data-topic="' + topicKey_(topic) + '"><div class="wrap">' +
    '<nav class="s-nav"><a href="/">→ ספריית הסיכומים</a><span>אילת שחק שול</span></nav>' +
    body +
    '<footer class="s-foot"><a href="/">→ לכל הסיכומים</a><span>נבנה עם Claude · אילת שחק שול · 2026<span class="dot" aria-hidden="true"></span></span></footer>' +
    '</div></body></html>';
}

// Every summary page passes through here on every deploy:
// share tags, shared stylesheet, topic color, no emoji.
function prepareSummaryForDeploy_(html, v) {
  let out = stripEmoji_(html);
  const key = topicKey_(v.topic || "General");
  if (!/href=["']\/summary\.css["']/i.test(out)) {
    // Legacy page: drop its own CSS, link the shared one, add a topic kicker.
    out = out.replace(/<style[\s\S]*?<\/style>/gi, "");
    out = out.replace(/<\/head>/i, FONT_LINK_ + '<link rel="stylesheet" href="/summary.css"></head>');
    out = out.replace(/<body([^>]*)>/i, function(m, attrs) {
      return '<body' + attrs.replace(/\s(class|data-topic)=["'][^"']*["']/gi, "") + ' class="legacy" data-topic="' + key + '">';
    });
    out = out.replace(/(<div class=["']hero["'][^>]*>)/i, '$1<span class="s-kicker">' + escHtml_(topicLabel_(v.topic || "General")) + '</span>');
  } else {
    out = out.replace(/<body([^>]*)>/i, function(m, attrs) {
      return '<body' + attrs.replace(/\sdata-topic=["'][^"']*["']/gi, "") + ' data-topic="' + key + '">';
    });
  }
  return addShareMeta_(out, v.title, v.videoId);
}

// Card details for the index page, read from the stored summary page.
function summaryCardInfo_(html, title) {
  const info = { dek: "", figure: "", figureLabel: "", quote: "" };
  if (html) {
    const art = html.match(/<article[^>]*class=["']summary["'][^>]*>/i);
    if (art) {
      const f = art[0].match(/data-figure=["']([^"']{1,8})["']/i);
      const fl = art[0].match(/data-figure-label=["']([^"']{1,24})["']/i);
      if (f) info.figure = f[1].trim();
      if (fl) info.figureLabel = fl[1].trim();
    }
    const dek = html.match(/class=["'](?:s-dek|subtitle)["'][^>]*>([\s\S]*?)<\/p>/i);
    if (dek) info.dek = dek[1].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
    const q = html.match(/<figure class=["']s-quote["'][^>]*>\s*<blockquote[^>]*>([\s\S]*?)<\/blockquote>/i);
    if (q) info.quote = q[1].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
  }
  if (!info.figure) {
    // Legacy pages: take the first number in the title and the word after it.
    // Years (2026) are not a key figure, so they are skipped.
    const re = /(\d+%?)(?:\s+([א-ת]{2,}))?/g;
    let m;
    while ((m = re.exec(String(title || "")))) {
      if (/^(19|20)\d\d$/.test(m[1])) continue;
      info.figure = m[1]; info.figureLabel = m[2] || "";
      break;
    }
  }
  info.dek = stripEmoji_(info.dek);
  return info;
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

    if (action === "subscribe" && email) {
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

// Deletion arrives by POST only, so the admin code never appears in a URL.
function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    var p = (e && e.parameter) || {};
    var vid  = (p.vid || "").toString().trim();
    var code = (p.code || "").toString();
    if (p.action !== "delete" || !vid) {
      output.setContent(JSON.stringify({status:"error",message:"Unknown action"}));
      return output;
    }
    // Server-side auth: ADMIN_CODE must be set AND match the code from the request.
    if (!CONFIG.ADMIN_CODE || code !== CONFIG.ADMIN_CODE) {
      output.setContent(JSON.stringify({status:"error",message:"Unauthorized"}));
      return output;
    }
    if (!/^[A-Za-z0-9_-]{11}$/.test(vid)) {
      output.setContent(JSON.stringify({status:"error",message:"Bad video id"}));
      return output;
    }
    deleteVideoFromLibrary(vid);
    output.setContent(JSON.stringify({status:"ok",deleted:vid}));
  } catch (err) {
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
  const htmlById = {};
  videos.forEach(function(v) { htmlById[v.videoId] = props.getProperty("vid_html_" + v.videoId) || ""; });
  const blobs = [
    Utilities.newBlob(buildIndexPage(videos, htmlById), "text/html", "index.html"),
    Utilities.newBlob(summaryCss_(), "text/css", "summary.css")
  ];

  videos.forEach(function(v) {
    const html = htmlById[v.videoId];
    if (html) blobs.push(Utilities.newBlob(prepareSummaryForDeploy_(html, v), "text/html", v.videoId + ".html"));
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
// REGENERATE EXISTING SUMMARIES (short format)
// ===========================================
// Rewrites every published summary with the current prompt. Sends NO emails:
// not to subscribers and not to the owner per video. Runs in small batches
// (one per run: Apps Script stops a run after 6 minutes), driven by its own time trigger,
// and redeploys the site once at the end.
//
// Run startRegeneration() once from the editor. Progress is in the log;
// regenerationStatus() shows what is left; stopRegeneration() cancels.
// The site is redeployed after every batch. A video that fails keeps its old
// page and is listed in a summary email to the owner at the end. The previous
// page of each video is kept in Drive as legacy-backup_<videoId>.html.txt.
const REGEN_QUEUE_KEY_ = "regen_queue";
const REGEN_FAILED_KEY_ = "regen_failed";
const REGEN_BATCH_ = 1; // one video per run: generate + fact check can take several minutes

function startRegeneration() {
  assertConfig_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getRange(CONFIG.START_ROW, 1, sheet.getLastRow() - CONFIG.START_ROW + 1, 6).getValues();
  const queue = [];
  data.forEach(function(row) {
    const url = (row[CONFIG.COL_URL - 1] || "").toString().trim();
    const status = (row[CONFIG.COL_STATUS - 1] || "").toString().trim();
    if (url && status.indexOf("Done") !== -1 && extractVideoId(url)) queue.push(extractVideoId(url));
  });
  const props = PropertiesService.getScriptProperties();
  props.setProperty(REGEN_QUEUE_KEY_, JSON.stringify(queue));
  props.setProperty(REGEN_FAILED_KEY_, "[]");
  stopRegenerationTrigger_();
  ScriptApp.newTrigger("regenerateBatch").timeBased().everyMinutes(5).create();
  Logger.log("Regeneration queued: " + queue.length + " summaries. First batch runs now.");
  regenerateBatch();
}

// Finds published summaries whose HTML structure is broken and rewrites them.
function regenerateMalformed() {
  assertConfig_();
  const props = PropertiesService.getScriptProperties();
  const rows = sheetRowsById_();
  const queue = Object.keys(rows).filter(function(vid) {
    const html = props.getProperty("vid_html_" + vid) || "";
    const art = html.match(/<article[\s\S]*<\/article>/i);
    return art && tagNestingError_(art[0]);
  });
  if (!queue.length) { Logger.log("No broken pages."); return; }
  props.setProperty(REGEN_QUEUE_KEY_, JSON.stringify(queue));
  props.setProperty(REGEN_FAILED_KEY_, "[]");
  stopRegenerationTrigger_();
  ScriptApp.newTrigger("regenerateBatch").timeBased().everyMinutes(5).create();
  Logger.log("Rewriting broken pages: " + queue.join(", "));
  regenerateBatch();
}

// Re-queue only the videos that failed in the last regeneration.
function retryFailedRegeneration() {
  assertConfig_();
  const props = PropertiesService.getScriptProperties();
  const failed = JSON.parse(props.getProperty(REGEN_FAILED_KEY_) || "[]");
  const queue = failed.map(function(f) { return String(f).split(":")[0].trim(); })
    .filter(function(v, i, arr) { return /^[A-Za-z0-9_-]{11}$/.test(v) && arr.indexOf(v) === i; });
  if (!queue.length) { Logger.log("Nothing to retry."); return; }
  props.setProperty(REGEN_QUEUE_KEY_, JSON.stringify(queue));
  props.setProperty(REGEN_FAILED_KEY_, "[]");
  stopRegenerationTrigger_();
  ScriptApp.newTrigger("regenerateBatch").timeBased().everyMinutes(5).create();
  Logger.log("Retrying " + queue.length + ": " + queue.join(", "));
  regenerateBatch();
}

function regenerationStatus() {
  const props = PropertiesService.getScriptProperties();
  Logger.log("Left: " + (props.getProperty(REGEN_QUEUE_KEY_) || "[]"));
  Logger.log("Failed: " + (props.getProperty(REGEN_FAILED_KEY_) || "[]"));
}

// Cancel a running regeneration. Pages already rewritten stay rewritten.
function stopRegeneration() {
  stopRegenerationTrigger_();
  PropertiesService.getScriptProperties().deleteProperty(REGEN_QUEUE_KEY_);
  Logger.log("Regeneration stopped.");
}

// Keeps the previous page before it is overwritten. Saved as .txt so the
// static-pages deploy (html/images only) never publishes it.
function backupLegacyHtml_(vid, html) {
  if (!html) return;
  try {
    const folder = CONFIG.DRIVE_FOLDER_ID ? DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID) : DriveApp.getRootFolder();
    const name = "legacy-backup_" + vid + ".html.txt";
    if (folder.getFilesByName(name).hasNext()) return; // keep the first (original) backup
    folder.createFile(Utilities.newBlob(html, "text/plain", name));
  } catch (e) { Logger.log("Legacy backup failed for " + vid + ": " + e.message); }
}

function stopRegenerationTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "regenerateBatch") ScriptApp.deleteTrigger(t);
  });
}

function regenerateBatch() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) { Logger.log("Busy (another run holds the lock) — next trigger will continue."); return; }
  try {
    const props = PropertiesService.getScriptProperties();
    const queue = JSON.parse(props.getProperty(REGEN_QUEUE_KEY_) || "[]");
    const failed = JSON.parse(props.getProperty(REGEN_FAILED_KEY_) || "[]");
    if (!queue.length) { finishRegeneration_(failed); return; }

    const rows = sheetRowsById_();
    const started = Date.now();
    let done = 0;
    while (queue.length && done < REGEN_BATCH_ && Date.now() - started < 3 * 60 * 1000) {
      const vid = queue.shift();
      const r = rows[vid];
      try {
        if (!r) throw new Error("row not found in sheet");
        const transcript = getYouTubeTranscript(r.url);
        if (!transcript) throw new Error("No transcript found");
        const html = generateValidatedHtml_(transcript, r.title, r.url, r.topic);
        backupLegacyHtml_(vid, props.getProperty("vid_html_" + vid));
        props.setProperty("vid_html_" + vid, html);
        props.setProperty("vid_title_he_" + vid, extractHebrewTitle(html) || r.title);
        try { saveToDrive(html, extractHebrewTitle(html) || r.title, vid); } catch (e) { Logger.log("Drive backup failed: " + e.message); }
        Logger.log("Regenerated: " + vid);
      } catch (err) {
        failed.push(vid + ": " + err.message);
        Logger.log("Regeneration failed for " + vid + ": " + err.message);
      }
      done++;
      // Re-read before writing: startRegeneration() may have reset the queue
      // while this video was being processed. Remove only this video.
      const stored = JSON.parse(props.getProperty(REGEN_QUEUE_KEY_) || "[]");
      const at = stored.indexOf(vid);
      if (at !== -1) stored.splice(at, 1);
      queue.length = 0; Array.prototype.push.apply(queue, stored);
      props.setProperty(REGEN_QUEUE_KEY_, JSON.stringify(queue));
      props.setProperty(REGEN_FAILED_KEY_, JSON.stringify(failed));
    }
    if (!queue.length) finishRegeneration_(failed);
    else redeployIndex(); // publish this batch now, so results can be checked as they come
  } finally {
    lock.releaseLock();
  }
}

function finishRegeneration_(failed) {
  stopRegenerationTrigger_();
  redeployIndex();
  PropertiesService.getScriptProperties().deleteProperty(REGEN_QUEUE_KEY_);
  const body = failed.length
    ? "<p>These kept their old page:</p><ul>" + failed.map(function(f) { return "<li>" + escHtml_(f) + "</li>"; }).join("") + "</ul>"
    : "<p>All summaries were rewritten in the short format.</p>";
  try {
    GmailApp.sendEmail(CONFIG.EMAIL_TO, "Summary regeneration finished", "", {
      htmlBody: '<div style="font-family:Arial,sans-serif">' + body + '<p><a href="' + siteUrl_() + '">' + siteUrl_() + '</a></p></div>',
      name: "YouTube Doc Bot"
    });
  } catch (e) { Logger.log("Finish email failed: " + e.message); }
  Logger.log("Regeneration finished. Failed: " + failed.length);
}

function sheetRowsById_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getRange(CONFIG.START_ROW, 1, sheet.getLastRow() - CONFIG.START_ROW + 1, 6).getValues();
  const out = {};
  data.forEach(function(row) {
    const url = (row[CONFIG.COL_URL - 1] || "").toString().trim();
    const vid = url ? extractVideoId(url) : null;
    if (!vid) return;
    out[vid] = {
      url: url,
      title: (row[CONFIG.COL_TITLE - 1] || "Video " + vid).toString().trim(),
      topic: (row[CONFIG.COL_TOPIC - 1] || "General").toString().trim()
    };
  });
  return out;
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
