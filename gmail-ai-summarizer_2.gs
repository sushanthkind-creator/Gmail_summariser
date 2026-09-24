/**
 * COLLEGE MAIL FILTER — IF/ELSE FORWARDING + DATE DETECTION
 * ----------------------------------------------------------
 * Every 15 minutes:
 *  1. PRIORITY SENDERS -> always forwarded (skipped only if the subject is an
 *     obvious promo like a webinar or internship, see SKIP_PRIORITY_PROMOS).
 *  2. EVERYONE ELSE -> forwarded ONLY if the subject clearly matches one of the
 *     6 categories: Quiz, Exam, Assignment Submission, Holiday,
 *     Suspension of Classes, Hackathon.
 *  3. You receive the FULL original email (with attachments), plus a header
 *     showing the category and the detected date/time.
 *  4. Date/time detection: Gemini if USE_AI_FOR_DATES is true, otherwise
 *     (or if Gemini fails) built-in date parsing. AI never decides what gets
 *     forwarded.
 *  5. Quiz / Exam / Assignment with a detected date -> Google Calendar event.
 *  6. Never forwards the same mail, or the same event, twice.
 *
 * SETUP:
 *   1. If USE_AI_FOR_DATES is true: Project Settings (gear) > Script Properties
 *      > GEMINI_API_KEY = <your key>   (never paste the key into this code)
 *   2. Run setup() once.
 *   3. Run markAllRecentAsDone() once to skip mail that's already in the inbox.
 */

// ============ CONFIG ============

const FORWARD_TO = "pitcrewf@gmail.com";

// false = no AI at all (no API key needed). true = Gemini reads dates/times.
const USE_AI_FOR_DATES = true;
const GEMINI_MODEL = "gemini-3.5-flash";

const PRIORITY_SENDERS = [
  "debanjali.sarkar@vitap.ac.in",
  "arun.yadav@vitap.ac.in",
  "ganesh.reddy@vitap.ac.in",
  "srinivasa.popuri@vitap.ac.in",
  "koteswarao.gorantla@vitap.ac.in",
  "sts_30084@vitap.ac.in",
  "dean.acad@vitap.ac.in",
  "vc@vitap.ac.in",
  "registrar@vitap.ac.in",
  "dean.scope@vitap.ac.in",
  "coe@vitap.ac.in",
];

// Skip obvious promotions even from priority senders
const SKIP_PRIORITY_PROMOS = true;

// Subject rules. Checked top to bottom; first match wins.
const CATEGORY_RULES = [
  ["Suspension of Classes", [/suspen\w*.*\bclass/, /\bclass\w*.*suspen/, /\bno class/,
                             /\bclass\w* (are |is |will be )?(cancel|not be held)/]],
  ["Holiday",               [/\bholiday/, /declared closed/, /remain closed/]],
  ["Hackathon",             [/hackathon/]],
  ["Quiz",                  [/\bquiz/]],
  ["Exam",                  [/\bexam/, /\b(cat|fat)\s*-?\s*\d/, /\b(cat|fat)\b/,
                             /\bmid[\s-]?term/, /\bend[\s-]?term/]],
  ["Assignment Submission", [/\bassignment/, /last date (to|for) submi/]],
];

// Subjects containing these are treated as promotions and never forwarded
const PROMO_WORDS = [
  "webinar", "workshop", "recruitment", "internship", "certification",
  "register now", "join the", "masterclass", "bootcamp",
];

const CALENDAR_CATEGORIES = ["Quiz", "Exam", "Assignment Submission"];

const PROCESSED_LABEL = "Processed/AIReviewed";
const LABEL_ROOT = "College/AI/";
const LOOKBACK_DAYS = 2;
const TZ = "Asia/Kolkata";

// ============ SETUP ============

function setup() {
  getOrCreateLabel_(PROCESSED_LABEL);
  CATEGORY_RULES.map(r => r[0]).concat(["Priority", "Duplicate"])
    .forEach(c => getOrCreateLabel_(LABEL_ROOT + c));

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "processInboxWithAI") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("processInboxWithAI").timeBased().everyMinutes(15).create();
  Logger.log("Setup complete. Runs every 15 min.");
}

// Run once: marks everything already in the inbox as handled, WITHOUT forwarding.
function markAllRecentAsDone() {
  const state = loadState_();
  let n = 0;
  GmailApp.search("newer_than:" + LOOKBACK_DAYS + "d", 0, 200).forEach(thread => {
    thread.getMessages().forEach(m => {
      if (!state.doneIds.has(m.getId())) { markDone_(state, m.getId()); n++; }
    });
  });
  saveState_(state);
  Logger.log("Marked " + n + " existing message(s) as done. Only new mail will be forwarded.");
}

// ============ MAIN ============

function processInboxWithAI() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log("Previous run still active, skipping."); return; }

  const state = loadState_();
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  let aiDown = !USE_AI_FOR_DATES || !apiKey;
  if (USE_AI_FOR_DATES && !apiKey) Logger.log("No GEMINI_API_KEY: using built-in date parsing.");

  try {
    const threads = GmailApp.search(
      "newer_than:" + LOOKBACK_DAYS + "d -from:me -category:promotions -category:social", 0, 100);
    const cutoff = Date.now() - LOOKBACK_DAYS * 86400000;
    const myEmail = Session.getActiveUser().getEmail().toLowerCase();

    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        const id = msg.getId();
        if (state.doneIds.has(id)) continue;

        try {
          const from = msg.getFrom().toLowerCase();
          const subject = msg.getSubject() || "(no subject)";

          if (msg.getDate().getTime() < cutoff || (myEmail && from.includes(myEmail))) {
            markDone_(state, id); continue;
          }

          // ---------- THE IF/ELSE DECISION ----------
          const isPriority = PRIORITY_SENDERS.some(s => from.includes(s.toLowerCase()));
          const decision = decide_(subject, isPriority);
          if (!decision.forward) { markDone_(state, id); continue; }

          // ---------- DATE/TIME (AI or built-in) ----------
          const body = msg.getPlainBody();
          let when = null, whenSource = "none";
          if (!aiDown) {
            try {
              when = extractDateWithAI_(apiKey, subject, body.slice(0, 6000));
              whenSource = "AI";
            } catch (e) {
              aiDown = true;
              Logger.log("Gemini unavailable, using built-in date parsing: " + e);
            }
          }
          if (!when) {
            when = extractDateTime_(subject) ;
            if (!when.date) when = extractDateTime_(body);
            whenSource = when.date ? "parser" : "none";
          }

          forwardAndRecord_(state, thread, msg, subject, isPriority, decision.category, when, whenSource);
          markDone_(state, id);
          safeLabel_(thread, PROCESSED_LABEL);
          saveState_(state);                       // cancel-safe
        } catch (e) {
          const n = (state.failCounts[id] || 0) + 1;
          state.failCounts[id] = n;
          Logger.log("Error on message (attempt " + n + "/3): " + e);
          if (n >= 3) { markDone_(state, id); delete state.failCounts[id]; }
          saveState_(state);
        }
      }
    }
  } finally {
    saveState_(state);
    lock.releaseLock();
  }
}

// Pure if/else: returns { forward, category }
function decide_(subject, isPriority) {
  const s = subject.toLowerCase();
  const promo = PROMO_WORDS.some(w => s.includes(w));
  let category = null;
  if (!promo) {
    for (const [cat, patterns] of CATEGORY_RULES) {
      if (patterns.some(re => re.test(s))) { category = cat; break; }
    }
  }
  if (isPriority) {
    if (promo && SKIP_PRIORITY_PROMOS) return { forward: false, category: null };
    return { forward: true, category: category };        // category may be null
  }
  return { forward: category !== null, category: category };
}

function forwardAndRecord_(state, thread, msg, subject, isPriority, category, when, whenSource) {
  const date = when.date, time = when.date ? when.time : null;

  // Same event already forwarded? (same category + date + time, or same subject)
  const key = eventKey_(category, date, time, subject);
  if (state.sentKeys[key]) {
    safeLabel_(thread, LABEL_ROOT + "Duplicate");
    Logger.log("'" + subject + "' -> duplicate of an event already forwarded, skipped.");
    return;
  }

  const tag = isPriority ? ("PRIORITY" + (category ? " · " + category : "")) : category;
  const whenText = date ? date + (time ? " at " + time : "") : "not detected";
  Logger.log("'" + subject + "' -> " + tag + " | when: " + whenText + " (" + whenSource + ")");

  if (category) safeLabel_(thread, LABEL_ROOT + category);
  if (isPriority) safeLabel_(thread, LABEL_ROOT + "Priority");

  const header =
    '<div style="font-family:Arial,sans-serif;font-size:13px;background:#f3f4f6;' +
    'border-left:4px solid #2563eb;padding:10px 12px;margin-bottom:12px">' +
    "<b>Category:</b> " + escapeHtml_(tag) + "<br>" +
    "<b>When:</b> " + escapeHtml_(whenText) + "<br>" +
    "<b>From:</b> " + escapeHtml_(msg.getFrom()) + "<br>" +
    "<b>Received:</b> " + Utilities.formatDate(msg.getDate(), TZ, "dd MMM yyyy, hh:mm a") +
    "</div>";

  const options = {
    htmlBody: header + msg.getBody(),
    name: "College Mail Filter",
  };
  const plain = "Category: " + tag + "\nWhen: " + whenText + "\nFrom: " + msg.getFrom() +
                "\n\n" + msg.getPlainBody().slice(0, 5000);

  try {
    options.attachments = msg.getAttachments();
    GmailApp.sendEmail(FORWARD_TO, "[" + tag + "] " + subject, plain, options);
  } catch (e) {
    // Attachments too large or unsupported: send without them
    delete options.attachments;
    options.htmlBody = header + "<p><i>(Attachments were too large to forward; " +
                       "open the original in your college inbox.)</i></p>" + msg.getBody();
    GmailApp.sendEmail(FORWARD_TO, "[" + tag + "] " + subject, plain, options);
  }
  state.sentKeys[key] = Date.now();
  Logger.log("  Forwarded to " + FORWARD_TO);

  if (category && CALENDAR_CATEGORIES.includes(category) && date) {
    try {
      createCalendarEvent_(category, subject, msg.getPlainBody().slice(0, 1000), date, time);
      Logger.log("  Calendar event created.");
    } catch (e) {
      Logger.log("  Calendar event failed: " + e);
    }
  }
}

// ============ DATE/TIME: AI ============

function extractDateWithAI_(apiKey, subject, body) {
  const today = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd (EEEE)");
  const prompt =
    "Today is " + today + " (India time). From this college email, extract the date and time " +
    "of the main event or deadline it announces.\n" +
    "- date: YYYY-MM-DD, resolving words like today, tomorrow, or weekday names. null if none.\n" +
    "- time: 24-hour HH:MM ONLY if a time is explicitly written. Never guess. null otherwise.\n\n" +
    "Subject: " + subject + "\nBody:\n" + body + "\n\n" +
    "Respond ONLY with JSON: {\"date\":null,\"time\":null}";

  const call = () => UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODEL +
    ":generateContent?key=" + apiKey, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0, responseMimeType: "application/json" },
      }),
      muteHttpExceptions: true,
    });

  let res = call();
  if ([429, 500, 502, 503, 504].includes(res.getResponseCode())) {
    Utilities.sleep(3000);
    res = call();
  }
  if (res.getResponseCode() !== 200) {
    throw new Error("HTTP " + res.getResponseCode() + ": " + res.getContentText().slice(0, 200));
  }
  const data = JSON.parse(res.getContentText());
  const parts = data.candidates && data.candidates[0] && data.candidates[0].content &&
                data.candidates[0].content.parts;
  const text = parts && parts[0] && parts[0].text;
  if (!text) throw new Error("Empty Gemini response");

  const out = JSON.parse(text.replace(/```json|```/g, "").trim());
  const date = validDate_(out.date);
  return { date: date, time: date ? validTime_(out.time) : null };
}

// ============ DATE/TIME: BUILT-IN PARSER (no AI) ============

const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
                 jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
const WEEKDAYS = { monday: 1, tuesday: 2, wednesday: 3, thursday: 4,
                   friday: 5, saturday: 6, sunday: 7 };
const MONTH_RE = "(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\\.?";

function extractDateTime_(text) {
  const t = String(text || "").toLowerCase();
  const today = todayParts_();
  let date = null, m;

  if ((m = t.match(/\b(\d{4})-(\d{2})-(\d{2})\b/))) {
    date = ymd_(+m[1], +m[2], +m[3]);
  } else if ((m = t.match(/\b(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})\b/))) {       // dd/mm/yyyy
    date = ymd_(+m[3], +m[2], +m[1]);
  } else if ((m = t.match(new RegExp("\\b(\\d{1,2})(?:st|nd|rd|th)?\\s*(?:of\\s+)?" + MONTH_RE +
                                     ",?\\s*(\\d{4})?")))) {                       // 25th September
    date = ymd_(m[3] ? +m[3] : guessYear_(today, MONTHS[m[2]], +m[1]), MONTHS[m[2]], +m[1]);
  } else if ((m = t.match(new RegExp("\\b" + MONTH_RE + "\\s+(\\d{1,2})(?:st|nd|rd|th)?\\b,?\\s*(\\d{4})?")))) {
    date = ymd_(m[3] ? +m[3] : guessYear_(today, MONTHS[m[1]], +m[2]), MONTHS[m[1]], +m[2]); // September 25th
  } else if (/\btoday\b|\btonight\b/.test(t)) {
    date = addDays_(today, 0);
  } else if (/\btomorrow\b/.test(t)) {
    date = addDays_(today, 1);
  } else if ((m = t.match(/\b(next\s+)?(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/))) {
    let diff = (WEEKDAYS[m[2]] - today.dow + 7) % 7;
    if (m[1] && diff === 0) diff = 7;
    date = addDays_(today, diff);
  }

  let time = null;
  if ((m = t.match(/\b(\d{1,2})(?:[:.](\d{2}))?\s*(a\.?m\.?|p\.?m\.?)(?![a-z])/))) {
    let h = +m[1];
    const min = m[2] || "00";
    const pm = m[3].startsWith("p");
    if (pm && h < 12) h += 12;
    if (!pm && h === 12) h = 0;
    time = validTime_(pad_(h) + ":" + min);
  } else if ((m = t.match(/\b([01]?\d|2[0-3]):([0-5]\d)\s*(hrs|hours)?\b/))) {
    time = pad_(+m[1]) + ":" + m[2];
  }
  return { date: date, time: date ? time : null };
}

function todayParts_() {
  const s = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd-u");   // u = 1 (Mon) .. 7 (Sun)
  const [y, mo, d, dow] = s.split("-").map(Number);
  return { y: y, m: mo, d: d, dow: dow };
}

function ymd_(y, m, d) {
  if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31)) return null;
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCMonth() !== m - 1) return null;          // rejects 31 Feb etc.
  return y + "-" + pad_(m) + "-" + pad_(d);
}

function addDays_(today, n) {
  const dt = new Date(Date.UTC(today.y, today.m - 1, today.d + n));
  return dt.getUTCFullYear() + "-" + pad_(dt.getUTCMonth() + 1) + "-" + pad_(dt.getUTCDate());
}

// No year written: assume this year, unless that's more than ~2 months in the past
function guessYear_(today, m, d) {
  const thisYear = Date.UTC(today.y, m - 1, d);
  const now = Date.UTC(today.y, today.m - 1, today.d);
  return (now - thisYear > 60 * 86400000) ? today.y + 1 : today.y;
}

function pad_(n) { return ("0" + n).slice(-2); }

// ============ STATE ============

function loadState_() {
  const p = PropertiesService.getScriptProperties();
  const doneList = JSON.parse(p.getProperty("DONE_IDS") || "[]");
  return {
    doneList: doneList,
    doneIds: new Set(doneList),
    sentKeys: JSON.parse(p.getProperty("SENT_KEYS") || "{}"),
    failCounts: JSON.parse(p.getProperty("FAIL_COUNTS") || "{}"),
  };
}

function markDone_(state, id) {
  if (state.doneIds.has(id)) return;
  state.doneIds.add(id);
  state.doneList.push(id);
}

function saveState_(state) {
  const p = PropertiesService.getScriptProperties();
  p.setProperty("DONE_IDS", JSON.stringify(state.doneList.slice(-400)));
  const monthAgo = Date.now() - 30 * 86400000;
  const fresh = {};
  Object.keys(state.sentKeys).filter(k => state.sentKeys[k] > monthAgo).slice(-100)
    .forEach(k => fresh[k] = state.sentKeys[k]);
  p.setProperty("SENT_KEYS", JSON.stringify(fresh));
  const fc = {};
  Object.keys(state.failCounts).slice(-50).forEach(k => fc[k] = state.failCounts[k]);
  p.setProperty("FAIL_COUNTS", JSON.stringify(fc));
}

function resetMemory() {
  const p = PropertiesService.getScriptProperties();
  ["DONE_IDS", "SENT_KEYS", "FAIL_COUNTS"].forEach(k => p.deleteProperty(k));
  Logger.log("Memory cleared.");
}

// ============ HELPERS ============

function validDate_(v) {
  const s = String(v || "").trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? ymd_(+m[1], +m[2], +m[3]) : null;
}

function validTime_(v) {
  const s = String(v || "").trim();
  if (!/^\d{2}:\d{2}$/.test(s)) return null;
  const [h, m] = s.split(":").map(Number);
  return (h <= 23 && m <= 59) ? s : null;
}

function normText_(s) {
  return String(s || "").toLowerCase()
    .replace(/^((re|fw|fwd)\s*:\s*)+/g, "")
    .replace(/[^a-z0-9]+/g, " ").trim();
}

// Quiz at 18:30 on the 25th from two different emails = one event.
// Without a time, fall back to the subject so different notices aren't merged.
function eventKey_(category, date, time, subject) {
  if (category && date && time) return [category, date, time].join("|");
  return [category || "priority", date || "nodate", normText_(subject)].join("|");
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function createCalendarEvent_(category, subject, details, date, time) {
  const cal = CalendarApp.getDefaultCalendar();
  const title = "[" + category + "] " + subject;
  const opts = { description: details };
  if (time) {
    const start = new Date(date + "T" + time + ":00+05:30");
    cal.createEvent(title, start, new Date(start.getTime() + 3600000), opts);
  } else {
    cal.createAllDayEvent(title, new Date(date + "T00:00:00+05:30"), opts);
  }
}

const LABEL_CACHE_ = {};
function getOrCreateLabel_(name) {
  if (!LABEL_CACHE_[name]) {
    LABEL_CACHE_[name] = GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
  }
  return LABEL_CACHE_[name];
}

function safeLabel_(thread, name) {
  try { thread.addLabel(getOrCreateLabel_(name)); }
  catch (e) { Logger.log("  (Couldn't add label '" + name + "': " + e + ")"); }
}
