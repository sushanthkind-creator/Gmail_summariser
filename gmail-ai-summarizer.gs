/**
 * COLLEGE MAIL — AI READ, CLASSIFY, SUMMARIZE, FORWARD
 * -----------------------------------------------------
 * 1. Only processes mail from PRIORITY_SENDERS, or mail whose subject
 *    contains one of SUBJECT_KEYWORDS.
 * 2. Gemini classifies it into a category and writes a short summary.
 * 3. PRIORITY_SENDERS mail is ALWAYS forwarded. Everyone else is only
 *    forwarded if the category is one of the 6 you care about.
 * 4. Quiz/Exam/Assignment with a date also get a calendar event.
 *
 * SETUP:
 *   1. Project Settings (gear icon) > Script Properties
 *      > Property = GEMINI_API_KEY, Value = <your real key>
 *      (NEVER paste the key into this code.)
 *   2. Run setup() once.
 */

// ============ CONFIG ============

const FORWARD_TO = "personal gmail id";
const GEMINI_MODEL = "gemini-3.5-flash"; // old gemini-2.0-flash is retired

const PRIORITY_SENDERS = [
  "faculty@vitap.ac.in",
];

const SUBJECT_KEYWORDS = [
  "exam", "quiz", "assignment", "holiday", "suspend", "hackathon",
  "class", "notice", "circular", "deadline",
];

const FORWARD_CATEGORIES = [
  "Quiz", "Exam", "Assignment Submission",
  "Holiday", "Suspension of Classes", "Hackathon",
];

const CALENDAR_CATEGORIES = ["Quiz", "Exam", "Assignment Submission"];

const PROCESSED_LABEL = "Processed/AIReviewed";
const MAX_BODY_CHARS = 8000;

// ============ SETUP (run once) ============

function setup() {
  getOrCreateLabel_(PROCESSED_LABEL);
  FORWARD_CATEGORIES.forEach(c => getOrCreateLabel_("College/AI/" + c));
  getOrCreateLabel_("College/AI/Other");
  getOrCreateLabel_("College/AI/Priority");

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "processInboxWithAI") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("processInboxWithAI").timeBased().everyMinutes(15).create();

  Logger.log("Setup complete. Trigger runs every 15 min.");
}

// ============ MAIN LOGIC ============

function processInboxWithAI() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("GEMINI_API_KEY not set in Script Properties.");

  const threads = GmailApp.search(buildCandidateQuery_(), 0, 30);
  Logger.log("Found " + threads.length + " unprocessed candidate mail(s).");

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const msg = messages[messages.length - 1];
    const subject = msg.getSubject();
    const body = msg.getPlainBody().slice(0, MAX_BODY_CHARS);
    const from = msg.getFrom().toLowerCase();
    const isPriority = PRIORITY_SENDERS.some(s => from.includes(s.toLowerCase()));

    let result;
    try {
      result = classifyWithGemini_(apiKey, subject, body);
    } catch (e) {
      // NOT marked processed, so it retries on the next run
      Logger.log("Gemini FAILED for '" + subject + "': " + e);
      return;
    }

    const category = FORWARD_CATEGORIES.includes(result.category) ? result.category : "Other";
    Logger.log("'" + subject + "' -> " + category + (isPriority ? " (PRIORITY)" : ""));

    thread.addLabel(getOrCreateLabel_("College/AI/" + category));
    if (isPriority) thread.addLabel(getOrCreateLabel_("College/AI/Priority"));

    if (isPriority || category !== "Other") {
      GmailApp.sendEmail(
        FORWARD_TO,
        "[" + (isPriority ? "PRIORITY" : category) + "] " + subject,
        "Category: " + category + "\n\nSummary:\n" + result.summary +
        "\n\n---\nFrom: " + msg.getFrom()
      );
      Logger.log("  Forwarded to " + FORWARD_TO);
    }

    if (CALENDAR_CATEGORIES.includes(category) && result.date) {
      try {
        createCalendarEvent_(category, subject, result);
        Logger.log("  Calendar event created for " + result.date);
      } catch (e) {
        Logger.log("  Calendar event failed: " + e);
      }
    }

    thread.addLabel(getOrCreateLabel_(PROCESSED_LABEL));
  });
}

// ============ ONE-TIME FIX ============
// Removes the processed label from the last 2 days of mail so anything
// that got stuck during earlier failed runs gets re-processed.
// Run once, then run processInboxWithAI.

function resetRecentProcessed() {
  const label = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!label) return;
  const threads = GmailApp.search("label:" + PROCESSED_LABEL + " newer_than:2d");
  threads.forEach(t => t.removeLabel(label));
  Logger.log("Reset " + threads.length + " thread(s). Now run processInboxWithAI.");
}

// ============ HELPERS ============

function buildCandidateQuery_() {
  const senderClauses = PRIORITY_SENDERS.map(s => "from:" + s);
  const subjectClauses = SUBJECT_KEYWORDS.map(k => "subject:" + k);
  const orClause = senderClauses.concat(subjectClauses).join(" OR ");
  return "(" + orClause + ") -label:" + PROCESSED_LABEL + " newer_than:2d";
}

function classifyWithGemini_(apiKey, subject, body) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/" +
    GEMINI_MODEL + ":generateContent?key=" + apiKey;

  const prompt = "You are classifying a college email.\n" +
    "Subject: " + subject + "\n" +
    "Body: " + body + "\n\n" +
    "Classify it into EXACTLY ONE of these categories: Quiz, Exam, Assignment Submission, " +
    "Holiday, Suspension of Classes, Hackathon, Other.\n" +
    "Then write a 2-3 sentence summary (dates, deadlines, action needed).\n\n" +
    "If the email mentions a specific date for the quiz/exam/deadline, give it as \"YYYY-MM-DD\" " +
    "(today is " + Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd") +
    "; resolve words like 'today' or 'tomorrow' from that). If a time is mentioned, give it as " +
    "24-hour \"HH:MM\". If not stated, use null.\n\n" +
    "Respond ONLY with JSON:\n" +
    "{\"category\": \"...\", \"summary\": \"...\", \"date\": \"YYYY-MM-DD or null\", \"time\": \"HH:MM or null\"}";

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, responseMimeType: "application/json" },
    }),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  const raw = response.getContentText();
  if (code !== 200) throw new Error("HTTP " + code + ": " + raw.slice(0, 300));

  const data = JSON.parse(raw);
  let text = null;
  if (data.candidates && data.candidates[0] && data.candidates[0].content &&
      data.candidates[0].content.parts && data.candidates[0].content.parts[0]) {
    text = data.candidates[0].content.parts[0].text;
  }
  if (!text) throw new Error("Empty Gemini response: " + raw.slice(0, 300));

  return JSON.parse(text.replace(/```json|```/g, "").trim());
}

function createCalendarEvent_(category, subject, result) {
  const calendar = CalendarApp.getDefaultCalendar();
  const title = "[" + category + "] " + subject;
  const options = { description: result.summary };

  if (result.time && result.time !== "null") {
    const start = new Date(result.date + "T" + result.time + ":00+05:30");
    const end = new Date(start.getTime() + 60 * 60 * 1000);
    calendar.createEvent(title, start, end, options);
  } else {
    calendar.createAllDayEvent(title, new Date(result.date + "T00:00:00+05:30"), options);
  }
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}
