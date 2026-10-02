/**
 * MESS MENU -> HEALTHY PLAN  (add-on to the college mail filter)
 * ---------------------------------------------------------------
 * When the monthly mess menu arrives from MM_SENDER:
 *  1. Reads the attached Excel menu (2-week rotation, e.g. "Tue 1, 15, 29").
 *  2. Splits combo cells ("Uggani + Mirchi Bajji"), drops non-veg options.
 *  3. RULES decide the obvious: fried, sweets, butter/cream, pickles -> skip;
 *     salads, dal, curd, roti, fruit, sprouts -> eat.
 *     AI only judges dishes rules can't call (e.g. "Guthi Vankaya Curry").
 *  4. Builds a day-by-day plan for the whole month. Sundays = cheat day.
 *  5. Adds a nutrition table from a FIXED lookup (never AI-generated).
 *  6. Emails you the plan as an Excel file.
 *
 * SETUP (one time):
 *  1. Add this as a SECOND file in the same Apps Script project
 *     (Files "+" > Script > name it messMenu). It uses the same
 *     GEMINI_API_KEY in Script Properties as the mail filter.
 *  2. Services "+" > Drive API > Add   (needed to read Excel files).
 *  3. Run setupMessMenu() once.
 *  4. To test on the latest menu mail right now: run testMessMenuOnLatest().
 */

// ============ CONFIG ============

const MM_SENDER = "cw.mh@vitap.ac.in";
const MM_FORWARD_TO = "pitcrewf@gmail.com";
const MM_SHEET_NAME = "Veg & Non-Veg";   // or "Special" if you're on the special mess
const MM_USE_AI = true;                  // false = rules only
const MM_GEMINI_MODEL = "gemini-3.5-flash";
const MM_MEALS = ["Breakfast", "Lunch", "Snacks", "Dinner"];
const MM_PROTEIN_MEALS = ["Breakfast", "Lunch", "Dinner"];
const MM_CHEAT_DOW = 0;                  // 0 = Sunday

// ============ FOOD RULES (first match wins, top to bottom) ============

const MM_NONVEG = /\b(egg|eggs|omlet|omelet|omelette|chicken|fish|mutton|prawns?|keema|non[\s-]?veg|nv)\b/;

const MM_RULES = [
  // ---- SKIP: sweets ----
  ["skip", "sweet", /\b(gulab|jamun|halwa|kheer|payasam|jalebi|jilebi|badusha|poornalu|kulfi|rasamalai|rasmalai|ras malai|kalakand|tukda|ice ?cream|custard|laddu|ladoo|burfi|barfi|kesari|sheera|sharbath?|sharbat|chocos|cake|pastry|brownie|lassi|basundi|rabri|meetha|jam|sweet|chikki|boondi|shrikhand|mysore pak)\b/],
  // ---- SKIP: heavy / refined / extra fat ----
  ["skip", "oily rice/noodles", /\b(biryani|pulav|pulao|fried rice|noodles)\b/],
  ["skip", "rich (butter, cream or cashew)", /\b(butter masala|makhani|malai|creamy|cream|kaju|cashew)\b/],
  ["skip", "refined flour", /\b(pav|white bread|naan|kulcha|maggi|bhature)\b/],
  ["skip", "layered with fat", /\blacha\b/],
  ["skip", "extra fat", /\b(butter|ghee|dalda)\b/],
  ["skip", "very salty and oily", /\b(pickle|avakaya)\b/],
  ["skip", "sugary / processed", /\b(tomato sauce|ketchup|mayonnaise|cheese)\b/],
  // ---- SKIP: fried ----
  ["skip", "fried", /\b(fry|fried|fries|vada|vadai|wada|b+h?a+j+i|samosa|puri|poori|cutlet|punugulu|bonda|fryums|chips|papad|65|manchuria|manchurian|pakoda|pakora|kachori|murukku|chakli|mixture|bhatura|vadiyalu)\b/],
  // ---- OPTIONAL: specific cases that must beat the EAT list ----
  ["optional", "plain milk is the better pick", /\b(tea|coffee|cofee)\b/],
  ["optional", "made with oil, fine occasionally", /\bparatha\b/],
  ["optional", "whole fruit has more fibre", /\bjuice\b/],
  ["optional", "protein, but high in fat", /\bpeanutbutter\b/],
  ["optional", "okay in small amounts", /\b(brownbread|cornflakes|mint sauce)\b/],
  // ---- Ambiguous: AI judges (default eat) ----
  ["generic", null, /\b(maharani|boiledfry)\b/],
  // ---- EAT: exact single words ----
  ["eat", null, /^(onions?|coriander|sabja seeds)$/],
  // ---- EAT ----
  ["eat", null, /\b(salad|sprouts|sundal|chick ?peas|soya|rajma|chana|chole|paneer|dal|pappu|lentil|sambar|rasam|curd|dahi|raitha|raita|buttermilk|majiga|milk|fruit|guava|banana|papaya|watermelon|muskmelon|apple|pomegranate|orange|grapes|soup|poriyal|roti|pulka|phulka|chapathi|chapati|ragi|raagi|multi ?grain|museli|muesli|oats|idli|tepla|thepla|pesarattu|palak|methi|spinach|bachali|gongura|amaranthus|sabja|lemon water|idly|kura|pulusu|kootu|mushroom|peas|sweetpotato|stirfry)\b/],
  // ---- OPTIONAL: fine in small amounts ----
  ["optional", "fine in small amounts", /\b(chutney|podi|rice|pulihora|upma|poha|dosa|uttapam|utappam|sevai|semiya|shavige|bath|bisbele)\b/],
  // ---- GENERIC: depends on how it's cooked -> AI decides ----
  ["generic", null, /\b(curry|masala|gravy|kurma|korma|sabji|sabzi|salan|jalfrezi|kolhapuri)\b/],
];

// ============ NUTRITION TABLE (fixed facts, never AI) ============

const MM_NUTRITION = [
  ["Paneer", /paneer/, "Protein, Calcium", "Builds and repairs muscle; strengthens bones"],
  ["Dal & lentils", /\b(dal|pappu|lentil|moong|pesara|tepla)\b/, "Protein, Fibre, Iron, Folate", "Muscle repair, steady energy, healthy blood"],
  ["Sambar", /\bsambar\b/, "Protein, Fibre", "Dal plus vegetables in one bowl"],
  ["Rajma", /\brajma\b/, "Protein, Fibre, Iron", "Keeps you full; supports healthy blood"],
  ["Chana / chickpeas", /\b(chana|chole|chick ?peas)\b/, "Protein, Fibre, Iron", "Keeps you full; supports healthy blood"],
  ["Soya", /\bsoya\b/, "Complete protein, Iron", "Has all essential amino acids, like dairy does"],
  ["Sprouts", /\bsprouts\b/, "Protein, Fibre, Vitamin C", "Easy protein boost to start the day"],
  ["Peanuts", /\b(peanut|groundnut)\b/, "Protein, Healthy fats, Vitamin E", "Energy and skin health"],
  ["Curd & buttermilk", /\b(curd|dahi|raitha|raita|buttermilk|majiga)\b/, "Protein, Calcium, Probiotics", "Good gut bacteria, better digestion"],
  ["Milk", /\bmilk\b/, "Protein, Calcium, Vitamin B12", "Bones; B12 is hard to get on a vegetarian diet"],
  ["Leafy greens", /\b(palak|spinach|methi|amaranthus|bachali|gongura|sorrel)\b/, "Iron, Vitamin K, Folate, Fibre", "Healthy blood, bone strength, blood clotting"],
  ["Carrot", /\bcarrot\b/, "Vitamin A (beta-carotene)", "Eye health and skin"],
  ["Beetroot", /\bbeetroot\b/, "Folate, Fibre, Nitrates", "Supports blood flow and stamina"],
  ["Cucumber", /\bcucumber\b/, "Water, small Vitamin K", "Hydration, very low calorie"],
  ["Tomato", /\btomato\b/, "Vitamin C, Lycopene", "Immunity; antioxidant"],
  ["Lemon", /\blemon\b/, "Vitamin C", "Immunity; helps you absorb iron from dal and greens"],
  ["Ragi", /\b(ragi|raagi)\b/, "Calcium, Fibre, Iron", "Bones and steady energy"],
  ["Whole wheat roti", /\b(roti|pulka|phulka|chapathi|chapati)\b/, "Fibre, Complex carbs", "Steady energy without a sugar spike"],
  ["Multigrain / muesli", /\b(multi ?grain|museli|muesli|oats)\b/, "Fibre, B vitamins", "Digestion and steady energy"],
  ["Mushroom", /\bmushroom\b/, "B vitamins, Selenium", "Energy metabolism; antioxidant"],
  ["Broccoli", /\bbrocc?ol+i\b/, "Vitamin C, Vitamin K, Fibre", "Immunity and bone health"],
  ["Sweet potato", /\bsweetpotato\b/, "Vitamin A, Fibre", "Eye health and digestion"],
  ["Guava", /\bguava\b/, "Vitamin C (very high), Fibre", "Immunity"],
  ["Banana", /\bbanana\b/, "Potassium, Quick energy", "Muscle and heart function"],
  ["Papaya", /\bpapaya\b/, "Vitamin C, Vitamin A", "Digestion and immunity"],
  ["Melons", /\b(watermelon|muskmelon)\b/, "Water, Vitamin A & C", "Hydration"],
  ["Other fruit", /\b(apple|pomegranate|grapes|orange)\b/, "Fibre, Vitamin C, Antioxidants", "Immunity and digestion"],
  ["Green peas", /\bpeas\b/, "Plant protein, Fibre", "Adds protein to veg dishes"],
  ["Beans & cabbage", /\b(beans|cabbage)\b/, "Fibre, Vitamin K", "Digestion and bone health"],
  ["Gourds", /\b(gourd|kakarakaya|dondakaya|beerakaya|potlakaya|sorakaya|dosakaya)\b/, "Fibre, Water", "Light on the stomach, low calorie"],
  ["Brinjal", /\b(baingan|vankaya|brinjal)\b/, "Fibre, Antioxidants", "Digestion"],
  ["Capsicum", /\bcapsicum\b/, "Vitamin C", "Immunity"],
  ["Sabja seeds", /\bsabja\b/, "Fibre", "Digestion; keeps you full"],
];

// ============ SETUP & TRIGGERS ============

function setupMessMenu() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "checkMessMenuMail") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkMessMenuMail").timeBased().everyHours(1).create();
  Logger.log("Mess menu checker installed: runs every hour.");
}

// Runs hourly: handles each new menu mail exactly once.
function checkMessMenuMail() {
  const p = PropertiesService.getScriptProperties();
  const done = JSON.parse(p.getProperty("MM_DONE_IDS") || "[]");
  const threads = GmailApp.search("from:" + MM_SENDER + " has:attachment newer_than:10d", 0, 10);

  for (const thread of threads) {
    for (const msg of thread.getMessages()) {
      if (done.includes(msg.getId())) continue;
      const atts = MM_menuAttachments_(msg);
      if (atts.length) {
        try {
          atts.forEach(a => MM_processAttachment_(a, msg.getDate()));
        } catch (e) {
          Logger.log("Mess menu failed: " + e);
          continue;                                  // retry next hour
        }
      }
      done.push(msg.getId());
      p.setProperty("MM_DONE_IDS", JSON.stringify(done.slice(-50)));
    }
  }
}

// Manual test: processes the most recent menu mail, even if already handled.
function testMessMenuOnLatest() {
  const threads = GmailApp.search("from:" + MM_SENDER + " has:attachment newer_than:60d", 0, 10);
  for (const thread of threads) {
    const msgs = thread.getMessages().reverse();
    for (const msg of msgs) {
      const atts = MM_menuAttachments_(msg);
      if (atts.length) {
        Logger.log("Using mail: '" + msg.getSubject() + "' (" + msg.getDate() + ")");
        MM_processAttachment_(atts[0], msg.getDate());
        return;
      }
    }
  }
  Logger.log("No menu mail with an Excel attachment found from " + MM_SENDER);
}

function MM_menuAttachments_(msg) {
  return msg.getAttachments().filter(a =>
    /\.xlsx?$/i.test(a.getName()) || /spreadsheet|excel/i.test(a.getContentType()));
}

function MM_processAttachment_(att, mailDate) {
  const read = MM_readXlsx_(att.copyBlob(), MM_SHEET_NAME);
  Logger.log("Reading sheet: " + read.sheetName);

  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  const aiFn = (MM_USE_AI && apiKey) ? (items => MM_aiClassify_(apiKey, items)) : null;

  const plan = MM_buildPlan(read.values, { refDate: mailDate }, aiFn);
  plan.warnings.forEach(w => Logger.log("Warning: " + w));
  Logger.log("Plan built: " + plan.days.length + " days, AI used: " + plan.aiUsed +
             ", items to check: " + plan.toCheck.length);

  const xlsx = MM_writeXlsx_(plan, read.sheetName);
  const body =
    "Your healthy mess plan for " + plan.monthName + " " + plan.year + " is attached.\n\n" +
    "- Green = eat, ~ = small amounts, Sundays are cheat days.\n" +
    "- Skipped items and the reason are listed for every day.\n" +
    "- The nutrition table at the bottom shows what you get from the plan.\n" +
    (plan.toCheck.length ? "\nCouldn't classify these, so check them yourself: " +
                           plan.toCheck.join(", ") + "\n" : "") +
    (plan.aiUsed ? "" : "\n(AI was off or unavailable, so curries were judged by default rules.)\n") +
    "\nGeneral nutrition info, not medical advice.";
  GmailApp.sendEmail(MM_FORWARD_TO, "Healthy mess plan: " + plan.monthName + " " + plan.year,
                     body, { attachments: [xlsx], name: "Mess Menu Planner" });
  Logger.log("Plan emailed to " + MM_FORWARD_TO);
}

// ============ CORE LOGIC (pure: no Gmail/Drive calls, fully testable) ============

function MM_buildPlan(values, opts, aiFn) {
  const warnings = [];
  const grid = values.map(r => r.map(c => String(c == null ? "" : c)));

  // Header row + meal columns
  const hdr = MM_findHeaderRow_(grid);
  if (hdr < 0) throw new Error("Couldn't find the header row (Day / Breakfast / Lunch ...).");
  const mealCols = {};
  grid[hdr].forEach((h, i) => {
    MM_MEALS.forEach(m => { if (h.toLowerCase().includes(m.toLowerCase())) mealCols[m] = i; });
  });

  // Month + year
  const ref = opts.refDate ? new Date(opts.refDate) : new Date();
  const months = ["january","february","march","april","may","june","july",
                  "august","september","october","november","december"];
  let month = -1;
  for (let r = 0; r < hdr && month < 0; r++) {
    const txt = grid[r].join(" ").toLowerCase();
    months.forEach((m, i) => { if (month < 0 && txt.includes(m)) month = i; });
  }
  if (month < 0) { month = ref.getMonth(); warnings.push("Month not found in title; using mail month."); }
  let year = ref.getFullYear();
  if (month < ref.getMonth() - 6) year++;
  if (month > ref.getMonth() + 6) year--;

  // Day blocks
  const anchorRe = /^\s*(mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?[\s,:-]*([\d\s,&and]+)$/i;
  const dowIdx = { sun: 0, mon: 1, tue: 2, wed: 3, thu: 4, fri: 5, sat: 6 };
  const blocks = [];
  for (let r = hdr + 1; r < grid.length; r++) {
    const rowText = grid[r].join(" ").toUpperCase();
    if (/INSTRUCTION|^\s*NOTE/.test(rowText)) break;
    const m = grid[r][0].match(anchorRe);
    if (m) {
      blocks.push({ dow: dowIdx[m[1].toLowerCase()], dates: (m[2].match(/\d{1,2}/g) || []).map(Number),
                    start: r, end: grid.length, cells: {} });
      if (blocks.length > 1) blocks[blocks.length - 2].end = r;
    } else if (blocks.length && rowText.trim() === "" ) {
      // blank separator rows are fine
    }
  }
  if (!blocks.length) throw new Error("Couldn't find day blocks like 'Tue 1, 15, 29'.");
  // stop last block at instructions row
  for (let r = blocks[blocks.length - 1].start; r < grid.length; r++) {
    if (/INSTRUCTION/.test(grid[r].join(" ").toUpperCase())) { blocks[blocks.length - 1].end = r; break; }
  }

  // Collect + split items per block/meal
  const unique = {};
  blocks.forEach(b => {
    MM_MEALS.forEach(meal => {
      const col = mealCols[meal];
      b.cells[meal] = [];
      if (col === undefined) return;
      for (let r = b.start; r < b.end; r++) {
        const raw = (grid[r][col] || "").trim();
        if (!raw) continue;
        MM_splitCell_(raw, b.dates).forEach(it => {
          b.cells[meal].push(it);
          unique[it.name] = true;
        });
      }
    });
  });

  // Classify: rules first, AI for generic/unknown
  const verdicts = {};
  const needAI = [];
  Object.keys(unique).forEach(name => {
    const v = MM_classifyByRules_(name);
    verdicts[name] = v;
    if (v.verdict === "generic" || v.verdict === "unknown") needAI.push(name);
  });

  let aiUsed = false;
  if (needAI.length && aiFn) {
    try {
      const ai = aiFn(needAI) || {};
      needAI.forEach(n => {
        const a = ai[n];
        if (a && ["eat", "optional", "skip"].includes(a.verdict)) {
          verdicts[n] = { verdict: a.verdict, reason: a.reason || null, source: "ai" };
        }
      });
      aiUsed = true;
    } catch (e) {
      warnings.push("AI classification failed, using defaults: " + e);
    }
  }
  // Defaults for anything still undecided
  const toCheck = [];
  needAI.forEach(n => {
    const v = verdicts[n];
    if (v.verdict === "generic") verdicts[n] = { verdict: "eat", reason: null, source: "default" };
    if (v.verdict === "unknown") { verdicts[n] = { verdict: "check", reason: "not recognised", source: "default" }; toCheck.push(n); }
  });

  // Expand blocks into real dates
  const days = [];
  blocks.forEach(b => {
    b.dates.forEach(d => {
      const dt = new Date(Date.UTC(year, month, d));
      if (dt.getUTCMonth() !== month) { warnings.push("Invalid date " + d + " skipped."); return; }
      if (dt.getUTCDay() !== b.dow) warnings.push("Date " + d + " doesn't match its weekday in the menu.");
      const cheat = dt.getUTCDay() === MM_CHEAT_DOW;
      const day = { date: dt.toISOString().slice(0, 10), dow: dt.getUTCDay(), cheat: cheat, meals: {}, notes: [] };

      MM_MEALS.forEach(meal => {
        const bucket = { eat: [], optional: [], check: [], skip: [], all: [] };
        b.cells[meal].forEach(it => {
          if (it.onlyDates && !it.onlyDates.includes(d)) return;
          const v = verdicts[it.name];
          bucket.all.push(it.label);
          if (v.verdict === "eat") bucket.eat.push(it.label);
          else if (v.verdict === "optional") bucket.optional.push(it.label);
          else if (v.verdict === "check") bucket.check.push(it.label);
          else bucket.skip.push(it.label + " (" + (v.reason || "skip") + ")");
        });
        day.meals[meal] = bucket;
      });

      if (cheat) {
        day.notes.push("Cheat day: everything's fair game.");
      } else {
        MM_PROTEIN_MEALS.forEach(meal => {
          const has = day.meals[meal].eat.some(x => MM_isProtein_(x));
          if (!has && day.meals[meal].all.length) {
            day.notes.push(meal + ": no protein pick, add curd, milk or sprouts if available.");
          }
        });
        const sn = day.meals.Snacks;
        if (sn && sn.all.length && !sn.eat.length) {
          day.notes.push("Snacks: nothing healthy, have fruit or roasted chana instead.");
        }
      }
      days.push(day);
    });
  });
  days.sort((a, b) => a.date < b.date ? -1 : 1);

  // Nutrition: count days each food is on your plan (eat items only)
  const counts = {};
  days.forEach(day => {
    const seen = {};
    MM_MEALS.forEach(meal => day.meals[meal].eat.forEach(label => {
      const t = MM_norm_(label);
      MM_NUTRITION.forEach(([food, re]) => { if (re.test(t)) seen[food] = true; });
    }));
    Object.keys(seen).forEach(f => counts[f] = (counts[f] || 0) + 1);
  });
  const nutrition = MM_NUTRITION
    .filter(([food]) => counts[food])
    .map(([food, , nutrients, benefit]) => ({ food: food, days: counts[food], nutrients: nutrients, benefit: benefit }))
    .sort((a, b) => b.days - a.days);

  return {
    monthName: months[month].charAt(0).toUpperCase() + months[month].slice(1),
    year: year, days: days, nutrition: nutrition, verdicts: verdicts,
    toCheck: toCheck, aiUsed: aiUsed, warnings: warnings,
  };
}

function MM_findHeaderRow_(grid) {
  for (let r = 0; r < Math.min(grid.length, 15); r++) {
    const t = grid[r].join(" ").toLowerCase();
    if (t.includes("breakfast") && t.includes("lunch")) return r;
  }
  return -1;
}

// Normalise text so rules match reliably ("butter milk" -> "buttermilk", etc.)
function MM_norm_(s) {
  return String(s).toLowerCase()
    .replace(/sweet\s*potato/g, "sweetpotato")
    .replace(/sweet\s*corn/g, "sweetcorn")
    .replace(/dal\s*fry/g, "dal tadka")
    .replace(/boiled\s*fry/g, "boiledfry")
    .replace(/butter\s*milk/g, "buttermilk")
    .replace(/peanut\s*butter/g, "peanutbutter")
    .replace(/stir[\s-]*fry/g, "stirfry")
    .replace(/brown\s*bread/g, "brownbread")
    .replace(/\s+/g, " ").trim();
}

// Split on top-level commas and "+", handle "a / b" alternatives, drop non-veg.
function MM_splitCell_(raw, blockDates) {
  const out = [];
  MM_splitTop_(raw, ",").forEach(part => {
    MM_splitTop_(part, "+").forEach(piece => {
      let p = piece.trim();
      if (!p) return;

      // Date-specific items like "Rasamalai (1, 29)"
      let onlyDates = null;
      const dm = p.match(/\(\s*(\d{1,2}(?:\s*,\s*\d{1,2})*)\s*\)\s*$/);
      if (dm) {
        const nums = dm[1].match(/\d{1,2}/g).map(Number);
        if (nums.every(n => blockDates.includes(n))) {
          onlyDates = nums;
          p = p.slice(0, dm.index).trim();
        }
      }

      // Alternatives: keep only vegetarian options
      const alts = MM_splitTop_(p, "/").map(a => a.trim()).filter(Boolean);
      const veg = alts.filter(a => !MM_NONVEG.test(MM_norm_(a)));
      if (!veg.length) return;
      let label = veg.join(" / ")
        .replace(/\(\s*(veg|non[\s-]?veg|nv)\s*\)/ig, "")
        .replace(/\s+/g, " ").trim();
      if (!label) return;
      out.push({ name: MM_norm_(label), label: label, onlyDates: onlyDates });
    });
  });
  return out;
}

function MM_splitTop_(s, sep) {
  const parts = [];
  let depth = 0, cur = "";
  for (const ch of s) {
    if (ch === "(") depth++;
    if (ch === ")") depth = Math.max(0, depth - 1);
    if (ch === sep && depth === 0) { parts.push(cur); cur = ""; } else cur += ch;
  }
  parts.push(cur);
  return parts;
}

function MM_classifyByRules_(name) {
  const t = MM_norm_(name);
  for (const [verdict, reason, re] of MM_RULES) {
    if (re.test(t)) return { verdict: verdict, reason: reason, source: "rule" };
  }
  return { verdict: "unknown", reason: null, source: "rule" };
}

function MM_isProtein_(label) {
  return /\b(paneer|dal|pappu|lentil|moong|pesara|tepla|sambar|rajma|chana|chole|chick ?peas|soya|sprouts|sundal|peanut|groundnut|curd|dahi|raitha|raita|buttermilk|majiga|milk|peas)\b/
    .test(MM_norm_(label));
}

// ============ AI: judges only the dishes rules couldn't ============

function MM_aiClassify_(apiKey, items) {
  const prompt =
    "A vegetarian student wants to eat healthy from a college mess in Andhra Pradesh, India. " +
    "Strict rules: no deep-fried food, no sweets, nothing heavy with lots of butter, cream or oil. " +
    "For each dish below, decide:\n" +
    "- eat: a healthy everyday choice\n- optional: fine in small amounts\n- skip: fried, sweet, or heavy\n" +
    "Give a reason of at most 5 words. Use the dish names exactly as given.\n\n" +
    "Dishes:\n" + items.map(i => "- " + i).join("\n") + "\n\n" +
    "Respond ONLY with JSON: {\"items\":[{\"name\":\"\",\"verdict\":\"\",\"reason\":\"\"}]}";

  const call = () => UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/" + MM_GEMINI_MODEL +
    ":generateContent?key=" + apiKey, {
      method: "post", contentType: "application/json", muteHttpExceptions: true,
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0, responseMimeType: "application/json" },
      }),
    });
  let res = call();
  if ([429, 500, 502, 503, 504].includes(res.getResponseCode())) { Utilities.sleep(5000); res = call(); }
  if (res.getResponseCode() !== 200) throw new Error("HTTP " + res.getResponseCode());

  const data = JSON.parse(res.getContentText());
  const parts = data.candidates && data.candidates[0] && data.candidates[0].content &&
                data.candidates[0].content.parts;
  const text = parts && parts[0] && parts[0].text;
  if (!text) throw new Error("Empty AI response");
  const parsed = JSON.parse(text.replace(/```json|```/g, "").trim());

  const map = {};
  (parsed.items || []).forEach(x => {
    if (x && x.name) map[MM_norm_(x.name)] = { verdict: String(x.verdict || "").toLowerCase(), reason: x.reason };
  });
  return map;
}

// ============ EXCEL IN / OUT (Google Drive) ============

function MM_readXlsx_(blob, preferredSheet) {
  let fileId;
  try {
    fileId = Drive.Files.create({ name: "mess-menu-tmp", mimeType: MimeType.GOOGLE_SHEETS }, blob).id;
  } catch (e) {
    if (typeof Drive === "undefined") {
      throw new Error("Drive API service not enabled. In Apps Script: Services (+) > Drive API > Add.");
    }
    fileId = Drive.Files.insert({ title: "mess-menu-tmp", mimeType: MimeType.GOOGLE_SHEETS },
                                blob, { convert: true }).id;
  }
  try {
    const ss = SpreadsheetApp.openById(fileId);
    let sh = ss.getSheetByName(preferredSheet);
    if (!sh) {
      sh = ss.getSheets().find(s => MM_findHeaderRow_(s.getDataRange().getDisplayValues()) >= 0)
           || ss.getSheets()[0];
    }
    return { values: sh.getDataRange().getDisplayValues(), sheetName: sh.getName() };
  } finally {
    DriveApp.getFileById(fileId).setTrashed(true);
  }
}

function MM_rowsForSheet_(plan) {
  const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const rows = [];
  plan.days.forEach(day => {
    const cells = MM_MEALS.map(meal => {
      const m = day.meals[meal];
      if (day.cheat) return m.all.join("\n") || "-";
      const lines = m.eat.concat(m.optional.map(x => "~ " + x), m.check.map(x => "? " + x));
      return lines.join("\n") || "-";
    });
    const skipped = day.cheat ? "" : MM_MEALS
      .filter(meal => day.meals[meal].skip.length)
      .map(meal => meal + ": " + day.meals[meal].skip.join(", ")).join("\n");
    const [y, mo, d] = day.date.split("-");
    rows.push([d + "/" + mo + "/" + y, dayNames[day.dow] + (day.cheat ? " (cheat day)" : "")]
              .concat(cells, [skipped, day.notes.join("\n")]));
  });
  return rows;
}

function MM_writeXlsx_(plan, sourceSheet) {
  const title = "Healthy Mess Plan - " + plan.monthName + " " + plan.year;
  const ss = SpreadsheetApp.create(title);
  const sh = ss.getSheets()[0];
  sh.setName("Healthy Plan");
  const COLS = 8;

  sh.getRange(1, 1, 1, COLS).merge().setValue(title + " (vegetarian, " + sourceSheet + " mess)")
    .setFontSize(14).setFontWeight("bold");
  sh.getRange(2, 1, 1, COLS).merge()
    .setValue("Green = eat  |  ~ = fine in small amounts  |  ? = check yourself  |  " +
              "Orange rows = Sunday cheat day  |  Yellow notes = something to add")
    .setFontStyle("italic").setFontColor("#555555");

  const header = ["Date", "Day"].concat(MM_MEALS, ["Skipped (why)", "Notes"]);
  sh.getRange(3, 1, 1, COLS).setValues([header])
    .setFontWeight("bold").setBackground("#1f4e78").setFontColor("#ffffff");

  const rows = MM_rowsForSheet_(plan);
  const body = sh.getRange(4, 1, rows.length, COLS);
  body.setValues(rows).setWrap(true).setVerticalAlignment("top");

  const bg = rows.map((r, i) => {
    const day = plan.days[i];
    if (day.cheat) return new Array(COLS).fill("#fde2c4");
    const noteBg = day.notes.length ? "#fff2a8" : "#ffffff";
    return ["#ffffff", "#ffffff", "#e2f4e2", "#e2f4e2", "#e2f4e2", "#e2f4e2", "#fbe4e4", noteBg];
  });
  body.setBackgrounds(bg);

  // Nutrition table
  let r = 4 + rows.length + 2;
  sh.getRange(r, 1, 1, COLS).merge().setValue("What you get from this plan (" + plan.monthName + ")")
    .setFontSize(13).setFontWeight("bold");
  r++;
  sh.getRange(r, 1, 1, 4).setValues([["Food", "Days on your plan", "Key nutrients", "What it does for you"]])
    .setFontWeight("bold").setBackground("#1f4e78").setFontColor("#ffffff");
  r++;
  if (plan.nutrition.length) {
    sh.getRange(r, 1, plan.nutrition.length, 4)
      .setValues(plan.nutrition.map(n => [n.food, n.days, n.nutrients, n.benefit]))
      .setWrap(true).setVerticalAlignment("top");
    r += plan.nutrition.length;
  }
  sh.getRange(r + 1, 1, 1, COLS).merge()
    .setValue("General nutrition information, not medical advice. Nutrients come from a fixed table, " +
              "not AI. Days count only items marked to eat.")
    .setFontStyle("italic").setFontColor("#777777");

  [95, 150, 210, 230, 170, 230, 260, 240].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(3);
  sh.getRange(1, 1, sh.getMaxRows(), COLS).setFontFamily("Arial");
  SpreadsheetApp.flush();

  const blob = UrlFetchApp.fetch(
    "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx",
    { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() } }
  ).getBlob().setName("Healthy_Mess_Plan_" + plan.monthName + "_" + plan.year + ".xlsx");
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  return blob;
}
