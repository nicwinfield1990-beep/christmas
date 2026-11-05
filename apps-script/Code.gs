/**
 * Christmas at Ours - Google Sheet backend.
 *
 * Setup (once):
 *  1. Create a new Google Sheet. Extensions > Apps Script. Paste this file over Code.gs.
 *  2. Set GUEST_CODE and HOST_CODE below to your own values.
 *  3. Run the `setup` function once (approve the permissions prompt). It creates the tabs and the menu.
 *  4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Deploy.
 *  5. Copy the Web app URL into API_URL at the top of christmas/index.html.
 */

const GUEST_CODE = "CHANGE_ME_GUEST";
const HOST_CODE = "CHANGE_ME_HOST";

const TABS = {
  RSVPs:    ["id", "household", "adults", "kids", "days", "arrive", "updated"],
  Hosts:    ["household_id", "household", "dietary", "dietary_detail", "note_to_hosts", "updated"],
  Meals:    ["id", "day", "slot", "title", "desc", "dishes", "status"],
  Ideas:    ["id", "mealId", "title", "note", "by", "at"],
  Votes:    ["household_id", "mealId", "ideaId"],
  Shopping: ["id", "name", "qty", "cat", "mealId", "note", "bringer", "got", "by", "at"],
};

const SEED_MEALS = [
  ["d24-dinner", "24", "dinner", "", "Not decided yet. Suggest something and vote.", "", "poll"],
  ["d25-breakfast", "25", "breakfast", "Smoked salmon & scrambled eggs", "", "", "confirmed"],
  ["d25-lunch", "25", "lunch", "Picky bits & canapés", "Full list to follow.", "", "confirmed"],
  ["d25-dinner", "25", "dinner", "Christmas dinner", "Full menu to follow.", "", "confirmed"],
  ["d26-breakfast", "26", "breakfast", "Bacon sarnies", "", "", "confirmed"],
];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(TABS[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, TABS[name].length).setFontWeight("bold");
    }
  });
  const meals = ss.getSheetByName("Meals");
  if (meals.getLastRow() === 1) SEED_MEALS.forEach(r => meals.appendRow(r));
  const first = ss.getSheets()[0];
  if (first.getName() === "Sheet1" && first.getLastRow() === 0) ss.deleteSheet(first);
}

/* ---------- sheet helpers ---------- */
function sheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function rows_(name) {
  const sh = sheet_(name), vals = sh.getDataRange().getValues(), head = vals.shift();
  return vals.filter(r => r.join("") !== "").map(r => { const o = {}; head.forEach((h, i) => o[h] = r[i]); return o; });
}
function findRow_(name, key, value) {
  const vals = sheet_(name).getDataRange().getValues(), col = vals[0].indexOf(key);
  for (let i = 1; i < vals.length; i++) if (String(vals[i][col]) === String(value)) return i + 1;
  return -1;
}
function upsert_(name, key, obj) {
  const sh = sheet_(name), head = TABS[name], row = head.map(h => obj[h] === undefined ? "" : obj[h]);
  const at = findRow_(name, key, obj[key]);
  if (at > 0) sh.getRange(at, 1, 1, head.length).setValues([row]); else sh.appendRow(row);
}
function patch_(name, id, fields) {
  const at = findRow_(name, "id", id); if (at < 0) return;
  const sh = sheet_(name), head = TABS[name];
  Object.keys(fields).forEach(k => { const c = head.indexOf(k); if (c >= 0) sh.getRange(at, c + 1).setValue(fields[k]); });
}
function remove_(name, key, value) { const at = findRow_(name, key, value); if (at > 0) sheet_(name).deleteRow(at); }
function clip_(s, n) { return String(s == null ? "" : s).slice(0, n); }
function uid_() { return Utilities.getUuid().slice(0, 8); }

/* ---------- read ---------- */
function state_(isHost) {
  const out = {
    meals: rows_("Meals").map(m => ({ id: m.id, day: String(m.day), slot: m.slot, title: m.title, desc: m.desc,
      dishes: String(m.dishes || "").split("\n").filter(Boolean), status: m.status || "tbc" })),
    rsvps: rows_("RSVPs").map(r => ({ id: r.id, household: r.household, adults: +r.adults || 0, kids: +r.kids || 0,
      days: String(r.days || "").split(",").filter(Boolean), arrive: r.arrive, updated: r.updated })),
    ideas: rows_("Ideas").map(i => ({ id: i.id, mealId: i.mealId, title: i.title, note: i.note, by: i.by, at: i.at })),
    votes: rows_("Votes").map(v => ({ hh: v.household_id, mealId: v.mealId, ideaId: v.ideaId })),
    items: rows_("Shopping").map(i => ({ id: i.id, name: i.name, qty: String(i.qty), cat: i.cat, mealId: i.mealId, note: i.note,
      bringer: i.bringer || null, got: i.got === true || i.got === "TRUE", by: i.by, at: i.at })),
    isHost: isHost,
  };
  if (isHost) out.hosts = rows_("Hosts").map(h => ({ hh: h.household_id, household: h.household,
    needs: String(h.dietary || "").split(",").filter(Boolean), detail: h.dietary_detail, note: h.note_to_hosts }));
  return out;
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.code !== GUEST_CODE && p.code !== HOST_CODE) return json_({ error: "bad_code" });
  return json_(state_(p.host === HOST_CODE));
}

/* ---------- write ---------- */
function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: "bad_request" }); }
  if (b.code !== GUEST_CODE && b.code !== HOST_CODE) return json_({ error: "bad_code" });
  const isHost = b.host === HOST_CODE;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const now = new Date().toISOString();
    switch (b.action) {
      case "rsvp": {
        const id = clip_(b.id, 20) || uid_();
        upsert_("RSVPs", "id", { id, household: clip_(b.household, 60), adults: +b.adults || 0, kids: +b.kids || 0,
          days: (b.days || []).map(String).join(","), arrive: clip_(b.arrive, 60), updated: now });
        upsert_("Hosts", "household_id", { household_id: id, household: clip_(b.household, 60), dietary: (b.needs || []).join(","),
          dietary_detail: clip_(b.detail, 400), note_to_hosts: clip_(b.note, 400), updated: now });
        return json_({ ok: true, id, mine: { needs: b.needs || [], detail: clip_(b.detail, 400), note: clip_(b.note, 400) } });
      }
      case "mine": { // a household's own private answers, to refill their form
        const h = rows_("Hosts").find(x => String(x.household_id) === String(b.id));
        return json_({ ok: true, mine: h ? { needs: String(h.dietary || "").split(",").filter(Boolean), detail: h.dietary_detail, note: h.note_to_hosts } : null });
      }
      case "addIdea":
        sheet_("Ideas").appendRow([uid_(), clip_(b.mealId, 20), clip_(b.title, 80), clip_(b.note, 80), clip_(b.by, 20), now]); break;
      case "deleteIdea": remove_("Ideas", "id", b.id); break;
      case "vote": {
        const sh = sheet_("Votes"), vals = sh.getDataRange().getValues();
        for (let i = vals.length - 1; i >= 1; i--) if (String(vals[i][0]) === String(b.hh) && vals[i][1] === b.mealId) sh.deleteRow(i + 1);
        if (b.ideaId) sh.appendRow([clip_(b.hh, 20), clip_(b.mealId, 20), clip_(b.ideaId, 20)]);
        break;
      }
      case "addItems":
        (b.items || []).slice(0, 40).forEach(i => sheet_("Shopping").appendRow([uid_(), clip_(i.name, 80), "'" + clip_(i.qty, 30), clip_(i.cat, 40),
          clip_(i.mealId, 20), clip_(i.note, 120), clip_(i.bringer, 20), false, clip_(b.by, 20), now]));
        break;
      case "updateItem": {
        const f = {}; if ("got" in b) f.got = !!b.got; if ("bringer" in b) f.bringer = clip_(b.bringer, 20);
        patch_("Shopping", b.id, f); break;
      }
      case "deleteItem": remove_("Shopping", "id", b.id); break;
      case "saveMeal":
        if (!isHost) return json_({ error: "host_only" });
        upsert_("Meals", "id", { id: clip_(b.id, 20), day: clip_(b.day, 4), slot: clip_(b.slot, 12), title: clip_(b.title, 80),
          desc: clip_(b.desc, 200), dishes: (b.dishes || []).map(d => clip_(d, 100)).join("\n"), status: clip_(b.status, 12) });
        break;
      default: return json_({ error: "unknown_action" });
    }
    SpreadsheetApp.flush();
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
