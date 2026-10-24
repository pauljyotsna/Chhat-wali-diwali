/**
 * Chhat Wali Diwali — RSVP backend
 * Paste this whole file into Extensions > Apps Script of your Google Sheet.
 * Then run setup() once, and deploy as a Web app (Execute as: Me, Who has access: Anyone).
 */

const SLOT_CAPS = {
  starter_veg: 2, starter_nonveg: 2, main_veg: 1, main_nonveg: 1,
  dal: 1, rice: 1, chaat: 1, dessert: 1, assign: 99
};
const RSVP_HEADERS = ["Timestamp", "Name", "Phone", "Coming", "Plus-ones", "Plus-one names", "Food",
  "Drinks", "Potluck slot", "Dish", "Fit colour", "Song request", "Birthday wish", "Golgappa score"];
const GALLERY_HEADERS = ["Timestamp", "By", "Caption", "Drive file ID", "Link"];
const SONG_HEADERS = ["Timestamp", "Song", "Requested by", "Artist", "Preview URL", "Artwork", "Track ID"];
const FOLDER_NAME = "Chhat Wali Diwali — Gossip Archive";

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let r = ss.getSheetByName("RSVPs") || ss.insertSheet("RSVPs");
  if (r.getLastRow() === 0) { r.appendRow(RSVP_HEADERS); r.setFrozenRows(1); r.getRange(1, 1, 1, RSVP_HEADERS.length).setFontWeight("bold"); }
  let g = ss.getSheetByName("Gallery") || ss.insertSheet("Gallery");
  if (g.getLastRow() === 0) { g.appendRow(GALLERY_HEADERS); g.setFrozenRows(1); g.getRange(1, 1, 1, GALLERY_HEADERS.length).setFontWeight("bold"); }
  sheet_("Songs", SONG_HEADERS);
  getFolder_();
  const s1 = ss.getSheetByName("Sheet1"); if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(s1);
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  if (sh.getLastColumn() < headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  return sh;
}

function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("FOLDER_ID");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty("FOLDER_ID", f.getId());
  return f;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function readRsvps_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RSVPs");
  const v = sh.getDataRange().getValues();
  v.shift();
  return v.map((row, i) => ({
    row: i + 2, name: String(row[1]), phone: String(row[2]), coming: String(row[3]), plus: Number(row[4]) || 0,
    slot: String(row[8]), dish: String(row[9]), fit: String(row[10]), song: String(row[11])
  }));
}

function state_() {
  const rs = readRsvps_();
  const st = { slots: {}, fit: {}, fitNames: {}, guests: [], dishes: [], songs: [], gallery: [], counts: { yes: 0, maybe: 0, no: 0, heads: 0 } };
  rs.forEach(r => {
    st.counts[r.coming] = (st.counts[r.coming] || 0) + 1;
    if (r.coming === "yes") {
      const first = r.name.trim().split(/\s+/)[0];
      st.counts.heads += 1 + r.plus;
      st.guests.push(first);
      if (r.dish) st.dishes.push(r.dish);
      if (r.fit) (st.fitNames[r.fit] = st.fitNames[r.fit] || []).push(first);
      if (r.slot) st.slots[r.slot] = (st.slots[r.slot] || 0) + 1;
      if (r.fit) st.fit[r.fit] = (st.fit[r.fit] || 0) + 1;
    }
    if (r.song) st.songs.push({ song: r.song, by: r.name.split(" ")[0] });
  });
  const sv = sheet_("Songs", SONG_HEADERS).getDataRange().getValues();
  sv.shift();
  sv.forEach(row => { if (row[1]) st.songs.push({ song: String(row[1]), by: String(row[2] || ""), artist: String(row[3] || ""), preview: String(row[4] || ""), artwork: String(row[5] || ""), trackId: String(row[6] || "") }); });
  const g = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gallery").getDataRange().getValues();
  g.shift();
  st.gallery = g.map(row => ({ by: String(row[1]), caption: String(row[2]), id: String(row[3]) }));
  return st;
}

function doGet(e) {
  return json_({ ok: true, state: state_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.type === "rsvp") return json_(saveRsvp_(d));
    if (d.type === "photo") return json_(savePhoto_(d));
    if (d.type === "song") return json_(saveSong_(d));
    return json_({ ok: false, error: "unknown_type" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function clean_(s, n) { return String(s || "").replace(/^[=+\-@]/, "'$&").slice(0, n || 300); }

function saveRsvp_(d) {
  const name = clean_(d.name, 80).trim();
  if (!name) return { ok: false, error: "name_required" };
  const rawPhone = String(d.phone || "").replace(/[^\d+]/g, "").slice(0, 16);
  const phone = rawPhone ? "'" + rawPhone : "";
  const rs = readRsvps_();
  // same person = same name (case-insensitive) and same phone (or both blank)
  const mine = rs.find(r => r.name.toLowerCase() === name.toLowerCase() && r.phone.replace(/[^\d+]/g, "") === rawPhone);
  const coming = ["yes", "maybe", "no"].indexOf(d.coming) >= 0 ? d.coming : "maybe";
  const slot = coming === "no" ? "" : (SLOT_CAPS[d.slot] ? d.slot : "");
  if (slot && coming === "yes") {
    const used = rs.filter(r => r !== mine && r.coming === "yes" && r.slot === slot).length;
    if (used >= SLOT_CAPS[slot]) return { ok: false, error: "slot_full" };
  }
  const row = [new Date(), name, phone, coming, Math.min(5, Number(d.plus) || 0), clean_(d.plusNames, 200), clean_(d.food, 20),
    (d.drinks || []).map(x => clean_(x, 20)).join(", "), slot, clean_(d.dish, 120), clean_(d.fit, 20),
    clean_(d.song, 120), clean_(d.wish, 600), Number(d.golgappa) || 0];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RSVPs");
  if (mine) sh.getRange(mine.row, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return { ok: true };
}

function saveSong_(d) {
  const song = clean_(d.song, 100).trim();
  if (!song) return { ok: false, error: "song_required" };
  const url = s => /^https:\/\/[\w.-]+\.(apple|mzstatic)\.com\//.test(String(s || "")) ? String(s).slice(0, 400) : "";
  sheet_("Songs", SONG_HEADERS).appendRow([new Date(), song, clean_(d.by, 40).trim(), clean_(d.artist, 100), url(d.preview), url(d.artwork), clean_(d.trackId, 20)]);
  return { ok: true };
}

function savePhoto_(d) {
  const m = String(d.image || "").match(/^data:(image\/(jpeg|png|webp));base64,(.+)$/);
  if (!m) return { ok: false, error: "bad_image" };
  const bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > 4 * 1024 * 1024) return { ok: false, error: "too_big" };
  const by = clean_(d.by, 60) || "Anonymous";
  const blob = Utilities.newBlob(bytes, m[1], "diwali_" + Date.now() + "_" + by.replace(/[^\w]/g, "") + ".jpg");
  const file = getFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gallery")
    .appendRow([new Date(), by, clean_(d.caption, 80), file.getId(), file.getUrl()]);
  return { ok: true };
}
