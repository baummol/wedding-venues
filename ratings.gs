/** @OnlyCurrentDoc */
// Backend for Molly & Zack's Venues: stores ratings, comments and ticked-off to-dos.
// Paste this whole file into Extensions > Apps Script in your Google Sheet, then deploy as a web app.

function tab_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); }
  return sh;
}
const ratings_ = () => tab_("Ratings", ["venue", "name", "stars", "comment", "updated"]);
const todos_ = () => tab_("Todos", ["id", "done", "updated"]);
const comments_ = () => tab_("Comments", ["venue", "name", "text", "created"]);

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// GET: every rating, plus the ids of to-dos that are ticked off
function doGet() {
  const ratings = ratings_().getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1])
    .map(r => ({ venue: r[0], name: r[1], stars: Number(r[2]), comment: r[3] || "" }));
  const done = todos_().getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1] === true).map(r => String(r[0]));
  const comments = comments_().getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1] && r[2])
    .map(r => ({ venue: r[0], name: r[1], text: r[2], created: r[3] ? new Date(r[3]).toISOString() : "" }));
  return json_({ ratings: ratings, done: done, comments: comments });
}

// POST: save one rating (replacing that person's earlier one), add one comment, or tick/untick one to-do
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.kind === "todo") {
      const id = String(d.id || "").slice(0, 200);
      if (!id) return json_({ ok: false, error: "missing id" });
      const sh = todos_(), values = sh.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === id) { sh.getRange(i + 1, 1, 1, 3).setValues([[id, !!d.done, new Date()]]); return json_({ ok: true }); }
      }
      sh.appendRow([id, !!d.done, new Date()]);
      return json_({ ok: true });
    }
    if (d.kind === "comment") {
      const cv = String(d.venue || "").slice(0, 80), cn = String(d.name || "").trim().slice(0, 40), ct = String(d.text || "").trim().slice(0, 1000);
      if (!cv || !cn || !ct) return json_({ ok: false, error: "missing venue, name or text" });
      comments_().appendRow([cv, cn, ct, new Date()]);
      return json_({ ok: true });
    }
    const venue = String(d.venue || "").slice(0, 80);
    const name = String(d.name || "").trim().slice(0, 40);
    const stars = Math.max(1, Math.min(5, Math.round(Number(d.stars) || 0)));
    const comment = String(d.comment || "").slice(0, 500);
    if (!venue || !name) return json_({ ok: false, error: "missing venue or name" });
    const sh = ratings_(), values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === venue && String(values[i][1]).trim().toLowerCase() === name.toLowerCase()) {
        sh.getRange(i + 1, 1, 1, 5).setValues([[venue, name, stars, comment, new Date()]]);
        return json_({ ok: true, updated: true });
      }
    }
    sh.appendRow([venue, name, stars, comment, new Date()]);
    return json_({ ok: true, added: true });
  } finally {
    lock.releaseLock();
  }
}
