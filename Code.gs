/**
 * Road to Vet — sync + coaching backend (Google Apps Script)
 * Paired with paige-vet-pathway/index.html
 *
 * Sheet tabs (create these exactly):
 *   Questions  — header row: id | subj | lvl | q | o1 | o2 | o3 | o4 | answer | ex | active
 *                answer = 1..4 (which option is correct). active = TRUE to serve it (blank/FALSE hides it).
 *                subj must be one of: bio chem phys vet math.  lvl = 1, 2 or 3.
 *   State      — header row: id | updatedAt | json
 *                one row per person (id "paige"); stores her pushed progress blob.
 *   Pending    — same header as Questions. The weekly loop writes DRAFT questions here.
 *                Nothing in Pending is ever served to the app — review, then copy good rows
 *                into Questions with active = TRUE.
 *
 * Deploy: Extensions > Apps Script > paste this > Deploy > New deployment >
 *   type "Web app" > Execute as "Me" > Who has access "Anyone" > Deploy.
 *   Copy the /exec URL into the app's Settings (the gear icon).
 */

var SUBJECTS = {bio:1, chem:1, phys:1, vet:1, math:1};

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'questions';
  if (action === 'state') return json(readState((e.parameter.id || 'paige')));
  return json(readQuestions());
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ok:false, error:'bad json'}); }
  if (body.type === 'state') return json(writeState(body));
  if (body.type === 'pending') return json(writePending(body.rows || []));
  if (body.type === 'publish') return json(writeQuestions(body.rows || []));
  return json({ok:false, error:'unknown type'});
}

function sheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === 'Questions' || name === 'Pending') {
      sh.appendRow(['id','subj','lvl','q','o1','o2','o3','o4','answer','ex','active']);
    } else if (name === 'State') {
      sh.appendRow(['id','updatedAt','json']);
    }
  }
  return sh;
}

function readQuestions() {
  var sh = sheet('Questions');
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0];
  var col = {};
  head.forEach(function(h, i) { col[String(h).trim().toLowerCase()] = i; });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var id = row[col.id];
    if (id === '' || id == null) continue;
    var active = row[col.active];
    if (active === false || String(active).toLowerCase() === 'false' || String(active).toLowerCase() === 'no') continue;
    var subj = String(row[col.subj]).trim();
    if (!SUBJECTS[subj]) continue;
    var opts = [row[col.o1], row[col.o2], row[col.o3], row[col.o4]]
                 .filter(function(o){ return o !== '' && o != null; })
                 .map(function(o){ return String(o); });
    if (opts.length < 2) continue;
    var ans = parseInt(row[col.answer], 10);
    if (!(ans >= 1 && ans <= opts.length)) continue;
    out.push({
      id: String(id),
      subj: subj,
      lvl: parseInt(row[col.lvl], 10) || 1,
      q: String(row[col.q]),
      o: opts,
      a: ans - 1,               // app expects 0-based
      ex: String(row[col.ex] || '')
    });
  }
  return out;
}

function readState(id) {
  var sh = sheet('State');
  var values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(id)) {
      var data = {};
      try { data = JSON.parse(values[r][2]); } catch (err) {}
      return { id: id, updatedAt: Number(values[r][1]) || 0, data: data };
    }
  }
  return { id: id, updatedAt: 0, data: null };
}

// True if a state blob carries real progress (any quiz history/streak, milestone, experience, why).
function stateHasProgress(d) {
  try {
    if (d && d.quiz && d.quiz.history && d.quiz.history.length) return true;
    if (d && d.quiz && d.quiz.streak && d.quiz.streak.count) return true;
    if (d && d.ms) { for (var k in d.ms) { if (d.ms[k]) return true; } }
    if (d && d.exp && d.exp.length) return true;
    if (d && d.why) return true;
    if (d && d.learn) { for (var k2 in d.learn) { if (d.learn[k2]) return true; } }
  } catch (e) {}
  return false;
}

function writeState(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet('State');
    var id = String(body.id || 'paige');
    var ts = Number(body.updatedAt) || Date.now();
    var jsonStr = JSON.stringify(body.data || {});
    var values = sh.getDataRange().getValues();
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][0]) === id) {
        var existingData = {};
        try { existingData = JSON.parse(values[r][2]); } catch (e) {}
        // never let a blank/empty device overwrite a record that already has real progress
        if (stateHasProgress(existingData) && !stateHasProgress(body.data)) {
          return { ok: true, note: 'blank push ignored, existing progress protected' };
        }
        if ((Number(values[r][1]) || 0) > ts) return { ok: true, note: 'stale, ignored' };
        sh.getRange(r + 1, 2).setValue(ts);
        sh.getRange(r + 1, 3).setValue(jsonStr);
        return { ok: true };
      }
    }
    sh.appendRow([id, ts, jsonStr]);
    return { ok: true, created: true };
  } finally {
    lock.releaseLock();
  }
}

// Publish verified questions straight to the LIVE Questions tab (active = TRUE).
// Used by the autonomous weekly job. Rejects malformed rows defensively so a bad
// generation can never write a broken question to Paige.
function writeQuestions(rows) {
  if (!rows.length) return { ok: true, added: 0 };
  var sh = sheet('Questions');
  var existing = {};
  var vals = sh.getDataRange().getValues();
  for (var r = 1; r < vals.length; r++) existing[String(vals[r][0])] = true;
  var added = 0;
  rows.forEach(function(row) {
    var subj = String(row.subj || '').trim();
    if (!SUBJECTS[subj]) return;
    var o = (row.o || []).filter(function(x){ return x !== '' && x != null; }).map(String);
    if (o.length < 2) return;
    var ans = row.answer != null ? parseInt(row.answer, 10) : (row.a != null ? row.a + 1 : 0);
    if (!(ans >= 1 && ans <= o.length)) return;
    if (!row.q || !row.ex) return;
    var id = String(row.id || ('gen_' + subj + '_' + Date.now() + '_' + added));
    if (existing[id]) return;
    existing[id] = true;
    sh.appendRow([ id, subj, parseInt(row.lvl,10)||1, String(row.q),
      o[0]||'', o[1]||'', o[2]||'', o[3]||'', ans, String(row.ex), true ]);
    added++;
  });
  return { ok: true, added: added };
}

function writePending(rows) {
  if (!rows.length) return { ok: true, added: 0 };
  var sh = sheet('Pending');
  var stamp = new Date();
  rows.forEach(function(r) {
    var o = r.o || [];
    sh.appendRow([
      r.id || ('gen_' + stamp.getTime() + '_' + Math.floor(Math.random()*1000)),
      r.subj || '', r.lvl || 1, r.q || '',
      o[0]||'', o[1]||'', o[2]||'', o[3]||'',
      r.answer || (r.a != null ? r.a + 1 : ''),
      r.ex || '', ''    // active blank — never served until reviewed
    ]);
  });
  return { ok: true, added: rows.length };
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
