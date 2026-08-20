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

// Weekly report recipients + a light key so the send/trigger URLs can't be hit by strangers.
var REPORT_TO = 'paige.linda.graham@gmail.com,brendonemmagraham@gmail.com';
var ADMIN_KEY = 'pfe-roadtovet-2026';

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'questions';
  if (action === 'state') return json(readState((e.parameter.id || 'paige')));
  if (action === 'report') {
    if (e.parameter.key !== ADMIN_KEY) return json({ok:false, error:'bad key'});
    sendWeeklyReport();
    return json({ok:true, sent:true, to:REPORT_TO});
  }
  if (action === 'installtrigger') {
    if (e.parameter.key !== ADMIN_KEY) return json({ok:false, error:'bad key'});
    return json(installWeeklyTrigger());
  }
  if (action === 'deactivate') {
    if (e.parameter.key !== ADMIN_KEY) return json({ok:false, error:'bad key'});
    return json(deactivateQuestion(e.parameter.id));
  }
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
  var toWrite = [];
  rows.forEach(function(row) {
    var subj = String(row.subj || '').trim();
    if (!SUBJECTS[subj]) return;
    var o = (row.o || []).filter(function(x){ return x !== '' && x != null; }).map(String);
    if (o.length < 2) return;
    var ans = row.answer != null ? parseInt(row.answer, 10) : (row.a != null ? row.a + 1 : 0);
    if (!(ans >= 1 && ans <= o.length)) return;
    if (!row.q || !row.ex) return;
    var id = String(row.id || ('gen_' + subj + '_' + Date.now() + '_' + toWrite.length));
    if (existing[id]) return;
    existing[id] = true;
    toWrite.push([ id, subj, String(parseInt(row.lvl,10)||1), String(row.q),
      o[0]||'', o[1]||'', o[2]||'', o[3]||'', String(ans), String(row.ex), 'TRUE' ]);
  });
  if (toWrite.length) {
    // format the target cells as plain text BEFORE writing so Sheets can't coerce
    // fraction-style options like "1/4" into dates, or "1 - 2" into numbers.
    var startRow = sh.getLastRow() + 1;
    var rng = sh.getRange(startRow, 1, toWrite.length, 11);
    rng.setNumberFormat('@');
    rng.setValues(toWrite);
  }
  return { ok: true, added: toWrite.length };
}

// Retire a question by id (set active = FALSE) so the app stops serving it. Used to pull
// a bad question without hunting for the row.
function deactivateQuestion(id) {
  if (!id) return { ok: false, error: 'no id' };
  var sh = sheet('Questions');
  var vals = sh.getDataRange().getValues();
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0]) === String(id)) {
      sh.getRange(r + 1, 11).setValue('FALSE');
      return { ok: true, deactivated: id };
    }
  }
  return { ok: false, error: 'not found' };
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

// ---- Weekly progress email ----

// Ordered pathway milestones (must match the app's ids) — lets the report name what she
// ticked off and what's next.
var MILESTONES = [
  {id:'y11a', t:'Choose your vet-pathway subjects'},
  {id:'y11b', t:'Lock in NCEA Level 1 literacy + numeracy'},
  {id:'y11c', t:"Write down your 'why'"},
  {id:'y11d', t:'Log your first hands-on animal session'},
  {id:'y11e', t:'Aim for a Merit or Excellence endorsement'},
  {id:'y12a', t:'Take Chemistry, Biology, Physics, Maths + English'},
  {id:'y12b', t:'Earn 14+ credits in NCEA Level 2 Physics'},
  {id:'y12c', t:'Earn 14+ credits in NCEA Level 2 Maths'},
  {id:'y12d', t:'Do a day at a small-animal vet clinic'},
  {id:'y12e', t:'Earn a Level 2 Merit/Excellence endorsement'},
  {id:'y13a', t:'Take Level 3 Chemistry + Level 3 Biology'},
  {id:'y13b', t:'Add Level 3 Physics and/or Maths'},
  {id:'y13c', t:'Gain University Entrance'},
  {id:'y13d', t:'Reach 20+ logged experience hours, 2+ settings'},
  {id:'y13e', t:"Apply to Massey's BVSc Pre-Selection"},
  {id:'prea', t:'Enrol in the Vet Pre-Selection semester'},
  {id:'preb', t:'Pass all four prerequisite papers'},
  {id:'prec', t:'Achieve a GPA of 5.0 or higher'},
  {id:'sela', t:'Sit the Casper assessment'},
  {id:'selb', t:'Sit the STAT assessment'},
  {id:'selc', t:'Complete the MMI'},
  {id:'seld', t:'Receive your offer into the professional phase'},
  {id:'profa', t:'Start the 5-year professional phase'},
  {id:'profb', t:'Complete clinical & EMS placements'},
  {id:'profc', t:'Graduate with your BVSc'},
  {id:'rega', t:'Register with the NZ Veterinary Council'}
];

// Rotating, factually-reliable tips (no invented claims).
var TIPS = [
  'Massey wants 14+ credits of NCEA Level 3 Chemistry and Biology — these quizzes build exactly that foundation.',
  "Vet selection isn't only grades: Casper and the MMI test how you think and communicate. Keep talking your reasoning out loud.",
  'Hands-on animal experience is gold for a vet application — a clinic day or a session on the farm both count.',
  'Strong algebra makes senior Chemistry and Physics much easier, so maths practice pays off twice.',
  'Consistency beats cramming — a few quizzes across several days builds mastery that sticks.',
  'Aim for Merit/Excellence at school: the high-marks habit is what a GPA of 5.0+ needs at Massey.'
];

function masteryLabel(v) {
  if (v < 1.35) return 'Beginner';
  if (v < 1.85) return 'Developing';
  if (v < 2.4)  return 'Confident';
  return 'Strong';
}

function subjOfId(id) {
  id = String(id);
  if (id.indexOf('gen_') === 0) { var p = id.split('_'); return p[1] || ''; }
  var m = id.match(/^[a-z]+/);
  return m ? m[0] : '';
}

function metaSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Meta');
  if (!sh) { sh = ss.insertSheet('Meta'); sh.appendRow(['key','json']); }
  return sh;
}
function readMeta(key) {
  var sh = metaSheet(); var v = sh.getDataRange().getValues();
  for (var r = 1; r < v.length; r++) { if (String(v[r][0]) === key) { try { return JSON.parse(v[r][1]); } catch (e) { return null; } } }
  return null;
}
function writeMeta(key, obj) {
  var sh = metaSheet(); var v = sh.getDataRange().getValues(); var js = JSON.stringify(obj);
  for (var r = 1; r < v.length; r++) { if (String(v[r][0]) === key) { sh.getRange(r + 1, 2).setValue(js); return; } }
  sh.appendRow([key, js]);
}

function installWeeklyTrigger() {
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(existing[i]);
  }
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(11).create();
  return { ok: true, installed: 'sendWeeklyReport every Monday 11:00 (script time zone: ' + Session.getScriptTimeZone() + ')' };
}

function sendWeeklyReport() {
  var st = readState('paige');
  var d = (st && st.data) || {};
  var q = d.quiz || {};
  var history = q.history || [];
  var ability = q.ability || {};
  var streak = (q.streak && q.streak.count) || 0;
  var missed = q.missed || [];
  var msNow = d.ms || {};
  var exp = d.exp || [];

  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var d7 = new Date(now.getTime() - 7 * 86400000);
  var d14 = new Date(now.getTime() - 14 * 86400000);
  function pdate(s) { return s ? new Date(s) : null; }

  // quiz activity: this week vs the week before (from dated history)
  var wkThis = history.filter(function (h) { var t = pdate(h && h.date); return t && t >= d7; });
  var wkPrev = history.filter(function (h) { var t = pdate(h && h.date); return t && t < d7 && t >= d14; });
  function acc(arr) { var c = 0, tt = 0; arr.forEach(function (h) { c += (h.correct || 0); tt += (h.total || 0); }); return tt ? Math.round(c / tt * 100) : null; }
  var qThis = wkThis.length, qPrev = wkPrev.length;
  var accThis = acc(wkThis), accPrev = acc(wkPrev);
  var totalQuizzes = history.length;

  // experience: this week + total (from dated entries)
  var expThis = exp.filter(function (e) { var t = pdate(e && e.date); return t && t >= d7; });
  var expHrsTotal = 0; exp.forEach(function (e) { expHrsTotal += (parseFloat(e.hours) || 0); });
  var settings = {}; exp.forEach(function (e) { if (e && e.oppId) settings[e.oppId] = 1; });

  // last week's snapshot for ability + milestone movement
  var snap = readMeta('weeksnap');
  var prevAb = (snap && snap.ability) ? snap.ability : null;
  var prevMs = (snap && snap.ms) ? snap.ms : {};

  var subs = [['bio','Biology'],['chem','Chemistry'],['phys','Physics'],['vet','Vet & Animal'],['math','Vet Maths']];
  var weakest = null, weakAb = 99, strongest = null, strongAb = -1, subjectBlocks = '', improvedList = [];
  subs.forEach(function (s) {
    var v = typeof ability[s[0]] === 'number' ? ability[s[0]] : 1.5;
    if (v < weakAb) { weakAb = v; weakest = s[1]; }
    if (v > strongAb) { strongAb = v; strongest = s[1]; }
    var pv = (prevAb && typeof prevAb[s[0]] === 'number') ? prevAb[s[0]] : null;
    var trendTxt, trendColor, note;
    if (pv == null) { trendTxt = 'baseline'; trendColor = '#5a6b66'; note = 'First read — this is the starting point.'; }
    else {
      var diff = v - pv;
      if (diff >= 0.1) { trendTxt = '▲ improving'; trendColor = '#2e9e6b'; note = 'Nice gains this week — keep it up.'; improvedList.push(s[1]); }
      else if (diff <= -0.1) { trendTxt = '▼ slipped'; trendColor = '#c9503a'; note = 'Dropped a little — worth a couple of quizzes here.'; }
      else { trendTxt = '→ steady'; trendColor = '#5a6b66'; note = (v >= 2.4) ? 'Strong and holding — keep it sharp.' : 'Holding steady — a good area to push next.'; }
    }
    var pct = Math.max(6, Math.min(100, Math.round((v - 1) / 2 * 100)));
    subjectBlocks +=
      '<div style="padding:10px 0;border-bottom:1px solid #eee">' +
        '<table style="width:100%;border-collapse:collapse"><tr>' +
          '<td style="font-size:14px;font-weight:700;color:#1c2b28">' + s[1] + '</td>' +
          '<td style="text-align:right;font-size:12px;color:#5a6b66">' + masteryLabel(v) + ' · <span style="color:' + trendColor + ';font-weight:700">' + trendTxt + '</span></td>' +
        '</tr></table>' +
        '<div style="background:#e4e1d6;border-radius:6px;height:8px;margin:6px 0"><div style="background:#0f5f57;height:8px;border-radius:6px;width:' + pct + '%"></div></div>' +
        '<div style="font-size:12px;color:#5a6b66">' + note + '</div>' +
      '</div>';
  });

  // to review — count missed by subject
  var revCount = {}; missed.forEach(function (id) { var sj = subjOfId(id); if (sj) revCount[sj] = (revCount[sj] || 0) + 1; });
  var revParts = subs.filter(function (s) { return revCount[s[0]]; }).map(function (s) { return revCount[s[0]] + ' ' + s[1]; });

  // milestones
  var doneCount = 0; MILESTONES.forEach(function (m) { if (msNow[m.id]) doneCount++; });
  var newly = MILESTONES.filter(function (m) { return msNow[m.id] && !prevMs[m.id]; });
  var next = null; for (var i = 0; i < MILESTONES.length; i++) { if (!msNow[MILESTONES[i].id]) { next = MILESTONES[i]; break; } }

  var dFrom = Utilities.formatDate(d7, tz, 'd MMM');
  var dTo = Utilities.formatDate(now, tz, 'd MMM');
  function delta(cur, prev, suffix) {
    if (prev == null || cur == null) return '';
    var diff = cur - prev;
    if (diff === 0) return '<div style="font-size:11px;color:#5a6b66">same as last week</div>';
    var up = diff > 0;
    return '<div style="font-size:11px;color:' + (up ? '#2e9e6b' : '#c9503a') + '">' + (up ? '▲ +' : '▼ ') + diff + (suffix || '') + ' vs last week</div>';
  }

  var rec = 'Do 2–3 ' + (weakest || 'mixed') + ' quizzes';
  if (expThis.length === 0) rec += ', and try to log one hands-on session on the farm';
  rec += '.';

  var tip = TIPS[(Math.floor(now.getTime() / 604800000)) % TIPS.length];

  var headline;
  if (qThis === 0) headline = 'a quiet week — a couple of quizzes will get you moving again.';
  else if (accThis != null && accThis >= 80) headline = 'strong week — ' + qThis + ' quiz' + (qThis > 1 ? 'zes' : '') + ' at ' + accThis + '%. That’s vet-school pace.';
  else headline = 'good effort — ' + qThis + ' quiz' + (qThis > 1 ? 'zes' : '') + ' this week' + (accThis != null ? ' at ' + accThis + '%' : '') + '.';

  var expHrsTotalStr = (expHrsTotal % 1 === 0) ? String(expHrsTotal) : expHrsTotal.toFixed(1);
  var nSettings = Object.keys(settings).length;

  var html =
  '<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:600px;margin:0 auto;background:#f5f3ec;padding:0 0 22px">' +
    '<div style="background:#0a463f;color:#fff;padding:22px 20px">' +
      '<div style="font-size:20px;font-weight:700">Paige’s Road to Vet</div>' +
      '<div style="font-size:13px;opacity:.85;margin-top:2px">Weekly review · ' + dFrom + ' – ' + dTo + '</div>' +
    '</div>' +
    '<div style="padding:16px 20px">' +

      '<div style="font-size:15px;color:#1c2b28;line-height:1.5;margin:2px 4px 14px">Hi Paige — ' + headline + '</div>' +

      '<div style="background:#fff;border:1px solid #e4e1d6;border-radius:14px;padding:16px;margin-bottom:14px">' +
        '<div style="font-size:15px;font-weight:700;color:#1c2b28;margin-bottom:10px">This week at a glance</div>' +
        '<table style="width:100%;border-collapse:collapse"><tr>' +
          '<td style="text-align:center;padding:6px;vertical-align:top"><div style="font-size:24px;font-weight:800;color:#0f5f57">' + qThis + '</div><div style="font-size:11px;color:#5a6b66">QUIZZES</div>' + delta(qThis, qPrev, '') + '</td>' +
          '<td style="text-align:center;padding:6px;vertical-align:top"><div style="font-size:24px;font-weight:800;color:#0f5f57">' + (accThis != null ? accThis + '%' : '–') + '</div><div style="font-size:11px;color:#5a6b66">AVG SCORE</div>' + delta(accThis, accPrev, '%') + '</td>' +
          '<td style="text-align:center;padding:6px;vertical-align:top"><div style="font-size:24px;font-weight:800;color:#0f5f57">' + streak + '</div><div style="font-size:11px;color:#5a6b66">DAY STREAK</div></td>' +
          '<td style="text-align:center;padding:6px;vertical-align:top"><div style="font-size:24px;font-weight:800;color:#0f5f57">' + expHrsTotalStr + '</div><div style="font-size:11px;color:#5a6b66">FARM HRS</div></td>' +
        '</tr></table>' +
      '</div>' +

      '<div style="background:#fff;border:1px solid #e4e1d6;border-radius:14px;padding:16px;margin-bottom:14px">' +
        '<div style="font-size:15px;font-weight:700;color:#1c2b28;margin-bottom:2px">How each subject is tracking</div>' +
        '<div style="font-size:12px;color:#5a6b66;margin-bottom:6px">Compared with last week</div>' +
        subjectBlocks +
        '<div style="font-size:12.5px;color:#0a463f;margin-top:10px">Strongest right now: <b>' + (strongest || '–') + '</b>' + (improvedList.length ? ' · Most improved: <b>' + improvedList.join(', ') + '</b>' : '') + '</div>' +
      '</div>' +

      (revParts.length ?
      '<div style="background:#fff;border:1px solid #e4e1d6;border-radius:14px;padding:16px;margin-bottom:14px">' +
        '<div style="font-size:15px;font-weight:700;color:#1c2b28;margin-bottom:6px">To revisit</div>' +
        '<div style="font-size:13px;color:#5a6b66;line-height:1.5">You have <b>' + missed.length + '</b> question' + (missed.length > 1 ? 's' : '') + ' saved to review — ' + revParts.join(', ') + '. They’ll keep coming back in your quizzes until you’ve got them.</div>' +
      '</div>' : '') +

      '<div style="background:#fff;border:1px solid #e4e1d6;border-radius:14px;padding:16px;margin-bottom:14px">' +
        '<div style="font-size:15px;font-weight:700;color:#1c2b28;margin-bottom:6px">Your pathway</div>' +
        '<div style="font-size:13px;color:#1c2b28"><b>' + doneCount + ' of 26</b> milestones done.</div>' +
        (newly.length ? '<div style="font-size:13px;color:#2e9e6b;margin-top:6px">🎉 Ticked off this week: ' + newly.map(function (m) { return m.t; }).join('; ') + '</div>' : '') +
        (next ? '<div style="font-size:13px;color:#5a6b66;margin-top:6px">Next up: <b style="color:#0a463f">' + next.t + '</b></div>' : '<div style="font-size:13px;color:#2e9e6b;margin-top:6px">Every milestone done — incredible.</div>') +
      '</div>' +

      '<div style="background:#fff;border:1px solid #e4e1d6;border-radius:14px;padding:16px;margin-bottom:14px">' +
        '<div style="font-size:15px;font-weight:700;color:#1c2b28;margin-bottom:6px">Hands-on experience</div>' +
        '<div style="font-size:13px;color:#1c2b28">' + expHrsTotalStr + ' hours logged in total across ' + nSettings + ' setting' + (nSettings === 1 ? '' : 's') + '.</div>' +
        '<div style="font-size:13px;color:' + (expThis.length ? '#2e9e6b' : '#c9503a') + ';margin-top:6px">' + (expThis.length ? (expThis.length + ' session' + (expThis.length > 1 ? 's' : '') + ' this week — brilliant.') : 'Nothing logged this week — even one session on the farm or at a clinic is real vet-school currency.') + '</div>' +
      '</div>' +

      '<div style="background:#e2efeb;border-radius:14px;padding:16px;margin-bottom:14px">' +
        '<div style="font-size:15px;font-weight:700;color:#0a463f;margin-bottom:6px">Focus for next week</div>' +
        '<div style="font-size:14px;color:#0a463f;line-height:1.5">' + rec + ' Fresh questions targeting <b>' + (weakest || 'your weak areas') + '</b> are already loaded.</div>' +
      '</div>' +

      '<div style="background:#f7edd6;border-radius:14px;padding:13px 15px;margin-bottom:12px;font-size:13px;color:#8a6d1f;line-height:1.5"><b>Worth knowing:</b> ' + tip + '</div>' +

      '<div style="text-align:center;font-size:13px;color:#5a6b66;padding:4px 10px">You’re building this one week at a time, Paige. Keep going. 🎓</div>' +
    '</div>' +
  '</div>';

  MailApp.sendEmail({
    to: REPORT_TO,
    subject: 'Paige’s week in review — ' + qThis + ' quiz' + (qThis === 1 ? '' : 'zes') + (accThis != null ? ' at ' + accThis + '%' : '') + ', focus: ' + (weakest || 'balanced'),
    htmlBody: html
  });

  // store this week's snapshot so next week can compare
  writeMeta('weeksnap', { date: Utilities.formatDate(now, tz, 'yyyy-MM-dd'), ability: ability, ms: msNow });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
