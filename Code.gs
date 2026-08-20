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

// ---- Weekly progress email ----

function masteryLabel(v) {
  if (v < 1.35) return 'Beginner';
  if (v < 1.85) return 'Developing';
  if (v < 2.4)  return 'Confident';
  return 'Strong';
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
  var review = (q.missed || []).length;

  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var weekAgo = new Date(now.getTime() - 7 * 86400000);
  var weekHist = history.filter(function (h) { return h && h.date && new Date(h.date) >= weekAgo; });
  var totalQuizzes = history.length;
  var quizzesThisWeek = weekHist.length;
  var wc = 0, wt = 0;
  weekHist.forEach(function (h) { wc += (h.correct || 0); wt += (h.total || 0); });
  var weekPct = wt ? Math.round(wc / wt * 100) : null;

  var subs = [['bio','Biology'],['chem','Chemistry'],['phys','Physics'],['vet','Vet & Animal'],['math','Vet Maths']];
  var weakest = null, weakAb = 99, masteryRows = '';
  subs.forEach(function (s) {
    var v = typeof ability[s[0]] === 'number' ? ability[s[0]] : 1.5;
    if (v < weakAb) { weakAb = v; weakest = s[1]; }
    var pct = Math.max(6, Math.min(100, Math.round((v - 1) / 2 * 100)));
    masteryRows +=
      '<tr>' +
      '<td style="padding:6px 8px;font-size:14px;color:#1c2b28;width:38%">' + s[1] + '</td>' +
      '<td style="padding:6px 8px;width:44%"><div style="background:#e4e1d6;border-radius:6px;height:10px"><div style="background:#0f5f57;height:10px;border-radius:6px;width:' + pct + '%"></div></div></td>' +
      '<td style="padding:6px 8px;font-size:12px;color:#5a6b66;text-align:right;width:18%">' + masteryLabel(v) + '</td>' +
      '</tr>';
  });

  var msDone = 0; if (d.ms) { for (var k in d.ms) { if (d.ms[k]) msDone++; } }
  var exp = d.exp || [];
  var expHours = 0; exp.forEach(function (e) { expHours += (parseFloat(e.hours) || 0); });

  var dateStr = Utilities.formatDate(now, tz, 'd MMMM yyyy');
  var weekLine = quizzesThisWeek
    ? (quizzesThisWeek + ' quiz' + (quizzesThisWeek > 1 ? 'zes' : '') + ' this week' + (weekPct != null ? ' at ' + weekPct + '% average' : ''))
    : 'No quizzes yet this week — jump back in!';

  var html =
  '<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:600px;margin:0 auto;background:#f5f3ec;padding:0 0 20px">' +
    '<div style="background:#0a463f;color:#fff;padding:22px 20px">' +
      '<div style="font-size:20px;font-weight:700">Paige’s Road to Vet</div>' +
      '<div style="font-size:13px;opacity:.85;margin-top:2px">Weekly progress · ' + dateStr + '</div>' +
    '</div>' +
    '<div style="padding:16px 20px">' +
      '<div style="background:#fff;border:1px solid #e4e1d6;border-radius:14px;padding:16px;margin-bottom:14px">' +
        '<div style="font-size:15px;font-weight:700;color:#1c2b28;margin-bottom:8px">This week</div>' +
        '<table style="width:100%;border-collapse:collapse"><tr>' +
          '<td style="text-align:center;padding:8px"><div style="font-size:26px;font-weight:800;color:#0f5f57">' + streak + '</div><div style="font-size:11px;color:#5a6b66">DAY STREAK</div></td>' +
          '<td style="text-align:center;padding:8px"><div style="font-size:26px;font-weight:800;color:#0f5f57">' + quizzesThisWeek + '</div><div style="font-size:11px;color:#5a6b66">QUIZZES</div></td>' +
          '<td style="text-align:center;padding:8px"><div style="font-size:26px;font-weight:800;color:#0f5f57">' + (weekPct != null ? weekPct + '%' : '–') + '</div><div style="font-size:11px;color:#5a6b66">AVG SCORE</div></td>' +
        '</tr></table>' +
        '<div style="font-size:13px;color:#5a6b66;text-align:center;margin-top:6px">' + weekLine + '</div>' +
      '</div>' +
      '<div style="background:#fff;border:1px solid #e4e1d6;border-radius:14px;padding:16px;margin-bottom:14px">' +
        '<div style="font-size:15px;font-weight:700;color:#1c2b28;margin-bottom:8px">Mastery by subject</div>' +
        '<table style="width:100%;border-collapse:collapse">' + masteryRows + '</table>' +
        '<div style="background:#e2efeb;border-radius:10px;padding:10px 12px;margin-top:12px;font-size:13px;color:#0a463f">' +
          '<b>Focus this week:</b> ' + (weakest || 'keep it balanced') + ' — your lowest subject right now. New questions here get added automatically.</div>' +
      '</div>' +
      '<div style="background:#fff;border:1px solid #e4e1d6;border-radius:14px;padding:16px;margin-bottom:14px">' +
        '<table style="width:100%;border-collapse:collapse"><tr>' +
          '<td style="text-align:center;padding:8px"><div style="font-size:22px;font-weight:800;color:#0f5f57">' + msDone + '/26</div><div style="font-size:11px;color:#5a6b66">MILESTONES</div></td>' +
          '<td style="text-align:center;padding:8px"><div style="font-size:22px;font-weight:800;color:#0f5f57">' + totalQuizzes + '</div><div style="font-size:11px;color:#5a6b66">TOTAL QUIZZES</div></td>' +
          '<td style="text-align:center;padding:8px"><div style="font-size:22px;font-weight:800;color:#0f5f57">' + (expHours % 1 === 0 ? expHours : expHours.toFixed(1)) + '</div><div style="font-size:11px;color:#5a6b66">EXPERIENCE HRS</div></td>' +
        '</tr></table>' +
      '</div>' +
      '<div style="text-align:center;font-size:13px;color:#5a6b66;padding:4px 10px">Every quiz and every hour on the farm is a step towards vet school. Keep going, Paige 🎓</div>' +
    '</div>' +
  '</div>';

  MailApp.sendEmail({
    to: REPORT_TO,
    subject: 'Paige’s Road to Vet — weekly progress (' + streak + '-day streak)',
    htmlBody: html
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
