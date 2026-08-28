/**
 * Backend for the "المحبة" club app, backed by this spreadsheet.
 *
 * SETUP:
 * 1. Create a new blank Google Sheet.
 * 2. Extensions > Apps Script. Delete any starter code, paste this whole file in.
 * 3. (Optional but recommended since multiple leaders can edit) Project Settings >
 *    Script Properties > add property PASSWORD with a value everyone will share,
 *    e.g. "club2026". If you skip this, anyone with the link can edit data.
 * 4. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    Deploy, authorize when prompted, and copy the Web App URL.
 * 5. Paste that URL into API_URL at the top of index.html's <script>.
 *
 * The four sheets (Teams / Children / Sessions / Scores) are created
 * automatically the first time the web app is called — you don't need to
 * make them by hand.
 */

function ensureSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = {
    Teams: ['id', 'name'],
    Children: ['id', 'name', 'teamId'],
    Sessions: ['id', 'date', 'label'],
    Scores: ['sessionId', 'childId', 'mass', 'sunday', 'participation', 'behavior', 'attendance']
  };
  Object.keys(defs).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(defs[name]);
    } else {
      // migration: add any newly-introduced columns to existing sheets without touching old rows
      var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      defs[name].forEach(function (col) {
        if (headers.indexOf(col) === -1) {
          sh.getRange(1, sh.getLastColumn() + 1).setValue(col);
          headers.push(col);
        }
      });
    }
  });
  var def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(def);
  }
}

function sheetAsObjects_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  return values.slice(1)
    .filter(function (r) { return r.some(function (c) { return c !== ''; }); })
    .map(function (r) {
      var o = {};
      headers.forEach(function (h, i) { o[h] = r[i]; });
      return o;
    });
}

function getAllData() {
  ensureSheets();
  var teams = sheetAsObjects_('Teams');
  var children = sheetAsObjects_('Children');
  var sessions = sheetAsObjects_('Sessions');
  var scoresRaw = sheetAsObjects_('Scores');
  var scores = {};
  scoresRaw.forEach(function (s) {
    var k = s.sessionId + '__' + s.childId;
    scores[k] = {
      mass: Number(s.mass) || 0,
      sunday: Number(s.sunday) || 0,
      participation: Number(s.participation) || 0,
      behavior: Number(s.behavior) || 0,
      attendance: Number(s.attendance) || 0
    };
  });
  return { teams: teams, children: children, sessions: sessions, scores: scores };
}

function checkPassword_(pw) {
  var real = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  if (!real) return true;
  return pw === real;
}

function addRow_(sheetName, obj) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var headers = sh.getDataRange().getValues()[0];
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sh.appendRow(row);
}

function removeRowsWhere_(sheetName, predicate) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  for (var i = values.length - 1; i >= 1; i--) {
    var o = {};
    headers.forEach(function (h, idx) { o[h] = values[i][idx]; });
    if (predicate(o)) sh.deleteRow(i + 1);
  }
}

function removeTeamCascade_(teamId) {
  var childrenSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Children');
  var values = childrenSh.getDataRange().getValues();
  var headers = values[0];
  var childIds = [];
  for (var i = 1; i < values.length; i++) {
    var o = {};
    headers.forEach(function (h, idx) { o[h] = values[i][idx]; });
    if (o.teamId === teamId) childIds.push(o.id);
  }
  removeRowsWhere_('Teams', function (r) { return r.id === teamId; });
  removeRowsWhere_('Children', function (r) { return r.teamId === teamId; });
  childIds.forEach(function (cid) {
    removeRowsWhere_('Scores', function (r) { return r.childId === cid; });
  });
}

function moveChild_(childId, teamId) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Children');
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var idIdx = headers.indexOf('id');
  var teamIdx = headers.indexOf('teamId');
  for (var i = 1; i < values.length; i++) {
    if (values[i][idIdx] === childId) {
      sh.getRange(i + 1, teamIdx + 1).setValue(teamId);
      break;
    }
  }
}

function setScoreRow_(sessionId, childId, cat, val) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Scores');
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var catIdx = headers.indexOf(cat);
  var rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === sessionId && values[i][1] === childId) { rowIndex = i; break; }
  }
  if (rowIndex === -1) {
    var row = headers.map(function (h) {
      if (h === 'sessionId') return sessionId;
      if (h === 'childId') return childId;
      if (h === cat) return val;
      return 0;
    });
    sh.appendRow(row);
  } else {
    sh.getRange(rowIndex + 1, catIdx + 1).setValue(val);
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  ensureSheets();
  return jsonOut_({ ok: true, data: getAllData() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureSheets();
    var body;
    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      return jsonOut_({ ok: false, error: 'bad_request' });
    }
    if (!checkPassword_(body.password)) {
      return jsonOut_({ ok: false, error: 'wrong_password' });
    }
    var p = body.payload || {};
    switch (body.action) {
      case 'ping':
        break;
      case 'addTeam':
        addRow_('Teams', { id: p.id, name: p.name });
        break;
      case 'removeTeam':
        removeTeamCascade_(p.id);
        break;
      case 'addChild':
        addRow_('Children', { id: p.id, name: p.name, teamId: p.teamId });
        break;
      case 'removeChild':
        removeRowsWhere_('Children', function (r) { return r.id === p.id; });
        removeRowsWhere_('Scores', function (r) { return r.childId === p.id; });
        break;
      case 'moveChild':
        moveChild_(p.id, p.teamId);
        break;
      case 'addSession':
        addRow_('Sessions', { id: p.id, date: p.date, label: p.label || '' });
        break;
      case 'removeSession':
        removeRowsWhere_('Sessions', function (r) { return r.id === p.id; });
        removeRowsWhere_('Scores', function (r) { return r.sessionId === p.id; });
        break;
      case 'setScore':
        setScoreRow_(p.sessionId, p.childId, p.cat, p.val);
        break;
      default:
        return jsonOut_({ ok: false, error: 'unknown_action' });
    }
    return jsonOut_({ ok: true, data: getAllData() });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
