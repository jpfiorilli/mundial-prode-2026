/**
 * PRODE MUNDIAL 2026 — Google Apps Script Backend
 */

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SECRET_TOKEN = 'CDh5gAB_KCmUfq8vbHA63pd1FVCYNMW9hI9FwV_vWV4';

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

/* ── CORS output helper ── */
function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── GET handler — serves iframe bridge page ── */
function doGet(e) {
  const bridge = e.parameter.bridge;
  const token  = e.parameter.token;

  /* Serve the bridge iframe page */
  if(bridge === '1') {
    const scriptToken = SECRET_TOKEN;
    const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body>
<script>
const SECRET = '${scriptToken}';
window.addEventListener('message', function(e) {
  var d = e.data;
  if(!d || d.token !== SECRET) return;
  var reqId = d.reqId;
  var action = d.action;

  function reply(result) {
    e.source.postMessage({reqId: reqId, result: result}, '*');
  }

  var url = '${SHEET_URL}';
  fetch(url, {
    method: 'POST',
    body: JSON.stringify(d),
    headers: {'Content-Type': 'text/plain'}
  })
  .then(function(r){ return r.json(); })
  .then(function(data){ reply(data); })
  .catch(function(err){ reply({error: err.message}); });
});
<\/script>
</body></html>`;
    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutput('<p>Prode 2026 API</p>');
}

/* ── POST handler ── */
function doPost(e) {
  try {
    const body  = JSON.parse(e.postData.contents);
    if(body.token !== SECRET_TOKEN) return jsonOut({error:'Unauthorised'});

    const action = body.action;
    if      (action === 'getAll')      return jsonOut(getAll());
    else if (action === 'savePlayer')  return jsonOut(savePlayer(body.player));
    else if (action === 'savePred')    return jsonOut(savePred(body.playerId, body.matchId, body.ph, body.pa));
    else if (action === 'saveResult')  return jsonOut(saveResult(body.matchId, body.rh, body.ra));
    else return jsonOut({error: 'Unknown action: ' + action});
  } catch(err) {
    return jsonOut({error: err.message});
  }
}

/* ── GET ALL ── */
function getAll() {
  return {
    players: getPlayers(),
    preds:   getPredictions(),
    results: getResults()
  };
}

/* ── PLAYERS ── */
function getPlayers() {
  const sheet = getSheet('players');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1)
    .filter(r => r[0] && String(r[0]) !== 'id')
    .map(r => ({id: String(r[0]), name: String(r[1]), email: String(r[2])}));
}

function savePlayer(player) {
  const sheet = getSheet('players');
  if (sheet.getLastRow() === 0)
    sheet.appendRow(['id','name','email','created_at']);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === player.id || rows[i][2] === player.email)
      return {ok: true, existing: true};
  }
  sheet.appendRow([player.id, player.name, player.email, new Date().toISOString()]);
  return {ok: true};
}

/* ── PREDICTIONS ── */
function getPredictions() {
  const sheet = getSheet('predictions');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return {};
  const out = {};
  rows.slice(1).forEach(r => {
    const playerId = String(r[0]);
    const matchId  = String(r[1]);
    if (!playerId || playerId === 'player_id') return;
    if (!out[playerId]) out[playerId] = {};
    out[playerId][matchId] = {ph: +r[2], pa: +r[3]};
  });
  return out;
}

function savePred(playerId, matchId, ph, pa) {
  const sheet = getSheet('predictions');
  if (sheet.getLastRow() === 0)
    sheet.appendRow(['player_id','match_id','pred_home','pred_away','saved_at']);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(playerId) && String(rows[i][1]) === String(matchId)) {
      sheet.getRange(i+1, 3, 1, 3).setValues([[ph, pa, new Date().toISOString()]]);
      return {ok: true, updated: true};
    }
  }
  sheet.appendRow([playerId, matchId, ph, pa, new Date().toISOString()]);
  return {ok: true};
}

/* ── RESULTS ── */
function getResults() {
  const sheet = getSheet('results');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return {};
  const out = {};
  rows.slice(1).forEach(r => {
    const matchId = String(r[0]);
    if (!matchId || matchId === 'match_id') return;
    out[matchId] = {rh: +r[1], ra: +r[2]};
  });
  return out;
}

function saveResult(matchId, rh, ra) {
  const sheet = getSheet('results');
  if (sheet.getLastRow() === 0)
    sheet.appendRow(['match_id','result_home','result_away','saved_at']);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(matchId)) {
      sheet.getRange(i+1, 2, 1, 3).setValues([[rh, ra, new Date().toISOString()]]);
      return {ok: true, updated: true};
    }
  }
  sheet.appendRow([matchId, rh, ra, new Date().toISOString()]);
  return {ok: true};
}


function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

/* ── CORS headers ── */
function setCORS(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'POST, GET')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function doOptions(e) {
  return setCORS(ContentService.createTextOutput(''));
}

/* ── Main entry point ── */
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);

    /* ── Token check — reject unauthorised requests ── */
    if(body.token !== SECRET_TOKEN){
      return setCORS(ContentService
        .createTextOutput(JSON.stringify({error:'Unauthorised'}))
        .setMimeType(ContentService.MimeType.JSON));
    }

    const action = body.action;
    let result;

    if      (action === 'getAll')      result = getAll();
    else if (action === 'savePlayer')  result = savePlayer(body.player);
    else if (action === 'savePred')    result = savePred(body.playerId, body.matchId, body.ph, body.pa);
    else if (action === 'saveResult')  result = saveResult(body.matchId, body.rh, body.ra);
    else result = {error: 'Unknown action: ' + action};

    return setCORS(ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON));

  } catch(err) {
    return setCORS(ContentService
      .createTextOutput(JSON.stringify({error: err.message}))
      .setMimeType(ContentService.MimeType.JSON));
  }
}

/* ── GET ALL: returns players, preds, results in one call ── */
function getAll() {
  const players     = getPlayers();
  const predictions = getPredictions();
  const results     = getResults();
  return { players, preds: predictions, results };
}

/* ── PLAYERS ── */
function getPlayers() {
  const sheet = getSheet('players');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1)
    .filter(r => r[0] && String(r[0]) !== 'id')
    .map(r => ({
      id:    String(r[0]),
      name:  String(r[1]),
      email: String(r[2]),
    }));
}

function savePlayer(player) {
  const sheet = getSheet('players');
  // Ensure header row exists
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'name', 'email', 'created_at']);
  }
  // Check if player already exists (by id or email)
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === player.id || rows[i][2] === player.email) {
      return {ok: true, existing: true};   // already registered
    }
  }
  sheet.appendRow([player.id, player.name, player.email, new Date().toISOString()]);
  return {ok: true};
}

/* ── PREDICTIONS ── */
function getPredictions() {
  const sheet = getSheet('predictions');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return {};
  const out = {};
  rows.slice(1).forEach(r => {
    const playerId = String(r[0]);
    const matchId  = String(r[1]);   // force string key — JS object keys are always strings
    const ph = r[2], pa = r[3];
    if (!playerId || playerId === 'player_id') return;
    if (!out[playerId]) out[playerId] = {};
    out[playerId][matchId] = {ph: +ph, pa: +pa};
  });
  return out;
}

function savePred(playerId, matchId, ph, pa) {
  const sheet = getSheet('predictions');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['player_id', 'match_id', 'pred_home', 'pred_away', 'saved_at']);
  }
  // Upsert: check if row exists
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(playerId) && String(rows[i][1]) === String(matchId)) {
      sheet.getRange(i + 1, 3, 1, 3).setValues([[ph, pa, new Date().toISOString()]]);
      return {ok: true, updated: true};
    }
  }
  sheet.appendRow([playerId, matchId, ph, pa, new Date().toISOString()]);
  return {ok: true};
}

/* ── RESULTS ── */
function getResults() {
  const sheet = getSheet('results');
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return {};
  const out = {};
  rows.slice(1).forEach(r => {
    const matchId = String(r[0]);
    if (!matchId || matchId === 'match_id') return;
    out[matchId] = {rh: +r[1], ra: +r[2]};
  });
  return out;
}

function saveResult(matchId, rh, ra) {
  const sheet = getSheet('results');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['match_id', 'result_home', 'result_away', 'saved_at']);
  }
  // Upsert
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(matchId)) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[rh, ra, new Date().toISOString()]]);
      return {ok: true, updated: true};
    }
  }
  sheet.appendRow([matchId, rh, ra, new Date().toISOString()]);
  return {ok: true};
}
