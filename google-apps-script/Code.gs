/**
 * PRODE MUNDIAL 2026 — Google Apps Script Backend
 * ─────────────────────────────────────────────────
 * HOW TO SET UP (5 minutes):
 *
 * 1. Go to https://sheets.new — create a new Google Sheet
 * 2. Name it "Prode Mundial 2026"
 * 3. Create 3 tabs (sheets) named exactly:
 *      players   |   predictions   |   results
 *
 * 4. In the sheet, go to Extensions → Apps Script
 * 5. Delete all existing code, paste this entire file
 * 6. Save (Ctrl+S), then click Deploy → New deployment
 * 7. Type: Web App
 *    Execute as: Me
 *    Who has access: Anyone
 * 8. Click Deploy → Copy the Web App URL
 * 9. Open the prode app → paste the URL when prompted
 *
 * That's it! All predictions and results sync automatically.
 */

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

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
  if (rows.length <= 1) return [];         // only header or empty
  return rows.slice(1).map(r => ({
    id:    r[0], 
    name:  r[1], 
    email: r[2],
  })).filter(p => p.id);
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
    const [playerId, matchId, ph, pa] = r;
    if (!playerId) return;
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
    const [matchId, rh, ra] = r;
    if (matchId) out[matchId] = {rh: +rh, ra: +ra};
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
