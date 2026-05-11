// ===== JHONG Withdrawal System v3.34 — Code.gs (Backend) =====
// This is the Google Apps Script backend.
// Deploy as Web App: Execute as Me | Who has access: Anyone
//
// SETUP:
//   1. Set SHEET_ID to your JHONG BACKEND spreadsheet ID
//   2. Set ADMIN_PIN to your own 4-digit value
//   3. Save, then Deploy → Manage deployments → New version → Deploy
//   4. Open the Web App URL — the app loads directly, no token needed
//
// HOW HTML SERVING WORKS:
//   doGet() with no ?action= param → serves index.html to the browser
//   doGet() with ?action=... param → handles API calls (data read/write)
// ================================================================

// ⚠ IMPORTANT: After pasting:
//   1. Edit ADMIN_PIN below to YOUR OWN value
//   2. Save (💾)
//   3. Deploy → Manage deployments → ✏ Edit → New version → Deploy
//   4. Copy the same URL into the webapp Settings panel
//      AND paste the SAME token + PINs into the webapp Settings.
// =====
// SECURITY MODEL:
//   • Anyone with the Web App URL can read data.
//   • USER role can READ and SAVE (no deletes).
//   • ADMIN role is required for any delete_* action.
//   • WITHDRAWAL and RECEIVED deletes are HARD — the row is permanently
//     removed from the sheet (no resurrection on Sync Now). All other
//     deletes (BEGINNING, SALESORDER, SERVED) remain soft for audit.
// =====
// Tabs in JHONG BACKEND:
//   ITEMCODE, WITHDRAWAL, RECEIVED, BEGINNING, SALESORDER, SPLIT, SERVED
//
// PASTE YOUR JHONG BACKEND SHEET ID BELOW.
var SHEET_ID = '18VqdFB_anOyzMA05DXtc7I0YKE84AnEELF35gnIDsFU';

// Admin PIN protects delete actions. Change this to your own 4-digit value.
var ADMIN_PIN = '7777';
var USER_PIN  = '1234';

var ITEMCODE_TAB        = 'ITEMCODE';
var CODEMAP_TAB         = 'CODEMAP';
var WITHDRAWAL_TAB      = 'WITHDRAWAL';
var RECEIVED_TAB        = 'RECEIVED';
var YARDSRECEIVED_TAB   = 'YARDSRECEIVED';
var BEGINNING_TAB       = 'BEGINNING';
var SALESORDER_TAB      = 'SALESORDER';
var SPLIT_TAB           = 'SPLIT';
var SERVED_TAB          = 'SERVED';
// Each header now ends with Deleted + DeletedAt + DeletedBy for soft-delete audit.
var WITHDRAWAL_HEADER    = ['Date','ItemCode','Description','Size','Qty','WithdrawalNo','Customer','Remarks','ID','Deleted','DeletedAt','DeletedBy'];
var RECEIVED_HEADER      = ['Date','ItemCode','Description','Size','Qty','MRR','Supplier','Remarks','ID','Deleted','DeletedAt','DeletedBy'];
var YARDSRECEIVED_HEADER = ['Date','ItemCode','Description','Size','Qty','MRR','Supplier','Remarks','ID','Deleted','DeletedAt','DeletedBy'];
var PARTIALROLLS_TAB    = 'PARTIALROLLS';
var PARTIALROLLS_HEADER = ['Date','ItemCode','Width','WidthUnit','Length','LengthUnit','Qty','Ref','Size','WithdrawalID','ID','Deleted','DeletedAt','DeletedBy'];
var PARTIALWITHDRAW_TAB    = 'PARTIALWITHDRAW';
var PARTIALWITHDRAW_HEADER = ['Date','ItemCode','Width','WidthUnit','Length','LengthUnit','Qty','Ref','Size','Note','ID','Deleted','DeletedAt','DeletedBy'];
var YARDSWITHDRAWN_TAB    = 'YARDSWITHDRAWN';
var YARDSWITHDRAWN_HEADER = ['Date','CoreNo','ItemID','Width','Yards','Ref','Customer','ID','Deleted','DeletedAt','DeletedBy'];
var DISPOSAL_TAB    = 'DISPOSAL';
var DISPOSAL_HEADER = ['Date','ItemCode','Width','WidthUnit','Length','LengthUnit','Qty','Remarks','ID','Deleted','DeletedAt','DeletedBy'];
var BEGINNING_HEADER  = ['Date','ItemCode','Description','Size','Qty','Notes','Deleted','DeletedAt','DeletedBy'];
var SALESORDER_HEADER = ['ID','Date','Time','OrderNo','Customer','Remarks','Items','Status','Deleted','DeletedAt','DeletedBy'];
var SPLIT_HEADER      = ['Date','SplitID','ParentCode','ParentDesc','ParentSize','ParentQty','ChildCode','ChildDesc','ChildSize','ChildQty','Note'];
var SERVED_HEADER     = ['Date','JobOrder','ItemID','Width','Length','Qty','Unit','Customer','Urgency','Status','Sales','ServedAt','RowKey','ID','Deleted','DeletedAt','DeletedBy'];

// Column indexes (1-based) for the soft-delete columns. Computed from header above.
function _delColIdx(header) { return header.indexOf('Deleted') + 1; }

function _safeText(s) {
  s = String(s == null ? '' : s);
  if (!s) return s;
  var c = s.charAt(0);
  if (c === '=' || c === '+' || c === '-' || c === '@') return "'" + s;
  return s;
}

function _isDeleteAction(action) {
  return action.indexOf('delete_') === 0;
}

function doGet(e) {
  var action = (e.parameter.action||'').toLowerCase();

  // No action param → serve the frontend HTML page
  if (!action) {
    var execUrl = ScriptApp.getService().getUrl();
    var tpl = HtmlService.createHtmlOutputFromFile('index');
    // Inject the real exec URL as a JS global so the frontend can call back
    // correctly. window.location.href inside the Apps Script sandbox is a
    // googleusercontent.com proxy URL — NOT the script.google.com exec URL.
    var content = tpl.getContent().replace(
      '<script>', '<script>\nvar APPS_SCRIPT_URL=' + JSON.stringify(execUrl) + ';\n'
    );
    return HtmlService.createHtmlOutput(content)
      .setTitle('JHONG Withdrawal System v3.34')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Deletes require the admin PIN
  if (_isDeleteAction(action)) {
    var pin = String(e.parameter.adminPin || '');
    if (!pin || pin !== ADMIN_PIN) {
      return _json({ok:false, error:'AUTH: admin PIN required for this action'});
    }
  }

  if (action === 'read_itemcode')      return _json({ok:true, rows: _getItemcodeSheet().getDataRange().getValues()});
  if (action === 'read_codemap')       return _json({ok:true, rows: _getCodemapCodes()});
  if (action === 'read_withdrawals')   return _json({ok:true, rows: _getWithdrawalSheet().getDataRange().getValues()});
  if (action === 'read_received')      return _json({ok:true, rows: _getReceivedSheet().getDataRange().getValues()});
  if (action === 'read_yardsreceived') return _json({ok:true, rows: _getYardsReceivedSheet().getDataRange().getValues()});
  if (action === 'read_beginning')     return _json({ok:true, rows: _getBeginningSheet().getDataRange().getValues()});
  if (action === 'read_salesorders')   return _json({ok:true, rows: _getSalesOrderSheet().getDataRange().getValues()});
  if (action === 'read_splits')        return _json({ok:true, rows: _getSplitSheet().getDataRange().getValues()});
  if (action === 'read_served')        return _json({ok:true, rows: _getServedSheet().getDataRange().getValues()});
  if (action === 'read_partialrolls')    return _json({ok:true, rows: _getPartialRollsSheet().getDataRange().getValues()});
  if (action === 'read_yardswithdrawn') return _json({ok:true, rows: _getYardsWithdrawnSheet().getDataRange().getValues()});

  if (e.parameter.payload) {
    var body;
    try { body = JSON.parse(e.parameter.payload); }
    catch(err) { return _json({ok:false, error:'Bad payload JSON: '+err}); }
    body.action = action;
    body._actor = String(e.parameter.actor || 'user');
    try { return _handleWrite(body); }
    catch(err2) { return _json({ok:false, error: String(err2 && err2.message || err2)}); }
  }
  return _json({ok:false, error:'Unknown action: '+action});
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = (body.action || '').toLowerCase();
    if (_isDeleteAction(action)) {
      var pin = String(body.adminPin || '');
      if (!pin || pin !== ADMIN_PIN) return _json({ok:false, error:'AUTH: admin PIN required'});
    }
    return _handleWrite(body);
  } catch(err) {
    return _json({ok:false, error: 'doPost error: '+String(err)});
  }
}

function _findRowById(sh, idColIndex1Based, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, idColIndex1Based, last - 1, 1).getValues();
  var target = String(id).trim();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === target) return i + 2;
  }
  return -1;
}

function _findRowByCode(sh, codeColIndex1Based, code) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var codes = sh.getRange(2, codeColIndex1Based, last - 1, 1).getValues();
  var target = String(code).trim();
  for (var i = 0; i < codes.length; i++) {
    if (String(codes[i][0]).trim() === target) return i + 2;
  }
  return -1;
}

function _genId(prefix) {
  return prefix + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
}

// Soft-delete: mark a row instead of removing it. Writes Deleted=TRUE,
// DeletedAt=ISO timestamp, DeletedBy=actor. Frontend filters out flagged rows.
function _softDeleteRow(sh, rowIdx, header, actor) {
  var delCol = _delColIdx(header); // 1-based column number for "Deleted"
  if (delCol < 1) throw new Error('Header has no Deleted column');
  var ts = new Date().toISOString();
  sh.getRange(rowIdx, delCol, 1, 3).setValues([[true, ts, String(actor||'admin')]]);
}

function _handleWrite(body) {
  var action = (body.action||'').toLowerCase();
  var actor  = String(body._actor || body.actor || 'user');

  if (action === 'save_withdrawal') {
    var r = body.record || {};
    if (!r.itemCode) return _json({ok:false, error:'Missing itemCode'});
    // Synthesized legacy IDs (whd_legacy_r*) are never written to the sheet,
    // so _findRowById will never find them. Treat them as "no ID" and generate
    // a real one so the row becomes trackable after the first edit.
    var rawId = String(r.id || '').trim();
    var id = (rawId && !/^whd_legacy_/.test(rawId)) ? rawId : _genId('whd');
    var sh = _getWithdrawalSheet();
    var existing = _findRowById(sh, 9, id);
    var row = [r.date||'', String(r.itemCode), r.desc||'', r.size||'', Number(r.qty)||0, r.withdrawalNo||'', r.customer||'', r.remarks||'', id, false, '', ''];
    if (existing !== -1) {
      sh.getRange(existing, 1, 1, row.length).setValues([row]);
    } else {
      // Legacy fallback: the row existed before the ID column was added.
      // Frontend sends _sheetRow (1-based) so we can overwrite it directly
      // instead of appending a duplicate.
      var legacyRow = parseInt(r._sheetRow || 0);
      if (legacyRow > 1) sh.getRange(legacyRow, 1, 1, row.length).setValues([row]);
      else sh.appendRow(row);
    }
    return _json({ok:true, id:id});
  }

  if (action === 'delete_withdrawal') {
    // HARD delete — permanently removes the row from the WITHDRAWAL sheet.
    // Supports two modes:
    //   1. id-based (preferred): finds the row by scanning column I for the ID.
    //   2. sheetRow-based (legacy fallback): used when the row has no ID in
    //      column I (pre-dates the ID column). The frontend sends the 1-based
    //      row number it read from the sheet so we can delete it directly.
    var sh = _getWithdrawalSheet();
    var rowIdx = -1;
    var did = String(body.id || '').trim();
    if (did) {
      rowIdx = _findRowById(sh, 9, did);
      if (rowIdx === -1) return _json({ok:false, error:'Withdrawal not found: '+did});
    } else {
      var sr = parseInt(body.sheetRow, 10);
      if (!sr || sr < 2) return _json({ok:false, error:'Missing id or valid sheetRow for delete_withdrawal'});
      if (sr > sh.getLastRow()) return _json({ok:false, error:'sheetRow '+sr+' out of range'});
      rowIdx = sr;
    }
    sh.deleteRow(rowIdx);
    return _json({ok:true, hard:true, row:rowIdx});
  }

  if (action === 'save_received') {
    var r2 = body.record || {};
    if (!r2.itemCode) return _json({ok:false, error:'Missing itemCode'});
    var rawRid = String(r2.id || '').trim();
    var rid = (rawRid && !/^rec_legacy_/.test(rawRid)) ? rawRid : _genId('rec');
    var rsh = _getReceivedSheet();
    var rexisting = _findRowById(rsh, 9, rid);
    var rrow = [r2.date||'', String(r2.itemCode), r2.desc||'', r2.size||'', Number(r2.qty)||0, r2.mrrNo||'', r2.supplier||'', r2.remarks||'', rid, false, '', ''];
    if (rexisting !== -1) {
      rsh.getRange(rexisting, 1, 1, rrow.length).setValues([rrow]);
    } else {
      var rLegacyRow = parseInt(r2._sheetRow || 0);
      if (rLegacyRow > 1) rsh.getRange(rLegacyRow, 1, 1, rrow.length).setValues([rrow]);
      else rsh.appendRow(rrow);
    }
    return _json({ok:true, id:rid});
  }

  if (action === 'save_yardsreceived') {
    var yr = body.record || {};
    if (!yr.itemCode) return _json({ok:false, error:'Missing itemCode'});
    var yrid = yr.id || _genId('yrc');
    var yrsh = _getYardsReceivedSheet();
    var yrexisting = _findRowById(yrsh, 9, yrid);
    var yrrow = [yr.date||'', String(yr.itemCode), yr.desc||'', yr.size||'', Number(yr.qty)||0, yr.mrrNo||'', yr.supplier||'', yr.remarks||'', yrid, false, '', ''];
    if (yrexisting === -1) yrsh.appendRow(yrrow);
    else yrsh.getRange(yrexisting, 1, 1, yrrow.length).setValues([yrrow]);
    return _json({ok:true, id:yrid});
  }

  if (action === 'delete_yardsreceived') {
    // HARD delete — permanently removes the row from the YARDSRECEIVED sheet.
    // Tries id first (column I), then falls back to mrrNo/coreNo (column F/6)
    // for rolls that were added before the ID column existed.
    var yrsh2 = _getYardsReceivedSheet();
    var yrdRowIdx = -1;
    var yrdId = String(body.id || '').trim();
    if (yrdId) {
      yrdRowIdx = _findRowById(yrsh2, 9, yrdId);
    }
    if (yrdRowIdx === -1) {
      // Fallback: find by mrrNo (Core No.) in column 6
      var yrdMrr = String(body.mrrNo || '').trim();
      if (yrdMrr) yrdRowIdx = _findRowByCode(yrsh2, 6, yrdMrr);
    }
    if (yrdRowIdx === -1) return _json({ok:false, error:'YARDSRECEIVED row not found for id='+yrdId+' mrrNo='+(body.mrrNo||'')});
    yrsh2.deleteRow(yrdRowIdx);
    return _json({ok:true, hard:true, row:yrdRowIdx});
  }

  if (action === 'save_partialroll') {
    var pr = body.record || {};
    if (!pr.itemCode) return _json({ok:false, error:'Missing itemCode'});
    var prid = pr.id || _genId('pr');
    var prsh = _getPartialRollsSheet();
    var prexisting = _findRowById(prsh, 11, prid);
    var prrow = [
      pr.date||'', String(pr.itemCode),
      pr.width||'', pr.widthUnit||'',
      pr.length||'', pr.lengthUnit||'',
      Number(pr.qty)||1, pr.ref||'', pr.size||'',
      pr.withdrawalId||'', prid, false, '', ''
    ];
    if (prexisting === -1) prsh.appendRow(prrow);
    else prsh.getRange(prexisting, 1, 1, prrow.length).setValues([prrow]);
    return _json({ok:true, id:prid});
  }

  if (action === 'delete_partialroll') {
    // HARD delete — id lives in column 11 (PARTIALROLLS_HEADER position).
    var dprid = String(body.id || '').trim();
    if (!dprid) return _json({ok:false, error:'Missing id for delete_partialroll'});
    var dprsh = _getPartialRollsSheet();
    var dprIdx = _findRowById(dprsh, 11, dprid);
    if (dprIdx === -1) return _json({ok:true, removed:false});
    dprsh.deleteRow(dprIdx);
    return _json({ok:true, hard:true, row:dprIdx});
  }

  // ===== PARTIALWITHDRAW =====
  if (action === 'save_partialwithdraw') {
    var pw = body.record || {};
    if (!pw.itemCode) return _json({ok:false, error:'Missing itemCode'});
    var pwid = pw.id || _genId('prwd');
    var pwsh = _getPartialWithdrawSheet();
    var pwexisting = _findRowById(pwsh, 11, pwid);
    var pwrow = [
      pw.date||'', String(pw.itemCode),
      pw.width||'', pw.widthUnit||'',
      pw.length||'', pw.lengthUnit||'',
      Number(pw.qty)||0, pw.ref||'', pw.size||'',
      _safeText(pw.note||''), pwid, false, '', ''
    ];
    if (pwexisting === -1) pwsh.appendRow(pwrow);
    else pwsh.getRange(pwexisting, 1, 1, pwrow.length).setValues([pwrow]);
    return _json({ok:true, id:pwid});
  }

  if (action === 'delete_partialwithdraw') {
    var dpwid = String(body.id || '').trim();
    if (!dpwid) return _json({ok:false, error:'Missing id for delete_partialwithdraw'});
    var dpwsh = _getPartialWithdrawSheet();
    var dpwIdx = _findRowById(dpwsh, 11, dpwid);
    if (dpwIdx === -1) return _json({ok:true, removed:false});
    dpwsh.deleteRow(dpwIdx);
    return _json({ok:true, hard:true, row:dpwIdx});
  }

  // ===== YARDSWITHDRAWN =====
  if (action === 'save_yardswithdrawn') {
    var yw = body.record || {};
    if (!yw.coreNo) return _json({ok:false, error:'Missing coreNo'});
    var ywid = yw.id || _genId('tmywd');
    var ywsh = _getYardsWithdrawnSheet();
    var ywexisting = _findRowById(ywsh, 8, ywid);
    var ywrow = [
      yw.date||'', String(yw.coreNo), String(yw.itemId||''),
      yw.width||'', Number(yw.yards)||0, yw.ref||'', _safeText(yw.customer||''),
      ywid, false, '', ''
    ];
    if (ywexisting === -1) ywsh.appendRow(ywrow);
    else ywsh.getRange(ywexisting, 1, 1, ywrow.length).setValues([ywrow]);
    return _json({ok:true, id:ywid});
  }

  if (action === 'delete_yardswithdrawn') {
    var dywid = String(body.id || '').trim();
    if (!dywid) return _json({ok:false, error:'Missing id for delete_yardswithdrawn'});
    var dywsh = _getYardsWithdrawnSheet();
    var dywIdx = _findRowById(dywsh, 8, dywid);
    if (dywIdx === -1) return _json({ok:true, removed:false});
    dywsh.deleteRow(dywIdx);
    return _json({ok:true, hard:true, row:dywIdx});
  }

  // ===== DISPOSAL =====
  if (action === 'save_disposal') {
    var dp = body.record || {};
    if (!dp.itemCode) return _json({ok:false, error:'Missing itemCode'});
    var dpid = dp.id || _genId('dsp');
    var dpsh = _getDisposalSheet();
    var dpexisting = _findRowById(dpsh, 9, dpid);
    var dprow = [
      dp.date||'', String(dp.itemCode),
      dp.width||'', dp.widthUnit||'',
      dp.length||'', dp.lengthUnit||'',
      Number(dp.qty)||0, _safeText(dp.remarks||''),
      dpid, false, '', ''
    ];
    if (dpexisting === -1) dpsh.appendRow(dprow);
    else dpsh.getRange(dpexisting, 1, 1, dprow.length).setValues([dprow]);
    return _json({ok:true, id:dpid});
  }

  if (action === 'delete_received') {
    // HARD delete — same rationale as delete_withdrawal above.
    var did2 = body.id || '';
    if (!did2) return _json({ok:false, error:'Missing id for delete_received'});
    var sh2 = _getReceivedSheet();
    var rowIdx2 = _findRowById(sh2, 9, did2);
    if (rowIdx2 === -1) return _json({ok:false, error:'Received not found: '+did2});
    sh2.deleteRow(rowIdx2);
    return _json({ok:true, hard:true, row:rowIdx2});
  }

  // Backfill missing IDs in column I (the ID column) of WITHDRAWAL / RECEIVED.
  // Also repairs the header row when legacy sheets are missing the ID/Deleted
  // columns, so the parser can correctly identify them on subsequent reads.
  if (action === 'backfill_withdrawal_ids' || action === 'backfill_received_ids') {
    var isWith = (action === 'backfill_withdrawal_ids');
    var bsh = isWith ? _getWithdrawalSheet() : _getReceivedSheet();
    var prefix = isWith ? 'whd' : 'rec';
    var canonicalHeader = isWith ? WITHDRAWAL_HEADER : RECEIVED_HEADER;

    // Step 1 — repair the header row so column I = 'ID', J = 'Deleted', etc.
    // Read whatever's currently in row 1, pad to canonical width, then write
    // any missing canonical names back into their fixed positions.
    var lastCol = Math.max(bsh.getLastColumn(), canonicalHeader.length);
    var headerVals = bsh.getRange(1, 1, 1, lastCol).getValues()[0];
    var headerChanged = false;
    for (var hi = 0; hi < canonicalHeader.length; hi++) {
      var have = String(headerVals[hi]||'').trim();
      if (!have) { headerVals[hi] = canonicalHeader[hi]; headerChanged = true; }
    }
    if (headerChanged) bsh.getRange(1, 1, 1, headerVals.length).setValues([headerVals]);

    // Step 2 — write a real ID into every blank cell in column I, rows 2+.
    var lastRow = bsh.getLastRow();
    if (lastRow < 2) return _json({ok:true, filled:0, headerRepaired:headerChanged});
    var idRange = bsh.getRange(2, 9, lastRow - 1, 1);
    var ids = idRange.getValues();
    var filled = 0;
    for (var bi = 0; bi < ids.length; bi++) {
      if (!String(ids[bi][0]||'').trim()) {
        ids[bi][0] = _genId(prefix);
        filled++;
      }
    }
    if (filled > 0) idRange.setValues(ids);
    return _json({ok:true, filled:filled, headerRepaired:headerChanged});
  }

  // ===== BEGINNING =====
  if (action === 'save_beginning_single') {
    var b = body.record || {};
    if (!b.code) return _json({ok:false, error:'Missing code'});
    var bsh = _getBeginningSheet();
    var bexisting = _findRowByCode(bsh, 2, b.code);
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Manila', 'yyyy-MM-dd');
    var qtyNum = Number(b.qty)||0;
    if (bexisting === -1) {
      var brow = [today, String(b.code), b.desc||'', b.size||'', qtyNum, b.notes||'', false, '', ''];
      bsh.appendRow(brow);
    } else {
      var existingDate = bsh.getRange(bexisting, 1).getValue() || today;
      var brow2 = [existingDate, String(b.code), b.desc||'', b.size||'', qtyNum, b.notes||'', false, '', ''];
      bsh.getRange(bexisting, 1, 1, brow2.length).setValues([brow2]);
    }
    return _json({ok:true, code: b.code, qty: qtyNum});
  }

  if (action === 'append_beginning_row') {
    var ab = body.record || {};
    if (!ab.code) return _json({ok:false, error:'Missing code'});
    var absh = _getBeginningSheet();
    var abDate = ab.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Manila', 'yyyy-MM-dd');
    var abQty = Number(ab.qty)||0;
    var abRow = [abDate, String(ab.code), ab.desc||'', ab.size||'', abQty, _safeText(ab.notes||''), false, '', ''];
    absh.appendRow(abRow);
    return _json({ok:true, code: ab.code, qty: abQty, row: absh.getLastRow()});
  }

  // ===== SPLIT (audit log) =====
  if (action === 'save_split') {
    var sp = body.record || {};
    if (!sp.parentCode || !sp.childCode) return _json({ok:false, error:'Missing parentCode/childCode'});
    var spsh = _getSplitSheet();
    var spDate = sp.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Manila', 'yyyy-MM-dd');
    var spRow = [
      spDate, String(sp.splitId || ''), String(sp.parentCode),
      sp.parentDesc || '', sp.parentSize || '', Number(sp.parentQty)||0,
      String(sp.childCode), sp.childDesc || '', sp.childSize || '',
      Number(sp.childQty)||0, _safeText(sp.note || '')
    ];
    spsh.appendRow(spRow);
    return _json({ok:true, splitId: sp.splitId, row: spsh.getLastRow()});
  }

  // ===== SERVED =====
  if (action === 'save_served') {
    var sv = body.record || {};
    if (!sv.rowKey) return _json({ok:false, error:'Missing rowKey'});
    var svsh = _getServedSheet();
    var rkIdx = _findRowByCode(svsh, 13, sv.rowKey);
    var svId = sv.id || _genId('srv');
    var svDate = sv.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Manila', 'yyyy-MM-dd');
    var svAt = sv.servedAt || new Date().toISOString();
    var svRow = [
      svDate, String(sv.jobOrder || ''), String(sv.itemId || ''),
      sv.width || '', sv.length || '', Number(sv.qty)||0,
      sv.unit || '', sv.customer || '', sv.urgency || '',
      sv.status || '', sv.sales || '', svAt,
      String(sv.rowKey), svId, false, '', ''
    ];
    if (rkIdx === -1) svsh.appendRow(svRow);
    else svsh.getRange(rkIdx, 1, 1, svRow.length).setValues([svRow]);
    return _json({ok:true, id: svId, rowKey: sv.rowKey, row: rkIdx === -1 ? svsh.getLastRow() : rkIdx});
  }

  if (action === 'delete_served') {
    var dRowKey = String(body.rowKey || '').trim();
    if (!dRowKey) return _json({ok:false, error:'Missing rowKey'});
    var dsvsh = _getServedSheet();
    var dIdx = _findRowByCode(dsvsh, 13, dRowKey);
    if (dIdx === -1) return _json({ok:true, removed:false});
    _softDeleteRow(dsvsh, dIdx, SERVED_HEADER, actor);
    return _json({ok:true, removed:true, soft:true});
  }

  if (action === 'save_beginning_bulk') {
    var items = body.items || [];
    var bsh4 = _getBeginningSheet();
    var todayB = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Manila', 'yyyy-MM-dd');
    var n = 0;
    items.forEach(function(it){
      if (!it.code) return;
      var existing = _findRowByCode(bsh4, 2, it.code);
      var qn = Number(it.qty)||0;
      if (existing === -1) {
        bsh4.appendRow([todayB, String(it.code), it.desc||'', it.size||'', qn, it.notes||'', false, '', '']);
      } else {
        var existDate = bsh4.getRange(existing, 1).getValue() || todayB;
        bsh4.getRange(existing, 1, 1, 9).setValues([[existDate, String(it.code), it.desc||'', it.size||'', qn, it.notes||'', false, '', '']]);
      }
      n++;
    });
    return _json({ok:true, count: n});
  }

  if (action === 'delete_beginning_single') {
    var dcode = body.code || '';
    var bsh3 = _getBeginningSheet();
    var bidx = _findRowByCode(bsh3, 2, dcode);
    if (bidx === -1) return _json({ok:false, error:'Beginning not found: '+dcode});
    _softDeleteRow(bsh3, bidx, BEGINNING_HEADER, actor);
    return _json({ok:true, soft:true});
  }

  // ===== SALESORDER =====
  if (action === 'save_salesorder') {
    var so = body.record || {};
    var soid = so.id || _genId('so');
    var ssh = _getSalesOrderSheet();
    var sexisting = _findRowById(ssh, 1, soid);
    var itemsJson = '';
    try { itemsJson = JSON.stringify(so.items || []); } catch(e) { itemsJson = '[]'; }
    var srow = [soid, so.date||'', so.time||'', so.orderNo||'', so.customer||'', so.remarks||'', itemsJson, (so.status||'pending'), false, '', ''];
    if (sexisting === -1) ssh.appendRow(srow);
    else ssh.getRange(sexisting, 1, 1, srow.length).setValues([srow]);
    return _json({ok:true, id: soid});
  }

  if (action === 'update_salesorder_status') {
    var sid = body.id || '';
    var newStatus = String(body.status || 'pending').toLowerCase();
    var ssh2 = _getSalesOrderSheet();
    var sidx = _findRowById(ssh2, 1, sid);
    if (sidx === -1) return _json({ok:false, error:'SalesOrder not found: '+sid});
    ssh2.getRange(sidx, 8, 1, 1).setValues([[newStatus]]);
    return _json({ok:true});
  }

  if (action === 'delete_salesorder') {
    var sid2 = body.id || '';
    var ssh3 = _getSalesOrderSheet();
    var sidx2 = _findRowById(ssh3, 1, sid2);
    if (sidx2 === -1) return _json({ok:false, error:'SalesOrder not found: '+sid2});
    _softDeleteRow(ssh3, sidx2, SALESORDER_HEADER, actor);
    return _json({ok:true, soft:true});
  }

  return _json({ok:false, error:'Unknown action: '+action});
}

function _openSS() {
  if (!SHEET_ID || SHEET_ID === 'PASTE_JHONG_BACKEND_SHEET_ID_HERE') {
    throw new Error('SHEET_ID is not set. Paste your JHONG BACKEND sheet ID.');
  }
  return SpreadsheetApp.openById(SHEET_ID);
}
function _getItemcodeSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(ITEMCODE_TAB);
  if (!sh) throw new Error('ITEMCODE tab not found in JHONG BACKEND.');
  return sh;
}
// Returns array of {code, desc} from CODEMAP columns A (code) and B (description).
function _getCodemapCodes() {
  var ss = _openSS();
  var sh = ss.getSheetByName(CODEMAP_TAB);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  // Columns A+B = columns 1 and 2
  var values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  var result = [];
  values.forEach(function(row) {
    var code = String(row[0] || '').trim();
    var desc = String(row[1] || '').trim();
    if (code) result.push({code: code, desc: desc});
  });
  return result;
}
function _getWithdrawalSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(WITHDRAWAL_TAB);
  if (!sh) { sh = ss.insertSheet(WITHDRAWAL_TAB); sh.appendRow(WITHDRAWAL_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(WITHDRAWAL_HEADER);
  return sh;
}
function _getReceivedSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(RECEIVED_TAB);
  if (!sh) { sh = ss.insertSheet(RECEIVED_TAB); sh.appendRow(RECEIVED_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(RECEIVED_HEADER);
  return sh;
}
function _getYardsReceivedSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(YARDSRECEIVED_TAB);
  if (!sh) { sh = ss.insertSheet(YARDSRECEIVED_TAB); sh.appendRow(YARDSRECEIVED_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(YARDSRECEIVED_HEADER);
  return sh;
}
function _getPartialRollsSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(PARTIALROLLS_TAB);
  if (!sh) { sh = ss.insertSheet(PARTIALROLLS_TAB); sh.appendRow(PARTIALROLLS_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(PARTIALROLLS_HEADER);
  return sh;
}
function _getPartialWithdrawSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(PARTIALWITHDRAW_TAB);
  if (!sh) { sh = ss.insertSheet(PARTIALWITHDRAW_TAB); sh.appendRow(PARTIALWITHDRAW_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(PARTIALWITHDRAW_HEADER);
  return sh;
}
function _getDisposalSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(DISPOSAL_TAB);
  if (!sh) { sh = ss.insertSheet(DISPOSAL_TAB); sh.appendRow(DISPOSAL_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(DISPOSAL_HEADER);
  return sh;
}
function _getYardsWithdrawnSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(YARDSWITHDRAWN_TAB);
  if (!sh) { sh = ss.insertSheet(YARDSWITHDRAWN_TAB); sh.appendRow(YARDSWITHDRAWN_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(YARDSWITHDRAWN_HEADER);
  return sh;
}
function _getBeginningSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(BEGINNING_TAB);
  if (!sh) { sh = ss.insertSheet(BEGINNING_TAB); sh.appendRow(BEGINNING_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(BEGINNING_HEADER);
  return sh;
}
function _getSalesOrderSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(SALESORDER_TAB);
  if (!sh) { sh = ss.insertSheet(SALESORDER_TAB); sh.appendRow(SALESORDER_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(SALESORDER_HEADER);
  return sh;
}
function _getSplitSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(SPLIT_TAB);
  if (!sh) { sh = ss.insertSheet(SPLIT_TAB); sh.appendRow(SPLIT_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(SPLIT_HEADER);
  return sh;
}
function _getServedSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(SERVED_TAB);
  if (!sh) { sh = ss.insertSheet(SERVED_TAB); sh.appendRow(SERVED_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(SERVED_HEADER);
  return sh;
}

function _txt(s){ return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); }

function _json(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}