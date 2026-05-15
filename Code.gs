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
var WITHDRAWAL_HEADER    = ['Date','ItemCode','Description','Width','W-UM','Length','L-UM','Qty','WithdrawalNo','Customer','Remarks','ID','Deleted','DeletedAt','DeletedBy'];
var RECEIVED_HEADER      = ['Date','ItemCode','Description','Size','Qty','MRR','Supplier','Remarks','ID','Deleted','DeletedAt','DeletedBy'];
var YARDSRECEIVED_HEADER = ['Date','ItemCode','Description','Size','Qty','MRR','Supplier','Remarks','ID','Deleted','DeletedAt','DeletedBy'];
var PARTIALROLLS_TAB    = 'PARTIALROLLS';
var PARTIALROLLS_HEADER = ['Date','ItemCode','Width','WidthUnit','Length','LengthUnit','Qty','Ref','Size','WithdrawalID','ID','Deleted','DeletedAt','DeletedBy'];
var PARTIALWITHDRAW_TAB    = 'PARTIALWITHDRAW';
var PARTIALWITHDRAW_HEADER = ['Date','ItemCode','Width','WidthUnit','Length','LengthUnit','Qty','Ref','Size','Note','ID','Deleted','DeletedAt','DeletedBy'];
var YARDSWITHDRAWN_TAB    = 'YARDSWITHDRAWN';
var YARDSWITHDRAWN_HEADER = ['Date','CoreNo','ItemID','Width','WidthUnit','Yards','LengthUnit','Ref','Customer','ID','Deleted','DeletedAt','DeletedBy'];
var DISPOSAL_TAB    = 'DISPOSAL';
var DISPOSAL_HEADER = ['Date','ItemCode','Width','WidthUnit','Length','LengthUnit','Qty','Remarks','ID','Deleted','DeletedAt','DeletedBy'];
var BEGINNING_HEADER  = ['Date','ItemCode','Description','Width','W-UM','Length','L-UM','Qty','Notes','Deleted','DeletedAt','DeletedBy'];
var SALESORDER_HEADER = ['ID','Date','Time','OrderNo','Customer','Remarks','Items','Status','Deleted','DeletedAt','DeletedBy'];
var SPLIT_HEADER      = ['Date','SplitID','ParentCode','ParentDesc','ParentSize','ParentQty','ChildCode','ChildDesc','ChildSize','ChildQty','WithdrawalID','ReceivedID','WithdrawalNo','Note'];
var SERVED_HEADER     = ['Date','JobOrder','ItemID','Width','Length','Qty','Unit','Customer','Urgency','Status','Sales','ServedAt','RowKey','ID','Deleted','DeletedAt','DeletedBy'];

// SPLIT-PARENT: records each parent roll consumed during a split (deducts from parent stock)
var SPLIT_PARENT_TAB    = 'SPLIT-PARENT';
var SPLIT_PARENT_HEADER = ['Date','SplitID','ItemCode','Description','Width','W-UM','Length','L-UM','Qty','WithdrawalNo','Customer','Note','ID','Deleted','DeletedAt','DeletedBy'];

// SPLIT-CHILD: records each batch of child rolls created during a split (adds to child stock)
var SPLIT_CHILD_TAB    = 'SPLIT-CHILD';
var SPLIT_CHILD_HEADER = ['Date','SplitID','ItemCode','Description','Width','W-UM','Length','L-UM','Qty','MRRNo','Supplier','Note','ID','Deleted','DeletedAt','DeletedBy'];

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
  if (action === 'read_split_parent')  return _json({ok:true, rows: _getSplitParentSheet().getDataRange().getValues()});
  if (action === 'read_split_child')   return _json({ok:true, rows: _getSplitChildSheet().getDataRange().getValues()});

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
    // ID column is now column 12 (1-based) after Width/W-UM/Length/L-UM replaced Size
    var existing = _findRowById(sh, 12, id);
    // Parse width/length — accept individual fields; fall back to splitting legacy Size string
    var _width  = String(r.width  || '').trim();
    var _wUnit  = String(r.widthUnit  || '').trim();
    var _length = String(r.length || '').trim();
    var _lUnit  = String(r.lengthUnit || '').trim();
    if (!_width && !_length && r.size) {
      var _sm = String(r.size).match(/^([0-9\/\-\.]+)\s*([a-zA-Z\"']*)\s*[xX×]\s*([0-9\/\-\.]+)\s*([a-zA-Z\"']*)$/);
      if (_sm) {
        _width = _sm[1] || ''; _wUnit = _sm[2] || '';
        _length = _sm[3] || ''; _lUnit = _sm[4] || '';
      } else {
        _width = String(r.size); // fallback: whole string in Width
      }
    }
    var row = [r.date||'', String(r.itemCode), r.desc||'', _width, _wUnit, _length, _lUnit, Number(r.qty)||0, r.withdrawalNo||'', r.customer||'', r.remarks||'', id, false, '', ''];
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
    // Supports three modes:
    //   1. id-based (preferred): finds the row by scanning column 12 for the ID.
    //   2. full-row scan (corruption recovery): if the col-12 scan misses, every
    //      cell in every data row is checked. This catches pre-fix split records
    //      where broken field mapping wrote the ID into the wrong column.
    //   3. sheetRow-based (legacy fallback): used when the row has no ID at all
    //      (pre-dates the ID column). The frontend sends the 1-based row number.
    var sh = _getWithdrawalSheet();
    var rowIdx = -1;
    var did = String(body.id || '').trim();
    if (did) {
      // Mode 1 — fast scan of the ID column only
      rowIdx = _findRowById(sh, 12, did);
      // Mode 2 — corruption-recovery: full-grid scan when col-12 misses
      if (rowIdx === -1) {
        var lastR = sh.getLastRow();
        var lastC = sh.getLastColumn();
        if (lastR >= 2 && lastC >= 1) {
          var allData = sh.getRange(2, 1, lastR - 1, lastC).getValues();
          outer: for (var ri = 0; ri < allData.length; ri++) {
            for (var ci = 0; ci < allData[ri].length; ci++) {
              if (String(allData[ri][ci]).trim() === did) {
                rowIdx = ri + 2;
                break outer;
              }
            }
          }
        }
      }
      if (rowIdx === -1) return _json({ok:false, error:'Withdrawal not found: '+did});
    } else {
      // Mode 3 — sheetRow fallback for truly ID-less legacy rows
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
    var ywexisting = _findRowById(ywsh, 10, ywid);
    var ywrow = [
      yw.date||'', String(yw.coreNo), String(yw.itemId||''),
      yw.width||'', yw.widthUnit||'in', Number(yw.yards)||0, yw.lengthUnit||'yds',
      yw.ref||'', _safeText(yw.customer||''),
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
    var dywIdx = _findRowById(dywsh, 10, dywid);
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
    var idRange = bsh.getRange(2, 12, lastRow - 1, 1);
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
    // Parse width/length — accept individual fields; fall back to splitting legacy Size string
    var _bWidth  = String(b.width  || '').trim();
    var _bWUnit  = String(b.widthUnit  || '').trim();
    var _bLength = String(b.length || '').trim();
    var _bLUnit  = String(b.lengthUnit || '').trim();
    if (!_bWidth && !_bLength && b.size) {
      var _bSm = String(b.size).match(/^([0-9\/\-\.]+)\s*([a-zA-Z\"']*)\s*[xX×]\s*([0-9\/\-\.]+)\s*([a-zA-Z\"']*)$/);
      if (_bSm) {
        _bWidth = _bSm[1] || ''; _bWUnit = _bSm[2] || '';
        _bLength = _bSm[3] || ''; _bLUnit = _bSm[4] || '';
      } else {
        _bWidth = String(b.size);
      }
    }
    if (bexisting === -1) {
      var brow = [today, String(b.code), b.desc||'', _bWidth, _bWUnit, _bLength, _bLUnit, qtyNum, b.notes||'', false, '', ''];
      bsh.appendRow(brow);
    } else {
      var existingDate = bsh.getRange(bexisting, 1).getValue() || today;
      var brow2 = [existingDate, String(b.code), b.desc||'', _bWidth, _bWUnit, _bLength, _bLUnit, qtyNum, b.notes||'', false, '', ''];
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
    // Parse width/length — accept individual fields; fall back to splitting legacy Size string
    var _abWidth  = String(ab.width  || '').trim();
    var _abWUnit  = String(ab.widthUnit  || '').trim();
    var _abLength = String(ab.length || '').trim();
    var _abLUnit  = String(ab.lengthUnit || '').trim();
    if (!_abWidth && !_abLength && ab.size) {
      var _abSm = String(ab.size).match(/^([0-9\/\-\.]+)\s*([a-zA-Z\"']*)\s*[xX×]\s*([0-9\/\-\.]+)\s*([a-zA-Z\"']*)$/);
      if (_abSm) {
        _abWidth = _abSm[1] || ''; _abWUnit = _abSm[2] || '';
        _abLength = _abSm[3] || ''; _abLUnit = _abSm[4] || '';
      } else {
        _abWidth = String(ab.size);
      }
    }
    var abRow = [abDate, String(ab.code), ab.desc||'', _abWidth, _abWUnit, _abLength, _abLUnit, abQty, _safeText(ab.notes||''), false, '', ''];
    absh.appendRow(abRow);
    return _json({ok:true, code: ab.code, qty: abQty, row: absh.getLastRow()});
  }

  // ===== SPLIT (delegates to dedicated function to avoid var-hoisting collisions) =====
  if (action === 'save_split') {
    return _handleSaveSplit(body.record || {});
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
      // Parse width/length — accept individual fields; fall back to splitting legacy Size string
      var _itWidth  = String(it.width  || '').trim();
      var _itWUnit  = String(it.widthUnit  || '').trim();
      var _itLength = String(it.length || '').trim();
      var _itLUnit  = String(it.lengthUnit || '').trim();
      if (!_itWidth && !_itLength && it.size) {
        var _itSm = String(it.size).match(/^([0-9\/\-\.]+)\s*([a-zA-Z\"']*)\s*[xX×]\s*([0-9\/\-\.]+)\s*([a-zA-Z\"']*)$/);
        if (_itSm) {
          _itWidth = _itSm[1] || ''; _itWUnit = _itSm[2] || '';
          _itLength = _itSm[3] || ''; _itLUnit = _itSm[4] || '';
        } else {
          _itWidth = String(it.size);
        }
      }
      if (existing === -1) {
        bsh4.appendRow([todayB, String(it.code), it.desc||'', _itWidth, _itWUnit, _itLength, _itLUnit, qn, it.notes||'', false, '', '']);
      } else {
        var existDate = bsh4.getRange(existing, 1).getValue() || todayB;
        bsh4.getRange(existing, 1, 1, 12).setValues([[existDate, String(it.code), it.desc||'', _itWidth, _itWUnit, _itLength, _itLUnit, qn, it.notes||'', false, '', '']]);
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

// =====================================================================
// SPLIT handler — extracted into its own function to give it a clean
// variable scope. Apps Script uses ES5 var-hoisting: all var declarations
// inside _handleWrite share one scope, so re-using names like sp, sh,
// existing etc. causes silent collisions that corrupt values mid-execution.
// Putting save_split here fixes that completely.
// =====================================================================
function _handleSaveSplit(sp) {
  if (!sp.parentCode) return _json({ok:false, error:'Missing parentCode'});
  if (!sp.childCode)  return _json({ok:false, error:'Missing childCode'});

  var tz        = Session.getScriptTimeZone() || 'Asia/Manila';
  var spDate    = sp.date || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var splitId   = String(sp.splitId      || _genId('SPLIT'));
  var whdlNo    = String(sp.withdrawalNo || splitId);
  var customer  = String(sp.customer     || 'Internal — Roll Split');
  var note      = _safeText(sp.note      || '');
  var pCode     = String(sp.parentCode);
  var pDesc     = String(sp.parentDesc   || '');
  var cCode     = String(sp.childCode);
  var cDesc     = String(sp.childDesc    || '');
  var pQty      = Number(sp.parentQty)   || 1;
  var cQty      = Number(sp.childQty)    || (pQty * 2);

  // Parse parent dimensions — prefer explicit fields, fall back to parentSize string
  var pW  = String(sp.parentWidth      || '').trim();
  var pWU = String(sp.parentWidthUnit  || '').trim();
  var pL  = String(sp.parentLength     || '').trim();
  var pLU = String(sp.parentLengthUnit || '').trim();
  if (!pW && !pL && sp.parentSize) {
    var psm = String(sp.parentSize).match(/^([0-9\/\-\.]+)\s*([a-zA-Z"']*)\s*[xX×]\s*([0-9\/\-\.]+)\s*([a-zA-Z"']*)$/);
    if (psm) { pW = psm[1]||''; pWU = psm[2]||''; pL = psm[3]||''; pLU = psm[4]||''; }
    else { pW = String(sp.parentSize); }
  }

  // Parse child dimensions — prefer explicit fields, fall back to childSize string
  var cW  = String(sp.childWidth      || '').trim();
  var cWU = String(sp.childWidthUnit  || '').trim();
  var cL  = String(sp.childLength     || '').trim();
  var cLU = String(sp.childLengthUnit || '').trim();
  if (!cW && !cL && sp.childSize) {
    var csm = String(sp.childSize).match(/^([0-9\/\-\.]+)\s*([a-zA-Z"']*)\s*[xX×]\s*([0-9\/\-\.]+)\s*([a-zA-Z"']*)$/);
    if (csm) { cW = csm[1]||''; cWU = csm[2]||''; cL = csm[3]||''; cLU = csm[4]||''; }
    else { cW = String(sp.childSize); }
  }

  var pSizeStr = (pW && pL) ? (pW + (pWU?' '+pWU:'') + ' × ' + pL + (pLU?' '+pLU:'')) : (sp.parentSize || '');
  var cSizeStr = (cW && cL) ? (cW + (cWU?' '+cWU:'') + ' × ' + cL + (cLU?' '+cLU:'')) : (sp.childSize  || '');

  var spParentId   = _genId('spp');
  var spChildId    = _genId('spc');

  // SPLIT-PARENT — parent roll consumed (treated as a withdrawal from balance
  // via frontend parseSplitParentRows merging into STATE.withdrawals)
  _getSplitParentSheet().appendRow([
    spDate, splitId, pCode, pDesc, pW, pWU, pL, pLU,
    pQty, whdlNo, customer, note, spParentId, false, '', ''
  ]);

  // SPLIT-CHILD — child rolls produced (treated as received items added to
  // balance via frontend parseSplitChildRows merging into STATE.received)
  _getSplitChildSheet().appendRow([
    spDate, splitId, cCode, cDesc, cW, cWU, cL, cLU,
    cQty, splitId, 'Internal — Roll Split', note, spChildId, false, '', ''
  ]);

  return _json({ok:true, splitId:splitId, splitParentId:spParentId, splitChildId:spChildId});
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
// Returns [{code, desc}] from CODEMAP col A (code) and col B (description).
function _getCodemapCodes() {
  var ss = _openSS();
  var sh = ss.getSheetByName(CODEMAP_TAB);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
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
function _getSplitParentSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(SPLIT_PARENT_TAB);
  if (!sh) { sh = ss.insertSheet(SPLIT_PARENT_TAB); sh.appendRow(SPLIT_PARENT_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(SPLIT_PARENT_HEADER);
  return sh;
}
function _getSplitChildSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(SPLIT_CHILD_TAB);
  if (!sh) { sh = ss.insertSheet(SPLIT_CHILD_TAB); sh.appendRow(SPLIT_CHILD_HEADER); return sh; }
  if (sh.getLastRow() === 0) sh.appendRow(SPLIT_CHILD_HEADER);
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

// =====================================================================
// STOCK MONITOR SHEET — Add-on for JHONG Withdrawal System v3.34
// =====================================================================
// HOW TO USE:
//   1. Open your JHONG BACKEND Google Spreadsheet.
//   2. Go to Extensions → Apps Script.
//   3. Paste this entire file's content at the BOTTOM of your existing
//      Code.gs (after the last line).
//   4. Save (💾).
//   5. Run `buildStockMonitorSheet` once manually:
//        → In the Apps Script editor, select function "buildStockMonitorSheet"
//          from the dropdown and click ▶ Run.
//      This creates the STOCKS_MONITOR sheet with all headers and formulas.
//   6. Set up auto-refresh (optional but recommended):
//        → Run `setupStockMonitorTrigger` once to schedule an hourly refresh.
//      Or just call `refreshStockMonitorSheet` any time you want a manual
//      refresh (it re-reads all live data and rewrites the sheet).
//
// HOW IT UPDATES IN REAL TIME:
//   The STOCKS_MONITOR sheet uses two update mechanisms:
//     a) Native Google Sheets IMPORTRANGE / named-range cross-tab formulas
//        for the TOTAL YARDS and PARTIAL ROLLS summary rows at the bottom.
//        These recalculate automatically whenever source data changes.
//     b) The Apps Script function `refreshStockMonitorSheet` rebuilds the
//        per-item stock rows on demand.  Install the time-based trigger
//        (via setupStockMonitorTrigger) to auto-refresh every hour, or call
//        it from doGet/doPost after any save action (see the hook below).
//
// FORMULA USED (mirrors the frontend exactly):
//   Available = origQty (ITEMCODE col G) + beginning + received − withdrawn
//   where:
//     beginning = sum of non-deleted BEGINNING rows for that item code
//     received  = sum of non-deleted RECEIVED rows
//                 + sum of non-deleted SPLIT-CHILD rows (child rolls produced
//                   by splits — uses new Width / W-UM / Length / L-UM format,
//                   qty at col I, Deleted at col N)
//     withdrawn = sum of non-deleted WITHDRAWAL rows
//                 + sum of non-deleted SPLIT-PARENT rows (parent rolls consumed
//                   by splits — same new column layout)
// =====================================================================

var STOCK_MONITOR_TAB = 'STOCKS_MONITOR';

// ── Public entry points ──────────────────────────────────────────────

/**
 * buildStockMonitorSheet()
 * Creates (or fully rebuilds) the STOCKS_MONITOR tab.
 * Safe to run multiple times — it clears and rewrites every time.
 * Call this once from the Apps Script editor after pasting this file.
 */
function buildStockMonitorSheet() {
  var ss = _openSS();

  // Create or clear the sheet
  var sh = ss.getSheetByName(STOCK_MONITOR_TAB);
  if (!sh) {
    sh = ss.insertSheet(STOCK_MONITOR_TAB);
    // Move it right after ITEMCODE for visibility
    var icIdx = ss.getSheetByName(ITEMCODE_TAB)
      ? ss.getSheetByName(ITEMCODE_TAB).getIndex()
      : 1;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(icIdx + 1);
  } else {
    sh.clearContents();
    sh.clearFormats();
    sh.clearNotes();
  }

  _writeStockMonitorContent(ss, sh);
  SpreadsheetApp.flush();
  Logger.log('STOCKS_MONITOR sheet built successfully.');
}

/**
 * refreshStockMonitorSheet()
 * Re-reads all inventory data and rewrites only the data rows
 * (keeps headers/footer intact). Call this from a time trigger or
 * hook it into doPost after any save action.
 */
function refreshStockMonitorSheet() {
  var ss = _openSS();
  var sh = ss.getSheetByName(STOCK_MONITOR_TAB);
  if (!sh) {
    buildStockMonitorSheet();
    return;
  }
  _writeStockMonitorContent(ss, sh);
  SpreadsheetApp.flush();
}

/**
 * setupStockMonitorTrigger()
 * Installs an hourly time-based trigger that calls refreshStockMonitorSheet.
 * Run this once from the Apps Script editor.
 * (Deletes any previous trigger for the same function first.)
 */
function setupStockMonitorTrigger() {
  // Remove any existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'refreshStockMonitorSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('refreshStockMonitorSheet')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('Hourly trigger for refreshStockMonitorSheet installed.');
}

// ── Core writer ──────────────────────────────────────────────────────

function _writeStockMonitorContent(ss, sh) {
  var tz = Session.getScriptTimeZone() || 'Asia/Manila';

  // ── 1. Read all source data ──────────────────────────────────────

  // ITEMCODE: [ItemCode, Description, Width, WidthUnit, Length, LengthUnit, Qty/Stock, ...]
  var icSh    = _getItemcodeSheet();
  var icData  = icSh.getLastRow() > 1
    ? icSh.getRange(2, 1, icSh.getLastRow() - 1, Math.max(7, icSh.getLastColumn())).getValues()
    : [];

  // BEGINNING: [Date, ItemCode, Desc, Width, W-UM, Length, L-UM, Qty, Notes, Deleted, ...]
  var begSh   = _getBeginningSheet();
  var begData = begSh.getLastRow() > 1
    ? begSh.getRange(2, 1, begSh.getLastRow() - 1, begSh.getLastColumn()).getValues()
    : [];

  // RECEIVED: [Date, ItemCode, Desc, Size, Qty, MRR, Supplier, Remarks, ID, Deleted, ...]
  var recSh   = _getReceivedSheet();
  var recData = recSh.getLastRow() > 1
    ? recSh.getRange(2, 1, recSh.getLastRow() - 1, recSh.getLastColumn()).getValues()
    : [];

  // WITHDRAWAL: [Date, ItemCode, Desc, Size, Qty, WdNo, Customer, Remarks, ID, Deleted, ...]
  var whdSh   = _getWithdrawalSheet();
  var whdData = whdSh.getLastRow() > 1
    ? whdSh.getRange(2, 1, whdSh.getLastRow() - 1, whdSh.getLastColumn()).getValues()
    : [];

  // YARDSRECEIVED: [Date, ItemCode, Desc, Size, Qty, MRR, Supplier, Remarks, ID, Deleted, ...]
  var yrSh    = _getYardsReceivedSheet();
  var yrData  = yrSh.getLastRow() > 1
    ? yrSh.getRange(2, 1, yrSh.getLastRow() - 1, yrSh.getLastColumn()).getValues()
    : [];

  // YARDSWITHDRAWN: [Date, CoreNo, ItemID, Width, Yards, Ref, Customer, ID, Deleted, ...]
  var ywSh    = _getYardsWithdrawnSheet();
  var ywData  = ywSh.getLastRow() > 1
    ? ywSh.getRange(2, 1, ywSh.getLastRow() - 1, ywSh.getLastColumn()).getValues()
    : [];

  // PARTIALROLLS: [Date, ItemCode, Width, WidthUnit, Length, LengthUnit, Qty, Ref, Size, WdID, ID, Deleted, ...]
  var prSh    = _getPartialRollsSheet();
  var prData  = prSh.getLastRow() > 1
    ? prSh.getRange(2, 1, prSh.getLastRow() - 1, prSh.getLastColumn()).getValues()
    : [];

  // PARTIALWITHDRAW: [Date, ItemCode, Width, WidthUnit, Length, LengthUnit, Qty, Ref, Size, Note, ID, Deleted, ...]
  var pwSh    = _getPartialWithdrawSheet();
  var pwData  = pwSh.getLastRow() > 1
    ? pwSh.getRange(2, 1, pwSh.getLastRow() - 1, pwSh.getLastColumn()).getValues()
    : [];

  // SPLIT-PARENT: [Date(0), SplitID(1), ItemCode(2), Description(3), Width(4),
  //   W-UM(5), Length(6), L-UM(7), Qty(8), WithdrawalNo(9), Customer(10),
  //   Note(11), ID(12), Deleted(13), DeletedAt(14), DeletedBy(15)]
  // NEW separate-dimension format — replaces the old single-Size column.
  var spSh    = _getSplitParentSheet();
  var spData  = spSh.getLastRow() > 1
    ? spSh.getRange(2, 1, spSh.getLastRow() - 1, spSh.getLastColumn()).getValues()
    : [];

  // SPLIT-CHILD: [Date(0), SplitID(1), ItemCode(2), Description(3), Width(4),
  //   W-UM(5), Length(6), L-UM(7), Qty(8), MRRNo(9), Supplier(10),
  //   Note(11), ID(12), Deleted(13), DeletedAt(14), DeletedBy(15)]
  // NEW separate-dimension format — replaces the old single-Size column.
  var scSh    = _getSplitChildSheet();
  var scData  = scSh.getLastRow() > 1
    ? scSh.getRange(2, 1, scSh.getLastRow() - 1, scSh.getLastColumn()).getValues()
    : [];

  // ── 2. Build lookup maps ─────────────────────────────────────────

  // Beginning: sum qty per itemCode (col B=1, qty=col H=7, Deleted=col J=9)
  var begMap = {};
  begData.forEach(function(r) {
    var code    = String(r[1] || '').trim();
    var deleted = r[9];                            // col J  (1-based col 10)
    if (!code || deleted === true || String(deleted).toLowerCase() === 'true') return;
    begMap[code] = (begMap[code] || 0) + (parseFloat(r[7]) || 0);
  });

  // Received: sum qty per itemCode (col B=1, qty=col E=4, Deleted=col J=9)
  // Plus SPLIT-CHILD rolls produced by splits (new W/W-UM/L/L-UM format,
  // qty in col I=8, Deleted in col N=13).
  var recMap = {};
  recData.forEach(function(r) {
    var code    = String(r[1] || '').trim();
    var deleted = r[9];
    if (!code || deleted === true || String(deleted).toLowerCase() === 'true') return;
    recMap[code] = (recMap[code] || 0) + (parseFloat(r[4]) || 0);
  });
  scData.forEach(function(r) {
    var code    = String(r[2] || '').trim();       // col C — ItemCode
    var deleted = r[13];                            // col N — Deleted
    if (!code || deleted === true || String(deleted).toLowerCase() === 'true') return;
    recMap[code] = (recMap[code] || 0) + (parseFloat(r[8]) || 0);  // col I — Qty
  });

  // Withdrawn: sum qty per itemCode (col B=1, qty=col E=4, Deleted=col J=9)
  // Plus SPLIT-PARENT rolls consumed by splits (new W/W-UM/L/L-UM format,
  // qty in col I=8, Deleted in col N=13).
  var whdMap = {};
  whdData.forEach(function(r) {
    var code    = String(r[1] || '').trim();
    var deleted = r[9];
    if (!code || deleted === true || String(deleted).toLowerCase() === 'true') return;
    whdMap[code] = (whdMap[code] || 0) + (parseFloat(r[4]) || 0);
  });
  spData.forEach(function(r) {
    var code    = String(r[2] || '').trim();       // col C — ItemCode
    var deleted = r[13];                            // col N — Deleted
    if (!code || deleted === true || String(deleted).toLowerCase() === 'true') return;
    whdMap[code] = (whdMap[code] || 0) + (parseFloat(r[8]) || 0);  // col I — Qty
  });

  // ── 3. ITEMCODE column detection (mirrors frontend parser) ───────
  // Frontend looks for a column whose header contains qty/stock/available/quantity
  // Fallback: col G (index 6) which is the standard 7th column.
  var icHeader = icSh.getLastRow() > 0
    ? icSh.getRange(1, 1, 1, Math.max(7, icSh.getLastColumn())).getValues()[0]
    : [];
  var origQtyIdx = 6; // default col G (0-based index 6)
  for (var hi = 0; hi < icHeader.length; hi++) {
    var hh = String(icHeader[hi] || '').toLowerCase();
    if (hh.indexOf('qty') !== -1 || hh.indexOf('stock') !== -1 ||
        hh.indexOf('available') !== -1 || hh.indexOf('quantity') !== -1 ||
        hh.indexOf('on hand') !== -1) {
      origQtyIdx = hi;
      break;
    }
  }

  // ── 4. Compute per-item stock rows ───────────────────────────────
  var itemRows = [];
  icData.forEach(function(r) {
    var code = String(r[0] || '').trim();
    if (!code) return;
    var desc     = String(r[1] || '').trim();
    var width    = r[2] !== undefined ? String(r[2]).trim() : '';
    var wUnit    = r[3] !== undefined ? String(r[3]).trim() : '';
    var length   = r[4] !== undefined ? String(r[4]).trim() : '';
    var lUnit    = r[5] !== undefined ? String(r[5]).trim() : '';
    var size     = (width && length) ? (width + ' x ' + length) : 'n/a';
    var origQty  = parseFloat(r[origQtyIdx] || 0) || 0;
    var beg      = begMap[code]  || 0;
    var rec      = recMap[code]  || 0;
    var whd      = whdMap[code]  || 0;
    var avail    = origQty + beg + rec - whd;

    var status;
    if (avail <= 0)      status = 'Out of Stock';
    else if (avail <= 3) status = 'Low Stock';
    else                 status = 'In Stock';

    itemRows.push([code, desc, size, origQty, beg, rec, whd, avail, status]);
  });

  // ── 5. Yards stock summary ───────────────────────────────────────
  // Group YARDSRECEIVED by itemCode; subtract YARDSWITHDRAWN (matched by CoreNo/ItemID)
  // YARDSRECEIVED: col B=itemCode(1), col E=qty yards(4), Deleted col J(9)
  var yrByItem = {};
  yrData.forEach(function(r) {
    var code    = String(r[1] || '').trim();
    var deleted = r[9];
    if (!code || deleted === true || String(deleted).toLowerCase() === 'true') return;
    if (!yrByItem[code]) yrByItem[code] = {received: 0, withdrawn: 0};
    yrByItem[code].received += parseFloat(r[4]) || 0;
  });

  // YARDSWITHDRAWN: col C=itemId(2), col E=yards(4), Deleted col I(8)
  ywData.forEach(function(r) {
    var code    = String(r[2] || '').trim();   // ItemID column
    var deleted = r[8];
    if (!code || deleted === true || String(deleted).toLowerCase() === 'true') return;
    if (!yrByItem[code]) yrByItem[code] = {received: 0, withdrawn: 0};
    yrByItem[code].withdrawn += parseFloat(r[4]) || 0;
  });

  var yardsRows = [];
  var totalYardsRec = 0, totalYardsWd = 0;
  Object.keys(yrByItem).sort().forEach(function(code) {
    var d       = yrByItem[code];
    var balance = d.received - d.withdrawn;
    totalYardsRec += d.received;
    totalYardsWd  += d.withdrawn;
    yardsRows.push([code, d.received, d.withdrawn, balance]);
  });

  // ── 6. Partial rolls summary ─────────────────────────────────────
  // PARTIALROLLS: Deleted col L(11). Count active rows per itemCode.
  // PARTIALWITHDRAW: Deleted col L(11). Count active withdrawals per itemCode.
  var prByItem = {};
  prData.forEach(function(r) {
    var code    = String(r[1] || '').trim();
    var deleted = r[11];
    if (!code || deleted === true || String(deleted).toLowerCase() === 'true') return;
    if (!prByItem[code]) prByItem[code] = {received: 0, withdrawn: 0};
    prByItem[code].received += parseFloat(r[6]) || 1; // qty col G(6), default 1 roll
  });
  pwData.forEach(function(r) {
    var code    = String(r[1] || '').trim();
    var deleted = r[11];
    if (!code || deleted === true || String(deleted).toLowerCase() === 'true') return;
    if (!prByItem[code]) prByItem[code] = {received: 0, withdrawn: 0};
    prByItem[code].withdrawn += parseFloat(r[6]) || 1;
  });

  var partialRows = [];
  var totalPrIn = 0, totalPrOut = 0;
  Object.keys(prByItem).sort().forEach(function(code) {
    var d       = prByItem[code];
    var balance = d.received - d.withdrawn;
    totalPrIn  += d.received;
    totalPrOut += d.withdrawn;
    partialRows.push([code, d.received, d.withdrawn, balance]);
  });

  // ── 7. Assemble all output rows ──────────────────────────────────
  var now         = Utilities.formatDate(new Date(), tz, 'MMM dd, yyyy  HH:mm:ss');
  var allRows     = [];
  var formats     = [];   // parallel array of background colours
  var fontColors  = [];
  var fontWeights = [];
  var hAligns     = [];   // horizontal alignment per row

  function row(cells, bg, fc, fw, ha) {
    allRows.push(cells);
    formats.push(bg    || '#FFFFFF');
    fontColors.push(fc  || '#000000');
    fontWeights.push(fw || 'normal');
    hAligns.push(ha    || 'left');
  }

  // ── Title block ──
  row(['JHONG WITHDRAWAL SYSTEM — STOCK MONITOR', '', '', '', '', '', '', '', ''],
      '#1B5E20', '#FFFFFF', 'bold', 'left');
  row(['Last refreshed: ' + now, '', '', '', '', '', '', '', ''],
      '#E8F5E9', '#2E7D32', 'normal', 'left');
  row([], '#FFFFFF');  // blank spacer

  // ── Section 1: MAIN STOCK ON HAND ──
  row(['■  MAIN STOCK — ON HAND  (formula: origQty + beginning + received − withdrawn)',
       '', '', '', '', '', '', '', ''],
      '#2E7D32', '#FFFFFF', 'bold', 'left');
  row(['Item Code', 'Description', 'Size', 'Orig Qty', 'Beginning', 'Received',
       'Withdrawn', 'Available', 'Status'],
      '#A5D6A7', '#1B5E20', 'bold', 'center');

  var mainStartRow = allRows.length + 1; // 1-based, will be used for totals formula
  var totalAvail = 0;
  itemRows.forEach(function(r) {
    var avail   = r[7];
    var status  = r[8];
    var rowBg   = (allRows.length % 2 === 0) ? '#F9FBE7' : '#FFFFFF';
    var statFc  = status === 'In Stock'    ? '#2E7D32'
                : status === 'Low Stock'   ? '#E65100'
                :                           '#C62828';
    // Columns: Code, Desc, Size, OrigQty, Beg, Rec, Whd, Avail, Status
    allRows.push(r);
    formats.push(rowBg);
    fontColors.push('#000000');
    fontWeights.push('normal');
    hAligns.push('left');
    // Override avail colour (col H = index 7) and status (col I = index 8) below
    totalAvail += avail;
  });
  var mainEndRow = allRows.length;

  // Totals row for main section
  row(['', 'TOTAL ITEMS: ' + itemRows.length, '', '', '', '', '', totalAvail, ''],
      '#C8E6C9', '#1B5E20', 'bold', 'right');
  row([], '#FFFFFF');  // blank spacer

  // ── Section 2: YARDS STOCK ──
  row(['▲  TOTAL YARDS STOCK  (YARDSRECEIVED − YARDSWITHDRAWN per item)',
       '', '', '', ''],
      '#1565C0', '#FFFFFF', 'bold', 'left');
  row(['Item Code / Item ID', 'Yards Received', 'Yards Withdrawn', 'Balance (yards)'],
      '#BBDEFB', '#0D47A1', 'bold', 'center');

  if (yardsRows.length === 0) {
    row(['(No yards stock data found)', '', '', ''], '#F5F5F5', '#9E9E9E', 'italic', 'left');
  } else {
    yardsRows.forEach(function(r) {
      var balance = r[3];
      var rowBg   = (allRows.length % 2 === 0) ? '#E3F2FD' : '#FFFFFF';
      allRows.push(r);
      formats.push(rowBg);
      fontColors.push(balance > 0 ? '#1565C0' : '#C62828');
      fontWeights.push('normal');
      hAligns.push('left');
    });
  }
  var yardsRec = yardsRows.length > 0 ? totalYardsRec : 0;
  var yardsWd  = yardsRows.length > 0 ? totalYardsWd  : 0;
  row(['TOTALS', yardsRec, yardsWd, yardsRec - yardsWd],
      '#90CAF9', '#0D47A1', 'bold', 'right');
  row([], '#FFFFFF');  // blank spacer

  // ── Section 3: PARTIAL ROLLS STOCK ──
  row(['◆  PARTIAL ROLLS STOCK  (PARTIALROLLS − PARTIALWITHDRAW per item)',
       '', '', '', ''],
      '#6A1B9A', '#FFFFFF', 'bold', 'left');
  row(['Item Code', 'Rolls In (partial)', 'Rolls Out (partial)', 'Balance (rolls)'],
      '#E1BEE7', '#4A148C', 'bold', 'center');

  if (partialRows.length === 0) {
    row(['(No partial rolls data found)', '', '', ''], '#F5F5F5', '#9E9E9E', 'italic', 'left');
  } else {
    partialRows.forEach(function(r) {
      var balance = r[3];
      var rowBg   = (allRows.length % 2 === 0) ? '#F3E5F5' : '#FFFFFF';
      allRows.push(r);
      formats.push(rowBg);
      fontColors.push(balance > 0 ? '#6A1B9A' : '#C62828');
      fontWeights.push('normal');
      hAligns.push('left');
    });
  }
  var prIn  = partialRows.length > 0 ? totalPrIn  : 0;
  var prOut = partialRows.length > 0 ? totalPrOut : 0;
  row(['TOTALS', prIn, prOut, prIn - prOut],
      '#CE93D8', '#4A148C', 'bold', 'right');

  // ── 8. Write everything to the sheet ────────────────────────────
  sh.clearContents();
  sh.clearFormats();

  // Determine column count from widest row
  var maxCols = 9; // main section has 9 cols
  var totalRows = allRows.length;
  if (totalRows === 0) return;

  // Pad every row to maxCols so setValues works
  var padded = allRows.map(function(r) {
    var arr = r.slice(0, maxCols);
    while (arr.length < maxCols) arr.push('');
    return arr;
  });

  sh.getRange(1, 1, totalRows, maxCols).setValues(padded);

  // Apply background colours (col A–I, but rows with only 4 data cols
  // just get the colour on all 9 cols — harmless).
  for (var ri = 0; ri < totalRows; ri++) {
    var rowRange = sh.getRange(ri + 1, 1, 1, maxCols);
    rowRange.setBackground(formats[ri]);
    rowRange.setFontColor(fontColors[ri]);
    rowRange.setFontWeight(fontWeights[ri]);
    if (hAligns[ri] !== 'left') {
      // Right-align numeric columns for data rows, header-center for header rows
      if (hAligns[ri] === 'center') {
        rowRange.setHorizontalAlignment('center');
      } else {
        sh.getRange(ri + 1, 4, 1, 6).setHorizontalAlignment('right');
      }
    }
  }

  // ── Status column colour override for main stock rows (col I = 9) ──
  // Find them by scanning the Status column values
  for (var ri2 = 0; ri2 < padded.length; ri2++) {
    var statusVal = padded[ri2][8];
    if (statusVal === 'In Stock')    sh.getRange(ri2+1, 9).setFontColor('#2E7D32').setFontWeight('bold');
    else if (statusVal === 'Low Stock')   sh.getRange(ri2+1, 9).setFontColor('#E65100').setFontWeight('bold');
    else if (statusVal === 'Out of Stock') sh.getRange(ri2+1, 9).setFontColor('#C62828').setFontWeight('bold');
  }

  // ── Format column widths ──
  sh.setColumnWidth(1, 200);  // Item Code
  sh.setColumnWidth(2, 280);  // Description
  sh.setColumnWidth(3, 110);  // Size
  sh.setColumnWidth(4, 90);   // Orig Qty
  sh.setColumnWidth(5, 90);   // Beginning
  sh.setColumnWidth(6, 90);   // Received
  sh.setColumnWidth(7, 90);   // Withdrawn
  sh.setColumnWidth(8, 95);   // Available
  sh.setColumnWidth(9, 120);  // Status

  // ── Reset frozen state + clear any existing merges from a previous run ──
  // Sheets refuses to merge across the frozen/non-frozen column boundary,
  // so frozen columns must be cleared before we re-apply the title merges.
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.getRange(1, 1, 2, maxCols).breakApart();

  // ── Title row: merge across all 9 cols ──
  sh.getRange(1, 1, 1, maxCols).merge().setFontSize(13);
  sh.getRange(2, 1, 1, maxCols).merge().setFontSize(10).setFontStyle('italic');

  // ── Freeze header rows AFTER the merges land (title + refreshed + blank
  //    + section header = 4 rows, then col header = row 5). ──
  sh.setFrozenRows(5);
  sh.setFrozenColumns(1);

  Logger.log('STOCKS_MONITOR written: ' + totalRows + ' rows, ' +
             itemRows.length + ' items, ' +
             yardsRows.length + ' yard items, ' +
             partialRows.length + ' partial-roll items.');
}

// =====================================================================
// OPTIONAL: Hook refreshStockMonitorSheet into every write action
// =====================================================================
// If you want the sheet to update automatically whenever ANY save
// action runs (not just on a timer), add this one line to the bottom
// of every `if (action === 'save_*')` block in _handleWrite():
//
//   refreshStockMonitorSheet();
//   return _json({ok:true, ...});
//
// For performance, you can instead call it only after the actions that
// affect stock (save_withdrawal, delete_withdrawal, save_received,
// delete_received, save_yardsreceived, delete_yardsreceived,
// save_yardswithdrawn, delete_yardswithdrawn, save_beginning_single,
// save_beginning_bulk, save_partialroll, delete_partialroll).
//
// =====================================================================
// MIGRATION: Split legacy SIZE column into Width / W-UM / Length / L-UM
// for the BEGINNING sheet (mirrors migrateWithdrawalSizeColumn)
// =====================================================================
// Run this ONCE from the Apps Script editor after deploying the updated
// BEGINNING_HEADER.
//
// HOW TO RUN:
//   1. Open your JHONG BACKEND spreadsheet → Extensions → Apps Script.
//   2. Select function "migrateBeginningSheet" from the dropdown.
//   3. Click ▶ Run.  Check the Execution Log for results.
// =====================================================================
function migrateBeginningSheet() {
  var sh = _getBeginningSheet();
  var lastRow = sh.getLastRow();
  if (lastRow < 1) { Logger.log('BEGINNING sheet is empty — nothing to migrate.'); return; }

  // Read header row
  var headerRange = sh.getRange(1, 1, 1, sh.getLastColumn());
  var header = headerRange.getValues()[0].map(function(h){ return String(h||'').trim(); });

  // Detect layout by inspecting column D header
  var colD = header[3] || ''; // 0-based index 3 = column D
  if (colD === 'Width') {
    Logger.log('BEGINNING sheet already uses the new Width/W-UM/Length/L-UM layout. No migration needed.');
    return;
  }

  // Old layout: [Date, ItemCode, Description, Size, Qty, Notes, Deleted, DeletedAt, DeletedBy]
  // New layout: [Date, ItemCode, Description, Width, W-UM, Length, L-UM, Qty, Notes, Deleted, DeletedAt, DeletedBy]
  // Strategy:
  //   1. Insert 3 blank columns after column D (position 4) to make room for W-UM, Length, L-UM
  //   2. Write new header row
  //   3. Parse each data row's old Size (now col D) into Width+W-UM+Length+L-UM

  Logger.log('Migrating ' + (lastRow - 1) + ' BEGINNING rows from SIZE → Width/W-UM/Length/L-UM...');

  // Step 1: Insert 3 columns after column D (col 4, 1-based)
  sh.insertColumnsAfter(4, 3);  // inserts cols E, F, G → old E-I shift to H-L

  // Step 2: Write the canonical header
  sh.getRange(1, 1, 1, BEGINNING_HEADER.length).setValues([BEGINNING_HEADER]);

  // Step 3: Read all data rows (now expanded after insert)
  if (lastRow < 2) {
    Logger.log('No data rows to migrate.');
    SpreadsheetApp.flush();
    return;
  }
  var totalCols = sh.getLastColumn();
  var data = sh.getRange(2, 1, lastRow - 1, totalCols).getValues();

  var migrated = 0;
  var sizeRe = /^([0-9\/\-\.]+)\s*([a-zA-Z"']*)\s*[xX×]\s*([0-9\/\-\.]+)\s*([a-zA-Z"']*)$/;

  data.forEach(function(row, ri) {
    var sizeVal = String(row[3] || '').trim();  // col D (0-based 3) = old Size
    if (!sizeVal) return;                        // nothing to parse
    var width = '', wUnit = '', length = '', lUnit = '';
    var m = sizeVal.match(sizeRe);
    if (m) {
      width  = m[1] || ''; wUnit  = m[2] || '';
      length = m[3] || ''; lUnit  = m[4] || '';
    } else {
      width = sizeVal;  // unparseable — store whole string in Width
    }
    // Write Width(D), W-UM(E), Length(F), L-UM(G) — columns 4,5,6,7 (1-based)
    sh.getRange(ri + 2, 4, 1, 4).setValues([[width, wUnit, length, lUnit]]);
    migrated++;
  });

  SpreadsheetApp.flush();
  Logger.log('BEGINNING migration complete. ' + migrated + ' rows updated.');
}

// =====================================================================
// MIGRATION: Split legacy SIZE column into Width / W-UM / Length / L-UM
// =====================================================================
// Run this ONCE from the Apps Script editor after deploying the updated
// WITHDRAWAL_HEADER.  It reads every row in the WITHDRAWAL sheet, checks
// whether columns D-G match the NEW header (Width/W-UM/Length/L-UM), and
// if the sheet still uses the old layout (Size in col D) it migrates each
// row in place.
//
// HOW TO RUN:
//   1. Open your JHONG BACKEND spreadsheet → Extensions → Apps Script.
//   2. Select function "migrateWithdrawalSizeColumn" from the dropdown.
//   3. Click ▶ Run.  Check the Execution Log for results.
// =====================================================================
function migrateWithdrawalSizeColumn() {
  var sh = _getWithdrawalSheet();
  var lastRow = sh.getLastRow();
  if (lastRow < 1) { Logger.log('WITHDRAWAL sheet is empty — nothing to migrate.'); return; }

  // Read header row
  var headerRange = sh.getRange(1, 1, 1, sh.getLastColumn());
  var header = headerRange.getValues()[0].map(function(h){ return String(h||'').trim(); });

  // Detect layout by inspecting column D header
  var colD = header[3] || ''; // 0-based index 3 = column D
  if (colD === 'Width') {
    Logger.log('WITHDRAWAL sheet already uses the new Width/W-UM/Length/L-UM layout. No migration needed.');
    return;
  }

  // Old layout: [Date, ItemCode, Description, Size, Qty, WithdrawalNo, Customer, Remarks, ID, Deleted, DeletedAt, DeletedBy]
  // New layout: [Date, ItemCode, Description, Width, W-UM, Length, L-UM, Qty, WithdrawalNo, Customer, Remarks, ID, Deleted, DeletedAt, DeletedBy]
  // Strategy:
  //   1. Insert 3 blank columns after column D (position 4) to make room for W-UM, Length, L-UM
  //   2. Write new header row
  //   3. Parse each data row's old Size (now col D) into Width+W-UM+Length+L-UM

  Logger.log('Migrating ' + (lastRow - 1) + ' WITHDRAWAL rows from SIZE → Width/W-UM/Length/L-UM...');

  // Step 1: Insert 3 columns after column D (col 4, 1-based)
  sh.insertColumnsAfter(4, 3);  // inserts cols E, F, G → old E-L shift to H-O

  // Step 2: Write the canonical header
  sh.getRange(1, 1, 1, WITHDRAWAL_HEADER.length).setValues([WITHDRAWAL_HEADER]);

  // Step 3: Read all data rows (now columns A–O after insert)
  if (lastRow < 2) {
    Logger.log('No data rows to migrate.');
    SpreadsheetApp.flush();
    return;
  }
  var totalCols = sh.getLastColumn();
  var data = sh.getRange(2, 1, lastRow - 1, totalCols).getValues();

  var migrated = 0;
  var sizeRe = /^([0-9\/\-\.]+)\s*([a-zA-Z"']*)\s*[xX×]\s*([0-9\/\-\.]+)\s*([a-zA-Z"']*)$/;

  data.forEach(function(row, ri) {
    var sizeVal = String(row[3] || '').trim();  // col D (0-based 3) = old Size
    if (!sizeVal) return;                        // nothing to parse
    var width = '', wUnit = '', length = '', lUnit = '';
    var m = sizeVal.match(sizeRe);
    if (m) {
      width  = m[1] || ''; wUnit  = m[2] || '';
      length = m[3] || ''; lUnit  = m[4] || '';
    } else {
      width = sizeVal;  // unparseable — store whole string in Width
    }
    // Write Width(D), W-UM(E), Length(F), L-UM(G) — columns 4,5,6,7 (1-based)
    sh.getRange(ri + 2, 4, 1, 4).setValues([[width, wUnit, length, lUnit]]);
    migrated++;
  });

  SpreadsheetApp.flush();
  Logger.log('Migration complete. ' + migrated + ' rows updated.');
}
