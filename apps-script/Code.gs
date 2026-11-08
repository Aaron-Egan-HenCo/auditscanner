/**
 * ScanAudit backend — Google Apps Script web app bound to a Google Sheet.
 *
 * Sheets (created automatically on first run):
 *   Audits : one row per saved audit session (summary + master/missing lists)
 *   Scans  : one row per scanned serial, keyed by AuditID
 *
 * Endpoints (deployed as Web App, "Execute as: Me", "Who has access: Anyone"):
 *   GET  ?action=ping
 *   GET  ?action=list[&limit=200]        -> audit summaries, newest first
 *   GET  ?action=get&id=<AuditID>        -> audit summary + its scans + missing list
 *   GET  ?action=all[&since=<ISO date>]  -> every scan row (for bulk export)
 *   POST (text/plain JSON body) {action:"save", audit:{...}, scans:[...]}
 *        Re-saving an existing AuditID replaces that audit's rows (no duplicates).
 */

const AUDITS_SHEET = 'Audits';
const SCANS_SHEET = 'Scans';

const AUDIT_HEADERS = [
  'AuditID', 'SavedAt', 'StartedAt', 'ScannedBy', 'Location', 'Notes',
  'MasterCount', 'Found', 'Extra', 'Missing', 'TotalScanned',
  'MasterList', 'MissingList', 'SaveCount'
];
const SCAN_HEADERS = [
  'AuditID', 'Serial', 'Status', 'ScannedAt', 'ScannedBy', 'Location'
];

// ---------- HTTP handlers ----------

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = p.action || 'ping';
    if (action === 'ping') return json_({ ok: true, ts: new Date().toISOString() });
    if (action === 'list') return json_({ ok: true, audits: listAudits_(parseInt(p.limit, 10) || 200) });
    if (action === 'get') return json_(getAudit_(p.id));
    if (action === 'many') return json_({ ok: true, results: getManyAudits_(String(p.ids || '').split(',').filter(Boolean)) });
    if (action === 'all') return json_({ ok: true, scans: allScans_(p.since) });
    return json_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    if (body.action !== 'save') return json_({ ok: false, error: 'Unknown action' });
    return json_(saveAudit_(body.audit, body.scans || []));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ---------- Core ----------

function saveAudit_(audit, scans) {
  if (!audit || !audit.auditId) throw new Error('Missing auditId');
  if (!audit.scannedBy || !String(audit.scannedBy).trim()) throw new Error('Missing scannedBy');

  const auditsSh = sheet_(AUDITS_SHEET, AUDIT_HEADERS);
  const scansSh = sheet_(SCANS_SHEET, SCAN_HEADERS);
  const id = String(audit.auditId);
  audit.scannedBy = String(audit.scannedBy).trim().replace(/\s+/g, ' ');
  audit.location = String(audit.location || '').trim();
  const savedAt = new Date();

  // Remove any existing scan rows for this audit (re-save = replace)
  removeRowsById_(scansSh, id);

  // Append scan rows
  if (scans.length) {
    const rows = scans.map(s => [
      id,
      text_(s.code),
      s.status === 'expected' ? 'MATCH' : s.status === 'added' ? 'ADDED' : 'EXTRA',
      date_(s.timestamp),
      audit.scannedBy,
      audit.location || ''
    ]);
    const start = scansSh.getLastRow() + 1;
    scansSh.getRange(start, 2, rows.length, 1).setNumberFormat('@'); // Serial column as text
    scansSh.getRange(start, 1, rows.length, SCAN_HEADERS.length).setValues(rows);
  }

  // Upsert audit summary row
  const existingRow = findRowById_(auditsSh, id);
  const prevCount = existingRow ? Number(auditsSh.getRange(existingRow, AUDIT_HEADERS.indexOf('SaveCount') + 1).getValue()) || 0 : 0;
  const summary = [
    id,
    savedAt,
    date_(audit.startedAt) || savedAt,
    audit.scannedBy,
    audit.location || '',
    audit.notes || '',
    Number(audit.masterCount) || 0,
    Number(audit.found) || 0,
    Number(audit.extra) || 0,
    Number(audit.missing) || 0,
    scans.length,
    (audit.masterList || []).join(';'),
    (audit.missingList || []).join(';'),
    prevCount + 1
  ];
  if (existingRow) {
    auditsSh.getRange(existingRow, 1, 1, summary.length).setValues([summary]);
  } else {
    auditsSh.appendRow(summary);
  }

  return { ok: true, auditId: id, savedAt: savedAt.toISOString(), rows: scans.length, updated: !!existingRow };
}

function listAudits_(limit) {
  const sh = sheet_(AUDITS_SHEET, AUDIT_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, AUDIT_HEADERS.length).getValues();
  return values
    .filter(r => r[0])
    .map(r => ({
      auditId: String(r[0]),
      savedAt: iso_(r[1]),
      startedAt: iso_(r[2]),
      scannedBy: r[3],
      location: r[4],
      notes: r[5],
      masterCount: r[6],
      found: r[7],
      extra: r[8],
      missing: r[9],
      totalScanned: r[10],
      saveCount: r[13]
    }))
    .sort((a, b) => (b.savedAt || '').localeCompare(a.savedAt || ''))
    .slice(0, limit);
}

function auditFromRow_(r) {
  return {
    auditId: String(r[0]), savedAt: iso_(r[1]), startedAt: iso_(r[2]),
    scannedBy: r[3], location: r[4], notes: r[5],
    masterCount: r[6], found: r[7], extra: r[8], missing: r[9], totalScanned: r[10],
    missingList: String(r[12] || '').split(';').filter(Boolean)
  };
}

function getAudit_(id) {
  if (!id) return { ok: false, error: 'Missing id' };
  const auditsSh = sheet_(AUDITS_SHEET, AUDIT_HEADERS);
  const row = findRowById_(auditsSh, id);
  if (!row) return { ok: false, error: 'Audit not found: ' + id };
  const audit = auditFromRow_(auditsSh.getRange(row, 1, 1, AUDIT_HEADERS.length).getValues()[0]);
  const scans = allScans_().filter(s => s.auditId === String(id));
  return { ok: true, audit: audit, scans: scans };
}

/** Several audits in one call; reads each sheet once. */
function getManyAudits_(ids) {
  const want = new Set(ids.map(String));
  const auditsSh = sheet_(AUDITS_SHEET, AUDIT_HEADERS);
  const last = auditsSh.getLastRow();
  const audits = last < 2 ? [] : auditsSh.getRange(2, 1, last - 1, AUDIT_HEADERS.length).getValues()
    .filter(r => want.has(String(r[0])))
    .map(auditFromRow_);
  const scansById = {};
  allScans_().forEach(s => {
    if (!want.has(s.auditId)) return;
    (scansById[s.auditId] = scansById[s.auditId] || []).push(s);
  });
  return audits.map(a => ({ ok: true, audit: a, scans: scansById[a.auditId] || [] }));
}

function allScans_(since) {
  const sh = sheet_(SCANS_SHEET, SCAN_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const sinceTs = since ? new Date(since).getTime() : 0;
  return sh.getRange(2, 1, last - 1, SCAN_HEADERS.length).getValues()
    .filter(r => r[0])
    .map(r => ({
      auditId: String(r[0]),
      serial: String(r[1]),
      status: r[2],
      scannedAt: iso_(r[3]),
      scannedBy: r[4],
      location: r[5]
    }))
    .filter(s => !sinceTs || (s.scannedAt && new Date(s.scannedAt).getTime() >= sinceTs));
}

// ---------- Helpers ----------

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (name === SCANS_SHEET) sh.getRange('B:B').setNumberFormat('@');
  }
  return sh;
}

function findRowById_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function removeRowsById_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const width = sh.getLastColumn();
  const data = sh.getRange(2, 1, last - 1, width).getValues();
  const keep = data.filter(r => String(r[0]) !== String(id));
  if (keep.length === data.length) return;
  sh.getRange(2, 1, last - 1, width).clearContent();
  if (keep.length) {
    sh.getRange(2, 2, keep.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, keep.length, width).setValues(keep);
  }
}

function text_(v) { return String(v == null ? '' : v).trim(); }

function iso_(v) {
  if (!v) return '';
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString();
  return String(v);
}

function date_(v) {
  if (!v) return '';
  const d = new Date(v);
  return isNaN(d.getTime()) ? '' : d;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor to create the sheets and grant permissions. */
function setup() {
  sheet_(AUDITS_SHEET, AUDIT_HEADERS);
  sheet_(SCANS_SHEET, SCAN_HEADERS);
}
