/**
 * ============================================================================
 *  APPROVAL PORTAL — Server-side logic (Code.gs)
 *  Google Apps Script + AppSheet REST API + Google Sign-In fallback
 * ============================================================================
 *
 *  NEW IN THIS REVISION — 2026 UPDATE (see the 7 requirements that produced
 *  this revision)
 *    1. Audit / staging sync — unchanged on the backend (sendStatusToExternal
 *       App / syncApprovalToStaging already carried al_comment/al_attachment
 *       correctly). The actual gap was on the CLIENT: it never called
 *       syncApprovalToStaging() at all. That's fixed in JavaScript.html, not
 *       here — see that file's header comment.
 *    2. Button state / perceived speed — pure frontend change.
 *    3. Real-time sync / blacklist — pure frontend change, PLUS a short
 *       server-side cache on fetchApprovalData() (below) so rapid client
 *       calls (manual Sync clicks, the poll) don't hammer AppSheet.
 *    4. Conditional upload visibility — pure frontend change.
 *    5. Hover nav sidebar — NEW getQuickLinks() below feeds it (app IDs
 *       only, never API keys) + pure frontend markup/CSS/JS.
 *    6. Font sizing — pure frontend (Styles.html) change.
 *    7. Backend performance — this app has no direct SpreadsheetApp/Sheets
 *       access anywhere (everything goes through AppSheet's REST API), so
 *       "batch getValues/setValues" doesn't literally apply here — that note
 *       from the prior revision still stands. What DOES apply and is NEW
 *       this revision:
 *         a. fetchApprovalData() is now backed by a short (8s) CacheService
 *            entry per approver email, so bursts of client calls within
 *            that window (Sync double-clicks, the 30s poll landing right
 *            after a manual sync, etc.) are served from cache instead of
 *            re-hitting AppSheet. Callers can pass forceRefresh=true to
 *            always bypass it (the client does this on initial load and on
 *            manual "Sync now" clicks).
 *         b. updateApprovalStatus() now busts that same cache entry for the
 *            acting approver right after its write succeeds, so a forced
 *            refresh immediately after a decision is guaranteed fresh. This
 *            is defense-in-depth — the primary "don't show me my own
 *            just-actioned request again" job is done client-side by the
 *            blacklist described in JavaScript.html.
 *       Nothing about what any function returns has changed — only how much
 *       AppSheet traffic it costs to get there.
 *
 *  (Prior revision notes retained below.)
 *
 *  PERFORMANCE (prior revision):
 *      1. updateApprovalStatus() used to call sendStatusToExternalApp()
 *         SYNCHRONOUSLY before returning, meaning the client waited on a
 *         second, cross-app AppSheet call it didn't actually need to block
 *         on. That sync is now split into its own function —
 *         syncApprovalToStaging() — which the client fires asynchronously
 *         right after updateApprovalStatus() confirms success. This
 *         function itself is now unchanged in every other way and returns
 *         as soon as the ONE write that matters (CentralizedApproval) is
 *         confirmed.
 *      2. handleApprovalAction() (a single-call convenience wrapper — not
 *         the primary UI path) used to trigger TWO separate row lookups
 *         for the same record: one inside updateApprovalStatus(), and
 *         another of its own before calling sendStatusToExternalApp()
 *         (which would then do a THIRD, internally). sendStatusToExternal-
 *         App() now accepts an optional pre-resolved row so that redundant
 *         third lookup is skipped whenever the caller already has it.
 *      3. resolveManagerAuth_() called appSheetFind_(managers) on every
 *         single auth check — which, for session-based (non-token)
 *         deployments, means EVERY server call re-fetches the managers
 *         table. It's now cached via CacheService for 5 minutes
 *         (getManagersCached_ / bustManagersCache).
 *      4. fetchApprovalData() used to pull every CentralizedApproval row
 *         and filter by AprEmail in Apps Script. It now asks AppSheet to
 *         pre-filter server-side via a Selector (case-insensitively, via
 *         LOWER(...)), with the original Apps-Script-side filter KEPT as a
 *         correctness safety net — if the Selector ever fails to match for
 *         any reason, results are identical to before, just slower for
 *         that one call. Behavior/correctness is unchanged either way.
 *
 *  al_comment / al_attachment ON STAGING TABLES (prior revision, unchanged):
 *    Exactly ONE function builds the row sent to StagingAllReq /
 *    StagingEntries / StagingXC — sendStatusToExternalApp() — and both
 *    fields live in that one shared rowData object, so all three tables get
 *    them identically. See the ★★★ marked block inside
 *    sendStatusToExternalApp() below.
 *
 *  UID-BASED ROW RESOLUTION: resolveApprovalRow_(requestId, uid) tries the
 *    unique `uid` column first, falls back to RequestID. Needed because
 *    XCLB (and potentially XCSB) can produce multiple rows sharing a
 *    RequestID. KNOWN LIMITATION: auditLogs has no uid column, so
 *    fetchAuditTrail()/logAuditEntry() remain RequestID-keyed.
 *  ATTACHMENTS: updateApprovalStatus() accepts an optional attachmentData
 *    for the pre-decision file picker (Drive upload, best-effort).
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// CONFIGURATION
// ---------------------------------------------------------------------------

const TABLES = {
  APPROVALS: 'CentralizedApproval',
  MANAGERS: 'managers',
  AUDIT_LOGS: 'auditLogs',
  STAGINGALLREQ: 'StagingAllReq',
  STAGINGENTRIES: 'StagingEntries',
  STAGINGXC: 'StagingXC'
};

const SESSION_TTL_SECONDS = 6 * 60 * 60; // 6 hours
const SESSION_CACHE_PREFIX = 'auth_session_';

const ROW_KEY_CANDIDATES = ['Row ID', 'RowID', 'uid', 'Uid', 'UID', 'Key'];

const DECISION_ATTACHMENTS_FOLDER_ID = '1FtxMfvW5k45zpCx-TCu7RKR96EyLsqsN';

// managers-table cache — see resolveManagerAuth_ / getManagersCached_.
const MANAGERS_CACHE_KEY = 'managers_directory_cache_v1';
const MANAGERS_CACHE_TTL_SECONDS = 300; // 5 minutes

// NEW (7. Performance): short-lived per-approver cache for fetchApprovalData.
const APPROVAL_DATA_CACHE_PREFIX = 'approval_data_v1_';
const APPROVAL_DATA_CACHE_TTL_SECONDS = 8;

// ---------------------------------------------------------------------------
// STAGING DESTINATION ROUTING
//   HR Request app  (StagingAllReq table)  ← JOB, OSR, AR, PRF
//   HR Connect app  (StagingEntries table) ← LEAVE, TRAF, OBOT
//   XC app          (StagingXC table)      ← XCLB, XCSB
// ---------------------------------------------------------------------------

const STAGING_APP_ID = 'ccba5fbc-1904-44ee-8706-222a618df974';
const STAGING_API_KEY = 'V2-RPuvq-CuScq-mmjAa-vVm3G-sLGep-Y9knD-MG0xh-xRsBV';

const HRCONNECT_APP_ID = '21a28128-1649-41bd-9828-d1db01ea53ef'; 
const HRCONNECT_API_KEY = 'V2-Zx0Wi-h3YZK-cEiid-7mrnk-yjXPc-J9GAH-aW3NM-eaq4X';

const XC_APP_ID = 'd133ff70-94a4-4fc9-8ea9-5fcbe8f81f4f';
const XC_API_KEY = 'V2-rumFG-7tgt8-VCiPD-9wVrq-StLWm-mRprs-VDuzI-ijdmi';

const MASTERLIST_APP_ID = XC_APP_ID;
const MASTERLIST_API_KEY = XC_API_KEY;
const MASTERLIST_TABLE = 'MASTERLIST';
const MASTERLIST_IMAGE_CACHE_PREFIX = 'emp_img_v6_';
const MASTERLIST_IMAGE_CACHE_TTL_SECONDS = 300; // 5 minutes

const STAGING_DESTINATIONS = {
  'JOB':   { appId: STAGING_APP_ID,   apiKey: STAGING_API_KEY,   tableName: TABLES.STAGINGALLREQ },
  'OSR':   { appId: STAGING_APP_ID,   apiKey: STAGING_API_KEY,   tableName: TABLES.STAGINGALLREQ },
  'AR':    { appId: STAGING_APP_ID,   apiKey: STAGING_API_KEY,   tableName: TABLES.STAGINGALLREQ },
  'PRF':   { appId: STAGING_APP_ID,   apiKey: STAGING_API_KEY,   tableName: TABLES.STAGINGALLREQ },
  'LEAVE': { appId: HRCONNECT_APP_ID, apiKey: HRCONNECT_API_KEY, tableName: TABLES.STAGINGENTRIES },
  'TRAF':  { appId: HRCONNECT_APP_ID, apiKey: HRCONNECT_API_KEY, tableName: TABLES.STAGINGENTRIES },
  'OBOT':  { appId: HRCONNECT_APP_ID, apiKey: HRCONNECT_API_KEY, tableName: TABLES.STAGINGENTRIES },
  'XCLB':  { appId: XC_APP_ID,        apiKey: XC_API_KEY,        tableName: TABLES.STAGINGXC },
  'XCSB':  { appId: XC_APP_ID,        apiKey: XC_API_KEY,        tableName: TABLES.STAGINGXC }
};

/**
 * Checks the `managers` table (via the existing 5-minute cache) rather than
 * a separate hardcoded list — managers IS the whitelist.
 */
function isUserAuthorized(email) {
  if (!email) return false;
  return resolveManagerAuth_(email).authorized === true;
}

function resolveStagingDestination_(appSource) {
  const key = (appSource || '').toString().trim().toUpperCase();
  return STAGING_DESTINATIONS[key] || null;
}

function getAppSheetConfig_() {
  const props = PropertiesService.getScriptProperties();
  const appId = props.getProperty('APPSHEET_APP_ID');
  const apiKey = props.getProperty('APPSHEET_API_KEY');

  if (!appId || !apiKey) {
    throw new Error(
      'AppSheet credentials are not configured. Set APPSHEET_APP_ID and ' +
      'APPSHEET_API_KEY in Project Settings → Script Properties.'
    );
  }
  return { appId: appId, apiKey: apiKey };
}

function getOAuthClientId() {
  return PropertiesService.getScriptProperties().getProperty('GOOGLE_OAUTH_CLIENT_ID') || '';
}

function getAppSheetUrl_(tableName) {
  const config = getAppSheetConfig_();
  return 'https://api.appsheet.com/api/v2/apps/' + config.appId +
    '/tables/' + encodeURIComponent(tableName) + '/Action';
}

function buildAttachmentUrl_(relativePath) {
  if (!relativePath) return '';
  if (/^https?:\/\//i.test(relativePath)) return relativePath;

  const config = getAppSheetConfig_();
  return 'https://www.appsheet.com/template/gettablefileurl' +
    '?appName=' + encodeURIComponent(config.appId) +
    '&tableName=' + encodeURIComponent(TABLES.APPROVALS) +
    '&fileName=' + encodeURIComponent(relativePath);
}

// ---------------------------------------------------------------------------
// LOW-LEVEL APPSHEET REQUEST HELPERS
// ---------------------------------------------------------------------------

function appSheetRequest_(tableName, action, rows, properties) {
  const config = getAppSheetConfig_();
  const url = getAppSheetUrl_(tableName);

  const payload = {
    Action: action,
    Properties: Object.assign({ Locale: 'en-US', Timezone: 'Asia/Manila' }, properties || {}),
    Rows: rows || []
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'ApplicationAccessKey': config.apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('AppSheet API error (' + code + ') on table "' + tableName + '": ' + text);
  }

  if (!text) return [];
  try {
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : (parsed.Rows || []);
  } catch (err) {
    throw new Error('Failed to parse AppSheet response for "' + tableName + '": ' + err.message);
  }
}

function appSheetFind_(tableName, selector) {
  const props = selector ? { Selector: selector } : {};
  return appSheetRequest_(tableName, 'Find', [], props);
}

function appSheetAdd_(tableName, rows) {
  return appSheetRequest_(tableName, 'Add', rows);
}

function appSheetEdit_(tableName, rows) {
  return appSheetRequest_(tableName, 'Edit', rows);
}

/** Finds a single CentralizedApproval row by RequestID via a server-side Selector. */
function findApprovalRowByRequestId_(requestId) {
  const safeId = String(requestId).replace(/"/g, '\\"');
  const selector = 'Filter(' + TABLES.APPROVALS + ', [RequestID]="' + safeId + '")';
  const rows = appSheetFind_(TABLES.APPROVALS, selector);
  return (rows && rows.length) ? rows[0] : null;
}

/** Finds a single CentralizedApproval row by its unique `uid` column. */
function findApprovalRowByUid_(uid) {
  const safeUid = String(uid).replace(/"/g, '\\"');
  const selector = 'Filter(' + TABLES.APPROVALS + ', [uid]="' + safeUid + '")';
  const rows = appSheetFind_(TABLES.APPROVALS, selector);
  return (rows && rows.length) ? rows[0] : null;
}

/** Resolves the exact CentralizedApproval row a caller means, preferring uid over RequestID. */
function resolveApprovalRow_(requestId, uid) {
  if (uid) {
    const byUid = findApprovalRowByUid_(uid);
    if (byUid) return byUid;
    Logger.log('resolveApprovalRow_: uid "%s" did not match any row — falling back to RequestID "%s".', uid, requestId);
  }
  return findApprovalRowByRequestId_(requestId);
}

/** Picks whichever key column is actually present on a fetched row. */
function detectRowKeyField_(row) {
  for (let i = 0; i < ROW_KEY_CANDIDATES.length; i++) {
    if (Object.prototype.hasOwnProperty.call(row, ROW_KEY_CANDIDATES[i])) {
      return ROW_KEY_CANDIDATES[i];
    }
  }
  return null;
}

/** Like appSheetRequest_, but targets a specific app/key instead of the Script Properties config. */
function appSheetRequestToApp_(appId, apiKey, tableName, action, rows, properties) {
  const url = 'https://api.appsheet.com/api/v2/apps/' + appId +
    '/tables/' + encodeURIComponent(tableName) + '/Action';

  const payload = {
    Action: action,
    Properties: Object.assign({ Locale: 'en-US', Timezone: 'Asia/Manila' }, properties || {}),
    Rows: rows || []
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'ApplicationAccessKey': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('AppSheet API error (' + code + ') on table "' + tableName + '": ' + text);
  }

  if (!text) return [];
  try {
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : (parsed.Rows || []);
  } catch (err) {
    throw new Error('Failed to parse AppSheet response for "' + tableName + '": ' + err.message);
  }
}

function appSheetFindInApp_(appId, apiKey, tableName, selector) {
  const props = selector ? { Selector: selector } : {};
  return appSheetRequestToApp_(appId, apiKey, tableName, 'Find', [], props);
}

function buildMasterlistFileUrl_(relativePath) {
  if (!relativePath) return '';
  if (/^https?:\/\//i.test(relativePath)) return relativePath;
  return 'https://www.appsheet.com/template/gettablefileurl' +
    '?appName=' + encodeURIComponent(MASTERLIST_APP_ID) +
    '&tableName=' + encodeURIComponent(MASTERLIST_TABLE) +
    '&fileName=' + encodeURIComponent(relativePath);
}

function findEmployeeImagePath_(employeeNo) {
  const safeEmpNo = String(employeeNo).replace(/"/g, '\\"');
  const selector = 'Filter(' + MASTERLIST_TABLE + ', [Emp No]="' + safeEmpNo + '")';
  const rows = appSheetFindInApp_(MASTERLIST_APP_ID, MASTERLIST_API_KEY, MASTERLIST_TABLE, selector);
  return (rows && rows.length) ? (rows[0].emp_img || '') : '';
}

function findEmployeeImagePathByEmail_(email) {
  const safeEmail = String(email).trim().replace(/"/g, '\\"');
  const selector = 'Filter(' + MASTERLIST_TABLE + ', LOWER([Email])=LOWER("' + safeEmail + '"))';
  const rows = appSheetFindInApp_(MASTERLIST_APP_ID, MASTERLIST_API_KEY, MASTERLIST_TABLE, selector);
  return (rows && rows.length) ? (rows[0].emp_img || '') : '';
}

const MASTERLIST_FILEID_CACHE_PREFIX = 'emp_fileid_v1_';
const MASTERLIST_FOLDERID_CACHE_PREFIX = 'emp_folderid_v1_';
const MASTERLIST_ID_CACHE_TTL_SECONDS = 21600; // 6 hours — file/folder ids rarely change

function getCachedFolderIdByName_(folderName) {
  const cache = CacheService.getScriptCache();
  const key = MASTERLIST_FOLDERID_CACHE_PREFIX + Utilities.base64EncodeWebSafe(folderName);
  const cached = cache.get(key);
  if (cached) return cached;

  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) return null;
  const id = folders.next().getId();
  try { cache.put(key, id, MASTERLIST_ID_CACHE_TTL_SECONDS); } catch (err) { /* non-fatal */ }
  return id;
}

function findEmployeeImageBlobFromDrive_(relativePath) {
  if (!relativePath) return null;
  const parts = relativePath.split('/');
  const fileName = parts[parts.length - 1];
  const folderName = parts.length > 1 ? parts[0] : 'MASTERLIST_Images';

  const cache = CacheService.getScriptCache();
  const fileIdKey = MASTERLIST_FILEID_CACHE_PREFIX + Utilities.base64EncodeWebSafe(relativePath);
  const cachedFileId = cache.get(fileIdKey);

  if (cachedFileId) {
    try {
      return DriveApp.getFileById(cachedFileId).getBlob();
    } catch (err) {
      cache.remove(fileIdKey); // file moved/renamed/deleted — fall through to a fresh search
    }
  }

  const folderId = getCachedFolderIdByName_(folderName);
  const folder = folderId ? DriveApp.getFolderById(folderId) : null;
  if (!folder) {
    Logger.log('Drive folder "%s" not found or not accessible to this script.', folderName);
    return null;
  }

  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    Logger.log('File "%s" not found inside Drive folder "%s".', fileName, folderName);
    return null;
  }
  const file = files.next();
  try { cache.put(fileIdKey, file.getId(), MASTERLIST_ID_CACHE_TTL_SECONDS); } catch (err) { /* non-fatal */ }
  return file.getBlob();
}

function employeeImageCacheKey_(keyType, keyValue) {
  return MASTERLIST_IMAGE_CACHE_PREFIX + keyType + '_' + Utilities.base64EncodeWebSafe(String(keyValue).trim().toLowerCase());
}

function resolveEmployeeImageDataUrl_(path) {
  if (!path) return '';
  const blob = findEmployeeImageBlobFromDrive_(path);
  if (!blob) return '';
  const contentType = blob.getContentType() || 'image/png';
  return 'data:' + contentType + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

/** Looks up an employee's photo from MASTERLIST (via Drive) by Employee No., cached 5 min. */
function getEmployeeImageUrl(employeeNo, sessionToken) {
  const auth = resolveAuth_(sessionToken);
  if (!auth.authorized) throw new Error(auth.reason || 'Unauthorized');
  if (!employeeNo) return '';

  const cache = CacheService.getScriptCache();
  const cacheKey = employeeImageCacheKey_('empno', employeeNo);
  const cached = cache.get(cacheKey);
  if (cached !== null) return cached;

  let dataUrl = '';
  try {
    dataUrl = resolveEmployeeImageDataUrl_(findEmployeeImagePath_(employeeNo));
  } catch (err) {
    Logger.log('getEmployeeImageUrl failed for "%s": %s', employeeNo, err.message);
  }

  try {
    cache.put(cacheKey, dataUrl, MASTERLIST_IMAGE_CACHE_TTL_SECONDS);
  } catch (err) {
    // Large photos may exceed CacheService's ~100KB/key limit — caching is
    // skipped then; correctness unaffected, just no 5-minute speed-up.
  }
  return dataUrl;
}

function getEmployeeImageUrlByEmail(email, sessionToken) {
  const auth = resolveAuth_(sessionToken);
  if (!auth.authorized) throw new Error(auth.reason || 'Unauthorized');
  if (!email) return '';

  const cache = CacheService.getScriptCache();
  const cacheKey = employeeImageCacheKey_('email', email);
  const cached = cache.get(cacheKey);
  if (cached !== null) return cached;

  let dataUrl = '';
  try {
    dataUrl = resolveEmployeeImageDataUrl_(findEmployeeImagePathByEmail_(email));
  } catch (err) {
    Logger.log('getEmployeeImageUrlByEmail failed for "%s": %s', email, err.message);
  }

  try {
    cache.put(cacheKey, dataUrl, MASTERLIST_IMAGE_CACHE_TTL_SECONDS);
  } catch (err) {
    // non-fatal, see note above.
  }
  return dataUrl;
}

// ---------------------------------------------------------------------------
// WEB APP ENTRY POINT
// ---------------------------------------------------------------------------

function doGet(e) {
  const mode = e && e.parameter ? e.parameter.pwa : null;

  // PWA manifest/service-worker routes are not gated — they carry no data.
  if (mode === 'manifest') {
    return ContentService
      .createTextOutput(JSON.stringify(getPwaManifest_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'sw') {
    return ContentService
      .createTextOutput(getServiceWorkerScript_())
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // ---- Whitelist gate ----
  let activeEmail = '';
  try {
    activeEmail = Session.getActiveUser().getEmail();
  } catch (err) {
    activeEmail = '';
  }

  // Session.getActiveUser().getEmail() returns '' when:
  //  - "Who has access" is set to "Anyone" (fully anonymous access), or
  //  - the visitor's domain privacy settings block sharing their identity
  //    even when they are signed into a Google account.
  // In that case we can't make a server-side decision yet — fall through to
  // the client-side Google Sign-In flow already built into this app
  // (see getUserAuthStatus / authenticateWithIdToken in this file), which
  // re-checks identity via an explicit sign-in token instead of the
  // ambient session. We only hard-deny here when we DO have an email and
  // it's not on the list.
  if (activeEmail && !isUserAuthorized(activeEmail)) {
    return HtmlService.createTemplateFromFile('AccessDenied')
      .evaluate()
      .setTitle('Access Denied')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Approval Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://www.gstatic.com/images/branding/product/1x/apps_script_48dp.png');
}

function getWebAppBaseUrl_() {
  return ScriptApp.getService().getUrl();
}

function getPwaManifest_() {
  const base = getWebAppBaseUrl_();
  return {
    name: 'Approval Dashboard',
    short_name: 'Approvals',
    start_url: base,
    scope: base,
    display: 'standalone',
    background_color: '#f4f5f7',
    theme_color: '#2f5cf0',
    icons: [
      { src: 'https://www.gstatic.com/images/branding/product/2x/apps_script_96dp.png', sizes: '96x96', type: 'image/png' },
      { src: 'https://www.gstatic.com/images/branding/product/2x/apps_script_192dp.png', sizes: '192x192', type: 'image/png' }
    ]
  };
}

function getServiceWorkerScript_() {
  return [
    "self.addEventListener('install', function (e) { self.skipWaiting(); });",
    "self.addEventListener('activate', function (e) { self.clients.claim(); });",
    "self.addEventListener('fetch', function (e) {",
    "  e.respondWith(fetch(e.request).catch(function () { return new Response('Offline — please reconnect.', { status: 503 }); }));",
    "});"
  ].join('\n');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------------
// AUTH — session path
// ---------------------------------------------------------------------------

function getUserAuthStatus() {
  let email = '';
  try {
    email = Session.getActiveUser().getEmail();
  } catch (err) {
    email = '';
  }

  if (!email) {
    return { authorized: false, reason: 'No active Google session was found.', needsSignIn: true };
  }
  return resolveManagerAuth_(email);
}

/**
 * The managers directory rarely changes, but this lookup used to hit
 * AppSheet on every single auth check — which, for session-based
 * (non-token) deployments, is EVERY server call. It's cached for
 * MANAGERS_CACHE_TTL_SECONDS via CacheService. Call bustManagersCache()
 * manually if you edit the managers table and don't want to wait out the
 * TTL.
 */
function getManagersCached_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(MANAGERS_CACHE_KEY);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (err) {
      // Corrupt cache entry — fall through to a fresh fetch below.
    }
  }

  const managers = appSheetFind_(TABLES.MANAGERS);
  try {
    cache.put(MANAGERS_CACHE_KEY, JSON.stringify(managers), MANAGERS_CACHE_TTL_SECONDS);
  } catch (err) {
    // CacheService has a ~100KB per-key limit. If the managers table is
    // large enough to exceed it, caching is simply skipped — auth still
    // works correctly, just without the speed-up.
    Logger.log('Could not cache managers directory (non-fatal): ' + err.message);
  }
  return managers;
}

/** Run manually (▶ in the editor) after editing the managers table if you don't want to wait out the 5-minute cache TTL. */
function bustManagersCache() {
  CacheService.getScriptCache().remove(MANAGERS_CACHE_KEY);
}

/** Shared lookup: is this email in the `managers` table? */
function resolveManagerAuth_(email) {
  let managers;
  try {
    managers = getManagersCached_();
  } catch (err) {
    return { authorized: false, reason: 'Unable to reach the manager directory: ' + err.message, email: email };
  }

  const match = managers.find(function (m) {
    const managerEmail = (m.Email || m.email || '').toString().trim().toLowerCase();
    return managerEmail === email.trim().toLowerCase();
  });

  if (!match) {
    return { authorized: false, reason: 'Unauthorized User', email: email };
  }

  return {
    authorized: true,
    email: email,
    name: match.Name || match.name || email,
    role: match.Role || match.role || 'Approver',
    department: match.Department || match.department || ''
  };
}

// ---------------------------------------------------------------------------
// AUTH — Google Sign-In (token) path
// ---------------------------------------------------------------------------

function authenticateWithIdToken(idToken) {
  const clientId = getOAuthClientId();
  if (!clientId) {
    return { authorized: false, reason: 'Google Sign-In is not configured for this app.' };
  }
  if (!idToken) {
    return { authorized: false, reason: 'Missing sign-in token.' };
  }

  const url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    return { authorized: false, reason: 'Could not verify your Google sign-in. Please try again.' };
  }

  const info = JSON.parse(response.getContentText());
  if (info.aud !== clientId) {
    return { authorized: false, reason: 'This sign-in token was not issued for this app.' };
  }
  if (info.email_verified !== 'true' && info.email_verified !== true) {
    return { authorized: false, reason: 'Your Google email address is not verified.' };
  }

  return startTokenSession_(info.email);
}

/**
 * "Switch Google Account" path. The client opens Google's account chooser
 * (GIS token client, prompt=select_account) and sends back the access token
 * for whichever account was picked. We verify it was issued to THIS app's
 * OAuth client, read the verified email from it, and start a token session
 * for that account exactly like authenticateWithIdToken does.
 */
function authenticateWithAccessToken(accessToken) {
  const clientId = getOAuthClientId();
  if (!clientId) {
    return { authorized: false, reason: 'Google Sign-In is not configured for this app.' };
  }
  if (!accessToken) {
    return { authorized: false, reason: 'Missing sign-in token.' };
  }

  const url = 'https://oauth2.googleapis.com/tokeninfo?access_token=' + encodeURIComponent(accessToken);
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    return { authorized: false, reason: 'Could not verify your Google sign-in. Please try again.' };
  }

  const info = JSON.parse(response.getContentText());
  if (info.aud !== clientId && info.azp !== clientId) {
    return { authorized: false, reason: 'This sign-in token was not issued for this app.' };
  }
  if (!info.email) {
    return { authorized: false, reason: 'Google did not share an email address for this account.' };
  }
  if (info.email_verified !== 'true' && info.email_verified !== true) {
    return { authorized: false, reason: 'Your Google email address is not verified.' };
  }

  return startTokenSession_(info.email);
}

/** Shared by both token sign-in paths: manager check + cached session token. */
function startTokenSession_(email) {
  const auth = resolveManagerAuth_(email);
  if (!auth.authorized) return auth;

  const sessionToken = Utilities.getUuid();
  CacheService.getScriptCache().put(SESSION_CACHE_PREFIX + sessionToken, JSON.stringify(auth), SESSION_TTL_SECONDS);

  return Object.assign({ sessionToken: sessionToken }, auth);
}

/** Resolves auth from either a cached token session or the active Google session. */
function resolveAuth_(sessionToken) {
  if (sessionToken) {
    const cached = CacheService.getScriptCache().get(SESSION_CACHE_PREFIX + sessionToken);
    if (!cached) {
      return { authorized: false, reason: 'Your session has expired. Please sign in again.', needsSignIn: true };
    }
    return JSON.parse(cached);
  }
  return getUserAuthStatus();
}

/** Ends a token-based session. No-op (but harmless) for session-based auth. */
function signOut(sessionToken) {
  if (sessionToken) {
    CacheService.getScriptCache().remove(SESSION_CACHE_PREFIX + sessionToken);
  }
  return { success: true };
}

// ---------------------------------------------------------------------------
// APPROVAL DATA
// ---------------------------------------------------------------------------

/** Builds the CacheService key for one approver's fetchApprovalData() result. */
function approvalDataCacheKey_(emailLower) {
  return APPROVAL_DATA_CACHE_PREFIX + Utilities.base64EncodeWebSafe(emailLower);
}

/**
 * NEW (7. Performance): busts the fetchApprovalData cache entry for one
 * approver. Called by updateApprovalStatus() right after its write succeeds
 * so a forced refresh immediately afterward is guaranteed fresh, even
 * though the client's own blacklist already hides the just-actioned request
 * regardless of what this returns.
 */
function invalidateApprovalDataCache_(email) {
  try {
    CacheService.getScriptCache().remove(approvalDataCacheKey_(email.trim().toLowerCase()));
  } catch (err) {
    Logger.log('Could not invalidate approval data cache (non-fatal): ' + err.message);
  }
}

/**
 * CHANGED (7. Performance): asks AppSheet to pre-filter by AprEmail
 * server-side via a Selector (case-insensitive, via LOWER(...)) instead of
 * always pulling every CentralizedApproval row. The original Apps-Script-
 * side filter is KEPT AS-IS below as a correctness safety net — if the
 * Selector ever fails to match for any reason, the fallback branch
 * (unfiltered Find + the untouched filter) produces byte-identical results
 * to the pre-optimization behavior.
 *
 * NEW (7. Performance): also backed by a short (APPROVAL_DATA_CACHE_TTL_
 * SECONDS) per-approver CacheService entry, to absorb bursts of client
 * calls (rapid Sync clicks, a poll landing right after a manual sync) that
 * would otherwise all hit AppSheet independently. Pass forceRefresh=true to
 * always bypass the cache (the client does this on initial load and on the
 * manual "Sync now" button).
 *
 * Nothing about what this function RETURNS has changed — only how much
 * data/traffic it costs to get there.
 *
 * @param {string} sessionToken
 * @param {boolean} [forceRefresh]
 */
function fetchApprovalData(sessionToken, forceRefresh) {
  const auth = resolveAuth_(sessionToken);
  if (!auth.authorized) throw new Error(auth.reason || 'Unauthorized');

  const emailLower = auth.email.trim().toLowerCase();
  const safeEmail = auth.email.trim().replace(/"/g, '\\"');
  const cache = CacheService.getScriptCache();
  const cacheKey = approvalDataCacheKey_(emailLower);

  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (err) {
        // Corrupt cache entry — fall through to a fresh fetch below.
      }
    }
  }

  let rows;
  try {
    const selector = 'Filter(' + TABLES.APPROVALS + ', LOWER([AprEmail])=LOWER("' + safeEmail + '"))';
    rows = appSheetFind_(TABLES.APPROVALS, selector);
  } catch (err) {
    Logger.log('Selector-based fetchApprovalData failed, falling back to unfiltered Find: ' + err.message);
    rows = appSheetFind_(TABLES.APPROVALS);
  }

  const normalized = rows
    .map(normalizeApprovalRow_)
    .filter(function (row) {
      return (row.AprEmail || '').trim().toLowerCase() === emailLower;
    });

  normalized.sort(function (a, b) {
    return new Date(b.CreatedAt).getTime() - new Date(a.CreatedAt).getTime();
  });

  try {
    cache.put(cacheKey, JSON.stringify(normalized), APPROVAL_DATA_CACHE_TTL_SECONDS);
  } catch (err) {
    // CacheService has a ~100KB per-key limit / value size cap. If a given
    // approver's queue is large enough to exceed it, caching is simply
    // skipped for that call — correctness is unaffected, just no speed-up.
    Logger.log('Could not cache approval data (non-fatal): ' + err.message);
  }

  return normalized;
}

function normalizeApprovalRow_(row) {
  const attachment = row.attachment || '';
  return {
    uid: row.uid || row.Uid || '',
    RequestID: row.RequestID || '',
    AppSource: row.AppSource || 'General',
    Title: row.Title || '(Untitled request)',
    RequesterName: row.RequesterName || 'Unknown',
    RequesterEmail: row.RequesterEmail || '',
    Status: row.Status || 'Pending',
    CreatedAt: row.CreatedAt || '',
    Description: row.Description || '',
    DetailsJSON: row.DetailsJSON || '',
    AprEmail: row.AprEmail || '',
    ExeEmail: row.ExeEmail || '',
    dept: row.dept || row.Department || '',
    section: row.section || '',
    attachment: attachment,
    attachmentUrl: buildAttachmentUrl_(attachment),
    sickUrl: row.sickUrl || '',
    company: row.company || ''
  };
}

/** Uploads a base64-encoded file (from the pre-decision attachment picker) to Drive. Best-effort. */
function saveDecisionAttachment_(attachmentData) {
  if (!attachmentData || !attachmentData.base64) return '';

  const bytes = Utilities.base64Decode(attachmentData.base64);
  const blob = Utilities.newBlob(
    bytes,
    attachmentData.mimeType || 'application/octet-stream',
    attachmentData.name || ('decision-attachment-' + Date.now())
  );

  let file;
  if (DECISION_ATTACHMENTS_FOLDER_ID) {
    file = DriveApp.getFolderById(DECISION_ATTACHMENTS_FOLDER_ID).createFile(blob);
  } else {
    file = DriveApp.createFile(blob);
  }

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // Sharing can be blocked by domain policy; the file still exists.
  }

  return file.getUrl();
}

/**
 * Approves or rejects a request. Writes CentralizedApproval and returns
 * immediately — see the top-of-file performance note: the staging-table
 * sync is a SEPARATE call (syncApprovalToStaging, below) that the client
 * fires right after this succeeds, so this function no longer waits on a
 * second, cross-app network call before returning.
 *
 * @param {string} requestId
 * @param {string} action  'Approved' or 'Rejected'
 * @param {{name:string, mimeType:string, base64:string}|null} attachmentData
 * @param {string} sessionToken
 * @param {string} [uid] Unique row identifier — see UID-BASED ROW RESOLUTION.
 * @param {string} [comment] The reviewer's decision comment — no longer
 *        used inside this function (see syncApprovalToStaging), kept as a
 *        parameter here purely so the client can pass everything it has in
 *        one place; omitting it is harmless.
 */
function updateApprovalStatus(requestId, action, attachmentData, sessionToken, uid, comment) {
  const auth = resolveAuth_(sessionToken);
  if (!auth.authorized) throw new Error(auth.reason || 'Unauthorized');

  if (action !== 'Approved' && action !== 'Rejected') {
    throw new Error('Unsupported action: ' + action);
  }

  const row = resolveApprovalRow_(requestId, uid);
  if (!row) throw new Error('Request ' + requestId + ' was not found.');

  const rowApprover = (row.AprEmail || '').toString().trim().toLowerCase();
  if (rowApprover !== auth.email.trim().toLowerCase()) {
    throw new Error('You are not the assigned approver for this request.');
  }

  const keyField = detectRowKeyField_(row);
  if (!keyField) {
    throw new Error(
      'Could not find a key column on "' + TABLES.APPROVALS + '" (checked: ' +
      ROW_KEY_CANDIDATES.join(', ') + '). Add your table\'s real key column ' +
      'name to ROW_KEY_CANDIDATES in Code.gs.'
    );
  }

  const updatePayload = {};
  updatePayload[keyField] = row[keyField];
  updatePayload['RequestID'] = requestId;
  updatePayload['Status'] = action;

  let decisionAttachmentUrl = '';
  if (attachmentData && attachmentData.base64) {
    try {
      decisionAttachmentUrl = saveDecisionAttachment_(attachmentData) || '';
      Logger.log('Upload OK: ' + decisionAttachmentUrl);
    } catch (err) {
      Logger.log('Decision attachment upload FAILED: ' + err.message + ' | stack: ' + err.stack);
    }
  } else {
    Logger.log('attachmentData missing or has no base64.');
  }

  appSheetEdit_(TABLES.APPROVALS, [updatePayload]);

  // NEW (7. Performance): bust this approver's short-lived queue cache so a
  // forced refresh right after this returns is guaranteed fresh.
  invalidateApprovalDataCache_(auth.email);

    return {
    success: true,
    requestId: requestId,
    status: action,
    actorName: auth.name,
    actorEmail: auth.email,
    uid: row.uid || row.Uid || uid || '',
    decisionAttachmentUrl: decisionAttachmentUrl
  };
}

/** Writes one auditLogs row. */
function logAuditEntry(requestId, action, comment, sessionToken, uid, attachmentUrl) {
  const auth = resolveAuth_(sessionToken);
  if (!auth.authorized) throw new Error(auth.reason || 'Unauthorized');

  const auditRow = {
    RequestID: requestId,
    ActorName: auth.name,
    ActorEmail: auth.email,
    Action: action,
    Timestamp: Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'"),
    Comment: comment || '',
    Attachment: attachmentUrl || ''
  };
  // if (uid) auditRow['uid'] = uid; // ← uncomment once auditLogs has a uid column

  Logger.log('Writing auditLogs row: ' + JSON.stringify(auditRow));
  const addResult = appSheetAdd_(TABLES.AUDIT_LOGS, [auditRow]);
  Logger.log('AppSheet Add result: ' + JSON.stringify(addResult));

  return { success: true };
}

/**
 * Pushes the approved/rejected request into the correct staging app/table
 * via AppSheet's Add action, so any Data Change bots scoped to that
 * destination fire normally.
 *
 * ★★★ STAGING TABLE FIELD MAPPING — StagingAllReq / StagingEntries / StagingXC ★★★
 * `rowData` below is the single object sent to whichever table
 * resolveStagingDestination_() resolves the AppSource to — al_comment and
 * al_attachment are populated HERE, ONCE, so all three tables get both
 * fields identically. (1. Audit Logs & Staging Tables Sync — the frontend
 * now always calls syncApprovalToStaging() with the SAME comment that was
 * just written to auditLogs, so al_comment on every staging table matches
 * the audit log entry exactly, and al_attachment continues to be derived
 * from the CentralizedApproval row's own attachment column, unchanged.)
 *
 * @param {string} [uid] Preferred row identifier for resolveApprovalRow_.
 * @param {Object} [preResolvedRow] If the caller already fetched this exact
 *        row (e.g. handleApprovalAction, which needs it anyway for its own
 *        AppSource lookup), pass it here to skip this function's internal
 *        row lookup entirely. Omit it and this behaves exactly as before
 *        (resolves the row itself).
 */
function sendStatusToExternalApp(appSource, requestId, action, comment, actorEmail, uid, preResolvedRow, decisionAttachmentUrl) {
  const row = preResolvedRow || resolveApprovalRow_(requestId, uid);
  if (!row) {
    throw new Error('Could not find CentralizedApproval row for ' + requestId + ' while syncing to staging.');
  }

  const resolvedSource = appSource || row.AppSource || '';
  const destination = resolveStagingDestination_(resolvedSource);

  if (!destination) {
    Logger.log(
      'No staging destination configured for AppSource "%s" (request %s) — skipping external sync.',
      resolvedSource, requestId
    );
    return null;
  }

  // ★★★ StagingAllReq / StagingEntries / StagingXC field mapping ★★★
  const rowData = {
    uid: row.uid || row.Uid || '',
    RequestID: requestId,
    AppSource: resolvedSource,
    Status: action,
    requesterEmail: row.RequesterEmail || '',
    aprEmail: row.AprEmail || '',
    exeEmail: row.ExeEmail || '',
    al_comment: comment || '',
    al_attachment: decisionAttachmentUrl || buildAttachmentUrl_(row.attachment) || ''
  };

  return appSheetRequestToApp_(destination.appId, destination.apiKey, destination.tableName, 'Add', [rowData]);
}

/**
 * The external staging-table sync, split out of updateApprovalStatus() so
 * the client can fire it asynchronously — call this immediately AFTER
 * updateApprovalStatus() succeeds, without waiting on it, so the UI is
 * never blocked by a cross-app network call it doesn't actually need for
 * its own state. A failure here can never look like a failed approve/reject
 * to the caller — the decision is already safely written to
 * CentralizedApproval by the time this runs. This is also the function
 * that carries al_comment/al_attachment through (via
 * sendStatusToExternalApp).
 *
 * NEW (1. Audit Logs & Staging Tables Sync): this was already correct on
 * the backend, but the frontend never actually called it. JavaScript.html's
 * proceedWithDecision() now fires this right after updateApprovalStatus()
 * succeeds, passing through the SAME comment the user typed (and that goes
 * into auditLogs via logAuditEntry), so StagingAllReq / StagingEntries /
 * StagingXC's al_comment always matches the audit trail.
 */
function syncApprovalToStaging(requestId, action, comment, sessionToken, uid, decisionAttachmentUrl) {
  const auth = resolveAuth_(sessionToken);
  if (!auth.authorized) throw new Error(auth.reason || 'Unauthorized');

  const row = resolveApprovalRow_(requestId, uid);
  const appSource = row ? row.AppSource : null;

  return sendStatusToExternalApp(appSource, requestId, action, comment, auth.email, row ? (row.uid || row.Uid) : uid, row, decisionAttachmentUrl);
}

/**
 * Convenience wrapper for any caller that wants EVERYTHING in one blocking
 * round trip (status write + audit log + staging sync). Unlike the primary
 * UI flow (updateApprovalStatus + async syncApprovalToStaging), this stays
 * fully synchronous/self-contained for any other caller relying on it —
 * only the redundant row re-fetch inside it was eliminated (see
 * sendStatusToExternalApp's preResolvedRow parameter).
 */
function handleApprovalAction(requestId, action, comment, attachmentData, sessionToken, uid) {
  const statusResult = updateApprovalStatus(requestId, action, attachmentData, sessionToken, uid, comment);

  // Fetched once, reused for both the AppSource lookup below and passed
  // directly into sendStatusToExternalApp() to avoid a third redundant
  // fetch of the same row.
  const row = resolveApprovalRow_(requestId, uid);
  const appSource = row ? row.AppSource : null;

  logAuditEntry(requestId, action, comment, sessionToken, uid);

  try {
    const auth = resolveAuth_(sessionToken);
    sendStatusToExternalApp(appSource, requestId, action, comment, auth.email, row ? (row.uid || row.Uid) : uid, row);
  } catch (err) {
    Logger.log('External sync failed in handleApprovalAction (non-fatal, decision already saved): ' + err.message);
  }

  return statusResult;
}

/** Requests created within the last 24h, or still Pending, for the notification bell. */
function fetchNotifications(sessionToken) {
  const auth = resolveAuth_(sessionToken);
  if (!auth.authorized) throw new Error(auth.reason || 'Unauthorized');

  const rows = fetchApprovalData(sessionToken);
  const cutoff = Date.now() - 24 * 60 * 60 * 1000;

  return rows
    .filter(function (row) {
      const created = new Date(row.CreatedAt).getTime();
      const isRecent = !isNaN(created) && created >= cutoff;
      const isPending = (row.Status || '').toLowerCase() === 'pending';
      return isRecent || isPending;
    })
    .slice(0, 20);
}

/** Chronological audit trail (approve/reject history) for a single request. Still RequestID-keyed. */
function fetchAuditTrail(requestId, sessionToken) {
  const auth = resolveAuth_(sessionToken);
  if (!auth.authorized) throw new Error(auth.reason || 'Unauthorized');
  if (!requestId) return [];

  const rows = appSheetFind_(TABLES.AUDIT_LOGS);

  return rows
    .filter(function (r) { return (r.RequestID || '') === requestId; })
    .map(function (r) {
      return {
        ActorName: r.ActorName || '',
        ActorEmail: r.ActorEmail || '',
        Action: r.Action || '',
        Timestamp: r.Timestamp || '',
        Comment: r.Comment || '',
        Attachment: r.Attachment || ''
      };
    })

    .sort(function (a, b) { return new Date(b.Timestamp).getTime() - new Date(a.Timestamp).getTime(); });
}

/**
 * NEW (5. Hover Navigation Sidebar): feeds the hover nav's quick-links list.
 * Deliberately returns only labels + AppSheet "start" URLs (which just open
 * the app's own sign-in-gated UI) — NEVER the ApplicationAccessKey values,
 * which must stay server-side only. Edit this list to add/remove/rename the
 * links shown in the sidebar; nothing else needs to change.
 */
function getQuickLinks() {
  return [
    { label: 'HR Request', url: 'https://www.appsheet.com/start/' + STAGING_APP_ID, icon: 'hr' },
    { label: 'HR Connect', url: 'https://www.appsheet.com/start/' + HRCONNECT_APP_ID, icon: 'connect' },
    { label: 'XC Exit Clearance', url: 'https://www.appsheet.com/start/' + XC_APP_ID, icon: 'xc' }, 
    { label: 'Ticketing System', url: 'https://www.appsheet.com/start/e264e075-31b0-46f9-88b4-5555cffd33d0#view=VIEW_DASHBOARD_USER',  icon: 'ticket'}
  ];
}

function testSendStatusToExternalApp() {
  sendStatusToExternalApp('JOB', 'JO2608172110', 'Approved', 'Test comment', 'test@example.com');
}

function testSendStatusToExternalApp_HRConnect() {
  sendStatusToExternalApp('LEAVE', 'LV2608252177', 'Approved', 'Test comment', 'test@example.com');
}

function testSendStatusToExternalApp_XC_WithCommentsAndAttachment(sampleRequestId, sampleComment) {
  sendStatusToExternalApp('XCLB', sampleRequestId, 'Approved', sampleComment || 'Test comment with attachment', 'test@example.com');
}

function testResolveApprovalRow_ByUid(sampleUid) {
  const row = resolveApprovalRow_(null, sampleUid);
  Logger.log('Resolved row for uid "%s": %s', sampleUid, JSON.stringify(row));
}

function inspectCentralizedApprovalRows() {
  const rows = appSheetFind_(TABLES.APPROVALS);
  Logger.log('Total rows: ' + rows.length);
  Logger.log('First 3 rows: ' + JSON.stringify(rows.slice(0, 3), null, 2));
}

function inspectAuditLogsRows() {
  const rows = appSheetFind_(TABLES.AUDIT_LOGS);
  Logger.log('Total rows: ' + rows.length);
  if (rows.length) {
    Logger.log('Columns on first row: ' + JSON.stringify(Object.keys(rows[0])));
  } else {
    Logger.log('auditLogs table is empty — check the table directly in AppSheet\'s Data editor instead.');
  }
}

function clearEmployeeImageCache() {
  CacheService.getScriptCache().removeAll(
    Object.keys(PropertiesService.getScriptProperties().getProperties())
  );
}

function testEmployeeImageLookup() {
  const employeeNo = 'RERI-1426'; // use a real Emp No from your table
  try {
    const rows = appSheetFindInApp_(MASTERLIST_APP_ID, MASTERLIST_API_KEY, MASTERLIST_TABLE,
      'Filter(' + MASTERLIST_TABLE + ', [Emp No]="' + employeeNo + '")');
    Logger.log('Rows found: ' + rows.length);
    Logger.log('First row: ' + JSON.stringify(rows[0]));
    if (rows.length) {
      const url = buildMasterlistFileUrl_(rows[0].emp_img);
      Logger.log('Built URL: ' + url);
    }
  } catch (err) {
    Logger.log('ERROR: ' + err.message);
  }
}

function testEmployeeImageDataUrl() {
  const path = 'MASTERLIST_Images/RERI-1426.emp_img.030242.png';
  const url = buildMasterlistFileUrl_(path);
  Logger.log('Fetching: ' + url);

  const response = UrlFetchApp.fetch(url, {
    headers: { 'ApplicationAccessKey': MASTERLIST_API_KEY },
    muteHttpExceptions: true
  });

  Logger.log('Response code: ' + response.getResponseCode());
  Logger.log('Response headers: ' + JSON.stringify(response.getHeaders()));
  Logger.log('Content-Type: ' + response.getBlob().getContentType());
  Logger.log('Byte length: ' + response.getBlob().getBytes().length);
  Logger.log('First 200 chars of body: ' + response.getContentText().substring(0, 200));
}

function testEmployeeImageFromDrive() {
  const path = 'MASTERLIST_Images/RERI-1426.emp_img.030242.png';
  const blob = findEmployeeImageBlobFromDrive_(path);
  if (!blob) { Logger.log('No blob returned — see log above for which step failed.'); return; }
  Logger.log('Got blob, content type: ' + blob.getContentType() + ', size: ' + blob.getBytes().length + ' bytes');
}

function diagnoseEmployeePhoto() {
  const employeeNo = 'RERI-1426'; // swap for a real Emp No you're testing with
  
  Logger.log('--- Step 1: Find row in MASTERLIST ---');
  try {
    const path = findEmployeeImagePath_(employeeNo);
    Logger.log('emp_img path: ' + path);
    
    if (!path) {
      Logger.log('No path returned — check Emp No exists / column name is still "Emp No" and "emp_img".');
      return;
    }
    
    Logger.log('--- Step 2: Fetch blob from Drive ---');
    const blob = findEmployeeImageBlobFromDrive_(path);
    if (!blob) {
      Logger.log('No blob returned — see log lines above for which step failed (folder not found / file not found).');
      return;
    }
    Logger.log('Got blob OK: ' + blob.getContentType() + ', ' + blob.getBytes().length + ' bytes');
    
    Logger.log('--- Step 3: Full getEmployeeImageUrl (requires a real session) ---');
    Logger.log('Skipping — this needs a sessionToken from an active browser session, test via the app itself.');
    
  } catch (err) {
    Logger.log('ERROR: ' + err.message);
    Logger.log('Stack: ' + err.stack);
  }
}
