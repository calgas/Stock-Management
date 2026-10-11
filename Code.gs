/**
 * STOCK MANAGEMENT — BACKEND (all-in-one, bound script)
 * getSpreadsheet_() uses SpreadsheetApp.getActiveSpreadsheet(), so this
 * only works pasted into Extensions > Apps Script on the actual sheet.
 *
 * Fresh install: run ADMIN_generateAppSecret() once, then create the
 * first Admin via ADMIN_setUserPassword (see USER MANAGEMENT below) —
 * no separate bootstrap step needed after that.
 *
 * CORS: frontend must POST with 'Content-Type: text/plain;charset=utf-8',
 * not 'application/json' — Apps Script Web Apps don't handle the
 * preflight application/json triggers. Body is still parsed as JSON here.
 */

var SESSION_DURATION_HOURS = 12;
var OTP_VALIDITY_MINUTES = 10;
var OTP_RESEND_COOLDOWN_SECONDS = 60;
var OTP_MAX_ATTEMPTS = 5;
var LOGIN_MAX_ATTEMPTS = 5;
var LOGIN_LOCKOUT_MINUTES = 15;
// Applies to new passwords only; existing ones keep working until changed.
var MIN_PASSWORD_LENGTH = 8;
var COMPANY_NAME = 'CALGAS CAPACITORS';
var COMPANY_LOGO_URL = 'https://calgas.github.io/Stock-Management/logo/CALGAS%20CAPACITORS-logo-768x240.jpg';

// Address every outbound mail should appear to come from. Apps Script
// cannot send as an arbitrary address: this only takes effect if it is a
// verified "Send mail as" alias on the Google account that owns this
// script. sendMail_ checks that at runtime and degrades gracefully.
// Stock may not be driven below zero. Issuing more than is on hand is
// almost always a missing Receipt rather than a real movement, and letting
// it through produces balances like -1,380 NOS that no physical count can
// ever reconcile. Set to true only as a deliberate, temporary escape hatch.
var ALLOW_NEGATIVE_STOCK = false;

var MAIL_FROM = 'noreply@calgas.in';
var MAIL_FROM_NAME = COMPANY_NAME + ' Stock Management';

/* =========================================================================
   OUTBOUND MAIL
   -------------------------------------------------------------------------
   Single funnel for every email this backend sends (password OTPs and the
   daily stock report). GmailApp is used rather than MailApp because it is
   the only one of the two that accepts a `from` address, and then only
   when that address is a verified alias on the owning account. If the
   alias is not present the send still goes out under the owner's own
   address rather than failing, so mail is never silently lost because of
   a Workspace misconfiguration.
   ========================================================================= */

var _mailAliases = null;
function canSendAsFrom_() {
  if (_mailAliases === null) {
    try { _mailAliases = GmailApp.getAliases() || []; } catch (err) { _mailAliases = []; }
  }
  return _mailAliases.indexOf(MAIL_FROM) !== -1;
}

// opts: { to, subject, body, htmlBody, attachments }
function sendMail_(opts) {
  var options = {
    name: MAIL_FROM_NAME,
    htmlBody: opts.htmlBody,
    replyTo: MAIL_FROM
  };
  if (opts.attachments) options.attachments = opts.attachments;
  if (canSendAsFrom_()) options.from = MAIL_FROM;
  GmailApp.sendEmail(opts.to, opts.subject, opts.body, options);
}

/* =========================================================================
   WEB APP ROUTER
   ========================================================================= */
var PUBLIC_ACTIONS = ['ping', 'login', 'requestPasswordReset', 'resetPasswordWithOtp'];

function doPost(e) { return handleRequest_(e); }
// GET is intentionally not routed to handleRequest_ — routing it would let
// login/reset run via a query string (?action=login&username=...&password=...),
// putting credentials in server logs and browser history. The frontend
// only ever POSTs, so this costs nothing functionally.
function doGet(e) {
  return jsonResponse_({ success: false, error: 'This API accepts POST requests only.' });
}

function handleRequest_(e) {
  var params = {};
  try {
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      params = e.parameter;
    }
  } catch (err) {
    return jsonResponse_({ success: false, error: 'Invalid request body — expected JSON' });
  }

  var action = params.action;
  if (!action) return jsonResponse_({ success: false, error: 'Missing action' });

  try {
    var session = null;
    if (PUBLIC_ACTIONS.indexOf(action) === -1) {
      session = validateSession_(params.token);
    }
    var data = routeAction_(action, params, session);
    return jsonResponse_({ success: true, data: data });
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

function routeAction_(action, params, session) {
  switch (action) {
    // No session: proves the deployment itself is reachable. Production
    // Tracker's ADMIN_checkStockLink calls it.
    case 'ping':
      return { ok: true, time: new Date() };
    case 'login':
      return login_(params);
    case 'logout':
      return logout_(session);
    case 'whoAmI':
      // Permissions come back here too: the app gates its buttons during
      // start-up, before the bootstrap payload has arrived.
      return { username: session.Username, role: session.Role, department: session.Department,
               permissions: sessionPermissions_(session) };

    case 'getDepartments':
      return getDepartments_();

    // Also bundled into getBootstrap; exposed separately so the lists can
    // be re-pulled after someone edits the Dropdowns tab, without forcing
    // a full master+stock refresh.
    case 'getDropdowns':
      return getDropdowns_();

    case 'changePassword':
      return changeOwnPassword_(session, params);

    // Bundles what the frontend needs right after login/refresh into one
    // exec instead of 5-6 separate calls — see getBootstrap_ for why.
    case 'getBootstrap':
      return getBootstrap_(session);

    case 'getRMMaster':
      return getMaster_(RM_CONFIG);
    case 'getFGMaster':
      return getMaster_(FG_CONFIG);

    case 'getRMCurrentStock':
      return getCurrentStock_(RM_CONFIG);
    case 'getFGCurrentStock':
      return getCurrentStock_(FG_CONFIG);

    case 'getAttachments':
      return getAttachments_(params);
    case 'getAttachmentFile':
      return getAttachmentFile_(params);
    case 'uploadAttachment':
      return uploadAttachment_(session, params);
    case 'removeAttachment':
      return removeAttachment_(session, params);

    case 'getRMLedger':
      return getLedger_(RM_CONFIG, params);
    case 'getFGLedger':
      return getLedger_(FG_CONFIG, params);

    // Read-only companion to the idempotency check. When a submit's outcome
    // never reached the client — the connection dropped, or the reply was
    // lost — the app asks here whether the entry actually landed, instead of
    // guessing or posting it again.
    case 'checkTransaction':
      return checkTransaction_(params);

    case 'addRMTransaction':
      requirePerm_(session, 'stock.rm.post');
      return addTransaction_(RM_CONFIG, session, params);
    case 'addFGTransaction':
      requirePerm_(session, 'stock.fg.post');
      return addTransaction_(FG_CONFIG, session, params);

    case 'getItemPrice':
      return getItemPrice_(params);

    case 'updateRMTransaction':
      return amendTransaction_(session, Object.assign({}, params, { key: 'rm' }));
    case 'updateFGTransaction':
      return amendTransaction_(session, Object.assign({}, params, { key: 'fg' }));
    case 'deleteRMTransaction':
      return deleteTransaction_(session, Object.assign({}, params, { key: 'rm' }));
    case 'deleteFGTransaction':
      return deleteTransaction_(session, Object.assign({}, params, { key: 'fg' }));

    case 'recalculateRMStock':
      requirePerm_(session, 'stock.recalculate');
      return recalcCurrentStock_(RM_CONFIG);
    case 'recalculateFGStock':
      requirePerm_(session, 'stock.recalculate');
      return recalcCurrentStock_(FG_CONFIG);

    case 'addRMItem':
      requirePerm_(session, 'stock.rm.items');
      return addMasterItem_(RM_CONFIG, session, params);
    case 'updateRMItem':
      requirePerm_(session, 'stock.rm.items');
      return updateMasterItem_(RM_CONFIG, session, params);
    case 'deleteRMItem':
      requirePerm_(session, 'stock.rm.retire');
      return setMasterItemActive_(RM_CONFIG, session, params, false);
    case 'reactivateRMItem':
      requirePerm_(session, 'stock.rm.retire');
      return setMasterItemActive_(RM_CONFIG, session, params, true);

    case 'addFGItem':
      requirePerm_(session, 'stock.fg.items');
      return addMasterItem_(FG_CONFIG, session, params);
    case 'updateFGItem':
      requirePerm_(session, 'stock.fg.items');
      return updateMasterItem_(FG_CONFIG, session, params);
    case 'deleteFGItem':
      requirePerm_(session, 'stock.fg.retire');
      return setMasterItemActive_(FG_CONFIG, session, params, false);
    case 'reactivateFGItem':
      requirePerm_(session, 'stock.fg.retire');
      return setMasterItemActive_(FG_CONFIG, session, params, true);

    case 'bulkAddRMItems':
      requirePerm_(session, 'stock.import');
      return bulkAddMasterItems_(RM_CONFIG, session, params);
    case 'bulkAddFGItems':
      requirePerm_(session, 'stock.import');
      return bulkAddMasterItems_(FG_CONFIG, session, params);

    // No role check — Export has never been role-restricted in this app
    // (anyone who can view a list can already see everything in it).
    case 'exportStyledWorkbook':
      requirePerm_(session, 'stock.export');
      return exportStyledWorkbook_(params);

    case 'getAccessControl':
      requireRole_(session, ['Admin']);
      return getAccessControl_();
    case 'setUserAccess':
      requireRole_(session, ['Admin']);
      return setUserAccess_(session, params);
    case 'setAccessTemplate':
      requireRole_(session, ['Admin']);
      return setAccessTemplate_(session, params);

    case 'getUsers':
      requireRole_(session, ['Admin']);
      return getUsers_();
    case 'addUser':
      requireRole_(session, ['Admin']);
      return addUser_(session, params);
    case 'updateUser':
      requireRole_(session, ['Admin']);
      return updateUser_(session, params);
    case 'deactivateUser':
      requireRole_(session, ['Admin']);
      return setUserActive_(session, params, false);
    case 'reactivateUser':
      requireRole_(session, ['Admin']);
      return setUserActive_(session, params, true);
    case 'adminResetUserPassword':
      requireRole_(session, ['Admin']);
      return requestPasswordReset_({ username: params.username }, session);

    case 'sendTestStockMail':
      requirePerm_(session, 'stock.mail.test');
      return sendTestStockMail_(session, params);

    case 'requestPasswordReset':
      return requestPasswordReset_(params, null);
    case 'resetPasswordWithOtp':
      return resetPasswordWithOtp_(params);

    default:
      throw new Error('Unknown action: ' + action);
  }
}

/* =========================================================================
   AUTH — login, sessions, passwords
   ========================================================================= */
function login_(params) {
  var identifier = String(params.username || '').trim();
  var password = String(params.password || '');
  if (!identifier || !password) throw new Error('username and password are required');

  var usersSheet = getSpreadsheet_().getSheetByName('Users');
  // Auto-creates FailedAttempts/LockedUntil/Salt on first use, same pattern
  // as the RM Supplier/InvoiceNo/InvoiceDate columns — no manual migration
  // needed. Salt is blank for accounts created before the salting upgrade;
  // hashPassword_ treats a blank salt as "use the legacy unsalted format"
  // so those accounts keep working until migrated below.
  ensureColumns_(usersSheet, ['FailedAttempts', 'LockedUntil', 'Salt']);
  var users = readSheetAsObjects_(usersSheet);
  var idLower = identifier.toLowerCase();
  // Accept a username or an email in the same field — email compared
  // case-insensitively since that's how every mail provider treats it.
  var user = users.filter(function (u) {
    return u.Username === identifier || (u.Email && String(u.Email).toLowerCase() === idLower);
  })[0];
  if (!user) throw new Error('Invalid username or password');
  var username = user.Username; // canonical from here on, regardless of what was typed
  if (!isActiveValue_(user.Active)) throw new Error('Account is inactive — ask an Admin to enable it');
  if (!user.PasswordHash) throw new Error('No password set yet for this account — use "Forgot password" to set one');

  if (user.LockedUntil && new Date(user.LockedUntil).getTime() > Date.now()) {
    throw new Error('Too many failed attempts — try again in a few minutes');
  }

  var hash = hashPassword_(password, getAppSecret_(), user.Salt);
  if (hash !== user.PasswordHash) {
    recordFailedLogin_(usersSheet, username, user);
    throw new Error('Invalid username or password');
  }
  clearLoginThrottle_(usersSheet, username);

  // Opportunistic migration: a correct legacy (unsalted) login is the one
  // moment we have the plaintext password in hand, so re-hash it with a
  // fresh per-user salt now instead of needing a separate migration pass.
  if (!user.Salt) {
    var salt = generateSalt_();
    var salted = hashPassword_(password, getAppSecret_(), salt);
    var userRowIdx = findRowIndexByValue_(usersSheet, 'Username', username);
    var userMap = headerIndexMap_(usersSheet);
    usersSheet.getRange(userRowIdx, userMap['Salt'] + 1).setValue(salt);
    usersSheet.getRange(userRowIdx, userMap['PasswordHash'] + 1).setValue(salted);
  }

  var now = new Date();
  var expires = new Date(now.getTime() + SESSION_DURATION_HOURS * 3600 * 1000);
  var token = Utilities.getUuid();

  var sessionsSheet = getOrCreateSessionsSheet_();
  // Added here rather than only at setup, so a sheet made before access
  // control existed gains the column on the next sign-in.
  ensureColumns_(sessionsSheet, ['Perms', 'App']);
  var app = sessionApp_(params.app);
  var sessionRow = {
    Token: token,
    Username: user.Username,
    Role: user.Role,
    Department: user.Department,
    CreatedAt: now,
    ExpiresAt: expires,
    // The resolved set travels with the session, so the sibling apps get it
    // from the row they already read to validate the token.
    Perms: grantedKeys_(resolveAccess_(user)).join(','),
    App: app
  };
  // One active sign-in per account in each app: this one replaces the
  // account's other sign-in to the same app, wherever it was — a phone left
  // signed in, another browser, a tab closed without signing out. Sign-ins
  // to the other apps are left alone, so opening ELE Tracker never closes
  // Stock Management. Locked, so two sign-ins at the same moment can't both
  // survive.
  withSessionsLock_(function () {
    endOtherSessions_(sessionsSheet, user.Username, app);
    appendRowByHeader_(sessionsSheet, sessionRow);
  });
  cacheSession_(sessionRow);

  updateLastLogin_(usersSheet, username, now);
  logAudit_(username, 'LOGIN', 'Users', username, '', '');

  return {
    app: app,
    token: token,
    name: user.Name,
    username: user.Username,
    role: user.Role,
    department: user.Department,
    expiresAt: expires,
    permissions: resolveAccess_(user)
  };
}

// Wrong-password path — increments the counter and, once it hits
// LOGIN_MAX_ATTEMPTS, sets a LockedUntil timestamp that login_ checks
// before ever touching the password hash.
function recordFailedLogin_(usersSheet, username, user) {
  var rowIdx = findRowIndexByValue_(usersSheet, 'Username', username);
  if (rowIdx === -1) return;
  var map = headerIndexMap_(usersSheet);
  var attempts = (Number(user.FailedAttempts) || 0) + 1;
  usersSheet.getRange(rowIdx, map['FailedAttempts'] + 1).setValue(attempts);
  if (attempts >= LOGIN_MAX_ATTEMPTS) {
    var lockUntil = new Date(Date.now() + LOGIN_LOCKOUT_MINUTES * 60 * 1000);
    usersSheet.getRange(rowIdx, map['LockedUntil'] + 1).setValue(lockUntil);
  }
}

// Correct-password path — resets the counter so a mistyped password once
// in a while never accumulates toward a lockout.
function clearLoginThrottle_(usersSheet, username) {
  var rowIdx = findRowIndexByValue_(usersSheet, 'Username', username);
  if (rowIdx === -1) return;
  var map = headerIndexMap_(usersSheet);
  usersSheet.getRange(rowIdx, map['FailedAttempts'] + 1).setValue(0);
  usersSheet.getRange(rowIdx, map['LockedUntil'] + 1).setValue('');
}

function logout_(session) {
  var sheet = getOrCreateSessionsSheet_();
  // Found and deleted under the lock: a sign-in elsewhere deletes rows too,
  // and a row number read outside it could point at someone else's session
  // by the time it's used.
  withSessionsLock_(function () {
    var rowIdx = findRowIndexByValue_(sheet, 'Token', session.Token);
    if (rowIdx > -1) sheet.deleteRow(rowIdx);
  });
  dropCachedSession_(session.Token);
  return { loggedOut: true };
}

// Every change to the Sessions tab that goes by row number — deleting a row,
// rewriting one — runs under this. Sign-in deletes rows, so a row number
// found outside the lock can shift onto another person's session before
// it's used: the wrong person signed out, or one person's role and access
// written onto someone else's row.
//
// Not re-entrant. Callers that already hold the script lock (none do today)
// must not call into it.
function withSessionsLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    try { SpreadsheetApp.flush(); } catch (err) {}
    lock.releaseLock();
  }
}

// Ends every session an account holds, in every app — on deactivation, and
// after a password reset or change, so whoever had the old password is out.
// keepToken spares the session making the change (a self-service change
// shouldn't sign the person out of the screen they're on).
function endAllSessionsFor_(username, reason, keepToken) {
  var sheet = getOrCreateSessionsSheet_();
  return withSessionsLock_(function () {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;
    var meta = sheetMeta_(sheet);
    var uCol = meta.map['Username'], tCol = meta.map['Token'];
    if (uCol === undefined || tCol === undefined) return 0;
    var values = sheet.getRange(2, 1, lastRow - 1, meta.lastCol).getValues();
    var cache = CacheService.getScriptCache();
    var ended = 0;
    // Bottom-up, so deleting a row doesn't shift the ones still to check.
    for (var i = values.length - 1; i >= 0; i--) {
      if (String(values[i][uCol]) !== String(username)) continue;
      var token = String(values[i][tCol]);
      if (keepToken && token === String(keepToken)) continue;
      sheet.deleteRow(i + 2);
      dropCachedSession_(token);
      try { cache.put(REVOKED_PREFIX + token, reason, 21600); } catch (err) {}
      ended++;
    }
    return ended;
  });
}

// Session lookup is on the hot path — every authenticated request pays for it.
// Reading the whole Sessions sheet each time made that cost grow with the
// number of logins ever issued, so the token is cached and the sheet is only
// consulted on a miss. The sheet stays authoritative: the cache is written on
// login, dropped on logout, and its TTL is capped so a revoked session cannot
// outlive the cache by more than SESSION_CACHE_TTL_SECONDS.
var SESSION_CACHE_TTL_SECONDS = 21600; // 6h, the CacheService maximum

function sessionCacheKey_(token) { return 'sess:' + token; }

function cacheSession_(session) {
  try {
    var msLeft = new Date(session.ExpiresAt).getTime() - Date.now();
    if (msLeft <= 0) return;
    var ttl = Math.min(SESSION_CACHE_TTL_SECONDS, Math.floor(msLeft / 1000));
    if (ttl > 0) CacheService.getScriptCache().put(sessionCacheKey_(session.Token), JSON.stringify(session), ttl);
  } catch (err) { /* cache is an optimisation — never fail a request over it */ }
}

function dropCachedSession_(token) {
  try { CacheService.getScriptCache().remove(sessionCacheKey_(token)); } catch (err) {}
}

// A session closed because its account signed in somewhere else is
// remembered for six hours — CacheService's longest — so the device it was
// taken from can be told why, rather than just "invalid session".
var REVOKED_PREFIX = 'revoked:';

// The apps that sign in through here. Each keeps its own session, so they
// can be open together — on one device or several — without closing each
// other. A client that doesn't say which app it is (an older version) is
// Stock Management, which is what every session was before this existed.
var SESSION_APPS = { stock: 'Stock Management', ele: 'ELE Tracker', pt: 'Production Tracker' };

function sessionApp_(app) {
  var key = String(app || '').trim().toLowerCase();
  return SESSION_APPS[key] ? key : 'stock';
}

function signedInElsewhere_(app) {
  return 'This account signed in to ' + (SESSION_APPS[app] || SESSION_APPS.stock) +
         ' on another device, so this session was closed — please log in again';
}

// Why a session was closed, for the device it was taken from. An app key
// means "signed in elsewhere"; the other reasons are spelled out here.
// Each starts "Invalid session" so every app's frontend recognises it as a
// sign-out and returns to the login screen.
var REVOKE_REASONS = {
  deactivated: 'Invalid session — this account was deactivated. Ask an Admin if that’s a mistake.',
  password: 'Invalid session — this account’s password was changed, so please log in again.'
};

function revokedMessage_(why) {
  return REVOKE_REASONS[why] || signedInElsewhere_(why);
}

function endOtherSessions_(sheet, username, app) {
  app = sessionApp_(app);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var meta = sheetMeta_(sheet);
  var uCol = meta.map['Username'], tCol = meta.map['Token'], aCol = meta.map['App'];
  if (uCol === undefined || tCol === undefined) return 0;
  var values = sheet.getRange(2, 1, lastRow - 1, meta.lastCol).getValues();
  var cache = CacheService.getScriptCache();
  var ended = 0;
  // Bottom-up, so deleting a row doesn't shift the ones still to check.
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][uCol]) !== String(username)) continue;
    // A row from before sessions knew their app is a Stock Management one.
    if (sessionApp_(aCol === undefined ? '' : values[i][aCol]) !== app) continue;
    var token = String(values[i][tCol]);
    sheet.deleteRow(i + 2);
    dropCachedSession_(token);
    try { cache.put(REVOKED_PREFIX + token, app, 21600); } catch (err) {}
    ended++;
  }
  return ended;
}

function validateSession_(token) {
  if (!token) throw new Error('Missing session token — please log in');

  var session = null;
  try {
    var hit = CacheService.getScriptCache().get(sessionCacheKey_(token));
    if (hit) session = JSON.parse(hit);
  } catch (err) { session = null; }

  if (!session) {
    var sessions = readSheetAsObjects_(getOrCreateSessionsSheet_());
    session = sessions.filter(function (s) { return s.Token === token; })[0];
    if (!session) {
      var why = null;
      try { why = CacheService.getScriptCache().get(REVOKED_PREFIX + token); } catch (err) {}
      throw new Error(why ? revokedMessage_(why) : 'Invalid session — please log in again');
    }
    cacheSession_(session);
  }

  // Checked against the stored expiry either way, so a cached entry can never
  // extend a session past the lifetime the sheet recorded for it. A blank or
  // unreadable expiry counts as expired; read as a date it would be NaN, and
  // NaN < now is false, so the session would never end.
  var expiresMs = new Date(session.ExpiresAt).getTime();
  if (isNaN(expiresMs) || expiresMs < Date.now()) {
    dropCachedSession_(token);
    throw new Error('Session expired — please log in again');
  }
  return session; // { Token, Username, Role, Department, CreatedAt, ExpiresAt }
}

// Posting for someone else's department is now a grant rather than a
// property of three hardcoded role names — which is what lets Role be
// anything the company uses (HOD, Management, and whatever comes next)
// without this function having to know about it.
function enforceDepartmentScope_(session, department) {
  if (can_(session, 'stock.post.any')) return;
  if (String(department) !== String(session.Department)) {
    throw new Error('You can only post transactions for your own department (' +
                    (session.Department || 'none set') + '). Ask an Admin for wider access.');
  }
}

function requireRole_(session, allowedRoles) {
  if (allowedRoles.indexOf(session.Role) === -1) {
    throw new Error('This action requires role: ' + allowedRoles.join(' or '));
  }
}

/* =========================================================================
   CORRECTING A LEDGER ENTRY

   A wrong entry used to mean opening the spreadsheet. Both of these do the
   same two things the sheet edit would have had to do by hand, in one
   locked step: change the ledger row, and move the running totals to match.

   Both are safe to send twice, which matters on a connection that drops:

     amend  — the new values are absolute, so a repeat computes a net
              change of zero and does nothing
     delete — the row is already gone, so a repeat simply reports that

   The entry's item cannot be changed here. Picking the wrong item is a
   different movement, not a corrected one: delete it and post it again,
   so both items' totals are rebuilt from their own ledger rows.
   ========================================================================= */

// Same script-wide lock the posting path takes, so a correction can't
// interleave with a transaction being posted for the same item.
function withTransactionLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    // Commit before the lock goes, so the next request reads the corrected
    // totals rather than the ones it is replacing.
    try { SpreadsheetApp.flush(); } catch (err) {}
    lock.releaseLock();
  }
}

function configForKey_(key) {
  if (key === 'rm') return RM_CONFIG;
  if (key === 'fg') return FG_CONFIG;
  throw new Error('Unknown stock type: ' + key);
}

function findLedgerRow_(config, txnId) {
  var sheet = getSpreadsheet_().getSheetByName(config.ledgerSheet);
  var id = String(txnId || '').trim();
  var row = readRowByValue_(sheet, 'TxnID', id);
  if (!row) throw new Error('Entry not found: ' + txnId);

  // Entries made before IDs got longer could, rarely, share one. Changing
  // "the entry with this ID" would then change whichever came first, so a
  // shared ID is refused rather than guessed at.
  var ids = sheet.getRange(2, row.map['TxnID'] + 1, sheet.getLastRow() - 1, 1).getValues();
  var count = ids.filter(function (r) { return String(r[0]) === id; }).length;
  if (count > 1) {
    throw new Error(count + ' ledger entries share the ID ' + id + ', so the app can\u2019t tell which one you mean. ' +
                    'Give one of them a different TxnID in the sheet, then try again.');
  }
  var obj = {};
  Object.keys(row.map).forEach(function (h) { obj[h] = row.values[row.map[h]]; });
  return { sheet: sheet, rowIdx: row.rowIdx, map: row.map, entry: obj };
}

// The balance an item would be left with after a net change, so a
// correction can be refused before anything is written rather than leaving
// stock negative.
function balanceAfter_(config, code, netIn, netOut) {
  var stockSheet = getSpreadsheet_().getSheetByName(config.stockSheet);
  var row = readRowByValue_(stockSheet, config.codeField, code);
  if (!row) return netIn - netOut;
  var num = function (header) {
    return (header in row.map) ? Number(row.values[row.map[header]]) || 0 : 0;
  };
  return num('OpeningStock') + (num(config.receivedField) + netIn) - (num(config.issuedField) + netOut);
}

function amendTransaction_(session, params) {
  var config = configForKey_(params.key);
  requirePerm_(session, 'stock.ledger.edit');

  return withTransactionLock_(function () {
    var found = findLedgerRow_(config, params.txnId);
    var entry = found.entry;
    var code = entry[config.codeField];
    // The entry as it stands must be one this person could have posted —
    // checking only the department it's being moved to would let someone
    // pull another department's entry into their own.
    if (entry.Department) enforceDepartmentScope_(session, entry.Department);

    var txnType = params.txnType === undefined ? entry.TxnType : String(params.txnType).trim();
    var validTypes = config.inTypes.concat(config.outTypes, config.adjustTypes);
    if (validTypes.indexOf(txnType) === -1) {
      throw new Error('Unknown transaction type: ' + txnType + '. Use one of: ' + validTypes.join(', '));
    }
    var qty = params.qty === undefined ? Number(entry.Qty) : Number(params.qty);
    if (isNaN(qty)) throw new Error('Quantity must be a number');
    if (config.adjustTypes.indexOf(txnType) === -1 && qty <= 0) throw new Error('Quantity must be greater than zero');

    var department = params.department === undefined ? entry.Department : String(params.department).trim();
    if (department) enforceDepartmentScope_(session, department);
    if (department !== String(entry.Department || '')) assertStockDepartment_(department);

    var before = txnDelta_(config, entry.TxnType, Number(entry.Qty) || 0);
    var after = txnDelta_(config, txnType, qty);
    var netIn = after.inDelta - before.inDelta;
    var netOut = after.outDelta - before.outDelta;

    if (!ALLOW_NEGATIVE_STOCK) {
      var projected = balanceAfter_(config, code, netIn, netOut);
      if (projected < 0) {
        throw new Error('That change would leave ' + code + ' at ' + projected +
                        '. Post the missing receipt first, or use a smaller quantity.');
      }
    }

    var rate = params.rate === undefined ? entry.Rate
             : (params.rate === '' ? '' : Number(params.rate));
    if (rate !== '' && isNaN(Number(rate))) throw new Error('Rate must be a number');
    var fields = {
      TxnType: txnType,
      Qty: qty,
      Rate: rate,
      Value: (rate !== '' && rate !== null && rate !== undefined) ? Number(rate) * qty : '',
      Department: department,
      Reference: params.reference === undefined ? entry.Reference : String(params.reference).trim(),
      Remarks: params.remarks === undefined ? entry.Remarks : String(params.remarks).trim(),
      EditedBy: session.Username,
      EditedAt: new Date()
    };
    ensureColumns_(found.sheet, ['EditedBy', 'EditedAt']);
    var fresh = findLedgerRow_(config, params.txnId); // columns may have moved
    writeRowCells_(fresh.sheet, fresh.rowIdx, stageCells_(fresh.map, fields));

    var masterRow = lookupMasterRow_(config, code) || {};
    var newBalance = (netIn === 0 && netOut === 0)
      ? balanceAfter_(config, code, 0, 0)
      : applyStockTotals_(config, code, netIn, netOut, masterRow);

    logAudit_(session.Username, 'EDIT_TXN', config.ledgerSheet, entry.TxnID,
              entry.TxnType + ' ' + entry.Qty + ' ' + (entry.Department || ''),
              txnType + ' ' + qty + ' ' + (department || ''));

    return { txnId: entry.TxnID, code: code, newBalance: newBalance };
  });
}

function deleteTransaction_(session, params) {
  var config = configForKey_(params.key);
  requirePerm_(session, 'stock.ledger.delete');

  return withTransactionLock_(function () {
    var found = findLedgerRow_(config, params.txnId);
    var entry = found.entry;
    var code = entry[config.codeField];
    if (entry.Department) enforceDepartmentScope_(session, entry.Department);

    var before = txnDelta_(config, entry.TxnType, Number(entry.Qty) || 0);
    if (!ALLOW_NEGATIVE_STOCK) {
      var projected = balanceAfter_(config, code, -before.inDelta, -before.outDelta);
      if (projected < 0) {
        throw new Error('Removing this receipt would leave ' + code + ' at ' + projected +
                        '. Reverse the issues that used it first.');
      }
    }

    var masterRow = lookupMasterRow_(config, code) || {};
    var newBalance = applyStockTotals_(config, code, -before.inDelta, -before.outDelta, masterRow);

    // Written to the audit log in full before the row goes, so a deletion
    // can be reconstructed from the log if it turns out to be wrong.
    logAudit_(session.Username, 'DELETE_TXN', config.ledgerSheet, entry.TxnID,
              JSON.stringify(entry), '');
    found.sheet.deleteRow(found.rowIdx);

    return { txnId: entry.TxnID, code: code, newBalance: newBalance };
  });
}

/* =========================================================================
   ACCESS CONTROL

   Two layers.

   1. Admin holds everything, always. It is not a grant and cannot be taken
      away by any template or override, so the last Admin can never lock
      everyone — including themselves — out of the app.

   2. Everyone else starts from the template for their department's
      category (Production, HOD, Manager, Management, Admin) and is then
      adjusted per account by an Admin.

   Both live in the sheet rather than in this file, so a new category or a
   changed default is an edit, not a deployment:

     Access_Templates   one row per category, one column per feature
     Users.Access       per-account overrides, e.g. "+stock.recalculate,-stock.rm.post"

   The resolved set is written onto the session row at sign-in, which is how
   ELE Tracker and Production Tracker get it: they already read that row to
   validate the token, so they pay nothing extra for it. Changing someone's
   access rewrites their live sessions too, so it takes effect on their next
   action instead of their next sign-in.

   Managing accounts is deliberately NOT in this catalog. It stays Admin-only
   and ungrantable — "no Team access" for anyone else, whatever their tier.
   ========================================================================= */

// The features an Admin can hand out, in the order the Admin Panel shows
// them. Adding one here is all that's needed for it to appear there.
var PERMISSIONS = [
  { key: 'stock.rm.post',    group: 'Stock',       label: 'Post RM transactions' },
  { key: 'stock.fg.post',    group: 'Stock',       label: 'Post FG transactions' },
  { key: 'stock.post.any',   group: 'Stock',       label: 'Post for any department' },
  { key: 'stock.rm.items',   group: 'Item master', label: 'Add & edit RM items' },
  { key: 'stock.fg.items',   group: 'Item master', label: 'Add & edit FG items' },
  { key: 'stock.rm.retire',  group: 'Item master', label: 'Retire & restore RM items' },
  { key: 'stock.fg.retire',  group: 'Item master', label: 'Retire & restore FG items' },
  { key: 'stock.import',     group: 'Item master', label: 'Bulk import items' },
  { key: 'stock.export',     group: 'Reports',     label: 'Export to Excel' },
  { key: 'stock.ledger.edit',  group: 'Ledger',      label: 'Correct a ledger entry' },
  { key: 'stock.ledger.delete',group: 'Ledger',      label: 'Delete a ledger entry' },
  { key: 'stock.recalculate',group: 'Maintenance', label: 'Recalculate balances' },
  { key: 'stock.mail.rm',    group: 'Mail',        label: 'Daily RM stock mail' },
  { key: 'stock.mail.fg',    group: 'Mail',        label: 'Daily FG stock mail' },
  { key: 'stock.mail.test',  group: 'Mail',        label: 'Send test mail' },

  // Production Tracker. Managed here with everything else, because the
  // people, departments and tiers are here; the session carries these to
  // Production Tracker's backend, which reads the same Sessions row.
  { key: 'pt.handover.send',    app: 'pt', group: 'Handovers',        label: 'Send and return from own department' },
  { key: 'pt.handover.receive', app: 'pt', group: 'Handovers',        label: 'Receive into own department' },
  { key: 'pt.handover.any',     app: 'pt', group: 'Handovers',        label: 'Act for any department' },
  { key: 'pt.wip.record',       app: 'pt', group: 'Work in progress', label: 'Record produced, scrapped and used' },
  { key: 'pt.view.all',         app: 'pt', group: 'Viewing',          label: 'See every department' },
  { key: 'pt.orders.place',     app: 'pt', group: 'Orders',           label: 'Place orders' },

  // ELE Tracker. Which stages count as "your department's" is set in ELE's
  // own Stage_Access tab, next to the stages themselves.
  { key: 'ele.stage.record', app: 'ele', group: 'Stages', label: 'Record your department\u2019s stages' },
  { key: 'ele.stage.any',    app: 'ele', group: 'Stages', label: 'Record any stage' }
];

// Which app each feature belongs to — the catalog entry's own app, or Stock
// Management's where it names none.
PERMISSIONS.forEach(function (p) { if (!p.app) p.app = 'stock'; });

function permissionKeys_() {
  return PERMISSIONS.map(function (p) { return p.key; });
}

// What a person RECEIVES rather than what they may DO. Admin holds every
// power by definition, but that must not mean every Admin is mailed every
// report whether they want it or not — so these follow the account's own
// choice for Admins too, and are off until ticked.
var SUBSCRIPTION_KEYS = ['stock.mail.rm', 'stock.mail.fg'];
function isSubscription_(key) { return SUBSCRIPTION_KEYS.indexOf(key) > -1; }

// Seeded once, then owned by the sheet. 'Default' covers an account whose
// department has no category yet, so a half-filled Departments tab can't
// leave someone with nothing.
// Production Tracker defaults: Management sees everything and enters
// nothing; a Manager acts for any department; everyone else works their
// own department's handovers.
var PT_OWN_DEPARTMENT = ['pt.handover.send', 'pt.handover.receive', 'pt.wip.record'];

// ELE Tracker defaults: everyone records their own department's stages; a
// Manager can record any; Management, which enters nothing, records none.
var ELE_OWN_DEPARTMENT = ['ele.stage.record'];

var DEFAULT_ACCESS_TEMPLATES = {
  'Admin':      permissionKeys_(),
  'Management': ['stock.export', 'stock.mail.rm', 'stock.mail.fg', 'pt.view.all'],
  'Manager':    ['stock.rm.post', 'stock.fg.post', 'stock.post.any', 'stock.rm.items', 'stock.fg.items',
                 'stock.import', 'stock.export', 'stock.mail.rm', 'stock.mail.fg',
                 'pt.handover.any', 'pt.view.all', 'pt.orders.place', 'ele.stage.any'].concat(PT_OWN_DEPARTMENT, ELE_OWN_DEPARTMENT),
  'HOD':        ['stock.rm.post', 'stock.fg.post', 'stock.export', 'stock.mail.rm', 'stock.mail.fg'].concat(PT_OWN_DEPARTMENT, ELE_OWN_DEPARTMENT),
  'Production': ['stock.rm.post', 'stock.fg.post', 'stock.export'].concat(PT_OWN_DEPARTMENT, ELE_OWN_DEPARTMENT),
  'Default':    ['stock.rm.post', 'stock.fg.post', 'stock.export'].concat(PT_OWN_DEPARTMENT, ELE_OWN_DEPARTMENT)
};

function getOrCreateAccessTemplates_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Access_Templates');
  var headers = ['Category'].concat(permissionKeys_());
  if (!sheet) {
    sheet = ss.insertSheet('Access_Templates');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sheet, headers.length);
  } else {
    // A feature added since the tab was made gets its column now — filled,
    // for every tier already there, from that tier's defaults. Left blank
    // it would read as "no" everywhere, and nobody below Admin would get a
    // new feature until each tier was ticked by hand.
    var before = sheetMeta_(sheet).map;
    var added = headers.filter(function (h) { return before[h] === undefined; });
    ensureColumns_(sheet, headers);
    if (added.length) {
      var meta = sheetMeta_(sheet);
      readSheetAsObjects_(sheet).forEach(function (r, i) {
        if (!r.Category) return;
        var granted = DEFAULT_ACCESS_TEMPLATES[r.Category] || DEFAULT_ACCESS_TEMPLATES['Default'];
        var fields = {};
        added.forEach(function (k) { fields[k] = granted.indexOf(k) > -1; });
        writeRowCells_(sheet, i + 2, stageCells_(meta.map, fields));
      });
    }
  }

  // Every category in use gets a row, seeded from the defaults above.
  var existing = {};
  readSheetAsObjects_(sheet).forEach(function (r) { if (r.Category) existing[r.Category] = true; });
  var wanted = Object.keys(DEFAULT_ACCESS_TEMPLATES);
  getDepartments_().forEach(function (d) { if (d.category && wanted.indexOf(d.category) === -1) wanted.push(d.category); });

  wanted.forEach(function (category) {
    if (existing[category]) return;
    var granted = DEFAULT_ACCESS_TEMPLATES[category] || DEFAULT_ACCESS_TEMPLATES['Default'];
    var row = { Category: category };
    permissionKeys_().forEach(function (k) { row[k] = granted.indexOf(k) > -1; });
    appendRowByHeader_(sheet, row);
  });
  return sheet;
}

// Read once per execution: resolving access for every account (Team &
// Access, the daily mail, a template save) used to re-read the tab, and
// re-run its create-and-migrate check, once per account. Anything that
// writes the tab clears this.
var _accessTemplatesCache_ = null;
function readAccessTemplates_() {
  if (_accessTemplatesCache_) return _accessTemplatesCache_;
  var out = {};
  readSheetAsObjects_(getOrCreateAccessTemplates_()).forEach(function (r) {
    if (!r.Category) return;
    var set = {};
    permissionKeys_().forEach(function (k) { set[k] = isTruthyFlag_(r[k]); });
    out[r.Category] = set;
  });
  _accessTemplatesCache_ = out;
  return out;
}

// "+stock.recalculate,-stock.rm.post" -> { 'stock.recalculate': true, 'stock.rm.post': false }
function parseAccessOverrides_(raw) {
  var out = {};
  String(raw || '').split(',').forEach(function (part) {
    part = part.trim();
    if (!part) return;
    var grant = part.charAt(0) !== '-';
    var key = part.replace(/^[+-]/, '').trim();
    if (key) out[key] = grant;
  });
  return out;
}

function stringifyAccessOverrides_(overrides) {
  return Object.keys(overrides || {})
    .map(function (k) { return (overrides[k] ? '+' : '-') + k; })
    .join(',');
}

// Template for the account's tier, then its own overrides on top.
function resolveAccess_(user, templates) {
  var granted = {};
  if (String(user.Role) === 'Admin') {
    var own = parseAccessOverrides_(user.Access);
    permissionKeys_().forEach(function (k) { granted[k] = isSubscription_(k) ? own[k] === true : true; });
    return granted;
  }
  templates = templates || readAccessTemplates_();
  var category = deptCategoryFor_(user.Department) || 'Default';
  var template = templates[category] || templates['Default'] || {};
  permissionKeys_().forEach(function (k) { granted[k] = !!template[k]; });

  var overrides = parseAccessOverrides_(user.Access);
  Object.keys(overrides).forEach(function (k) {
    if (permissionKeys_().indexOf(k) > -1) granted[k] = overrides[k];
  });
  return granted;
}

function grantedKeys_(granted) {
  return Object.keys(granted).filter(function (k) { return granted[k]; });
}

// Admin is absolute; everyone else is whatever their session was given.
function can_(session, key) {
  if (!session) return false;
  if (String(session.Role) === 'Admin') return true;
  return String(session.Perms || '').split(',').indexOf(key) > -1;
}

function requirePerm_(session, key) {
  if (can_(session, key)) return;
  var label = PERMISSIONS.filter(function (p) { return p.key === key; })[0];
  throw new Error('Your account doesn\'t have access to ' + (label ? label.label.toLowerCase() : key) +
                  '. Ask an Admin to enable it.');
}

/* =========================================================================
   ACCESS MIGRATION

   Run ADMIN_migrateAccess() once from the editor after pasting this file.
   It is safe to run again: it only ever adds what is missing.

   The point of it is that nobody's access changes on the day it goes in.
   Mail subscriptions used to live in DailyMailRM/DailyMailFG; those become
   overrides so each account's resolved access matches exactly what it had,
   whatever its tier's template happens to say.
   ========================================================================= */

// What each role could do before this system existed. The migration keeps
// anything an account would otherwise lose, because a tier's defaults are a
// starting point for the future — not a reason for someone to find a button
// missing on the morning this goes in.
var LEGACY_ROLE_ACCESS = {
  'Admin': ['stock.rm.post', 'stock.fg.post', 'stock.post.any', 'stock.rm.items', 'stock.fg.items',
            'stock.rm.retire', 'stock.fg.retire', 'stock.import', 'stock.export', 'stock.recalculate', 'stock.mail.test'],
  'Manager': ['stock.rm.post', 'stock.fg.post', 'stock.post.any', 'stock.rm.items', 'stock.fg.items',
              'stock.import', 'stock.export'],
  'Supervisor': ['stock.rm.post', 'stock.fg.post', 'stock.export']
};

function ADMIN_migrateAccess() {
  var ss = getSpreadsheet_();
  var usersSheet = ss.getSheetByName('Users');

  // The carry-over below happens once. Run again later, it would re-grant
  // every old role power an Admin has since taken away, and read blank mail
  // columns as "unsubscribe". Access_Templates is the first thing this
  // function creates, so finding it already there means that has been done.
  var firstRun = !ss.getSheetByName('Access_Templates');

  ensureColumns_(usersSheet, ['Access']);
  ensureColumns_(getOrCreateSessionsSheet_(), ['Perms']);
  getOrCreateAccessTemplates_();

  var templates = readAccessTemplates_();
  var users = readSheetAsObjects_(usersSheet);
  var report = { firstRun: firstRun, accounts: 0, mailKept: 0, kept: [], widened: [], categories: Object.keys(templates) };

  if (firstRun) {
    users.forEach(function (u) {
      if (!u.Username) return;
      var overrides = parseAccessOverrides_(u.Access);
      var mailPairs = [['DailyMailRM', 'stock.mail.rm'], ['DailyMailFG', 'stock.mail.fg']];

      if (String(u.Role) === 'Admin') {
        // An Admin's powers need no carrying over; what they chose to
        // receive by mail does.
        mailPairs.forEach(function (pair) {
          if (isTruthyFlag_(u[pair[0]])) { overrides[pair[1]] = true; report.mailKept++; }
        });
      } else {
        var category = deptCategoryFor_(u.Department) || 'Default';
        var template = templates[category] || templates['Default'] || {};

        // Anything the account could do under the old role rules that its
        // new tier doesn't grant is kept as an override. Anything the tier
        // adds is left in place and reported, so the widening is visible.
        var legacy = LEGACY_ROLE_ACCESS[String(u.Role)] || [];
        legacy.forEach(function (k) {
          if (!template[k]) { overrides[k] = true; report.kept.push(u.Username + ' ' + k); }
        });
        permissionKeys_().forEach(function (k) {
          if (template[k] && legacy.indexOf(k) === -1 && !isSubscription_(k)) report.widened.push(u.Username + ' ' + k);
        });

        // Keep today's mail exactly as it is, however the template votes —
        // but only where the old column says something. A blank cell is no
        // instruction at all.
        mailPairs.forEach(function (pair) {
          var cell = u[pair[0]];
          if (cell === '' || cell === undefined || cell === null) return;
          var had = isTruthyFlag_(cell);
          if (had !== !!template[pair[1]]) { overrides[pair[1]] = had; report.mailKept++; }
          else delete overrides[pair[1]];
        });
      }

      var row = readRowByValue_(usersSheet, 'Username', u.Username);
      if (row) {
        writeRowCells_(usersSheet, row.rowIdx, stageCells_(row.map, { Access: stringifyAccessOverrides_(overrides) }));
        report.accounts++;
      }
    });
  } else {
    report.note = 'Already migrated — every account\u2019s access was left exactly as it is. ' +
                  'Only the tabs and columns were checked, and live sessions refreshed.';
  }

  // Live sessions get their resolved set now, so no one has to sign in again.
  var fresh = readAccessTemplates_();
  readSheetAsObjects_(usersSheet).forEach(function (u) {
    if (u.Username) refreshSessionsFor_(u.Username, resolveAccess_(u, fresh));
  });

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

// What a given account can do right now, and why. Handy when someone says
// a button is missing.
function ADMIN_explainAccess(username) {
  var u = readSheetAsObjects_(getSpreadsheet_().getSheetByName('Users'))
    .filter(function (r) { return r.Username === username; })[0];
  if (!u) { Logger.log('No such user: ' + username); return null; }
  var category = deptCategoryFor_(u.Department) || 'Default';
  var out = {
    username: u.Username, role: u.Role, department: u.Department, category: category,
    overrides: parseAccessOverrides_(u.Access),
    effective: grantedKeys_(resolveAccess_(u))
  };
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

/* =========================================================================
   ADMIN PANEL API — Admin only, and not grantable
   ========================================================================= */

// The effective set for whoever is asking. Admin is absolute; everyone
// else gets what their session was given.
function sessionPermissions_(session) {
  var out = {};
  var admin = session && String(session.Role) === 'Admin';
  var granted = String((session && session.Perms) || '').split(',');
  permissionKeys_().forEach(function (k) { out[k] = admin || granted.indexOf(k) > -1; });
  return out;
}

function getAccessControl_() {
  var templates = readAccessTemplates_();
  var users = readSheetAsObjects_(getSpreadsheet_().getSheetByName('Users')).map(function (u) {
    return {
      username: u.Username,
      name: u.Name,
      role: u.Role,
      department: u.Department,
      category: deptCategoryFor_(u.Department) || 'Default',
      email: u.Email || '',
      phone: u.Phone || '',
      active: isActiveValue_(u.Active),
      overrides: parseAccessOverrides_(u.Access),
      effective: resolveAccess_(u, templates)
    };
  });
  return {
    permissions: PERMISSIONS,
    templates: templates,
    departments: getDepartments_(),
    users: users
  };
}

// Overrides only — the template stays the account's baseline, so a category
// whose defaults change later still moves the accounts that follow it.
function setUserAccess_(session, params) {
  var username = String(params.username || '').trim();
  if (!username) throw new Error('Username is required');

  var sheet = getSpreadsheet_().getSheetByName('Users');
  ensureColumns_(sheet, ['Access']);
  var row = readRowByValue_(sheet, 'Username', username);
  if (!row) throw new Error('User not found: ' + username);
  var user = {};
  Object.keys(row.map).forEach(function (h) { user[h] = row.values[row.map[h]]; });

  var before = String(user.Access || '');
  var wanted = params.permissions || {};
  var templates = readAccessTemplates_();
  // Starts from what the account already has, and changes only the switches
  // that were sent. Team & Access shows one app at a time, so a save from
  // the Production Tracker view must leave Stock Management's overrides
  // exactly as they were — rebuilding from scratch would wipe them.
  var overrides = parseAccessOverrides_(user.Access);

  if (String(user.Role) === 'Admin') {
    // Every power is theirs already and can't be narrowed; only what they
    // receive by mail is theirs to choose. Anything else sent is ignored.
    SUBSCRIPTION_KEYS.forEach(function (k) {
      if (!(k in wanted)) return;
      if (wanted[k] === true) overrides[k] = true; else delete overrides[k];
    });
  } else {
    // Anything matching the template is stored as inheritance rather than
    // as an override, so the row stays readable and keeps following its tier.
    var category = deptCategoryFor_(user.Department) || 'Default';
    var template = templates[category] || templates['Default'] || {};
    permissionKeys_().forEach(function (k) {
      if (!(k in wanted)) return;
      if (!!wanted[k] !== !!template[k]) overrides[k] = !!wanted[k];
      else delete overrides[k];
    });
  }

  writeRowCells_(sheet, row.rowIdx, stageCells_(row.map, { Access: stringifyAccessOverrides_(overrides) }));
  user.Access = stringifyAccessOverrides_(overrides);
  var granted = resolveAccess_(user, templates);
  refreshSessionsFor_(username, granted);
  if (before !== user.Access) logAudit_(session.Username, 'ACCESS', 'Users', username, before || '(follows tier)', user.Access || '(follows tier)');
  return { username: username, overrides: overrides, effective: granted };
}

function setAccessTemplate_(session, params) {
  var category = String(params.category || '').trim();
  if (!category) throw new Error('Category is required');
  var sheet = getOrCreateAccessTemplates_();
  var row = readRowByValue_(sheet, 'Category', category);
  var wanted = params.permissions || {};
  var previous = readAccessTemplates_()[category] || {};
  if (!row) {
    var fresh = { Category: category };
    permissionKeys_().forEach(function (k) { fresh[k] = !!wanted[k]; });
    appendRowByHeader_(sheet, fresh);
  } else {
    var fields = {};
    permissionKeys_().forEach(function (k) { if (k in wanted) fields[k] = !!wanted[k]; });
    writeRowCells_(sheet, row.rowIdx, stageCells_(row.map, fields));
  }
  _accessTemplatesCache_ = null; // the tab just changed
  // Everyone following this tier is affected, so their live sessions are
  // brought up to date rather than left on yesterday's rules.
  var templates = readAccessTemplates_();
  readSheetAsObjects_(getSpreadsheet_().getSheetByName('Users')).forEach(function (u) {
    if (!u.Username || String(u.Role) === 'Admin') return;
    if ((deptCategoryFor_(u.Department) || 'Default') !== category) return;
    refreshSessionsFor_(u.Username, resolveAccess_(u, templates));
  });
  // Only what actually changed, both ways — a whole row of TRUE/FALSE says
  // nothing about which switch someone flipped.
  var was = {}, now = {};
  permissionKeys_().forEach(function (k) {
    if (k in wanted && !!wanted[k] !== !!previous[k]) { was[k] = !!previous[k]; now[k] = !!wanted[k]; }
  });
  if (Object.keys(now).length) logAudit_(session.Username, 'ACCESS_TEMPLATE', 'Access_Templates', category, JSON.stringify(was), JSON.stringify(now));
  return { category: category, permissions: templates[category] };
}

// Rewrites the live sessions of one account so a change of access takes
// effect on their next action, not their next sign-in.
// `identity` optionally rewrites Role/Department too, for when those change.
function refreshSessionsFor_(username, granted, identity) {
  var sheet = getOrCreateSessionsSheet_();
  ensureColumns_(sheet, ['Perms']);
  var fields = { Perms: grantedKeys_(granted).join(',') };
  if (identity) {
    if (identity.Role !== undefined) fields.Role = identity.Role;
    if (identity.Department !== undefined) fields.Department = identity.Department;
  }
  // Read and written under the lock, so a sign-in deleting rows at the same
  // moment can't shift these rows and land this account's role and access
  // on someone else's session.
  return withSessionsLock_(function () {
    var meta = sheetMeta_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;
    var values = sheet.getRange(2, 1, lastRow - 1, meta.lastCol).getValues();
    var touched = 0;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][meta.map['Username']]) !== String(username)) continue;
      writeRowCells_(sheet, i + 2, stageCells_(meta.map, fields));
      dropCachedSession_(String(values[i][meta.map['Token']]));
      touched++;
    }
    return touched;
  });
}

/* =========================================================================
   DEPARTMENTS

   The sheet owns two things per department: its name, which is what a stock
   movement is booked to and what an account belongs to, and its category,
   which is the tier that decides default access — Admin, Management,
   Manager, HOD, Production.

   Header spellings are tolerated rather than assumed. Renaming Dept_Name to
   DeptName (or back) is the kind of edit that should not silently empty
   every department dropdown in the app.
   ========================================================================= */

// Read once per execution. deptCategoryFor_ runs once per account whenever
// access is resolved for a list of people, and each call used to re-read
// this tab. Nothing in this file writes it.
var _departmentsCache_ = null;
function getDepartments_() {
  if (!_departmentsCache_) _departmentsCache_ = readDepartments_();
  return _departmentsCache_;
}

function readDepartments_() {
  var sheet = getSpreadsheet_().getSheetByName('Departments');
  if (!sheet) return [];
  return readSheetAsObjects_(sheet).map(function (r) {
    var category = r.Dept_category || r.Dept_Category || r.DeptCategory || r.Category || '';
    var flag = [r.Stock_Use, r.StockUse, r['Stock Use']].filter(function (v) {
      return v !== undefined && v !== null && v !== '';
    })[0];
    return {
      id: r.DeptID || r.Dept_ID || r.DeptId || '',
      name: r.Dept_Name || r.DeptName || r.Name || '',
      category: category,
      // Whether stock is ever booked to this department. Set per department
      // in a Stock_Use column; where that's blank, the tiers that are about
      // people rather than places — Admin, Management, Manager — say no.
      stockUse: flag !== undefined ? isTruthyFlag_(flag) : NON_STOCK_CATEGORIES.indexOf(category) === -1
    };
  }).filter(function (d) { return d.name; });
}

var NON_STOCK_CATEGORIES = ['Admin', 'Management', 'Manager'];

// A department marked as not using stock can't have stock booked to it.
// A name that isn't in the tab at all is let through, so an entry booked
// before a department was renamed can still be corrected.
function assertStockDepartment_(name) {
  if (!name) return;
  var d = getDepartments_().filter(function (x) { return x.name === name; })[0];
  if (d && !d.stockUse) {
    throw new Error('Stock isn\u2019t booked to ' + name + ' — pick the department where the material is used or kept.');
  }
}

// The tier an account inherits from the department it belongs to.
function deptCategoryFor_(deptName) {
  if (!deptName) return '';
  var hit = getDepartments_().filter(function (d) { return d.name === deptName; })[0];
  return hit ? hit.category : '';
}

function getOrCreateSessionsSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Sessions');
  if (!sheet) {
    var headers = ['Token', 'Username', 'Role', 'Department', 'CreatedAt', 'ExpiresAt', 'Perms'];
    sheet = ss.insertSheet('Sessions');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sheet, headers.length);
  }
  return sheet;
}

function updateLastLogin_(usersSheet, username, when) {
  var rowIdx = findRowIndexByValue_(usersSheet, 'Username', username);
  if (rowIdx === -1) return;
  var map = headerIndexMap_(usersSheet);
  usersSheet.getRange(rowIdx, map['LastLogin'] + 1).setValue(when);
}

// salt is optional for backward compatibility: accounts created before this
// upgrade have PasswordHash = HMAC(password, secret) with no salt suffix.
// Passing a falsy salt reproduces that exact legacy message format so old
// hashes keep validating; a truthy salt switches to the salted format.
// login_ migrates a legacy account to a salted hash the next time it logs
// in successfully (see below) — no bulk migration pass needed.
function hashPassword_(password, secretKey, salt) {
  var message = salt ? (password + '|' + salt) : password;
  var rawHmac = Utilities.computeHmacSha256Signature(message, secretKey);
  return rawHmac.map(function (byte) {
    var v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function generateSalt_() {
  return Utilities.getUuid();
}

// The secret lives in Script Properties, which only people who can edit this
// script can see. In the Settings tab, anyone who could view the spreadsheet
// could read it, and with the PasswordHash column try passwords offline.
// The tab is still read as a fallback, so nothing breaks before
// ADMIN_moveAppSecretToProperties() has been run.
var APP_SECRET_MOVED = 'MOVED_TO_SCRIPT_PROPERTIES';

function getAppSecret_() {
  var stored = PropertiesService.getScriptProperties().getProperty('APP_SECRET');
  if (stored) return stored;
  var settings = readSheetAsObjects_(getSpreadsheet_().getSheetByName('Settings'));
  var row = settings.filter(function (s) { return s.Key === 'APP_SECRET'; })[0];
  if (!row || !row.Value || row.Value === 'SET_IN_PHASE_2' || row.Value === APP_SECRET_MOVED) {
    throw new Error('APP_SECRET not set. Run ADMIN_generateAppSecret() from the Apps Script editor first.');
  }
  return row.Value;
}

function setSettingValue_(sheet, key, value) {
  var rowIdx = findRowIndexByValue_(sheet, 'Key', key);
  if (rowIdx === -1) {
    appendRowByHeader_(sheet, { Key: key, Value: value });
  } else {
    var map = headerIndexMap_(sheet);
    sheet.getRange(rowIdx, map['Value'] + 1).setValue(value);
  }
}

// ---------- Manual admin tools (run from the Apps Script editor) ----------

// Fresh install only: a new secret invalidates every password already set.
function ADMIN_generateAppSecret() {
  var secret = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('APP_SECRET', secret);
  setSettingValue_(getSpreadsheet_().getSheetByName('Settings'), 'APP_SECRET', APP_SECRET_MOVED);
  Logger.log('APP_SECRET generated and saved in Script Properties. Next: set the admin password.');
}

// Run once from the editor. Copies the existing secret from the Settings tab
// into Script Properties, checks the copy, and only then blanks the tab's
// copy. The value is unchanged, so every password keeps working.
// Before running it, keep a copy of the secret somewhere safe (a password
// manager): it's the one thing that can't be recovered if lost.
function ADMIN_moveAppSecretToProperties() {
  var props = PropertiesService.getScriptProperties();
  var settingsSheet = getSpreadsheet_().getSheetByName('Settings');
  var row = readSheetAsObjects_(settingsSheet).filter(function (s) { return s.Key === 'APP_SECRET'; })[0];
  var inSheet = row ? String(row.Value || '') : '';
  var already = props.getProperty('APP_SECRET');

  if (!inSheet || inSheet === APP_SECRET_MOVED) {
    Logger.log(already ? 'Already moved — the secret is in Script Properties.' : 'No secret found in the Settings tab or Script Properties.');
    return;
  }
  if (already && already !== inSheet) {
    throw new Error('Script Properties already holds a different APP_SECRET. Nothing was changed — check which one the live passwords use.');
  }
  props.setProperty('APP_SECRET', inSheet);
  if (props.getProperty('APP_SECRET') !== inSheet) throw new Error('Copy did not verify. Nothing was blanked.');
  setSettingValue_(settingsSheet, 'APP_SECRET', APP_SECRET_MOVED);
  Logger.log('APP_SECRET moved to Script Properties; the Settings tab copy is blanked.');
}

// Break-glass fallback — set a user's password directly from the Apps
// Script editor. Normal path is: Admin creates the user in the Users
// panel with no password, then the user (or Admin, via "Send reset")
// sets one through the OTP email flow. Use this only if email delivery
// is broken and someone needs in immediately.
function ADMIN_setUserPassword(username, plainPassword) {
  var sheet = getSpreadsheet_().getSheetByName('Users');
  var rowIdx = findRowIndexByValue_(sheet, 'Username', username);
  if (rowIdx === -1) {
    Logger.log('User not found: ' + username);
    return;
  }
  ensureColumns_(sheet, ['Salt']);
  var salt = generateSalt_();
  var hash = hashPassword_(plainPassword, getAppSecret_(), salt);
  var map = headerIndexMap_(sheet);
  sheet.getRange(rowIdx, map['PasswordHash'] + 1).setValue(hash);
  sheet.getRange(rowIdx, map['Salt'] + 1).setValue(salt);
  sheet.getRange(rowIdx, map['Active'] + 1).setValue(true);
  Logger.log('Password set for ' + username + ' — they can log in now.');
}

// Optional — RM_Stock_Ledger now grows its Supplier/InvoiceNo/InvoiceDate
// columns automatically on first RM receipt, and PasswordResets is
// created lazily on first password reset. Run this if you'd rather have
// both in place proactively. Idempotent — safe to run any time.
function ADMIN_migrateSchema() {
  var ss = getSpreadsheet_();
  ensureColumns_(ss.getSheetByName('RM_Stock_Ledger'), ['Supplier', 'InvoiceNo', 'InvoiceDate']);
  getOrCreatePasswordResetsSheet_();
  Logger.log('Schema up to date.');
}

function ADMIN_cleanupExpiredSessions() {
  var sheet = getOrCreateSessionsSheet_();
  // Read and deleted under the lock — a sign-in at 2am deletes rows too.
  var removed = withSessionsLock_(function () {
    var data = sheet.getDataRange().getValues();
    var expIdx = data[0].indexOf('ExpiresAt');
    var now = Date.now();
    var count = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      var exp = new Date(data[i][expIdx]).getTime();
      // A blank expiry is treated as expired here too (see validateSession_).
      if (isNaN(exp) || exp < now) {
        sheet.deleteRow(i + 1);
        count++;
      }
    }
    return count;
  });
  Logger.log('Removed ' + removed + ' expired session(s).');
}

// Run once from the Apps Script editor. Installs a nightly trigger so
// ADMIN_cleanupExpiredSessions() (and expired OTPs) get swept
// automatically — Sessions/PasswordResets no longer grow unbounded.
function ADMIN_installNightlyCleanupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'ADMIN_nightlyCleanup_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('ADMIN_nightlyCleanup_').timeBased().everyDays(1).atHour(2).create();
  Logger.log('Nightly cleanup trigger installed (runs ~2am script timezone).');
}
function ADMIN_nightlyCleanup_() {
  ADMIN_cleanupExpiredSessions();
  cleanupExpiredPasswordResets_();
  nightlyDriftCheck_();
}

/* =========================================================================
   BALANCE DRIFT CHECK
   -------------------------------------------------------------------------
   Balances are maintained incrementally by applyStockDelta_, which is fast
   but means the stored CurrentBalance and the sum of the ledger can diverge
   — a half-completed write, a manual edit in the sheet, a restored row. That
   divergence is silent: the numbers still look plausible.

   This recomputes from the ledger in memory and reports any mismatch. It
   deliberately does NOT correct anything. An automatic correction on a
   nightly trigger could quietly overwrite a legitimate manual fix and would
   destroy the evidence of whatever caused the drift. A human decides, then
   triggers a recalculation from the app (Admin → Recalculate stock).
   ========================================================================= */

var DRIFT_TOLERANCE = 0.001; // float noise from fractional Kg/Ltr quantities

function computeLedgerBalances_(config) {
  var ss = getSpreadsheet_();
  var ledgerRows = readSheetAsObjects_(ss.getSheetByName(config.ledgerSheet));
  var totals = {};
  for (var i = 0; i < ledgerRows.length; i++) {
    var r = ledgerRows[i];
    var code = String(r[config.codeField] || '').trim();
    if (!code) continue;
    var qty = Number(r.Qty) || 0;
    if (totals[code] === undefined) totals[code] = 0;
    if (config.inTypes.indexOf(r.TxnType) > -1) totals[code] += qty;
    else if (config.outTypes.indexOf(r.TxnType) > -1) totals[code] -= qty;
    else totals[code] += qty; // Adjustment carries its own sign
  }
  return totals;
}

function findDrift_(config) {
  var ledgerTotals = computeLedgerBalances_(config);
  var stockRows = readSheetAsObjects_(getSpreadsheet_().getSheetByName(config.stockSheet));
  var drift = [];
  for (var i = 0; i < stockRows.length; i++) {
    var row = stockRows[i];
    var code = String(row[config.codeField] || '').trim();
    if (!code) continue;
    var opening = Number(row.OpeningStock) || 0;
    var stored = Number(row.CurrentBalance) || 0;
    var expected = opening + (ledgerTotals[code] || 0);
    if (Math.abs(expected - stored) > DRIFT_TOLERANCE) {
      drift.push({ code: code, stored: stored, expected: expected, diff: stored - expected });
    }
  }
  return drift;
}

/* =========================================================================
   HOLIDAYS

   Dates listed in the Holidays column of the Dropdowns tab, and every day
   named in its Weekly Off column, get no automatic mail — neither the
   morning stock workbook nor the drift alert. Nothing is
   lost by skipping: stock and drift are states, not events, so the next
   working day's mail reports them as they then stand. Mail sent by hand
   (Send test mail) is never held back.

   Compared as calendar days in the spreadsheet's own time zone, so a date
   typed as 27/09/2026 matches the morning of 27 September in India rather
   than the UTC instant that happens to fall on the 26th.
   ========================================================================= */

function holidayToday_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Dropdowns');
  if (!sheet) return null;
  var meta = sheetMeta_(sheet);
  var col = function (names) {
    for (var n = 0; n < names.length; n++) if (meta.map[names[n]] !== undefined) return meta.map[names[n]];
    return undefined;
  };
  var dateCol = col(['Holidays', 'Holiday']);
  var reasonCol = col(['Holiday Reason', 'Holiday_Reason', 'HolidayReason']);
  var weekCol = col(['Weekly Off', 'Weekly_Off', 'WeeklyOff']);
  if (dateCol === undefined && weekCol === undefined) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var tz = ss.getSpreadsheetTimeZone ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
  var now = new Date();
  var today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var rows = sheet.getRange(2, 1, lastRow - 1, meta.lastCol).getValues();

  // A dated holiday first, so a festival that falls on the weekly off-day is
  // reported by its own name.
  if (dateCol !== undefined) {
    for (var i = 0; i < rows.length; i++) {
      var cell = rows[i][dateCol];
      if (cell === '' || cell === null) continue;
      var d = cell instanceof Date ? cell : parseDay_(cell);
      if (isNaN(d.getTime())) continue;
      if (Utilities.formatDate(d, tz, 'yyyy-MM-dd') === today) {
        return { date: today, reason: reasonCol !== undefined ? String(rows[i][reasonCol] || '') : '' };
      }
    }
  }

  // A day name in Weekly Off — "Sunday", or "Sun" — holds mail every week,
  // so the Holidays list only needs the festivals and one-off closures.
  if (weekCol !== undefined) {
    var dayName = Utilities.formatDate(now, tz, 'EEEE');
    var lowered = dayName.toLowerCase();
    for (var w = 0; w < rows.length; w++) {
      var v = String(rows[w][weekCol] || '').trim().toLowerCase();
      if (v.length >= 3 && lowered.indexOf(v) === 0) return { date: today, reason: 'Weekly off (' + dayName + ')' };
    }
  }
  return null;
}

function nightlyDriftCheck_() {
  try {
    var report = [];
    [RM_CONFIG, FG_CONFIG].forEach(function (config) {
      var drift = findDrift_(config);
      if (drift.length) report.push({ label: config.key.toUpperCase(), rows: drift });
    });
    if (!report.length) return; // silence is the good outcome — no daily noise
    var holiday = holidayToday_();
    if (holiday) { Logger.log('Drift alert held for holiday: ' + (holiday.reason || holiday.date)); return; }

    var html = '<p>Stored balances disagree with the ledger for the items below. ' +
      'Nothing has been changed automatically — review, then use ' +
      'Recalculate stock in the app if the ledger is correct.</p>';
    report.forEach(function (section) {
      html += '<h3>' + section.label + '</h3><table border="1" cellpadding="6" cellspacing="0">' +
        '<tr><th>Code</th><th>Stored</th><th>Ledger implies</th><th>Difference</th></tr>';
      section.rows.forEach(function (d) {
        html += '<tr><td>' + escapeHtml_(d.code) + '</td><td>' + d.stored + '</td><td>' + d.expected +
          '</td><td>' + (d.diff > 0 ? '+' : '') + d.diff + '</td></tr>';
      });
      html += '</table>';
    });

    var admins = readSheetAsObjects_(getSpreadsheet_().getSheetByName('Users'))
      .filter(function (u) { return u.Role === 'Admin' && isActiveValue_(u.Active) && u.Email; });

    admins.forEach(function (admin) {
      sendMail_({
        to: admin.Email,
        subject: '[' + COMPANY_NAME + '] Stock balance drift detected',
        body: 'Stored balances disagree with the ledger. Open the HTML version for details.',
        htmlBody: html
      });
    });
  } catch (err) {
    // A failure here must never take down session/OTP cleanup above it.
    console.error('nightlyDriftCheck_ failed: ' + err.message);
  }
}

/* =========================================================================
   USER MANAGEMENT — Admin-only. PasswordHash is never sent to the client;
   new users are created with no password (Active but can't log in) and
   set their own via the OTP reset flow below, or an Admin can trigger
   that same email for them with "Send reset".
   ========================================================================= */
// Role is a label now. What an account can do comes from its department's
// tier and its own overrides, so the company can call people whatever it
// likes — HOD, Management, and whatever comes next — without a code change.
//
// The one exception is Admin, which means everything. Known names are
// matched regardless of case so "admin" can't quietly produce an account
// that looks like an Admin and isn't one.
var KNOWN_ROLES = ['Admin', 'Management', 'Manager', 'HOD', 'Supervisor'];

function normalizeRole_(raw) {
  var role = String(raw || '').replace(/\s+/g, ' ').trim();
  if (!role) throw new Error('Role is required');
  if (role.length > 40) throw new Error('Role is too long — keep it under 40 characters');
  var known = KNOWN_ROLES.filter(function (r) { return r.toLowerCase() === role.toLowerCase(); })[0];
  return known || role;
}

// An account's tier comes from its department, so without one there is
// nothing to base its access on except the Default row.
function requireDepartmentFor_(role, department) {
  if (role !== 'Admin' && !department) {
    throw new Error('Department is required — it decides which access tier this person starts from');
  }
}

// Changing the last Admin into anything else would leave nobody able to
// open Team & Access, and the only way back would be editing the sheet.
function assertNotLastAdmin_(sheet, username, newRole) {
  if (newRole === 'Admin') return;
  var users = readSheetAsObjects_(sheet);
  var target = users.filter(function (u) { return u.Username === username; })[0];
  if (!target || String(target.Role) !== 'Admin') return;
  var others = users.filter(function (u) {
    return u.Username !== username && String(u.Role) === 'Admin' && isActiveValue_(u.Active);
  });
  if (others.length === 0) {
    throw new Error('This is the last active Admin. Make someone else Admin first, then change this role.');
  }
}

function sanitizeUser_(u) {
  var copy = {};
  Object.keys(u).forEach(function (k) { if (k !== 'PasswordHash') copy[k] = u[k]; });
  copy.HasPassword = !!u.PasswordHash;
  return copy;
}

function getUsers_() {
  var sheet = getSpreadsheet_().getSheetByName('Users');
  var users = readSheetAsObjects_(sheet);
  return users.map(sanitizeUser_);
}

function addUser_(session, params) {
  var username = String(params.Username || '').trim();
  var name = String(params.Name || '').trim();
  var role = normalizeRole_(params.Role);
  var department = String(params.Department || '').trim();
  var email = String(params.Email || '').trim();

  if (!username) throw new Error('Username is required');
  if (!name) throw new Error('Name is required');
  requireDepartmentFor_(role, department);
  if (!email) throw new Error('Email is required — it\u2019s where their password-reset code goes');

  var sheet = getSpreadsheet_().getSheetByName('Users');
  if (findRowIndexByValue_(sheet, 'Username', username) > -1) {
    throw new Error('Username already exists: ' + username);
  }

  var userId = generateId_('USR');
  appendRowByHeader_(sheet, {
    UserID: userId, Name: name, Username: username, PasswordHash: '',
    Role: role, Department: department, Email: email, Phone: params.Phone || '',
    Active: true, CreatedDate: new Date(), LastLogin: ''
  });
  logAudit_(session.Username, 'ADD_USER', 'Users', username, '', JSON.stringify({ name: name, role: role, department: department }));

  // Send them straight to "set your password" rather than emailing a
  // temporary one — one less secret in transit. User is still created
  // even if the email fails; emailSent tells the caller which happened
  // so the UI isn't stuck claiming a mail was sent when it wasn't.
  var emailSent = true;
  try { requestPasswordReset_({ username: username }, session); } catch (err) { emailSent = false; }

  return { username: username, userId: userId, emailSent: emailSent };
}

function updateUser_(session, params) {
  var username = String(params.Username || '').trim();
  if (!username) throw new Error('Username is required');

  var sheet = getSpreadsheet_().getSheetByName('Users');
  var rowIdx = findRowIndexByValue_(sheet, 'Username', username);
  if (rowIdx === -1) throw new Error('Unknown user: ' + username);

  var map = headerIndexMap_(sheet);
  var current = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
  var role = params.Role !== undefined ? normalizeRole_(params.Role) : String(current[map['Role']]);
  var department = params.Department !== undefined ? String(params.Department).trim() : String(current[map['Department']] || '');
  requireDepartmentFor_(role, department);
  assertNotLastAdmin_(sheet, username, role);
  if (params.Role !== undefined) params.Role = role;

  // Username and UserID are immutable identifiers. Only fields that actually
  // change are written — in one guarded range write — and those same fields
  // are what the audit log records, before and after.
  var editable = ['Name', 'Role', 'Department', 'Email', 'Phone'];
  var changes = {}, was = {}, now = {};
  editable.forEach(function (key) {
    if (params[key] === undefined || !(key in map)) return;
    var oldVal = current[map[key]];
    if (String(oldVal === null || oldVal === undefined ? '' : oldVal) === String(params[key])) return;
    changes[key] = params[key]; was[key] = oldVal; now[key] = params[key];
  });
  if (Object.keys(changes).length) writeRowCells_(sheet, rowIdx, stageCells_(map, changes));

  // Role and department are cached on the session row, and the department
  // decides the tier — so both, and the access they resolve to, are pushed
  // to the person's open sessions. Demoting an Admin takes effect on their
  // next action, not whenever their session happens to expire.
  var updated = {};
  var fresh = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
  Object.keys(map).forEach(function (h) { updated[h] = fresh[map[h]]; });
  refreshSessionsFor_(username, resolveAccess_(updated), { Role: updated.Role, Department: updated.Department });

  if (Object.keys(now).length) logAudit_(session.Username, 'UPDATE_USER', 'Users', username, JSON.stringify(was), JSON.stringify(now));
  return { username: username, role: updated.Role };
}

function setUserActive_(session, params, active) {
  var username = String(params.username || params.Username || '').trim();
  if (!username) throw new Error('Username is required');
  if (!active && username === session.Username) throw new Error('You can\u2019t deactivate your own account');

  var sheet = getSpreadsheet_().getSheetByName('Users');
  var rowIdx = findRowIndexByValue_(sheet, 'Username', username);
  if (rowIdx === -1) throw new Error('Unknown user: ' + username);

  var map = headerIndexMap_(sheet);
  var row = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
  var targetRole = row[map['Role']];

  if (!active && targetRole === 'Admin') {
    var allUsers = readSheetAsObjects_(sheet);
    var activeAdmins = allUsers.filter(function (u) { return u.Role === 'Admin' && isActiveValue_(u.Active); });
    if (activeAdmins.length <= 1) throw new Error('Can\u2019t deactivate the last active Admin');
  }

  sheet.getRange(rowIdx, map['Active'] + 1).setValue(active);
  // Deactivating used to leave every open session working for up to
  // SESSION_DURATION_HOURS, in all three apps. Ending them here signs the
  // person out on their next action; ELE and Production Tracker read the
  // same Sessions tab, so it reaches them too.
  var signedOut = active ? 0 : endAllSessionsFor_(username, 'deactivated');
  logAudit_(session.Username, active ? 'REACTIVATE_USER' : 'DEACTIVATE_USER', 'Users', username,
            active ? 'inactive' : 'active', active ? 'active' : 'inactive');
  return { username: username, active: active, sessionsEnded: signedOut };
}

function isActiveValue_(val) {
  if (val === false) return false;
  if (typeof val === 'string' && val.toUpperCase() === 'FALSE') return false;
  return true;
}

// Opposite default from isActiveValue_ on purpose: Active defaults to true
// so a blank cell doesn't silently deactivate someone, but an opt-in flag
// like a blank switch must default to false — otherwise the moment a new
// column is added, every existing user would suddenly be switched on to
// something they never asked for.
function isTruthyFlag_(val) {
  if (val === true) return true;
  if (typeof val === 'string' && val.toUpperCase() === 'TRUE') return true;
  return false;
}

/* =========================================================================
   PASSWORD RESET — 6-digit OTP emailed to the address on file in Users.
   Self-service from the login screen, or Admin-triggered from the Users
   panel (adminResetUserPassword). Either way the code only ever goes to
   the user's own inbox, never back to the caller.
   ========================================================================= */
function getOrCreatePasswordResetsSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('PasswordResets');
  if (!sheet) {
    var headers = ['Username', 'OTP', 'CreatedAt', 'ExpiresAt', 'Attempts'];
    sheet = ss.insertSheet('PasswordResets');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sheet, headers.length);
  }
  return sheet;
}

// From a UUID rather than Math.random(): Utilities.getUuid() is drawn from a
// secure random source, Math.random() is not.
function generateOtp_() {
  var n = parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16);
  return String(100000 + (n % 900000)); // 100000-999999
}

// adminSession present => an Admin explicitly triggered this from the
// Users page, so specific errors (unknown user, no email on file) are
// fine — they already have that visibility. adminSession null => the
// public "forgot password" endpoint, which always returns the same
// generic response so an anonymous caller can't use it to enumerate
// which usernames exist or which ones are missing an email.
function requestPasswordReset_(params, adminSession) {
  var identifier = String(params.username || '').trim();
  if (!identifier) throw new Error('Username is required');

  var usersSheet = getSpreadsheet_().getSheetByName('Users');
  var idLower = identifier.toLowerCase();
  var user = readSheetAsObjects_(usersSheet).filter(function (u) {
    return u.Username === identifier || (u.Email && String(u.Email).toLowerCase() === idLower);
  })[0];
  var genericResult = { requested: true };

  if (!user || !isActiveValue_(user.Active)) {
    if (adminSession) throw new Error('Unknown or inactive user: ' + identifier);
    return genericResult;
  }
  var username = user.Username; // canonical from here on
  if (!user.Email) {
    if (adminSession) throw new Error('No email on file for ' + username + ' — add one before resetting.');
    return genericResult;
  }

  var sheet = getOrCreatePasswordResetsSheet_();
  var rowIdx = findRowIndexByValue_(sheet, 'Username', username);
  var now = new Date();

  if (rowIdx > -1) {
    var existing = readSheetAsObjects_(sheet).filter(function (r) { return r.Username === username; })[0];
    if (existing && (now.getTime() - new Date(existing.CreatedAt).getTime()) < OTP_RESEND_COOLDOWN_SECONDS * 1000) {
      if (adminSession) throw new Error('A code was just sent — wait a bit before requesting another.');
      return genericResult;
    }
    sheet.deleteRow(rowIdx);
  }

  var otp = generateOtp_();
  var expires = new Date(now.getTime() + OTP_VALIDITY_MINUTES * 60 * 1000);
  appendRowByHeader_(sheet, { Username: username, OTP: otp, CreatedAt: now, ExpiresAt: expires, Attempts: 0 });

  sendOtpEmail_(user.Email, user.Name || username, otp);
  logAudit_(adminSession ? adminSession.Username : username, 'REQUEST_PASSWORD_RESET', 'Users', username, '', '');
  return genericResult;
}

function resetPasswordWithOtp_(params) {
  var identifier = String(params.username || '').trim();
  var otp = String(params.otp || '').trim();
  var newPassword = String(params.newPassword || '');
  if (!identifier || !otp) throw new Error('Username and code are required');
  if (newPassword.length < MIN_PASSWORD_LENGTH) throw new Error('New password must be at least ' + MIN_PASSWORD_LENGTH + ' characters');

  var usersSheet = getSpreadsheet_().getSheetByName('Users');
  var idLower = identifier.toLowerCase();
  var user = readSheetAsObjects_(usersSheet).filter(function (u) {
    return u.Username === identifier || (u.Email && String(u.Email).toLowerCase() === idLower);
  })[0];
  // Same message whether the account doesn't exist or just has no pending
  // reset — doesn't tell a guesser which case they hit.
  if (!user) throw new Error('No reset in progress for this account — request a new code');
  var username = user.Username; // canonical from here on

  var sheet = getOrCreatePasswordResetsSheet_();
  var rowIdx = findRowIndexByValue_(sheet, 'Username', username);
  if (rowIdx === -1) throw new Error('No reset in progress for this account — request a new code');

  // Re-check Active here, not just back when the OTP was requested — an
  // Admin could have deactivated this account in the few minutes between
  // the code being emailed and being used. Without this check, a
  // just-deactivated user could still complete the reset, and the
  // Active=true a few lines down would silently reactivate them.
  if (!isActiveValue_(user.Active)) {
    sheet.deleteRow(rowIdx);
    throw new Error('This account was deactivated — contact an Admin before resetting the password');
  }

  var map = headerIndexMap_(sheet);
  var row = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
  var expiresAt = row[map['ExpiresAt']];
  var attempts = Number(row[map['Attempts']]) || 0;

  if (new Date(expiresAt).getTime() < Date.now()) {
    sheet.deleteRow(rowIdx);
    throw new Error('That code expired — request a new one');
  }
  if (attempts >= OTP_MAX_ATTEMPTS) {
    sheet.deleteRow(rowIdx);
    throw new Error('Too many incorrect attempts — request a new code');
  }
  if (String(row[map['OTP']]) !== otp) {
    sheet.getRange(rowIdx, map['Attempts'] + 1).setValue(attempts + 1);
    throw new Error('Incorrect code');
  }

  var userRowIdx = findRowIndexByValue_(usersSheet, 'Username', username);
  // Account could theoretically be removed between requesting and
  // completing a reset — clean up the pending OTP row either way.
  if (userRowIdx === -1) { sheet.deleteRow(rowIdx); throw new Error('Account no longer exists'); }
  ensureColumns_(usersSheet, ['Salt']);
  var newSalt = generateSalt_();
  var userMap = headerIndexMap_(usersSheet);
  usersSheet.getRange(userRowIdx, userMap['PasswordHash'] + 1).setValue(hashPassword_(newPassword, getAppSecret_(), newSalt));
  usersSheet.getRange(userRowIdx, userMap['Salt'] + 1).setValue(newSalt);
  usersSheet.getRange(userRowIdx, userMap['Active'] + 1).setValue(true);

  sheet.deleteRow(rowIdx);
  // Whoever had the old password is signed out everywhere.
  endAllSessionsFor_(username, 'password');
  logAudit_(username, 'PASSWORD_RESET', 'Users', username, '', '');
  return { reset: true };
}

// Self-service password change for the signed-in user. Distinct from the
// admin reset and the OTP flow: it proves ownership with the current
// password instead of an emailed code, so it never touches the
// PasswordResets sheet. Always re-salts, which also quietly upgrades any
// account still on a legacy unsalted hash.
function changeOwnPassword_(session, params) {
  var currentPassword = String(params.currentPassword || '');
  var newPassword = String(params.newPassword || '');
  if (!currentPassword) throw new Error('Enter your current password');
  if (newPassword.length < MIN_PASSWORD_LENGTH) throw new Error('New password must be at least ' + MIN_PASSWORD_LENGTH + ' characters');
  if (newPassword === currentPassword) throw new Error('New password must be different from the current one');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var usersSheet = getSpreadsheet_().getSheetByName('Users');
    var rowIdx = findRowIndexByValue_(usersSheet, 'Username', session.Username);
    if (rowIdx === -1) throw new Error('Account no longer exists');

    ensureColumns_(usersSheet, ['Salt']);
    var map = headerIndexMap_(usersSheet);
    var row = usersSheet.getRange(rowIdx, 1, 1, usersSheet.getLastColumn()).getValues()[0];
    var storedHash = row[map['PasswordHash']];
    var storedSalt = row[map['Salt']];
    if (!storedHash) throw new Error('No password set yet for this account — use "Forgot password" instead');

    var secret = getAppSecret_();
    // Salted first, then the legacy unsalted format, matching login_ so an
    // account that hasn't logged in since the salt upgrade can still change
    // its password rather than being locked out of this screen.
    var ok = storedSalt
      ? hashPassword_(currentPassword, secret, storedSalt) === storedHash
      : hashPassword_(currentPassword, secret, '') === storedHash;
    if (!ok) throw new Error('Current password is incorrect');

    var newSalt = generateSalt_();
    usersSheet.getRange(rowIdx, map['PasswordHash'] + 1).setValue(hashPassword_(newPassword, secret, newSalt));
    usersSheet.getRange(rowIdx, map['Salt'] + 1).setValue(newSalt);

    logAudit_(session.Username, 'PASSWORD_CHANGED', 'Users', session.Username, '', '');
  } finally {
    lock.releaseLock();
  }
  // Other sessions end; this one stays, so the person isn't thrown out of
  // the screen they just used. Outside the lock above, which
  // endAllSessionsFor_ takes for itself.
  var signedOut = endAllSessionsFor_(session.Username, 'password', session.Token);
  return { changed: true, otherSessionsEnded: signedOut };
}

function cleanupExpiredPasswordResets_() {
  var sheet = getOrCreatePasswordResetsSheet_();
  withSessionsLock_(function () {
    var data = sheet.getDataRange().getValues();
    var expIdx = data[0].indexOf('ExpiresAt');
    var now = Date.now();
    for (var i = data.length - 1; i >= 1; i--) {
      if (new Date(data[i][expIdx]).getTime() < now) sheet.deleteRow(i + 1);
    }
  });
}

function sendOtpEmail_(email, name, otp) {
  var subject = COMPANY_NAME + ' \u2014 your password reset code';
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:420px;margin:0 auto">' +
    '<img src="' + COMPANY_LOGO_URL + '" alt="' + COMPANY_NAME + '" style="height:40px;margin-bottom:16px">' +
    '<p>Hi ' + escapeHtml_(name) + ',</p>' +
    '<p>Your verification code for ' + COMPANY_NAME + ' Stock Management is:</p>' +
    '<p style="font-size:28px;font-weight:bold;letter-spacing:4px;margin:16px 0">' + otp + '</p>' +
    '<p style="color:#5B615F;font-size:13px">This code expires in ' + OTP_VALIDITY_MINUTES + ' minutes. ' +
    'If you didn\u2019t request this, you can ignore this email.</p>' +
    '</div>';
  sendMail_({
    to: email,
    subject: subject,
    body: 'Your ' + COMPANY_NAME + ' password reset code is ' + otp + '. It expires in ' + OTP_VALIDITY_MINUTES + ' minutes.',
    htmlBody: html
  });
}

/* =========================================================================
   DAILY STOCK REPORT EMAILS
   -------------------------------------------------------------------------
   Opt-in per person, as the "Daily RM stock mail" and "Daily FG stock mail"
   switches in Team & Access — the same matrix as every other feature, and
   for Admins too, who hold every power but receive mail only if they choose
   to. sendDailyStockMailAll_ runs on a schedule (see
   ADMIN_installDailyStockMailTrigger below); sendTestStockMail_ lets an
   Admin trigger a one-off send on demand to check the format before
   trusting the schedule.
   ========================================================================= */

function statusSortPriority_(status) {
  if (status === 'Critical') return 0;
  if (status === 'Low') return 1;
  return 2;
}

// Current Stock doesn't carry an Active column (only Master does), so this
// cross-references the two, then sorts Critical/Low to the top — the
// items someone actually needs to see first.
function activeStockRows_(config) {
  var activeCodes = {};
  getMaster_(config).forEach(function (m) { if (isActiveValue_(m.Active)) activeCodes[m[config.codeField]] = true; });
  return getCurrentStock_(config)
    .filter(function (r) { return activeCodes[r[config.codeField]]; })
    .sort(function (a, b) { return statusSortPriority_(a.StockStatus) - statusSortPriority_(b.StockStatus); });
}

// Shared styling engine — every Excel output this app produces (the daily
// mail attachment and, from here on, the app's own Export buttons) goes
// through this one function, so they all look like the same product
// instead of two different styling systems drifting apart over time.
//
// opts:
//   title, subtitle   — merged banner rows across the top (subtitle optional)
//   columns           — [{ key, label, width, numFmt, align }]
//   rows              — plain objects keyed by column.key
//   statusColumnKey   — optional; if set, cells in that column get a
//                       colored background/text based on their value
//                       (Critical/Low/OK — same palette the web app uses)
function writeStyledSheet_(sheet, opts) {
  var columns = opts.columns;
  var numCols = columns.length;
  var BRAND_NAVY = '#0F172A', BRAND_BLUE = '#1D4ED8', BORDER = '#CBD5E1';
  var STATUS_COLORS = {
    Critical: { bg: '#FEF2F2', fg: '#EF4444' },
    Low: { bg: '#FFFBEB', fg: '#F59E0B' },
    OK: { bg: '#ECFDF5', fg: '#10B981' }
  };

  var rowCursor = 1;

  sheet.getRange(rowCursor, 1, 1, numCols).merge()
    .setValue(opts.title).setFontSize(15).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground(BRAND_NAVY).setVerticalAlignment('middle');
  sheet.setRowHeight(rowCursor, 34);
  rowCursor++;

  if (opts.subtitle) {
    sheet.getRange(rowCursor, 1, 1, numCols).merge()
      .setValue(opts.subtitle).setFontSize(10).setFontColor('#94A3B8')
      .setBackground(BRAND_NAVY).setVerticalAlignment('middle');
    sheet.setRowHeight(rowCursor, 22);
    rowCursor++;
  }

  var headerRow = rowCursor;
  sheet.getRange(headerRow, 1, 1, numCols).setValues([columns.map(function (c) { return c.label; })])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(BRAND_BLUE).setVerticalAlignment('middle');
  sheet.setRowHeight(headerRow, 26);
  rowCursor++;

  var dataStartRow = rowCursor;
  if (opts.rows.length === 0) {
    sheet.getRange(rowCursor, 1, 1, numCols).merge()
      .setValue('No data.').setFontColor('#94A3B8').setFontStyle('italic');
    rowCursor++;
  } else {
    var values = opts.rows.map(function (r) { return columns.map(function (c) { return r[c.key] !== undefined ? r[c.key] : ''; }); });
    sheet.getRange(dataStartRow, 1, values.length, numCols).setValues(values);

    columns.forEach(function (c, idx) {
      var colRange = sheet.getRange(dataStartRow, idx + 1, values.length, 1);
      if (c.numFmt === '#,##0') colRange.setNumberFormats(quantityFormats_(values, idx));
      else if (c.numFmt) colRange.setNumberFormat(c.numFmt);
      if (c.align) colRange.setHorizontalAlignment(c.align);
    });

    // Alternating row shading — a real Sheets banding, so it survives the
    // xlsx export as native Excel banded-rows formatting, not a one-off
    // per-cell hack.
    sheet.getRange(dataStartRow, 1, values.length, numCols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

    if (opts.statusColumnKey) {
      var statusColIdx = columns.findIndex(function (c) { return c.key === opts.statusColumnKey; });
      if (statusColIdx > -1) {
        opts.rows.forEach(function (r, i) {
          var palette = STATUS_COLORS[r[opts.statusColumnKey]];
          if (palette) sheet.getRange(dataStartRow + i, statusColIdx + 1).setBackground(palette.bg).setFontColor(palette.fg).setFontWeight('bold');
        });
      }
    }
    rowCursor += values.length;
  }

  var tableRange = sheet.getRange(headerRow, 1, rowCursor - headerRow, numCols);
  tableRange.setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);

  columns.forEach(function (c, idx) { if (c.width) sheet.setColumnWidth(idx + 1, c.width); });

  sheet.setFrozenRows(headerRow);
  if (opts.rows.length > 0) tableRange.createFilter();
}

// Column definitions shared by the daily mail report and the app's own
// stock export, so "what a stock report looks like" is defined once.
function stockReportColumns_(config) {
  var catLabel = config.key === 'rm' ? 'Category' : 'Family';
  var inLabel = config.key === 'rm' ? 'Received' : 'Produced';
  var outLabel = config.key === 'rm' ? 'Issued' : 'Dispatched';
  return [
    { key: config.codeField, label: 'Code', width: 110 },
    { key: config.nameField, label: 'Name', width: 220 },
    { key: config.categoryField, label: catLabel, width: 120 },
    { key: config.unitField, label: 'Unit', width: 65, align: 'center' },
    { key: 'OpeningStock', label: 'Opening', width: 85, numFmt: '#,##0', align: 'right' },
    { key: config.receivedField, label: inLabel, width: 85, numFmt: '#,##0', align: 'right' },
    { key: config.issuedField, label: outLabel, width: 85, numFmt: '#,##0', align: 'right' },
    { key: 'CurrentBalance', label: 'Balance', width: 90, numFmt: '#,##0', align: 'right' },
    { key: config.reorderField, label: 'Reorder Level', width: 100, numFmt: '#,##0', align: 'right' },
    { key: 'StockStatus', label: 'Status', width: 90, align: 'center' },
    { key: 'LastUpdated', label: 'Last Updated', width: 105, numFmt: 'yyyy-mm-dd' }
  ];
}

// Builds one temporary spreadsheet with one tab per entry in `sheets`,
// exports it as a real .xlsx Blob, then deletes the temporary file —
// nothing lingers in Drive. This is the one part of the feature that
// needs Drive access, purely for this create-export-delete round trip.
// sheets: [{ name, title, subtitle, columns, rows, statusColumnKey }]
function buildStyledWorkbookBlob_(sheets, filenameBase) {
  var ss = SpreadsheetApp.create('tmp-calgas-export-' + new Date().getTime());
  var ssId = ss.getId();
  try {
    var firstSheet = ss.getSheets()[0];
    sheets.forEach(function (s, idx) {
      var sheet = idx === 0 ? firstSheet.setName(s.name) : ss.insertSheet(s.name);
      writeStyledSheet_(sheet, s);
    });

    SpreadsheetApp.flush(); // make sure the writes above are actually committed before we export

    // DriveApp.getFileById(...).getAs(MimeType.MICROSOFT_EXCEL) doesn't
    // reliably convert a native Google Sheet to .xlsx on the V8 runtime —
    // it throws "Converting from application/vnd.google-apps.spreadsheet
    // ... is not supported". Fetching Sheets' own export endpoint is the
    // standard workaround; it needs no setup beyond what's already
    // authorized (Drive access is already granted from the trash step
    // below — this just adds one more scope, for UrlFetchApp itself, on
    // the next authorization prompt).
    var exportUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';
    var response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Could not export the stock report as .xlsx (HTTP ' + response.getResponseCode() + ')');
    }
    var blob = response.getBlob();
    blob.setName((filenameBase || 'Export') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.xlsx');
    return blob;
  } finally {
    DriveApp.getFileById(ssId).setTrashed(true);
  }
}

// Entry point for the app's own Export buttons (RM/FG stock list, item
// history). The client already knows which rows to export (it's already
// filtered/fetched them) and just needs them turned into a properly
// styled file — this is a pure formatting service, not a data source, so
// it doesn't re-derive anything from the sheets itself. Row/sheet caps
// here are a cost guard (this is reachable by any authenticated session,
// not just Admins), not a security boundary — the output is a throwaway
// temp file either way, never written back to real data.
function exportStyledWorkbook_(params) {
  var sheets = params.sheets;
  if (!Array.isArray(sheets) || sheets.length === 0) throw new Error('Nothing to export');
  if (sheets.length > 5) throw new Error('Too many sheets in one export');
  sheets.forEach(function (s) {
    if (!Array.isArray(s.columns) || s.columns.length === 0 || !Array.isArray(s.rows)) throw new Error('Invalid export payload');
    if (s.rows.length > 5000) throw new Error('That\u2019s ' + s.rows.length + ' rows \u2014 too many for one export. Narrow it down with a filter first.');
  });
  var blob = buildStyledWorkbookBlob_(sheets, params.filename || 'CALGAS_Export');
  return { filename: blob.getName(), base64: Utilities.base64Encode(blob.getBytes()) };
}

// "N critical, M low" style line for one report, or "all clear".
function alertSummaryLine_(rows, label) {
  var critical = rows.filter(function (r) { return r.StockStatus === 'Critical'; }).length;
  var low = rows.filter(function (r) { return r.StockStatus === 'Low'; }).length;
  if (!critical && !low) return label + ': all clear';
  var parts = [];
  if (critical) parts.push(critical + ' critical');
  if (low) parts.push(low + ' low');
  return label + ': ' + parts.join(', ');
}

function stockReportSheetDef_(config, rows) {
  var dateLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
  return {
    name: config.key === 'rm' ? 'RM Stock' : 'FG Stock',
    title: COMPANY_NAME + ' \u2014 ' + (config.key === 'rm' ? 'RM' : 'FG') + ' Current Stock',
    subtitle: dateLabel + ' \u00b7 ' + rows.length + ' active item' + (rows.length === 1 ? '' : 's'),
    columns: stockReportColumns_(config),
    rows: rows,
    statusColumnKey: 'StockStatus'
  };
}

// rmRows/fgRows are null for whichever report this user didn't opt into.
// blob is passed in (rather than built here) so sendDailyStockMailAll_
// can build each distinct RM/FG/both combination once and reuse it across
// every subscriber who wants that same combination.
function sendStockReportToUser_(user, rmRows, fgRows, blob) {
  var dateLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
  var both = !!(rmRows && fgRows);
  var reportLabel = both ? 'RM & FG Stock Report' : (rmRows ? 'RM Stock Report' : 'FG Stock Report');
  var subject = COMPANY_NAME + ' \u2014 Daily ' + reportLabel + ' \u2014 ' + dateLabel;

  var alerts = [];
  if (rmRows) alerts.push(alertSummaryLine_(rmRows, 'RM'));
  if (fgRows) alerts.push(alertSummaryLine_(fgRows, 'FG'));

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto">' +
    '<img src="' + COMPANY_LOGO_URL + '" alt="' + COMPANY_NAME + '" style="height:40px;margin-bottom:16px">' +
    '<p>Hi ' + escapeHtml_(user.Name || user.Username) + ',</p>' +
    '<p>Here\u2019s the daily stock report for ' + dateLabel + '. Full detail is in the attached workbook' + (both ? ' (RM and FG on separate tabs).' : '.') + '</p>' +
    '<p style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:8px;padding:12px 16px;font-size:13.5px;color:#334155">' + alerts.join(' &nbsp;&middot;&nbsp; ') + '</p>' +
    '<p style="color:#5B615F;font-size:12.5px;margin-top:24px">This is an automated report. Manage your subscription under Team &amp; Access in the app.</p>' +
    '</div>';

  sendMail_({
    to: user.Email,
    subject: subject,
    body: reportLabel + ' for ' + dateLabel + ' is attached.',
    htmlBody: html,
    attachments: [blob]
  });
}

// Runs daily via the trigger installed by ADMIN_installDailyStockMailTrigger.
// One combined email per qualifying user. RM/FG rows and the resulting
// workbook are each built at most once per run (not once per recipient) —
// there are only three possible combinations (RM-only / FG-only / both),
// so this caches by that combination rather than re-reading Master +
// Current Stock and re-exporting a Drive file per subscriber.
function sendDailyStockMailAll_() {
  var holiday = holidayToday_();
  if (holiday) {
    Logger.log('Daily stock mail skipped — holiday: ' + (holiday.reason || holiday.date));
    return { sent: 0, failed: 0, skipped: 'holiday', reason: holiday.reason };
  }
  var usersSheet = getSpreadsheet_().getSheetByName('Users');
  var templates = readAccessTemplates_();
  var users = readSheetAsObjects_(usersSheet).filter(function (u) {
    if (!isActiveValue_(u.Active) || !u.Email) return false;
    var access = resolveAccess_(u, templates);
    u.__wantRM = !!access['stock.mail.rm'];
    u.__wantFG = !!access['stock.mail.fg'];
    return u.__wantRM || u.__wantFG;
  });

  if (users.length === 0) { Logger.log('Daily stock mail: no subscribed users.'); return { sent: 0, failed: 0 }; }

  var needsRM = users.some(function (u) { return u.__wantRM; });
  var needsFG = users.some(function (u) { return u.__wantFG; });
  var rmRows = needsRM ? activeStockRows_(RM_CONFIG) : null;
  var fgRows = needsFG ? activeStockRows_(FG_CONFIG) : null;

  var blobCache = {};
  function blobFor(wantRM, wantFG) {
    var key = (wantRM ? 'R' : '') + (wantFG ? 'F' : '');
    if (!blobCache[key]) {
      var sheets = [];
      if (wantRM) sheets.push(stockReportSheetDef_(RM_CONFIG, rmRows));
      if (wantFG) sheets.push(stockReportSheetDef_(FG_CONFIG, fgRows));
      blobCache[key] = buildStyledWorkbookBlob_(sheets, 'CALGAS_Stock_Report');
    }
    return blobCache[key];
  }

  var sent = 0, failed = 0;
  users.forEach(function (u) {
    var wantRM = u.__wantRM, wantFG = u.__wantFG;
    try {
      sendStockReportToUser_(u, wantRM ? rmRows : null, wantFG ? fgRows : null, blobFor(wantRM, wantFG));
      sent++;
    } catch (err) {
      failed++;
      Logger.log('Daily stock mail failed for ' + u.Username + ': ' + err.message);
    }
  });

  Logger.log('Daily stock mail: ' + sent + ' sent, ' + failed + ' failed.');
  return { sent: sent, failed: failed };
}

// Admin-triggered manual send — to a specified user, or to the calling
// Admin themselves if no target is given. Sends whatever that user
// currently has checked; doesn't override their subscription.
function sendTestStockMail_(session, params) {
  var targetUsername = params.username ? String(params.username).trim() : session.Username;
  var usersSheet = getSpreadsheet_().getSheetByName('Users');
  var user = readSheetAsObjects_(usersSheet).filter(function (u) { return u.Username === targetUsername; })[0];
  if (!user) throw new Error('Unknown user: ' + targetUsername);
  if (!user.Email) throw new Error(targetUsername + ' has no email on file');

  var access = resolveAccess_(user);
  var wantRM = !!access['stock.mail.rm'];
  var wantFG = !!access['stock.mail.fg'];
  if (!wantRM && !wantFG) throw new Error(targetUsername + ' doesn\u2019t have RM or FG daily mail enabled yet');

  var rmRows = wantRM ? activeStockRows_(RM_CONFIG) : null;
  var fgRows = wantFG ? activeStockRows_(FG_CONFIG) : null;
  var sheets = [];
  if (wantRM) sheets.push(stockReportSheetDef_(RM_CONFIG, rmRows));
  if (wantFG) sheets.push(stockReportSheetDef_(FG_CONFIG, fgRows));
  var blob = buildStyledWorkbookBlob_(sheets, 'CALGAS_Stock_Report');
  sendStockReportToUser_(user, rmRows, fgRows, blob);
  return { sentTo: user.Email };
}

// Run once from the Apps Script editor. This is the first feature that
// touches Drive (buildStyledWorkbookBlob_'s create/export/delete round
// trip), so running this will prompt a fresh authorization the first time
// — that's expected, not an error. Same install pattern as
// ADMIN_installNightlyCleanupTrigger.
function ADMIN_installDailyStockMailTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyStockMailAll_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyStockMailAll_').timeBased().everyDays(1).atHour(6).nearMinute(30).create();
  Logger.log('Daily stock mail trigger installed (runs ~6:30am script timezone).');
}

/* =========================================================================
   SHARED STOCK LOGIC (RM and FG both run through this, see configs below)
   ========================================================================= */
function getMaster_(config) {
  return readSheetAsObjects_(getSpreadsheet_().getSheetByName(config.masterSheet));
}

// Every one of getRMMaster/getFGMaster/getRMCurrentStock/getFGCurrentStock/
// getDepartments/getUsers is its own Web App exec — each pays its own
// validateSession_() (a Sessions-sheet read) and script-boot cost even
// though they all ran in parallel from the client. Bundling them into one
// call cuts that overhead from 5-6x down to 1x on every login/refresh.
// Role gating for `users` happens here (server-side, off the validated
// session) rather than trusting a client-sent flag.
function getBootstrap_(session) {
  var data = {
    departments: getDepartments_(),
    // Resolved on every load rather than only at sign-in, so the buttons a
    // person sees match what the backend will allow right now — access can
    // change while they are signed in.
    permissions: sessionPermissions_(session),
    dropdowns: getDropdowns_(),
    rmMaster: getMaster_(RM_CONFIG),
    fgMaster: getMaster_(FG_CONFIG),
    rmStock: getCurrentStock_(RM_CONFIG),
    fgStock: getCurrentStock_(FG_CONFIG)
  };
  if (session.Role === 'Admin') data.users = getUsers_();
  return data;
}

// The Dropdowns tab is column-oriented, not row-oriented: each header
// (RM_Category, FG_Category, FG_Family, Unit, Supplier, OLCustomer) owns
// an independent vertical list, and the columns are different lengths, so
// readSheetAsObjects_ (which assumes every row is one record) is the wrong
// tool here — row 5 of Supplier has nothing to do with row 5 of Unit.
// This reads the block raw and transposes it into { header: [values] },
// dropping blanks and de-duplicating while preserving the sheet's own
// ordering (the person maintaining the tab controls the order they see).
// Returns {} rather than throwing if the tab is missing, so an older
// deployment without the tab keeps working on its hardcoded fallbacks.
function getDropdowns_() {
  var sheet = getSpreadsheet_().getSheetByName('Dropdowns');
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return {};

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0];
  var out = {};

  for (var c = 0; c < lastCol; c++) {
    var header = String(headers[c] || '').trim();
    if (!header) continue;
    var seen = {};
    var list = [];
    for (var r = 1; r < lastRow; r++) {
      var v = values[r][c];
      if (v === '' || v === null || v === undefined) continue;
      v = String(v).trim();
      if (!v || seen[v]) continue;
      seen[v] = true;
      list.push(v);
    }
    out[header] = list;
  }
  return out;
}

function lookupMasterRow_(config, code) {
  var sheet = getSpreadsheet_().getSheetByName(config.masterSheet);
  var row = readRowByValue_(sheet, config.codeField, code);
  if (!row) return null;
  var obj = {};
  Object.keys(row.map).forEach(function (header) { obj[header] = row.values[row.map[header]]; });
  return obj;
}

function getCurrentStock_(config) {
  return readSheetAsObjects_(getSpreadsheet_().getSheetByName(config.stockSheet));
}

// Newest first, optionally filtered to one code, capped at `limit` (default
// 200) so a full-history request can't blow up the response payload.
// Rows are appended chronologically (see appendRowByHeader_/addTransaction_),
// so any one item's transactions are scattered throughout the whole sheet
// mixed in with every other item's — there's no contiguous block to read.
// The old version dealt with that by reading + object-converting the
// ENTIRE ledger on every single history lookup, then filtering down to one
// item's rows. That cost scales with the total size of the ledger (every
// item, ever), not with how many entries the item being viewed actually
// has — so it gets slower for everyone's history lookups as the ledger
// grows, even for an item with only a handful of transactions.
//
// This scans backward from the newest rows in chunks and stops as soon as
// it has `limit` matches. Since new transactions are appended at the
// bottom, an item's most recent activity — which is what the History
// panel actually shows — is almost always found within the first chunk
// or two, so the common case reads a small slice of the sheet instead of
// all of it. Worst case (an old/inactive item, or an unmatched code) still
// scans the whole sheet, exactly like before — never slower, often much
// faster.
//
// params.dateFrom / dateTo (yyyy-mm-dd) and params.enteredBy (substring,
// case-insensitive) are optional filters applied during the same scan —
// they don't cost anything beyond the scan itself. A dateFrom in
// particular lets the scan stop early: once a row is older than dateFrom,
// every row further back is even older (append-only ordering), so there's
// nothing left to find.
/* =========================================================================
   LAST PRICE PAID

   The rate on the most recent receipt that carried one — what was last
   actually paid. The entry form fills its Rate box from the item's
   StandardCost in the master (already loaded, so no call is needed) and uses
   this only as context beside it, or in its place for an item with no
   StandardCost set. If no receipt ever had a rate, the most recent entry of
   any type with one is used instead, and the answer says which it was, so
   the form can say where the figure came from rather than presenting a
   guess as a fact.

   Read from the ledger each time rather than cached on the stock row, so a
   corrected or deleted receipt can never leave a stale price behind.
   ========================================================================= */

function getItemPrice_(params) {
  var config = configForKey_(params.key);
  var code = String(params.code || '').trim();
  if (!code) throw new Error('Item code is required');

  var sheet = getSpreadsheet_().getSheetByName(config.ledgerSheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { code: code, rate: null };

  var meta = sheetMeta_(sheet);
  var col = function (h) { return meta.map[h]; };
  var codeIdx = col(config.codeField), rateIdx = col('Rate'), typeIdx = col('TxnType');
  if (codeIdx === undefined || rateIdx === undefined) return { code: code, rate: null };

  var fallback = null;
  var CHUNK = 500;
  var cursor = lastRow;
  while (cursor >= 2) {
    var start = Math.max(2, cursor - CHUNK + 1);
    var rows = sheet.getRange(start, 1, cursor - start + 1, meta.lastCol).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      var r = rows[i];
      if (String(r[codeIdx]) !== code) continue;
      var rate = Number(r[rateIdx]);
      if (r[rateIdx] === '' || r[rateIdx] === null || isNaN(rate) || rate <= 0) continue;

      var found = {
        code: code,
        rate: rate,
        txnType: r[typeIdx],
        date: r[col('Date')] || r[col('Timestamp')] || '',
        supplier: col('Supplier') !== undefined ? r[col('Supplier')] : '',
        invoiceNo: col('InvoiceNo') !== undefined ? r[col('InvoiceNo')] : ''
      };
      if (config.inTypes.indexOf(String(r[typeIdx])) > -1) {
        found.source = 'receipt';
        return found;
      }
      if (!fallback) { found.source = 'other'; fallback = found; }
    }
    cursor = start - 1;
  }
  return fallback || { code: code, rate: null };
}

// Quantities and balances: a whole number stays whole (1,234), one with
// decimals shows them, up to four (0.1364, 64.0281). A single whole-number
// format for the column rounded 0.1364 kg of solder wire to "0" in exports.
function quantityFormats_(values, idx) {
  return values.map(function (row) {
    var v = row[idx];
    var fractional = typeof v === 'number' && Math.abs(v - Math.round(v)) > 1e-9;
    return [fractional ? '#,##0.0###' : '#,##0'];
  });
}

var LEDGER_MAX_LIMIT = 5000;

function getLedger_(config, params) {
  var sheet = getSpreadsheet_().getSheetByName(config.ledgerSheet);
  var code = params.code ? String(params.code).trim() : '';
  // Capped: an uncapped limit let one request read the whole ledger and run
  // into Apps Script's time limit. 5,000 is the item-history export's size.
  var limit = Math.min(Number(params.limit) || 200, LEDGER_MAX_LIMIT);
  var dateFrom = params.dateFrom ? new Date(params.dateFrom + 'T00:00:00') : null;
  var dateTo = params.dateTo ? new Date(params.dateTo + 'T23:59:59') : null;
  var userFilter = params.enteredBy ? String(params.enteredBy).trim().toLowerCase() : '';
  // Filtering these server-side matters: `limit` caps rows *returned*, so
  // trimming by type/department on the client would silently drop matches
  // beyond the cap. Applied here, the limit counts only rows that match.
  var typeFilter = params.txnType ? String(params.txnType).trim() : '';
  var deptFilter = params.department ? String(params.department).trim().toLowerCase() : '';
  // Older pages: the screens ask for the newest rows first, then for the
  // next ones by skipping the matches they already have.
  var skip = Math.max(0, Math.floor(Number(params.skip) || 0));
  var skipped = 0;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var codeColIdx = headers.indexOf(config.codeField);
  var tsColIdx = headers.indexOf('Timestamp');
  var userColIdx = headers.indexOf('EnteredBy');
  var typeColIdx = headers.indexOf('TxnType');
  var deptColIdx = headers.indexOf('Department');

  var CHUNK_SIZE = 500;
  var matches = [];
  var rowCursor = lastRow; // 1-based, inclusive — walks backward toward row 2
  var pastRangeStart = false;

  while (rowCursor >= 2 && matches.length < limit && !pastRangeStart) {
    var chunkStart = Math.max(2, rowCursor - CHUNK_SIZE + 1);
    var values = sheet.getRange(chunkStart, 1, rowCursor - chunkStart + 1, lastCol).getValues();

    for (var i = values.length - 1; i >= 0; i--) {
      var row = values[i];
      if (code && codeColIdx !== -1 && row[codeColIdx] !== code) continue;

      var ts = null;
      if (tsColIdx !== -1) { var d = new Date(row[tsColIdx]); if (!isNaN(d.getTime())) ts = d; }

      if (dateFrom && ts && ts < dateFrom) { pastRangeStart = true; break; }
      if (dateTo && ts && ts > dateTo) continue;
      if (userFilter && userColIdx !== -1 && String(row[userColIdx] || '').toLowerCase().indexOf(userFilter) === -1) continue;
      if (typeFilter && typeColIdx !== -1 && String(row[typeColIdx] || '') !== typeFilter) continue;
      if (deptFilter && deptColIdx !== -1 && String(row[deptColIdx] || '').toLowerCase().indexOf(deptFilter) === -1) continue;
      if (skipped < skip) { skipped++; continue; }

      var obj = {};
      for (var c = 0; c < headers.length; c++) { if (headers[c]) obj[headers[c]] = row[c]; }
      matches.push(obj);
      if (matches.length >= limit) break;
    }

    rowCursor = chunkStart - 1;
  }

  // Belt-and-suspenders: appends should already be in timestamp order, but
  // re-sort explicitly so this is correct even after a manual sheet edit.
  matches.sort(function (a, b) { return new Date(b.Timestamp || 0).getTime() - new Date(a.Timestamp || 0).getTime(); });

  // How many files each entry has, so a row can show its paperclip without
  // a request per row. The tab is read once, whatever the number of rows.
  var counts = attachmentCounts_(config.key);
  matches.forEach(function (m) { m.AttachmentCount = counts[m.TxnID] || 0; });

  // Asked for by an item's history, which checks the two agree. They must be
  // read in the same request: a balance from one moment and a history from
  // another disagree whenever someone posts in between, and that looks
  // exactly like corruption when it's nothing of the kind.
  if (params.withStock && code) {
    var stockRow = readRowByValue_(getSpreadsheet_().getSheetByName(config.stockSheet), config.codeField, code);
    var stock = null;
    if (stockRow) {
      stock = {};
      ['OpeningStock', 'CurrentBalance', config.receivedField, config.issuedField].forEach(function (h) {
        if (h in stockRow.map) stock[h] = stockRow.values[stockRow.map[h]];
      });
    }
    return { rows: matches, stock: stock };
  }
  return matches;
}

/* =========================================================================
   ATTACHMENTS

   Invoices, delivery notes, photos — any number of images or PDFs on any
   entry, never required. The files live in a Drive folder of the account
   that owns this script, one subfolder per month, and are never shared by
   link: the app fetches them through its own sign-in, so floor staff need
   no Google account and an invoice can't leak through a forwarded URL.

   The Attachments tab says which file belongs to which entry. A removed
   file is marked, not destroyed — like a deleted entry, it stays on record.
   Production Tracker adds rows here too: when a handover is received and
   posts an entry, the handover's files are recorded on that entry.
   ========================================================================= */

var ATTACH_MAX_FILES = 10;
var ATTACH_MAX_BYTES = 10 * 1024 * 1024;
var ATTACH_TYPES = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/gif': 'gif', 'application/pdf': 'pdf' };
var ATTACH_FOLDER_NAME = 'CALGAS Stock Attachments';
var ATTACH_HEADERS = ['AttachmentID', 'Ledger', 'TxnID', 'FileID', 'FileName', 'MimeType', 'Size', 'Thumb',
                      'UploadedBy', 'UploadedAt', 'Status', 'RemovedBy', 'RemovedAt', 'RemoveReason', 'ClientTxnId', 'Source'];

function getOrCreateAttachmentsSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Attachments');
  if (!sheet) {
    sheet = ss.insertSheet('Attachments');
    sheet.getRange(1, 1, 1, ATTACH_HEADERS.length).setValues([ATTACH_HEADERS]);
    styleHeader_(sheet, ATTACH_HEADERS.length);
  } else {
    // Source arrived with Production Tracker's files; a tab made before it
    // gains the column here.
    ensureColumns_(sheet, ATTACH_HEADERS);
  }
  return sheet;
}

function ledgerConfigFor_(ledger) {
  var k = String(ledger || '').toLowerCase();
  if (k === 'rm') return RM_CONFIG;
  if (k === 'fg') return FG_CONFIG;
  throw new Error('ledger must be RM or FG');
}

function findLedgerEntry_(config, txnId) {
  var row = readRowByValue_(getSpreadsheet_().getSheetByName(config.ledgerSheet), 'TxnID', String(txnId || '').trim());
  if (!row) return null;
  var out = {};
  Object.keys(row.map).forEach(function (h) { out[h] = row.values[row.map[h]]; });
  return out;
}

// Adding a file to an entry is allowed to whoever could have posted it:
// the ledger's post permission, for that entry's department.
function requireCanAttach_(session, config, entry) {
  requirePerm_(session, config.key === 'rm' ? 'stock.rm.post' : 'stock.fg.post');
  if (!can_(session, 'stock.post.any') && String(entry.Department) !== String(session.Department)) {
    throw new Error('You can only add files to your own department\u2019s entries (' + (session.Department || 'none set') +
                    '). Ask an Admin for wider access.');
  }
}

function readAttachments_() {
  var sheet = getSpreadsheet_().getSheetByName('Attachments');
  return sheet ? readSheetAsObjects_(sheet) : [];
}

function activeAttachments_(ledgerKey, txnId) {
  var L = String(ledgerKey).toUpperCase();
  return readAttachments_().filter(function (a) {
    return String(a.Ledger).toUpperCase() === L && String(a.TxnID) === String(txnId) && a.Status === 'Active';
  });
}

function attachmentCounts_(ledgerKey) {
  var L = String(ledgerKey).toUpperCase();
  var counts = {};
  readAttachments_().forEach(function (a) {
    if (a.Status === 'Active' && String(a.Ledger).toUpperCase() === L) counts[a.TxnID] = (counts[a.TxnID] || 0) + 1;
  });
  return counts;
}

function publicAttachment_(a) {
  return {
    attachmentId: a.AttachmentID, ledger: a.Ledger, txnId: a.TxnID, fileName: a.FileName, mimeType: a.MimeType,
    size: Number(a.Size) || 0, thumb: a.Thumb || '', uploadedBy: a.UploadedBy, uploadedAt: a.UploadedAt,
    // Where a file came from when it wasn't added here — "Production
    // Tracker HO-00012" for the paperwork of the handover whose receipt
    // posted this entry.
    source: a.Source || ''
  };
}

// A name that's safe as a Drive file name and in a cell, keeping the
// extension the file really has.
function cleanFileName_(name, mime) {
  var base = String(name || '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 120);
  var ext = ATTACH_TYPES[mime];
  if (!base) base = 'file';
  if (!new RegExp('\\.' + ext + '$', 'i').test(base) && !(ext === 'jpg' && /\.jpe?g$/i.test(base))) {
    base = base.replace(/\.[A-Za-z0-9]{1,5}$/, '') + '.' + ext;
  }
  return base;
}

function formatBytes_(n) {
  return n >= 1048576 ? (Math.round(n / 104857.6) / 10) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
}

// The month's folder under the one attachments folder, made the first time
// it's needed. The folder's ID is kept in Settings, so renaming or moving
// it in Drive is fine.
function attachmentsFolder_(when) {
  var settings = getSpreadsheet_().getSheetByName('Settings');
  var saved = readSheetAsObjects_(settings).filter(function (r) { return r.Key === 'ATTACHMENTS_FOLDER_ID'; })[0];
  var root = null;
  if (saved && saved.Value) {
    try { root = DriveApp.getFolderById(String(saved.Value)); } catch (err) { root = null; }
  }
  if (!root) {
    root = DriveApp.createFolder(ATTACH_FOLDER_NAME);
    setSettingValue_(settings, 'ATTACHMENTS_FOLDER_ID', root.getId());
  }
  var month = Utilities.formatDate(when, getSpreadsheet_().getSpreadsheetTimeZone(), 'yyyy-MM-dd').slice(0, 7);
  var existing = root.getFoldersByName(month);
  return existing.hasNext() ? existing.next() : root.createFolder(month);
}

function findAttachmentByClientTxnId_(clientTxnId) {
  return readAttachments_().filter(function (a) { return String(a.ClientTxnId) === clientTxnId; })[0] || null;
}

// One file per request, so a phone on a weak signal sends one photo at a
// time and a failure costs only that photo. The checks run under the lock,
// the slow Drive write outside it — other people's entries don't wait on a
// large PDF — then the record is made under the lock again, re-checking in
// case a retry of this same upload got there first.
function uploadAttachment_(session, params) {
  var config = ledgerConfigFor_(params.ledger);
  var txnId = String(params.txnId || '').trim();
  var clientTxnId = String(params.clientTxnId || '').trim();
  var mime = String(params.mimeType || '').toLowerCase();
  if (!ATTACH_TYPES[mime]) throw new Error('Only photos (JPEG, PNG, WebP, GIF) and PDFs can be attached.');
  var data = String(params.data || '');
  if (!data) throw new Error('The file is empty.');
  if (Math.floor(data.length * 3 / 4) > ATTACH_MAX_BYTES * 1.01) {
    throw new Error('That file is larger than ' + formatBytes_(ATTACH_MAX_BYTES) + ', the most one file can be.');
  }
  var name = cleanFileName_(params.fileName, mime);
  var thumb = String(params.thumb || '');
  if (thumb.length > 45000 || !/^[A-Za-z0-9+\/=]*$/.test(thumb)) thumb = '';

  var sheet, folder;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sheet = getOrCreateAttachmentsSheet_();
    if (clientTxnId) {
      var prior = findAttachmentByClientTxnId_(clientTxnId);
      if (prior) return publicAttachment_(prior);
    }
    var entry = findLedgerEntry_(config, txnId);
    if (!entry) throw new Error('Entry ' + txnId + ' wasn\u2019t found \u2014 it may have been deleted.');
    requireCanAttach_(session, config, entry);
    if (activeAttachments_(config.key, txnId).length >= ATTACH_MAX_FILES) {
      throw new Error('This entry already has ' + ATTACH_MAX_FILES + ' files, the most one entry can hold.');
    }
    folder = attachmentsFolder_(new Date());
  } finally {
    lock.releaseLock();
  }

  var bytes = Utilities.base64Decode(data);
  if (bytes.length > ATTACH_MAX_BYTES) {
    throw new Error('That file is larger than ' + formatBytes_(ATTACH_MAX_BYTES) + ', the most one file can be.');
  }
  var file = folder.createFile(Utilities.newBlob(bytes, mime, txnId + ' - ' + name));

  var rec;
  lock.waitLock(30000);
  try {
    if (clientTxnId) {
      var raced = findAttachmentByClientTxnId_(clientTxnId);
      if (raced) {
        try { file.setTrashed(true); } catch (err) {}
        return publicAttachment_(raced);
      }
    }
    if (activeAttachments_(config.key, txnId).length >= ATTACH_MAX_FILES) {
      try { file.setTrashed(true); } catch (err) {}
      throw new Error('This entry already has ' + ATTACH_MAX_FILES + ' files, the most one entry can hold.');
    }
    rec = {
      AttachmentID: generateId_('ATT'), Ledger: config.key.toUpperCase(), TxnID: txnId, FileID: file.getId(),
      FileName: name, MimeType: mime, Size: bytes.length, Thumb: thumb, UploadedBy: session.Username,
      UploadedAt: new Date(), Status: 'Active', ClientTxnId: clientTxnId
    };
    appendRowByHeader_(sheet, rec);
  } finally {
    lock.releaseLock();
  }
  logAudit_(session.Username, 'ATTACH_FILE', config.key.toUpperCase() + ' Ledger', txnId, '', name + ' (' + formatBytes_(bytes.length) + ')');
  return publicAttachment_(rec);
}

// Anyone who can read the ledger can see its files.
function getAttachments_(params) {
  var config = ledgerConfigFor_(params.ledger);
  return activeAttachments_(config.key, String(params.txnId || '').trim()).map(publicAttachment_);
}

function getAttachmentFile_(params) {
  var id = String(params.attachmentId || '').trim();
  var a = readAttachments_().filter(function (r) { return r.AttachmentID === id; })[0];
  if (!a || a.Status !== 'Active') throw new Error('That file isn\u2019t available any more.');
  var blob = DriveApp.getFileById(String(a.FileID)).getBlob();
  return { attachmentId: id, fileName: a.FileName, mimeType: a.MimeType, size: Number(a.Size) || 0,
           data: Utilities.base64Encode(blob.getBytes()) };
}

// Taking a file off an entry is a correction, so it needs the same
// permission as correcting an entry, for that entry's department. The file
// itself is kept in Drive.
function removeAttachment_(session, params) {
  var id = String(params.attachmentId || '').trim();
  var sheet = getOrCreateAttachmentsSheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var row = readRowByValue_(sheet, 'AttachmentID', id);
    if (!row) throw new Error('That file wasn\u2019t found.');
    var a = {};
    Object.keys(row.map).forEach(function (h) { a[h] = row.values[row.map[h]]; });
    if (a.Status === 'Removed') return { attachmentId: id, alreadyDone: true };
    requirePerm_(session, 'stock.ledger.edit');
    var entry = findLedgerEntry_(ledgerConfigFor_(a.Ledger), a.TxnID);
    if (entry) enforceDepartmentScope_(session, entry.Department);
    else if (!can_(session, 'stock.post.any')) throw new Error('That entry no longer exists; only someone who can post for any department can change its files.');
    writeRowCells_(sheet, row.rowIdx, stageCells_(row.map, {
      Status: 'Removed', RemovedBy: session.Username, RemovedAt: new Date(), RemoveReason: String(params.reason || '').trim()
    }));
    logAudit_(session.Username, 'REMOVE_FILE', String(a.Ledger).toUpperCase() + ' Ledger', a.TxnID, a.FileName,
              'Removed' + (params.reason ? ': ' + params.reason : ''));
    return { attachmentId: id, removed: true };
  } finally {
    lock.releaseLock();
  }
}

// Reads the stored balance for a code without touching it. Used when a
// duplicate post is detected and the original result has to be reconstructed.
// Quantities are stored as floats, so an error message built straight from
// one can read "56.560000000000002". Trim to 3 dp and drop trailing zeros.
function trimQty_(n) {
  return String(Math.round(Number(n) * 1000) / 1000);
}

function readCurrentBalance_(config, code) {
  var stockSheet = getSpreadsheet_().getSheetByName(config.stockSheet);
  var rowIdx = findRowIndexByValue_(stockSheet, config.codeField, code);
  if (rowIdx === -1) return 0;
  var map = headerIndexMap_(stockSheet);
  return Number(stockSheet.getRange(rowIdx, map['CurrentBalance'] + 1).getValue()) || 0;
}

// Scans the tail of the ledger for a ClientTxnId. Bounded, because a duplicate
// can only arrive seconds after the original — scanning the whole ledger would
// make every post O(ledger size) to defend against a rare case.
// Only the ClientTxnId column is scanned, so the window costs one round trip
// whatever its size. It's wide enough that an entry left unresolved
// overnight — a tablet that lost signal at the end of a shift — is still
// recognised the next morning instead of being written twice.
var IDEMPOTENCY_SCAN_ROWS = 2000;

function checkTransaction_(params) {
  var clientTxnId = String(params.clientTxnId || '').trim();
  if (!clientTxnId) throw new Error('clientTxnId is required');

  // Which ledger to search. The key is sent by the client; without it both
  // are checked, since an ID is unique across them either way.
  var configs = params.key === 'rm' ? [RM_CONFIG]
              : params.key === 'fg' ? [FG_CONFIG]
              : [RM_CONFIG, FG_CONFIG];

  for (var i = 0; i < configs.length; i++) {
    var sheet = getSpreadsheet_().getSheetByName(configs[i].ledgerSheet);
    if (!sheet) continue;
    var prior = findByClientTxnId_(sheet, configs[i], clientTxnId);
    if (prior) {
      prior.found = true;
      prior.key = configs[i].key;
      return prior;
    }
  }
  return { found: false };
}

function findByClientTxnId_(ledgerSheet, config, clientTxnId) {
  var lastRow = ledgerSheet.getLastRow();
  if (lastRow < 2) return null;
  var meta = sheetMeta_(ledgerSheet);
  var idIdx = meta.map['ClientTxnId'];
  if (idIdx === undefined) return null;

  var startRow = Math.max(2, lastRow - IDEMPOTENCY_SCAN_ROWS + 1);
  var ids = ledgerSheet.getRange(startRow, idIdx + 1, lastRow - startRow + 1, 1).getValues();

  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) !== clientTxnId) continue;
    var values = ledgerSheet.getRange(startRow + i, 1, 1, meta.lastCol).getValues()[0];
    var row = {};
    for (var c = 0; c < meta.headers.length; c++) if (meta.headers[c]) row[meta.headers[c]] = values[c];
    return {
      txnId: row.TxnID,
      code: row[config.codeField],
      txnType: row.TxnType,
      qty: Number(row.Qty),
      newBalance: readCurrentBalance_(config, row[config.codeField]),
      duplicate: true
    };
  }
  return null;
}

function addTransaction_(config, session, params) {
  var code = String(params.code || '').trim();
  var txnType = params.txnType;
  var qty = Number(params.qty);
  var department = params.department || session.Department || '';
  var reference = params.reference || '';
  var remarks = params.remarks || '';
  var rate = params.rate !== undefined && params.rate !== '' ? Number(params.rate) : '';

  if (!code) throw new Error('code is required');
  if (config.allTypes.indexOf(txnType) === -1) {
    throw new Error('Invalid txnType — must be one of: ' + config.allTypes.join(', '));
  }
  if (!qty || isNaN(qty)) throw new Error('qty must be a non-zero number');
  if (config.adjustTypes.indexOf(txnType) === -1 && qty <= 0) {
    throw new Error('qty must be positive for ' + txnType);
  }
  if (rate !== '' && isNaN(rate)) throw new Error('rate must be a number');

  enforceDepartmentScope_(session, department);
  assertStockDepartment_(department);

  var masterRow = lookupMasterRow_(config, code);
  if (!masterRow) throw new Error('Unknown code: ' + code);

  // Ledger append + balance update run as one locked unit. Without this,
  // two people posting for the same item within the same moment (e.g. two
  // Supervisors at shift start) could both read the old balance and the
  // second write would silently clobber the first — a lost update. A
  // single script-wide lock is coarser than per-item locking, but Apps
  // Script only offers script/user/document-level locks, and transaction
  // volume here doesn't justify building a per-item queue.
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ledgerSheet = getSpreadsheet_().getSheetByName(config.ledgerSheet);

    // Idempotency. The client retries a timed-out request, but a timeout only
    // means the response was lost — the append may well have committed. If the
    // same ClientTxnId is already in the ledger, return the original outcome
    // rather than posting the movement twice. Checked inside the lock so a
    // concurrent duplicate cannot slip between this read and the append.
    var clientTxnId = String(params.clientTxnId || '').trim();
    if (clientTxnId) {
      ensureColumns_(ledgerSheet, ['ClientTxnId']);
      var prior = findByClientTxnId_(ledgerSheet, config, clientTxnId);
      if (prior) return prior;
    }

    // No-negative-stock guard. This sits inside the lock and before the
    // append on purpose: checked outside it, two Issues posted in the same
    // moment could each read a sufficient balance and both commit, taking
    // the item negative anyway. A client-side check is a courtesy to the
    // operator, never the boundary — this is the boundary.
    // The item's stock row is read once here and reused by both the overdraw
    // guard and the balance update below, which used to read it separately.
    var stockSheet = getSpreadsheet_().getSheetByName(config.stockSheet);
    var stockRow = readRowByValue_(stockSheet, config.codeField, code);
    var available = stockRow && ('CurrentBalance' in stockRow.map)
      ? Number(stockRow.values[stockRow.map['CurrentBalance']]) || 0
      : 0;

    if (!ALLOW_NEGATIVE_STOCK) {
      var outQty = 0;
      if (config.outTypes.indexOf(txnType) > -1) outQty = qty;
      else if (config.adjustTypes.indexOf(txnType) > -1 && qty < 0) outQty = -qty;

      if (outQty > 0) {
        // Tolerance mirrors DRIFT_TOLERANCE so issuing the entire balance
        // still works: float noise must not make 56.56 out of 56.56 look
        // like an overdraw.
        if (outQty - available > DRIFT_TOLERANCE) {
          throw new Error(
            'Not enough stock. ' + code + ' has ' + trimQty_(available) + ' on hand and this ' +
            'would remove ' + trimQty_(outQty) + '. Post the missing Receipt first, or reduce the quantity.'
          );
        }
      }
    }

    var txnId = generateId_(config.txnIdPrefix);
    var now = new Date();
    var value = rate !== '' ? rate * qty : '';

    var rowObj = {
      TxnID: txnId,
      Date: params.date ? parseDay_(params.date) : now,
      TxnType: txnType,
      Qty: qty,
      Unit: masterRow[config.unitField] || '',
      Rate: rate,
      Value: value,
      Department: department,
      Reference: reference,
      Remarks: remarks,
      EnteredBy: session.Username,
      Timestamp: now
    };
    rowObj[config.codeField] = code;
    if (clientTxnId) rowObj.ClientTxnId = clientTxnId;

    // Optional procurement details on RM receipts (Supplier/InvoiceNo/
    // InvoiceDate). Auto-creates the columns on first use rather than
    // requiring ADMIN_migrateSchema() to have been run already — and
    // scoped to RM via config.key so FG never touches this regardless of
    // what the frontend sends.
    if (config.key === 'rm' && (params.supplier || params.invoiceNo || params.invoiceDate)) {
      ensureColumns_(ledgerSheet, ['Supplier', 'InvoiceNo', 'InvoiceDate']);
      if (params.supplier) rowObj.Supplier = params.supplier;
      if (params.invoiceNo) rowObj.InvoiceNo = params.invoiceNo;
      if (params.invoiceDate) rowObj.InvoiceDate = parseDay_(params.invoiceDate);
    }
    appendRowByHeader_(ledgerSheet, rowObj);

    var newBalance = applyStockDelta_(config, code, txnType, qty, masterRow, stockSheet, stockRow);

    logAudit_(session.Username, 'ADD_TRANSACTION', config.ledgerSheet, txnId, '',
      JSON.stringify({ code: code, txnType: txnType, qty: qty, department: department }));

    var result = { txnId: txnId, code: code, txnType: txnType, qty: qty, newBalance: newBalance };
    // Adjustments are allowed to push a balance negative (correcting an
    // over-counted receipt, etc.) but it's easy to fat-finger, so flag it
    // back to the caller instead of letting it pass silently — the
    // frontend surfaces this as a warning toast.
    if (newBalance < 0) {
      result.warning = 'Balance for ' + code + ' is now negative (' + newBalance + ') — double-check this entry.';
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}

// stockSheet/stockRow are passed in by addTransaction_, which has already
// read them; other callers can omit both and they're read here.
// What a transaction does to the two running totals. Split out so an entry
// can be undone by applying the same pair negated — which is what amending
// and deleting a ledger row rely on.
function txnDelta_(config, txnType, qty) {
  qty = Number(qty) || 0;
  if (config.inTypes.indexOf(txnType) > -1) return { inDelta: qty, outDelta: 0 };
  if (config.outTypes.indexOf(txnType) > -1) return { inDelta: 0, outDelta: qty };
  // Adjustment — signed qty: positive increases stock, negative decreases it.
  return qty >= 0 ? { inDelta: qty, outDelta: 0 } : { inDelta: 0, outDelta: -qty };
}

function applyStockDelta_(config, code, txnType, qty, masterRow, stockSheet, stockRow) {
  var d = txnDelta_(config, txnType, qty);
  return applyStockTotals_(config, code, d.inDelta, d.outDelta, masterRow, stockSheet, stockRow);
}

// Deltas may be negative here: that is how a correction takes an entry back
// out of the running totals.
function applyStockTotals_(config, code, inDelta, outDelta, masterRow, stockSheet, stockRow) {
  stockSheet = stockSheet || getSpreadsheet_().getSheetByName(config.stockSheet);
  if (stockRow === undefined) stockRow = readRowByValue_(stockSheet, config.codeField, code);
  var now = new Date();

  if (!stockRow) {
    var balance = inDelta - outDelta;
    var rowObj = {
      OpeningStock: 0,
      CurrentBalance: balance,
      StockStatus: computeStockStatus_(balance, masterRow[config.reorderField]),
      LastUpdated: now
    };
    rowObj[config.codeField] = code;
    rowObj[config.nameField] = masterRow[config.nameField] || '';
    rowObj[config.categoryField] = masterRow[config.categoryField] || '';
    rowObj[config.unitField] = masterRow[config.unitField] || '';
    rowObj[config.reorderField] = masterRow[config.reorderField] || 0;
    rowObj[config.receivedField] = inDelta;
    rowObj[config.issuedField] = outDelta;
    appendRowByHeader_(stockSheet, rowObj);
    return balance;
  }

  var map = stockRow.map;
  var cellNum = function (header) {
    return (header in map) ? Number(stockRow.values[map[header]]) || 0 : 0;
  };
  var currentIn = cellNum(config.receivedField);
  var currentOut = cellNum(config.issuedField);
  var opening = cellNum('OpeningStock');
  // Master is the source of truth for the reorder level; the copy on this
  // row is a mirror that can lag behind an edit. Reading master here means
  // the recomputed status always reflects the current threshold, and the
  // mirror gets corrected on the way through.
  var reorder = Number(masterRow[config.reorderField]) || 0;

  var newIn = currentIn + inDelta;
  var newOut = currentOut + outDelta;
  var newBalance = opening + newIn - newOut;

  var changes = {};
  var stage = function (header, value) { if (header in map) changes[map[header]] = value; };
  stage(config.receivedField, newIn);
  stage(config.issuedField, newOut);
  stage('CurrentBalance', newBalance);
  stage(config.reorderField, reorder);
  stage('StockStatus', computeStockStatus_(newBalance, reorder));
  stage('LastUpdated', now);
  writeRowCells_(stockSheet, stockRow.rowIdx, changes);

  return newBalance;
}

function computeStockStatus_(balance, reorderLevel) {
  var reorder = Number(reorderLevel) || 0;
  if (balance <= 0) return 'Critical';
  if (reorder > 0 && balance <= reorder) return 'Low';
  return 'OK';
}

// Full rebuild of the current-stock totals straight from Master + Ledger.
// Safe to run any time — it's a repair tool, not part of the normal flow.
// Every master item gets a row; rows stay where they are; columns other
// than the stock figures are left alone.
function recalcCurrentStock_(config) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = getSpreadsheet_();
    var masterRows = readSheetAsObjects_(ss.getSheetByName(config.masterSheet));
    var ledgerRows = readSheetAsObjects_(ss.getSheetByName(config.ledgerSheet));
    var stockSheet = ss.getSheetByName(config.stockSheet);

    // The columns this rebuild owns. Any other column on the tab — a note,
    // a formula someone added — is kept, row for row, and so is the column
    // order. The previous version cleared the whole tab and wrote these
    // eleven back, which deleted anything else.
    var headers = [config.codeField, config.nameField, config.categoryField, config.unitField,
      'OpeningStock', config.receivedField, config.issuedField, 'CurrentBalance',
      config.reorderField, 'StockStatus', 'LastUpdated'];
    var hadHeader = stockSheet.getLastRow() > 0;
    ensureColumns_(stockSheet, headers);
    if (!hadHeader) styleHeader_(stockSheet, headers.length);
    var meta = sheetMeta_(stockSheet);
    var lastCol = meta.lastCol;

    // The tab as it stands, formulas included. Every row stays where it is
    // and only the columns above are rewritten, so a formula in another
    // column still points at the row it was written for. Plain text is
    // re-escaped on the way back: a cell reading "=x" as text would
    // otherwise come back as a formula.
    var oldCount = Math.max(0, stockSheet.getLastRow() - 1);
    var oldRange = oldCount ? stockSheet.getRange(2, 1, oldCount, lastCol) : null;
    var oldValues = oldRange ? oldRange.getValues() : [];
    var oldFormulas = oldRange ? oldRange.getFormulas() : [];
    var codeIdx = meta.map[config.codeField];

    var totals = {};
    ledgerRows.forEach(function (t) {
      var code = t[config.codeField];
      if (!code) return;
      if (!totals[code]) totals[code] = { inQty: 0, outQty: 0 };
      var qty = Number(t.Qty) || 0;
      if (config.inTypes.indexOf(t.TxnType) > -1) {
        totals[code].inQty += qty;
      } else if (config.outTypes.indexOf(t.TxnType) > -1) {
        totals[code].outQty += qty;
      } else {
        if (qty >= 0) totals[code].inQty += qty; else totals[code].outQty += -qty;
      }
    });

    var masterByCode = {};
    var masterOrder = [];
    masterRows.forEach(function (m) {
      var code = m[config.codeField];
      if (code === '' || code === null || code === undefined || masterByCode[code]) return;
      masterByCode[code] = m;
      masterOrder.push(code);
    });

    var now = new Date();
    var itemsWithActivity = 0;
    var notInMaster = 0;
    var duplicates = [];
    var onTab = {};

    // The owned cells of one row, rebuilt from the ledger. `item` gives the
    // name, category, unit and reorder level: the master row, or for a code
    // no longer in Master, the tab's own row, so its opening stock and
    // history aren't dropped.
    var rebuild = function (cells, code, item, opening) {
      var t = totals[code];
      if (t) itemsWithActivity++;
      t = t || { inQty: 0, outQty: 0 };
      var reorder = Number(item[config.reorderField]) || 0;
      var balance = opening + t.inQty - t.outQty;
      var owned = {
        OpeningStock: opening,
        CurrentBalance: balance,
        StockStatus: computeStockStatus_(balance, reorder),
        LastUpdated: now
      };
      owned[config.codeField] = code;
      owned[config.nameField] = item[config.nameField];
      owned[config.categoryField] = item[config.categoryField];
      owned[config.unitField] = item[config.unitField];
      owned[config.reorderField] = reorder;
      owned[config.receivedField] = t.inQty;
      owned[config.issuedField] = t.outQty;
      headers.forEach(function (h) {
        var v = owned[h];
        cells[meta.map[h]] = sheetSafe_(v === undefined || v === null ? '' : v);
      });
      return cells;
    };

    // Rows already on the tab, in place. Opening balances are not reset:
    // they're entered independently of any transaction, so the tab's own
    // figure is kept.
    var values = oldValues.map(function (row, i) {
      var cells = row.map(function (v, c) { return oldFormulas[i][c] || sheetSafe_(v); });
      var code = row[codeIdx];
      if (code === '' || code === null || code === undefined) return cells; // a note row, left as is
      if (onTab[code]) duplicates.push(String(code));
      onTab[code] = true;
      var item = masterByCode[code];
      if (!item) {
        notInMaster++;
        item = {};
        meta.headers.forEach(function (h, c) { if (h) item[h] = row[c]; });
      }
      return rebuild(cells, code, item, Number(row[meta.map['OpeningStock']]) || 0);
    });

    // Every master item gets a row, whether or not it has ever had a
    // transaction — a new item with no movements yet must not vanish from
    // Current Stock on a recalc. Those not on the tab yet go at the end.
    var added = 0;
    masterOrder.forEach(function (code) {
      if (onTab[code]) return;
      values.push(rebuild(new Array(lastCol).fill(''), code, masterByCode[code], 0));
      added++;
    });
    var outRows = values.filter(function (cells) {
      var code = cells[codeIdx];
      return code !== '' && code !== null && code !== undefined;
    });

    // One write over the rows in place. The previous version cleared the
    // whole tab first, which left it empty whenever the run stopped between
    // the clear and the write, and deleted every column it didn't own.
    if (values.length > 0) {
      stockSheet.getRange(2, 1, values.length, lastCol).setValues(values);
    }

    var result = { itemsRecalculated: outRows.length, itemsWithNoActivity: outRows.length - itemsWithActivity };
    if (added) result.itemsAdded = added;
    if (notInMaster) result.itemsNotInMaster = notInMaster;
    if (duplicates.length) {
      result.warning = (result.warning ? result.warning + ' ' : '') + 'These codes appear on more than one row of ' +
        config.stockSheet + ', and each copy now shows the same figures: ' + duplicates.join(', ') +
        '. Delete the extra rows by hand.';
    }
    // Not a hard limit — this is a rare, manual, Admin-only repair tool
    // and it should still attempt the work. But Apps Script executions
    // have a hard 6-minute ceiling, and a huge ledger with no explanation
    // would just fail as an opaque timeout. 20,000 rows is comfortably
    // past anything this ever needs to touch at normal usage, so it's a
    // reasonable "you're in unusual territory" line, not a real limit.
    if (ledgerRows.length > 20000) {
      result.warning = 'The ' + config.ledgerSheet + ' ledger has grown to ' + ledgerRows.length + ' rows — this recalculation may take a while, and a very large ledger risks Apps Script\u2019s execution time limit.' +
        (result.warning ? ' ' + result.warning : '');
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
   ITEM MASTER CRUD — soft delete only (Active -> FALSE); ledger rows that
   reference a deactivated code stay intact and historically accurate.
   ========================================================================= */
function buildMasterRowFromParams_(params) {
  // ID_No. is assigned once by addMasterItem_ and never editable afterward —
  // excluded here so it can't be overwritten even by a hand-crafted request.
  var excluded = { action: 1, token: 1, code: 1, Active: 1, CreatedDate: 1, CreatedBy: 1, 'ID_No.': 1 };
  var rowObj = {};
  Object.keys(params).forEach(function (key) {
    if (!excluded[key] && params[key] !== undefined) rowObj[key] = params[key];
  });
  return rowObj;
}

function addMasterItem_(config, session, params) {
  var code = String(params[config.codeField] || params.code || '').trim();
  if (!code) throw new Error(config.codeField + ' is required');

  // Locked because generateNextMasterId_ reads the current max ID before
  // writing the next one — without a lock, two people adding an item at
  // the same moment could both land on the same ID_No.
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSpreadsheet_().getSheetByName(config.masterSheet);
    if (findRowIndexByValue_(sheet, config.codeField, code) > -1) {
      throw new Error('Code already exists: ' + code);
    }
    if (!params[config.nameField]) throw new Error(config.nameField + ' is required');

    var rowObj = buildMasterRowFromParams_(params);
    rowObj[config.codeField] = code;
    rowObj[config.idField] = generateNextMasterId_(config, sheet);
    rowObj['Active'] = true;
    rowObj['CreatedDate'] = new Date();
    rowObj['CreatedBy'] = session.Username;
    if (!rowObj[config.unitField]) rowObj[config.unitField] = 'NOS';

    appendRowByHeader_(sheet, rowObj);
    logAudit_(session.Username, 'ADD_ITEM', config.masterSheet, code, '', JSON.stringify(rowObj));
    return { code: code, idNo: rowObj[config.idField] };
  } finally {
    lock.releaseLock();
  }
}

// Sequential, human-readable IDs (RM-00001, FG-00001, ...) — separate from
// generateId_'s timestamp+random IDs, which aren't meant to be read by a
// person. Auto-creates the ID_No. column if it isn't there yet, so this
// doesn't depend on the column already existing at column A.
function generateNextMasterId_(config, sheet) {
  ensureColumns_(sheet, [config.idField]);
  var maxNum = 0;
  readSheetAsObjects_(sheet).forEach(function (r) {
    var m = String(r[config.idField] || '').match(/(\d+)\s*$/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  var padded = ('00000' + (maxNum + 1)).slice(-5);
  return config.idPrefix + '-' + padded;
}

// Batch version of appendRowByHeader_ — writes every row in a single
// setValues() call instead of one appendRow per row, which is what
// actually matters once this is dozens or hundreds of rows (a real bulk
// import) rather than the single-row case appendRowByHeader_ handles.
function appendRowsByHeader_(sheet, rowObjs) {
  if (rowObjs.length === 0) return;
  var map = headerIndexMap_(sheet);
  var lastCol = sheet.getLastColumn();
  var startRow = sheet.getLastRow() + 1;
  var values = rowObjs.map(function (rowObj) {
    var row = new Array(lastCol).fill('');
    Object.keys(rowObj).forEach(function (key) { if (key in map) row[map[key]] = rowObj[key]; });
    return row;
  });
  sheet.getRange(startRow, 1, values.length, lastCol).setValues(values);
}

// Bulk item creation — mirrors addMasterItem_ row-for-row (same required
// fields, same sequential ID_No. scheme, same Active/CreatedDate/CreatedBy
// stamping) but writes every row in two setValues() calls total (one for
// Master, one for Current Stock) instead of looping the single-item path
// N times, and validates the whole batch up front so a few bad rows don't
// block the good ones.
//
// Also creates each item's Current Stock row directly, seeded from its
// Opening Stock — a freshly-imported item would otherwise be invisible in
// the Stock tables/dashboard until its first transaction (that's true of
// the single-item Add flow too, not something bulk import introduces, but
// it matters more here since a real import is often the entire reason
// someone's opening balances need to show up immediately).
function bulkAddMasterItems_(config, session, params) {
  var rows = params.rows;
  if (!Array.isArray(rows) || rows.length === 0) throw new Error('No rows to import');
  if (rows.length > 500) throw new Error('Import is capped at 500 rows at a time \u2014 split into smaller batches');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var masterSheet = getSpreadsheet_().getSheetByName(config.masterSheet);
    var stockSheet = getSpreadsheet_().getSheetByName(config.stockSheet);

    var existingCodes = {};
    readSheetAsObjects_(masterSheet).forEach(function (m) { existingCodes[m[config.codeField]] = true; });

    // Same scheme as generateNextMasterId_, computed once so a batch of N
    // rows gets N consecutive IDs instead of re-scanning the sheet before
    // every single row.
    ensureColumns_(masterSheet, [config.idField]);
    var maxNum = 0;
    readSheetAsObjects_(masterSheet).forEach(function (r) {
      var m = String(r[config.idField] || '').match(/(\d+)\s*$/);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    });

    var now = new Date();
    var seenInBatch = {};
    var masterRowObjs = [];
    var stockRowObjs = [];
    var skipped = [];

    rows.forEach(function (raw, idx) {
      var code = String(raw[config.codeField] || raw.code || '').trim();
      var name = String(raw[config.nameField] || raw.name || '').trim();
      var rowLabel = 'Row ' + (idx + 2); // +2 = header row + 1-based row numbering, matches what the user sees in Excel

      if (!code) { skipped.push({ row: rowLabel, code: code, reason: 'Missing code' }); return; }
      if (!name) { skipped.push({ row: rowLabel, code: code, reason: 'Missing name' }); return; }
      if (existingCodes[code]) { skipped.push({ row: rowLabel, code: code, reason: 'Code already exists' }); return; }
      if (seenInBatch[code]) { skipped.push({ row: rowLabel, code: code, reason: 'Duplicate within this file' }); return; }
      seenInBatch[code] = true;

      maxNum++;
      var idNo = config.idPrefix + '-' + ('00000' + maxNum).slice(-5);
      var opening = Number(raw.OpeningStock) || 0;
      // The transaction guard can't see imports, so the same floor is
      // applied here — otherwise a negative Opening column would seed a
      // negative balance the guard would then refuse to let anyone fix.
      if (!ALLOW_NEGATIVE_STOCK && opening < 0) {
        skipped.push({ row: rowLabel, code: code, reason: 'Opening stock cannot be negative' });
        return;
      }
      var reorder = Number(raw[config.reorderField]) || 0;
      var unit = String(raw[config.unitField] || 'NOS').trim();
      var category = String(raw[config.categoryField] || '').trim();

      var masterRow = {};
      masterRow[config.codeField] = code;
      masterRow[config.nameField] = name;
      masterRow[config.categoryField] = category;
      masterRow[config.unitField] = unit;
      masterRow[config.reorderField] = reorder;
      masterRow[config.idField] = idNo;
      masterRow.Active = true;
      masterRow.CreatedDate = now;
      masterRow.CreatedBy = session.Username;
      masterRowObjs.push(masterRow);

      var stockRow = {
        OpeningStock: opening,
        CurrentBalance: opening,
        StockStatus: computeStockStatus_(opening, reorder),
        LastUpdated: now
      };
      stockRow[config.codeField] = code;
      stockRow[config.nameField] = name;
      stockRow[config.categoryField] = category;
      stockRow[config.unitField] = unit;
      stockRow[config.reorderField] = reorder;
      stockRow[config.receivedField] = 0;
      stockRow[config.issuedField] = 0;
      stockRowObjs.push(stockRow);
    });

    if (masterRowObjs.length > 0) {
      appendRowsByHeader_(masterSheet, masterRowObjs);
      appendRowsByHeader_(stockSheet, stockRowObjs);
      var codes = masterRowObjs.map(function (r) { return r[config.codeField]; });
      logAudit_(session.Username, 'BULK_ADD_ITEMS', config.masterSheet, masterRowObjs.length + ' item(s)', '', JSON.stringify(codes));
    }

    return { added: masterRowObjs.length, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

function updateMasterItem_(config, session, params) {
  var code = String(params[config.codeField] || params.code || '').trim();
  if (!code) throw new Error(config.codeField + ' is required');

  // Locked for the same reason as addMasterItem_ — two Managers editing the
  // same item at the same instant shouldn't be able to race and silently
  // drop one of the edits.
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSpreadsheet_().getSheetByName(config.masterSheet);
    var rowIdx = findRowIndexByValue_(sheet, config.codeField, code);
    if (rowIdx === -1) throw new Error('Unknown code: ' + code);

    var map = headerIndexMap_(sheet);
    var before = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];

    var rowObj = buildMasterRowFromParams_(params);
    // The code is how the row was found; it isn't rewritten.
    delete rowObj[config.codeField];
    // One merged write through writeRowCells_, which also escapes text that
    // starts with = + - or @. The per-field setValue this replaces skipped
    // that, so a name typed as "=..." became a live formula.
    writeRowCells_(sheet, rowIdx, stageCells_(map, rowObj));

    logAudit_(session.Username, 'UPDATE_ITEM', config.masterSheet, code, JSON.stringify(before), JSON.stringify(rowObj));

    // Current_Stock keeps its own copy of the reorder level and a derived
    // StockStatus, and both are only refreshed when a transaction is posted.
    // Editing the reorder level here would otherwise leave that row stale
    // indefinitely — the item would keep reporting "Low" against the old
    // threshold (or stop reporting Low against a new one) until someone
    // happened to post a transaction or ran a full Recalculate.
    if (config.reorderField in rowObj) {
      syncCurrentStockReorder_(config, code, rowObj[config.reorderField]);
    }

    return { code: code };
  } finally {
    lock.releaseLock();
  }
}

// Pushes a changed reorder level onto the item's Current_Stock row and
// recomputes its status from the balance already recorded there. No-ops
// when the item has no stock row yet (nothing has been posted for it).
function syncCurrentStockReorder_(config, code, reorderLevel) {
  var stockSheet = getSpreadsheet_().getSheetByName(config.stockSheet);
  if (!stockSheet) return;
  var rowIdx = findRowIndexByValue_(stockSheet, config.codeField, code);
  if (rowIdx === -1) return;

  var map = headerIndexMap_(stockSheet);
  var reorder = Number(reorderLevel) || 0;
  var balance = Number(stockSheet.getRange(rowIdx, map['CurrentBalance'] + 1).getValue()) || 0;

  if (config.reorderField in map) stockSheet.getRange(rowIdx, map[config.reorderField] + 1).setValue(reorder);
  if ('StockStatus' in map) stockSheet.getRange(rowIdx, map['StockStatus'] + 1).setValue(computeStockStatus_(balance, reorder));
}

function setMasterItemActive_(config, session, params, active) {
  var code = String(params[config.codeField] || params.code || '').trim();
  if (!code) throw new Error(config.codeField + ' is required');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSpreadsheet_().getSheetByName(config.masterSheet);
    var rowIdx = findRowIndexByValue_(sheet, config.codeField, code);
    if (rowIdx === -1) throw new Error('Unknown code: ' + code);

    var map = headerIndexMap_(sheet);
    sheet.getRange(rowIdx, map['Active'] + 1).setValue(active);

    logAudit_(session.Username, active ? 'REACTIVATE_ITEM' : 'DEACTIVATE_ITEM', config.masterSheet, code,
              active ? 'inactive' : 'active', active ? 'active' : 'inactive');
    return { code: code, active: active };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================================
   RM / FG CONFIGS — the only thing that differs between the two
   ========================================================================= */
var RM_CONFIG = {
  masterSheet: 'RM_Master',
  ledgerSheet: 'RM_Stock_Ledger',
  stockSheet: 'RM_Current_Stock',
  codeField: 'RM_Code',
  idField: 'ID_No.',
  idPrefix: 'RM',
  nameField: 'RM_Name',
  categoryField: 'Category',
  unitField: 'Unit',
  reorderField: 'ReorderLevel',
  receivedField: 'TotalReceipts',
  issuedField: 'TotalIssues',
  inTypes: ['Receipt', 'Return'],
  outTypes: ['Issue'],
  adjustTypes: ['Adjustment'],
  allTypes: ['Receipt', 'Issue', 'Adjustment', 'Return'],
  txnIdPrefix: 'RMT',
  key: 'rm'
};

var FG_CONFIG = {
  masterSheet: 'FG_Master',
  ledgerSheet: 'FG_Stock_Ledger',
  stockSheet: 'FG_Current_Stock',
  codeField: 'FG_Code',
  idField: 'ID_No.',
  idPrefix: 'FG',
  nameField: 'FG_Name',
  categoryField: 'FG_Family',
  unitField: 'Unit',
  reorderField: 'ReorderLevel',
  receivedField: 'TotalProduced',
  issuedField: 'TotalIssued',
  inTypes: ['Production', 'Return'],
  outTypes: ['Dispatch'],
  adjustTypes: ['Adjustment'],
  allTypes: ['Production', 'Dispatch', 'Adjustment', 'Return'],
  txnIdPrefix: 'FGT',
  key: 'fg'
};

/* =========================================================================
   GENERIC SHEET HELPERS
   -------------------------------------------------
   Everything here is header-driven: functions look up columns by header
   name, not position, so reordering or inserting a column in the Sheet
   never breaks the script.
   ========================================================================= */
function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function styleHeader_(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols).setFontWeight('bold').setBackground('#f1efe8');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, numCols);
}

// Appends any header in `names` that the sheet doesn't already have.
// Existing columns/data are left exactly where they are.
function ensureColumns_(sheet, names) {
  var meta = sheetMeta_(sheet);
  var lastCol = meta.lastCol;
  var added = false;
  names.forEach(function (name) {
    if (name in meta.map) return;
    lastCol += 1;
    sheet.getRange(1, lastCol).setValue(name);
    meta.map[name] = lastCol - 1;
    added = true;
  });
  if (added) invalidateSheetMeta_(sheet);
}

// { headerName: zeroBasedColumnIndex }
// Every getValue/setValue/getValues below is a network round trip to the
// Sheets service, and they dominate how long a request takes. One posted
// transaction used to make around forty of them — the header row alone was
// re-read five times. These three helpers cut that down: headers are read
// once per execution, a row is read whole rather than cell by cell, and the
// cells that change are written back in as few ranges as possible.
var _sheetMetaCache_ = {};
function sheetMeta_(sheet) {
  var key = sheet.getSheetId();
  if (_sheetMetaCache_[key]) return _sheetMetaCache_[key];
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var map = {};
  headers.forEach(function (h, i) { if (h) map[h] = i; });
  var meta = { map: map, headers: headers, lastCol: lastCol };
  _sheetMetaCache_[key] = meta;
  return meta;
}
// Called whenever a column is added, so the cached header row can't go stale
// within an execution.
function invalidateSheetMeta_(sheet) { delete _sheetMetaCache_[sheet.getSheetId()]; }

function headerIndexMap_(sheet) { return sheetMeta_(sheet).map; }

// Finds a row by a key column and returns it whole, so the caller doesn't
// pay a round trip per cell it needs. Returns null when there's no match.
function readRowByValue_(sheet, headerName, value) {
  var meta = sheetMeta_(sheet);
  if (!(headerName in meta.map)) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var keys = sheet.getRange(2, meta.map[headerName] + 1, lastRow - 1, 1).getValues();
  var target = String(value);
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) !== target) continue;
    var rowIdx = i + 2;
    return {
      rowIdx: rowIdx,
      map: meta.map,
      lastCol: meta.lastCol,
      values: sheet.getRange(rowIdx, 1, 1, meta.lastCol).getValues()[0]
    };
  }
  return null;
}

// Writes several cells of one row, merging adjacent columns into a single
// range write. Columns that aren't being changed are never rewritten, so a
// formula in a neighbouring column survives untouched.
// Sheets reads a cell that starts with = + - or @ as a formula. Written as
// plain text by this script — an access override like "+stock.recalculate",
// or a remark someone types starting with "=" — it becomes #ERROR! at best,
// and at worst a formula that runs inside the spreadsheet. A leading
// apostrophe tells Sheets "this is text"; Sheets drops the apostrophe itself,
// so reading the cell back gives exactly the string that was written.
function sheetSafe_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}

// A form sends a date as "2026-09-24". new Date() reads that as midnight
// UTC, which Sheets then shows in India as 05:30 — the right day with a
// spurious time, and the wrong day anywhere west of Greenwich. Built from
// its parts instead, it is midnight in the script's own time zone.
function parseDay_(v) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v || '').trim());
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date(v);
}

function writeRowCells_(sheet, rowIdx, changes) {
  var cols = Object.keys(changes).map(Number).sort(function (a, b) { return a - b; });
  var i = 0;
  while (i < cols.length) {
    var startAt = i;
    while (i + 1 < cols.length && cols[i + 1] === cols[i] + 1) i++;
    var run = cols.slice(startAt, i + 1);
    sheet.getRange(rowIdx, run[0] + 1, 1, run.length)
         .setValues([run.map(function (c) { return sheetSafe_(changes[c]); })]);
    i++;
  }
}

// Whole sheet (minus header row) as an array of { headerName: value } objects
// Builds the { columnIndex: value } map writeRowCells_ wants from
// { header: value }, skipping headers this sheet doesn't have.
function stageCells_(map, fields) {
  var changes = {};
  Object.keys(fields).forEach(function (header) {
    if (map[header] !== undefined) changes[map[header]] = fields[header];
  });
  return changes;
}

function readSheetAsObjects_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) {
      if (h) obj[h] = row[i];
    });
    return obj;
  });
}

// Appends one row, mapping rowObj keys to the matching header column.
// Keys that don't match a header are silently dropped; missing headers
// are left blank.
function appendRowByHeader_(sheet, rowObj) {
  var meta = sheetMeta_(sheet);
  var map = meta.map;
  var row = new Array(meta.lastCol).fill('');
  Object.keys(rowObj).forEach(function (key) {
    if (key in map) row[map[key]] = sheetSafe_(rowObj[key]);
  });
  sheet.appendRow(row);
}

// 1-based sheet row number of the first row where headerName === value,
// or -1 if not found.
function findRowIndexByValue_(sheet, headerName, value) {
  var map = headerIndexMap_(sheet);
  if (!(headerName in map)) return -1;
  var col = map[headerName] + 1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) return i + 2;
  }
  return -1;
}

// The time to the second, then a random tail. Three random digits gave two
// entries in the same second a 1-in-900 chance of sharing an ID — tolerable
// while IDs were only labels, not now that correcting or deleting an entry
// finds it by ID. Six hex characters make that about 1 in 16 million.
// Earlier IDs keep their shape and stay valid; lookups are by exact string.
function generateId_(prefix) {
  var tz = Session.getScriptTimeZone();
  var ts = Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss');
  var rand = Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  return prefix + '-' + ts + '-' + rand;
}

function logAudit_(user, action, module, recordId, oldValue, newValue) {
  try {
    var sheet = getSpreadsheet_().getSheetByName('AuditLog');
    appendRowByHeader_(sheet, {
      Timestamp: new Date(),
      User: user,
      Action: action,
      Module: module,
      RecordID: recordId,
      OldValue: oldValue,
      NewValue: newValue
    });
  } catch (err) {
    // Never let an audit-log failure break the actual operation.
    Logger.log('Audit log failed: ' + err.message);
  }
}

// For names and codes placed in email HTML. They come from the sheet, so a
// name containing < or & would otherwise break the layout or inject markup.
function escapeHtml_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}