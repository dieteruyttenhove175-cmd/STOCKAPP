/* =========================================================
   11_SessionService.gs
   Refactor: session core service
   Doel:
   - centrale sessielaag voor code-login
   - admin Google-auth blijft ondersteund als fallback
   - sessies lezen / aanmaken / ongeldig maken / opschonen
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / config / helpers
   --------------------------------------------------------- */

function getSessionsTab_() {
  return TABS.SESSIONS || 'Sessies';
}

function getSessionHours_() {
  return safeNumber((APP_CONFIG && APP_CONFIG.SESSION_HOURS) || 12, 12);
}

function makeSessionId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'SID-' + stamp + '-' + makeUuidId().slice(0, 8).toUpperCase();
}

function getNowRaw_() {
  return Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function getExpiryRawFromNow_(hours) {
  var dt = new Date();
  dt.setHours(dt.getHours() + safeNumber(hours, getSessionHours_()));
  return Utilities.formatDate(dt, TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function isIsoDateTimeExpired_(rawValue) {
  var raw = safeText(rawValue);
  if (!raw) return true;
  return raw < getNowRaw_();
}

function withSessionLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------
   Mapping / read layer
   --------------------------------------------------------- */

function mapSessionRow(row) {
  return {
    sessionId: safeText(row.SessionID || row.SessionId || row.ID),
    userCode: safeText(row.UserCode || row.GebruikerCode),
    userName: safeText(row.UserName || row.GebruikerNaam || row.Naam),
    userRole: safeText(row.UserRole || row.Rol),
    loginEmail: normalizeLoginEmail(row.LoginEmail || row.Email || ''),
    authMode: safeText(row.AuthMode || row.Auth || 'code'),
    active: row.Active === undefined ? isTrue(row.Actief) : isTrue(row.Active),
    createdAt: safeText(row.CreatedAt || row.AangemaaktOp),
    createdAtRaw: safeText(row.CreatedAtRaw || row.AangemaaktOpRaw || row.CreatedAt || row.AangemaaktOp),
    expiresAt: safeText(row.ExpiresAt || row.VervaltOp),
    expiresAtRaw: safeText(row.ExpiresAtRaw || row.VervaltOpRaw || row.ExpiresAt || row.VervaltOp),
    invalidatedAt: safeText(row.InvalidatedAt || row.OngeldigOp),
    invalidatedReason: safeText(row.InvalidatedReason || row.OngeldigReden),
  };
}

function getAllSessions_() {
  return readObjectsSafe(getSessionsTab_())
    .map(mapSessionRow)
    .sort(function (a, b) {
      return (
        safeText(b.createdAtRaw).localeCompare(safeText(a.createdAtRaw)) ||
        safeText(b.sessionId).localeCompare(safeText(a.sessionId))
      );
    });
}

function getSessionById(sessionId) {
  var id = safeText(sessionId);
  if (!id) return null;

  return getAllSessions_().find(function (item) {
    return safeText(item.sessionId) === id;
  }) || null;
}

function isSessionActive_(session) {
  if (!session) return false;
  if (!isTrue(session.active)) return false;
  if (isIsoDateTimeExpired_(session.expiresAtRaw)) return false;
  return true;
}

function getActiveSessionById_(sessionId) {
  var session = getSessionById(sessionId);
  return isSessionActive_(session) ? session : null;
}

function getUserBySessionId(sessionId) {
  var session = getActiveSessionById_(sessionId);
  if (!session) return null;

  return {
    code: session.userCode,
    naam: session.userName,
    rol: session.userRole,
    email: session.loginEmail,
    authMode: session.authMode,
    sessionId: session.sessionId,
  };
}

/* ---------------------------------------------------------
   Create / invalidate
   --------------------------------------------------------- */

function buildSessionObject_(payload) {
  payload = payload || {};

  var createdAtRaw = getNowRaw_();
  var expiresAtRaw = getExpiryRawFromNow_(getSessionHours_());

  return {
    SessionID: makeSessionId_(),
    UserCode: safeText(payload.userCode),
    UserName: safeText(payload.userName),
    UserRole: safeText(payload.userRole),
    LoginEmail: normalizeLoginEmail(payload.loginEmail),
    AuthMode: safeText(payload.authMode || 'code'),
    Active: true,
    CreatedAt: toDisplayDateTime(createdAtRaw),
    CreatedAtRaw: createdAtRaw,
    ExpiresAt: toDisplayDateTime(expiresAtRaw),
    ExpiresAtRaw: expiresAtRaw,
    InvalidatedAt: '',
    InvalidatedReason: '',
  };
}

function createSessionForUser(payload) {
  payload = payload || {};

  if (!safeText(payload.userCode)) throw new Error('UserCode is verplicht voor sessie.');
  if (!safeText(payload.userRole)) throw new Error('UserRole is verplicht voor sessie.');

  return withSessionLock_(function () {
    var obj = buildSessionObject_(payload);
    appendObjects(getSessionsTab_(), [obj]);
    return mapSessionRow(obj);
  });
}

function invalidateSessionById(sessionId, reason) {
  var id = safeText(sessionId);
  if (!id) throw new Error('SessionId ontbreekt.');

  return withSessionLock_(function () {
    var table = getAllValues(getSessionsTab_());
    if (!table.length) {
      return { invalidated: false, sessionId: id };
    }

    var headerRow = table[0];
    var dataRows = table.slice(1);
    var changed = false;
    var nowRaw = getNowRaw_();

    var newRows = dataRows.map(function (row) {
      var obj = rowToObject(headerRow, row);
      var mapped = mapSessionRow(obj);

      if (safeText(mapped.sessionId) !== id) {
        return row;
      }

      if (!isTrue(mapped.active)) {
        return row;
      }

      obj.Active = false;
      obj.InvalidatedAt = toDisplayDateTime(nowRaw);
      obj.InvalidatedReason = safeText(reason || 'MANUAL');
      changed = true;

      return buildRowFromHeaders(headerRow, obj);
    });

    if (changed) {
      writeFullTable(getSessionsTab_(), headerRow, newRows);
    }

    return {
      invalidated: changed,
      sessionId: id,
    };
  });
}

/* ---------------------------------------------------------
   Cleanup / maintenance
   --------------------------------------------------------- */

function cleanupExpiredSessions() {
  return withSessionLock_(function () {
    var table = getAllValues(getSessionsTab_());
    if (!table.length) {
      return { updatedCount: 0 };
    }

    var headerRow = table[0];
    var dataRows = table.slice(1);
    var updatedCount = 0;
    var nowRaw = getNowRaw_();

    var newRows = dataRows.map(function (row) {
      var obj = rowToObject(headerRow, row);
      var mapped = mapSessionRow(obj);

      if (!isTrue(mapped.active)) {
        return row;
      }
      if (!isIsoDateTimeExpired_(mapped.expiresAtRaw)) {
        return row;
      }

      obj.Active = false;
      obj.InvalidatedAt = toDisplayDateTime(nowRaw);
      obj.InvalidatedReason = 'EXPIRED';
      updatedCount += 1;

      return buildRowFromHeaders(headerRow, obj);
    });

    if (updatedCount) {
      writeFullTable(getSessionsTab_(), headerRow, newRows);

      writeAudit({
        actie: 'CLEANUP_EXPIRED_SESSIONS',
        actor: { naam: 'System', rol: 'System', email: '' },
        documentType: 'Sessies',
        documentId: 'EXPIRED',
        details: {
          updatedCount: updatedCount,
        },
      });
    }

    return {
      updatedCount: updatedCount,
    };
  });
}

function syncOpenSessionsToNewLoginEmail(userCode, newLoginEmail) {
  var code = safeText(userCode);
  var email = normalizeLoginEmail(newLoginEmail);

  if (!code) throw new Error('UserCode ontbreekt.');
  if (!email) throw new Error('Nieuwe login e-mail ontbreekt.');

  return withSessionLock_(function () {
    var table = getAllValues(getSessionsTab_());
    if (!table.length) {
      return { updatedCount: 0 };
    }

    var headerRow = table[0];
    var dataRows = table.slice(1);
    var updatedCount = 0;

    var newRows = dataRows.map(function (row) {
      var obj = rowToObject(headerRow, row);
      var mapped = mapSessionRow(obj);

      if (safeText(mapped.userCode) !== code) {
        return row;
      }
      if (!isTrue(mapped.active)) {
        return row;
      }
      if (isIsoDateTimeExpired_(mapped.expiresAtRaw)) {
        return row;
      }

      obj.LoginEmail = email;
      updatedCount += 1;
      return buildRowFromHeaders(headerRow, obj);
    });

    if (updatedCount) {
      writeFullTable(getSessionsTab_(), headerRow, newRows);
    }

    return {
      updatedCount: updatedCount,
      userCode: code,
      loginEmail: email,
    };
  });
}

/* ---------------------------------------------------------
   Google admin fallback
   --------------------------------------------------------- */

function resolveAdminGoogleUser_() {
  if (typeof getCurrentAdminUserRecord !== 'function') {
    return null;
  }

  var adminUser = getCurrentAdminUserRecord();
  if (!adminUser) return null;

  return {
    code: safeText(adminUser.code || adminUser.userCode || adminUser.Code),
    naam: safeText(adminUser.naam || adminUser.userName || adminUser.Naam),
    rol: safeText(adminUser.rol || adminUser.userRole || adminUser.Rol),
    email: normalizeLoginEmail(adminUser.email || adminUser.loginEmail || adminUser.Email),
    authMode: 'admin_google',
    sessionId: '',
  };
}

function resolveCustomSessionUser_(sessionId) {
  var user = getUserBySessionId(sessionId);
  if (!user) return null;

  return {
    code: safeText(user.code),
    naam: safeText(user.naam),
    rol: safeText(user.rol),
    email: normalizeLoginEmail(user.email),
    authMode: safeText(user.authMode || 'code'),
    sessionId: safeText(user.sessionId),
  };
}

/* ---------------------------------------------------------
   Main auth context
   --------------------------------------------------------- */

function requireLoggedInUser(sessionId) {
  cleanupExpiredSessions();

  var sessionUser = resolveCustomSessionUser_(sessionId);
  if (sessionUser) {
    return sessionUser;
  }

  var adminUser = resolveAdminGoogleUser_();
  if (adminUser) {
    return adminUser;
  }

  throw new Error('Geen geldige login gevonden.');
}

function getSessionContext(payload) {
  payload = payload || {};

  var sessionId = safeText(payload.sessionId || payload.sid);
  var sessionUser = resolveCustomSessionUser_(sessionId);
  if (sessionUser) {
    return {
      authenticated: true,
      user: sessionUser,
      authMode: sessionUser.authMode,
      sessionId: sessionUser.sessionId,
    };
  }

  var adminUser = resolveAdminGoogleUser_();
  if (adminUser) {
    return {
      authenticated: true,
      user: adminUser,
      authMode: adminUser.authMode,
      sessionId: '',
    };
  }

  return {
    authenticated: false,
    user: null,
    authMode: '',
    sessionId: '',
  };
}
