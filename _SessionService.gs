/* =========================================================
   11_SessionService.gs — sessies / huidige gebruiker
   ========================================================= */

function getCurrentSessionEmailSafe() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

function getAuthModeForUser(user) {
  if (!user) return '';
  return user.rol === ROLE.ADMIN ? 'admin_google' : 'custom_session';
}

function getEffectiveLoginEmail(user) {
  return normalizeLoginEmail((user && (user.loginEmail || user.email)) || '');
}

function getSessionDurationHours() {
  return Number(APP_CONFIG.SESSION_HOURS || 12);
}

function makeSessionId() {
  return makeUuidId('S');
}

function formatSessionDateTime_(dateValue) {
  return Utilities.formatDate(dateValue, getAppTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function withSessionLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getAllSessions() {
  return readObjectsSafe(TABS.SESSIONS).map(mapSession);
}

function getActiveSessions() {
  const now = new Date();

  return getAllSessions().filter(function (session) {
    if (!session || !session.actief) return false;

    const expires = new Date(session.verlooptOp);
    if (isNaN(expires.getTime())) return false;

    return expires.getTime() > now.getTime();
  });
}

function getUserBySessionId(sessionId) {
  const sid = safeText(sessionId);
  if (!sid) return null;

  const session = getActiveSessions().find(function (x) {
    return x.sessionId === sid;
  });
  if (!session) return null;

  const users = getAllUsers();
  return users.find(function (x) {
    return getEffectiveLoginEmail(x) === session.loginEmail;
  }) || null;
}

function getCurrentUserRecord() {
  const email = getCurrentSessionEmailSafe();
  if (!email) return null;

  return getAllUsers().find(function (user) {
    return user.email === email;
  }) || null;
}

function getCurrentAdminUserRecord() {
  const user = getCurrentUserRecord();
  if (!user || !user.active) return null;
  return user.rol === ROLE.ADMIN ? user : null;
}

function requireLoggedInUser(sessionId) {
  const sessionUser = sessionId ? getUserBySessionId(sessionId) : null;
  if (sessionUser) return sessionUser;

  const adminUser = getCurrentAdminUserRecord();
  if (adminUser) return adminUser;

  throw new Error('Geen geldige sessie. Log opnieuw in.');
}

function createSessionForUser(user) {
  return withSessionLock_(function () {
    const now = new Date();
    const expires = new Date(now.getTime() + getSessionDurationHours() * 60 * 60 * 1000);
    const sessionId = makeSessionId();

    appendObjects(TABS.SESSIONS, [{
      SessionID: sessionId,
      LoginEmail: getEffectiveLoginEmail(user),
      Naam: safeText(user && user.naam),
      Rol: safeText(user && user.rol),
      TechniekerCode: safeText(user && user.techniekerCode),
      AangemaaktOp: formatSessionDateTime_(now),
      VerlooptOp: formatSessionDateTime_(expires),
      Actief: 'Ja'
    }]);

    return {
      sessionId: sessionId,
      aangemaaktOp: formatSessionDateTime_(now),
      verlooptOp: formatSessionDateTime_(expires)
    };
  });
}

function invalidateSessionById(sessionId) {
  const sid = safeText(sessionId);
  if (!sid) return false;

  return withSessionLock_(function () {
    const sheet = getSheetOrThrow(TABS.SESSIONS);
    const values = sheet.getDataRange().getValues();
    if (!values || !values.length) return false;

    const headers = values[0].map(function (h) { return safeText(h); });
    const col = getColMap(headers);

    if (col['SessionID'] === undefined || col['Actief'] === undefined) return false;

    for (let i = 1; i < values.length; i++) {
      if (safeText(values[i][col['SessionID']]) !== sid) continue;

      const activeCellValue = safeText(values[i][col['Actief']]);
      if (activeCellValue === 'Nee') return true;

      sheet.getRange(i + 1, col['Actief'] + 1).setValue('Nee');
      return true;
    }

    return false;
  });
}

function syncOpenSessionsToNewLoginEmail(oldLoginEmail, newLoginEmail) {
  const oldValue = normalizeLoginEmail(oldLoginEmail);
  const newValue = normalizeLoginEmail(newLoginEmail);

  if (!oldValue || !newValue || oldValue === newValue) return;

  withSessionLock_(function () {
    const sheet = getSheetOrThrow(TABS.SESSIONS);
    const values = sheet.getDataRange().getValues();
    if (!values || !values.length) return;

    const headers = values[0].map(function (h) { return safeText(h); });
    const col = getColMap(headers);

    if (col['LoginEmail'] === undefined || col['Actief'] === undefined) return;

    for (let i = 1; i < values.length; i++) {
      const isActive = isTrue(values[i][col['Actief']]);
      if (!isActive) continue;

      const rowLoginEmail = normalizeLoginEmail(values[i][col['LoginEmail']]);
      if (rowLoginEmail !== oldValue) continue;

      sheet.getRange(i + 1, col['LoginEmail'] + 1).setValue(newValue);
    }
  });
}

function cleanupExpiredSessions() {
  return withSessionLock_(function () {
    const sheet = getSheetOrThrow(TABS.SESSIONS);
    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return 0;

    const headers = values[0].map(function (h) { return safeText(h); });
    const col = getColMap(headers);

    if (
      col['SessionID'] === undefined ||
      col['VerlooptOp'] === undefined ||
      col['Actief'] === undefined
    ) {
      return 0;
    }

    const now = new Date();
    let changed = 0;

    for (let i = 1; i < values.length; i++) {
      if (!isTrue(values[i][col['Actief']])) continue;

      const expires = new Date(values[i][col['VerlooptOp']]);
      if (isNaN(expires.getTime())) continue;
      if (expires.getTime() > now.getTime()) continue;

      sheet.getRange(i + 1, col['Actief'] + 1).setValue('Nee');
      changed++;
    }

    return changed;
  });
}

function buildAuthTemplateContext(user, sessionId) {
  return {
    currentSessionId: safeText(sessionId),
    currentUserName: safeText(user && user.naam),
    currentUserRole: safeText(user && user.rol),
    currentLoginEmail: getEffectiveLoginEmail(user),
    currentAuthMode: getAuthModeForUser(user)
  };
}
