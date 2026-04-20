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

function getAllSessions() {
  return readObjectsSafe(TABS.SESSIONS).map(mapSession);
}

function getActiveSessions() {
  const now = new Date();

  return getAllSessions().filter(session => {
    if (!session.actief) return false;

    const expires = new Date(session.verlooptOp);
    if (isNaN(expires)) return false;

    return expires.getTime() > now.getTime();
  });
}

function getUserBySessionId(sessionId) {
  const sid = safeText(sessionId);
  if (!sid) return null;

  const session = getActiveSessions().find(x => x.sessionId === sid);
  if (!session) return null;

  const users = getAllUsers();
  return users.find(x => getEffectiveLoginEmail(x) === session.loginEmail) || null;
}

function getCurrentUserRecord() {
  const email = getCurrentSessionEmailSafe();
  if (!email) return null;

  return getAllUsers().find(user => user.email === email) || null;
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
  const now = new Date();
  const expires = new Date(now.getTime() + getSessionDurationHours() * 60 * 60 * 1000);
  const sessionId = makeSessionId();

  appendObjects(TABS.SESSIONS, [{
    SessionID: sessionId,
    LoginEmail: getEffectiveLoginEmail(user),
    Naam: user.naam,
    Rol: user.rol,
    TechniekerCode: user.techniekerCode,
    AangemaaktOp: Utilities.formatDate(now, getAppTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    VerlooptOp: Utilities.formatDate(expires, getAppTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    Actief: 'Ja'
  }]);

  return {
    sessionId,
    aangemaaktOp: Utilities.formatDate(now, getAppTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    verlooptOp: Utilities.formatDate(expires, getAppTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  };
}

function invalidateSessionById(sessionId) {
  const sid = safeText(sessionId);
  if (!sid) return false;

  const sheet = getSheetOrThrow(TABS.SESSIONS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) return false;

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);
  if (col['SessionID'] === undefined || col['Actief'] === undefined) return false;

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['SessionID']]) !== sid) continue;
    values[i][col['Actief']] = 'Nee';
    updated = true;
    break;
  }

  if (updated && values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return updated;
}

function syncOpenSessionsToNewLoginEmail(oldLoginEmail, newLoginEmail) {
  const oldValue = normalizeLoginEmail(oldLoginEmail);
  const newValue = normalizeLoginEmail(newLoginEmail);

  if (!oldValue || !newValue || oldValue === newValue) return;

  const sheet = getSheetOrThrow(TABS.SESSIONS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) return;

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  if (col['LoginEmail'] === undefined || col['Actief'] === undefined) return;

  let changed = false;

  for (let i = 1; i < values.length; i++) {
    if (!isTrue(values[i][col['Actief']])) continue;
    if (normalizeLoginEmail(values[i][col['LoginEmail']]) !== oldValue) continue;

    values[i][col['LoginEmail']] = newValue;
    changed = true;
  }

  if (changed && values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }
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