/* =========================================================
   10_AuthService.gs — login / logout / eigen login wijzigen
   ========================================================= */

function makeLoginFailureId() {
  return makeStampedId('LF');
}

function safeWriteAuthAudit(action, role, actor, documentType, documentId, details) {
  if (typeof writeAudit === 'function') {
    writeAudit(action, role, actor, documentType, documentId, details);
  }
}

function writeLoginFailure(loginEmail, enteredCode, reason, matchedUser) {
  appendObjects(TABS.LOGIN_FAILURES, [{
    FoutID: makeLoginFailureId(),
    Tijdstip: nowStamp(),
    LoginEmail: normalizeLoginEmail(loginEmail),
    IngevoerdeCode: safeText(enteredCode),
    Reden: safeText(reason),
    MatchGebruiker: safeText(matchedUser)
  }]);
}

function validateLoginPayload(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const loginEmail = normalizeLoginEmail(payload.loginEmail);
  const loginCode = safeText(payload.loginCode);

  if (!loginEmail) throw new Error('Login-email ontbreekt.');
  if (!loginCode) throw new Error('Code ontbreekt.');

  return { loginEmail, loginCode };
}

function loginWithCode(payload) {
  const parsed = validateLoginPayload(payload);
  const loginEmail = parsed.loginEmail;
  const loginCode = parsed.loginCode;

  const users = getAllUsers();
  const matchedByEmail = users.find(user => getEffectiveLoginEmail(user) === loginEmail);

  if (!matchedByEmail) {
    writeLoginFailure(loginEmail, loginCode, 'Onbekende login-email', '');
    throw new Error('Ongeldige login-email of code.');
  }

  if (matchedByEmail.rol === ROLE.ADMIN) {
    writeLoginFailure(
      loginEmail,
      loginCode,
      'Admin probeerde code-login',
      matchedByEmail.naam || matchedByEmail.email
    );
    throw new Error('Admin meldt aan via Google-account, niet via code.');
  }

  if (safeText(matchedByEmail.loginCode) !== loginCode) {
    writeLoginFailure(
      loginEmail,
      loginCode,
      'Foute code',
      matchedByEmail.naam || matchedByEmail.email
    );
    throw new Error('Ongeldige login-email of code.');
  }

  const session = createSessionForUser(matchedByEmail);
  updateUserLastLogin(matchedByEmail);

  safeWriteAuthAudit(
    'Login',
    matchedByEmail.rol,
    matchedByEmail.naam || getEffectiveLoginEmail(matchedByEmail),
    'Sessie',
    session.sessionId,
    { loginEmail: getEffectiveLoginEmail(matchedByEmail) }
  );

  return {
    success: true,
    sessionId: session.sessionId,
    authMode: 'custom_session',
    user: {
      naam: matchedByEmail.naam,
      rol: matchedByEmail.rol,
      techniekerCode: matchedByEmail.techniekerCode,
      loginEmail: getEffectiveLoginEmail(matchedByEmail)
    }
  };
}

function logoutBySession(payload) {
  const sessionId = getPayloadSessionId(payload);
  if (!sessionId) return { success: true };

  const changed = invalidateSessionById(sessionId);

  safeWriteAuthAudit(
    'Logout',
    'Sessie',
    sessionId,
    'Sessie',
    sessionId,
    { invalidated: changed }
  );

  return { success: true };
}

function changeOwnLoginAccess(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const currentCode = safeText(payload.currentCode);
  const newLoginEmailRaw = safeText(payload.newLoginEmail);
  const newCodeRaw = safeText(payload.newCode);

  const user = requireLoggedInUser(sessionId);

  if (user.rol === ROLE.ADMIN) {
    throw new Error('Admin gebruikt Google-authenticatie en wijzigt hier geen code.');
  }

  if (!currentCode) {
    throw new Error('Huidige code ontbreekt.');
  }

  if (safeText(user.loginCode) !== currentCode) {
    throw new Error('Huidige code is fout.');
  }

  const oldLoginEmail = getEffectiveLoginEmail(user);
  const newLoginEmail = normalizeLoginEmail(newLoginEmailRaw || oldLoginEmail);
  const newCode = newCodeRaw || safeText(user.loginCode);

  if (!newLoginEmail) {
    throw new Error('Nieuwe login-email ontbreekt.');
  }

  if (!isLikelyEmail(newLoginEmail)) {
    throw new Error('Nieuwe login-email is ongeldig.');
  }

  if (!newCode || newCode.length < 4) {
    throw new Error('Nieuwe code moet minstens 4 tekens hebben.');
  }

  const duplicate = getAllUsers().find(other =>
    other.active &&
    other.rol !== ROLE.ADMIN &&
    getEffectiveLoginEmail(other) === newLoginEmail &&
    getEffectiveLoginEmail(other) !== oldLoginEmail
  );

  if (duplicate) {
    throw new Error('Deze login-email is al in gebruik.');
  }

  updateOwnLoginSettings(user, newLoginEmail, newCode);

  safeWriteAuthAudit(
    'Logingegevens gewijzigd',
    user.rol,
    user.naam || oldLoginEmail,
    'Gebruiker',
    oldLoginEmail,
    {
      newLoginEmail: newLoginEmail,
      codeChanged: newCode !== safeText(user.loginCode)
    }
  );

  return {
    success: true,
    loginEmail: newLoginEmail,
    message: 'Logingegevens aangepast.'
  };
}