/* =========================================================
   10_AuthService.gs
   Refactor: authentication core service
   Doel:
   - duidelijke loginflow voor code-login
   - admins blijven Google-auth only
   - mislukte pogingen apart loggen
   - eigen logininstellingen veilig aanpassen
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / helpers
   --------------------------------------------------------- */

function getUsersTab_() {
  return TABS.USERS || 'Gebruikers';
}

function getLoginFailuresTab_() {
  return TABS.LOGIN_FAILURES || 'LoginFailures';
}

function makeLoginFailureId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'LGF-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function mapAuthUser_(row) {
  var email = normalizeLoginEmail(
    row.LoginEmail || row.LoginMail || row.Login || row.Email || row.Mail || ''
  );

  return {
    code: safeText(row.Code || row.GebruikerCode || row.UserCode),
    naam: safeText(row.Naam || row.Name),
    rol: safeText(row.Rol || row.Role),
    actief: row.Actief === undefined ? true : isTrue(row.Actief),
    email: safeText(row.Email || row.Mail),
    loginEmail: email,
    loginCode: safeText(row.LoginCode || row.CodeLogin || row.Pin || ''),
  };
}

function getAllAuthUsers_() {
  return readObjectsSafe(getUsersTab_()).map(mapAuthUser_);
}

function getActiveAuthUsers_() {
  return getAllAuthUsers_().filter(function (item) {
    return item.actief;
  });
}

function getAuthUserByEffectiveLoginEmail_(loginEmail) {
  var target = normalizeLoginEmail(loginEmail);
  if (!target) return null;

  return getActiveAuthUsers_().find(function (item) {
    return normalizeLoginEmail(item.loginEmail || item.email) === target;
  }) || null;
}

function isAdminUser_(user) {
  return safeText(user && user.rol) === safeText(ROLE.ADMIN || 'Admin');
}

function isCodeLoginAllowedForUser_(user) {
  return !!user && !isAdminUser_(user);
}

function assertValidLoginEmail_(value) {
  var email = normalizeLoginEmail(value);
  if (!email) {
    throw new Error('Login e-mail is verplicht.');
  }
  if (!isValidEmail(email)) {
    throw new Error('Ongeldig e-mailadres.');
  }
  return email;
}

function assertValidLoginCode_(value) {
  var code = safeText(value);
  if (!code) {
    throw new Error('Logincode is verplicht.');
  }
  if (code.length < 4) {
    throw new Error('Logincode moet minstens 4 tekens bevatten.');
  }
  return code;
}

function writeLoginFailure_(payload) {
  payload = payload || {};

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  appendObjects(getLoginFailuresTab_(), [{
    LoginFailureID: makeLoginFailureId_(),
    LoginEmail: normalizeLoginEmail(payload.loginEmail),
    Reason: safeText(payload.reason),
    PayloadInfo: safeJson(payload.extra || {}),
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
  }]);
}

function ensureUserAdminSupport_() {
  return {
    hasUpdateUserLastLogin: typeof updateUserLastLogin === 'function',
    hasSyncOpenSessionsToNewLoginEmail: typeof syncOpenSessionsToNewLoginEmail === 'function',
  };
}

/* ---------------------------------------------------------
   Login / logout
   --------------------------------------------------------- */

function validateLoginPayload_(payload) {
  payload = payload || {};
  return {
    loginEmail: assertValidLoginEmail_(payload.loginEmail),
    loginCode: assertValidLoginCode_(payload.loginCode),
  };
}

function buildLoginSuccessResponse_(user, session) {
  return {
    ok: true,
    sessionId: safeText(session && (session.sessionId || session.SessionID || session.id)),
    user: {
      code: safeText(user.code),
      naam: safeText(user.naam),
      rol: safeText(user.rol),
      loginEmail: safeText(user.loginEmail),
    },
  };
}

function loginWithCode(payload) {
  var normalized = validateLoginPayload_(payload);
  var user = getAuthUserByEffectiveLoginEmail_(normalized.loginEmail);

  if (!user) {
    writeLoginFailure_({
      loginEmail: normalized.loginEmail,
      reason: 'USER_NOT_FOUND',
    });
    throw new Error('Onbekende login of foutieve code.');
  }

  if (!isCodeLoginAllowedForUser_(user)) {
    writeLoginFailure_({
      loginEmail: normalized.loginEmail,
      reason: 'ADMIN_GOOGLE_ONLY',
      extra: { rol: user.rol },
    });
    throw new Error('Admins melden aan via Google-account.');
  }

  if (safeText(user.loginCode) !== safeText(normalized.loginCode)) {
    writeLoginFailure_({
      loginEmail: normalized.loginEmail,
      reason: 'INVALID_CODE',
      extra: { userCode: user.code },
    });
    throw new Error('Onbekende login of foutieve code.');
  }

  if (typeof createSessionForUser !== 'function') {
    throw new Error('Session service ontbreekt. Werk eerst het sessionblok in.');
  }

  var session = createSessionForUser({
    userCode: user.code,
    userName: user.naam,
    userRole: user.rol,
    loginEmail: user.loginEmail,
  });

  var support = ensureUserAdminSupport_();
  if (support.hasUpdateUserLastLogin) {
    updateUserLastLogin(user.code);
  }

  writeAudit({
    actie: 'LOGIN_WITH_CODE',
    actor: {
      code: user.code,
      naam: user.naam,
      rol: user.rol,
      email: user.loginEmail,
    },
    documentType: 'Sessie',
    documentId: safeText(session && (session.sessionId || session.SessionID || session.id)),
    details: {
      loginEmail: user.loginEmail,
    },
  });

  return buildLoginSuccessResponse_(user, session);
}

function logoutBySession(payload) {
  payload = payload || {};

  var sessionId = safeText(payload.sessionId || payload.sid);
  if (!sessionId) {
    throw new Error('SessionId ontbreekt.');
  }

  if (typeof invalidateSessionById !== 'function') {
    throw new Error('Session service ontbreekt. Werk eerst het sessionblok in.');
  }

  invalidateSessionById(sessionId);

  writeAudit({
    actie: 'LOGOUT_BY_SESSION',
    actor: {
      naam: safeText(payload.actorName || ''),
      rol: safeText(payload.actorRole || ''),
      email: safeText(payload.actorEmail || ''),
    },
    documentType: 'Sessie',
    documentId: sessionId,
    details: {},
  });

  return {
    ok: true,
    sessionId: sessionId,
  };
}

/* ---------------------------------------------------------
   Eigen logininstellingen aanpassen
   --------------------------------------------------------- */

function getUserSheetTable_() {
  var table = getAllValues(getUsersTab_());
  if (!table.length) {
    throw new Error('Gebruikerstab is leeg of ongeldig.');
  }
  return table;
}

function getEditableUserRowByCode_(userCode) {
  var table = getUserSheetTable_();
  var headerRow = table[0];
  var dataRows = table.slice(1);

  for (var i = 0; i < dataRows.length; i += 1) {
    var obj = rowToObject(headerRow, dataRows[i]);
    var mapped = mapAuthUser_(obj);
    if (safeText(mapped.code) === safeText(userCode)) {
      return {
        headerRow: headerRow,
        rowIndex: i,
        obj: obj,
        mapped: mapped,
        rows: dataRows,
      };
    }
  }

  return null;
}

function assertLoginEmailUniqueForOtherUsers_(loginEmail, currentUserCode) {
  var target = normalizeLoginEmail(loginEmail);

  var duplicate = getActiveAuthUsers_().find(function (item) {
    return safeText(item.code) !== safeText(currentUserCode) &&
      normalizeLoginEmail(item.loginEmail || item.email) === target;
  });

  if (duplicate) {
    throw new Error('Deze login e-mail is al in gebruik.');
  }
}

function changeOwnLoginAccess(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!actor || !safeText(actor.code)) {
    throw new Error('Geen geldige gebruiker gevonden.');
  }

  if (isAdminUser_(actor)) {
    throw new Error('Admins passen hun login niet via code-login aan.');
  }

  var currentCode = assertValidLoginCode_(payload.currentLoginCode);
  var nextLoginEmail = assertValidLoginEmail_(payload.newLoginEmail || payload.loginEmail);
  var nextLoginCode = assertValidLoginCode_(payload.newLoginCode || payload.loginCode);

  var editable = getEditableUserRowByCode_(actor.code);
  if (!editable) {
    throw new Error('Gebruiker niet gevonden.');
  }

  if (safeText(editable.mapped.loginCode) !== currentCode) {
    throw new Error('Huidige logincode is fout.');
  }

  assertLoginEmailUniqueForOtherUsers_(nextLoginEmail, actor.code);

  editable.obj.LoginEmail = nextLoginEmail;
  editable.obj.LoginCode = nextLoginCode;

  var rebuiltRows = editable.rows.map(function (row, index) {
    if (index !== editable.rowIndex) {
      return row;
    }
    return buildRowFromHeaders(editable.headerRow, editable.obj);
  });

  writeFullTable(getUsersTab_(), editable.headerRow, rebuiltRows);

  var support = ensureUserAdminSupport_();
  if (support.hasSyncOpenSessionsToNewLoginEmail) {
    syncOpenSessionsToNewLoginEmail(actor.code, nextLoginEmail);
  }

  writeAudit({
    actie: 'CHANGE_OWN_LOGIN_ACCESS',
    actor: actor,
    documentType: 'Gebruiker',
    documentId: actor.code,
    details: {
      loginEmail: nextLoginEmail,
    },
  });

  return {
    ok: true,
    userCode: actor.code,
    loginEmail: nextLoginEmail,
  };
}