/* =========================================================
   13_UserAdminService.gs
   Refactor: user admin + login management
   Doel:
   - centrale user directory
   - lookup op login-email en contact-email
   - laatste login bijwerken
   - admin loginbeheer
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / mapping
   --------------------------------------------------------- */

function getUserAdminUsersTab_() {
  return TABS.USERS || 'Gebruikers';
}

function getUserAdminLoginFailuresTab_() {
  return TABS.LOGIN_FAILURES || 'LoginFailures';
}

function mapUserAdminUser_(row) {
  return {
    code: safeText(row.Code || row.GebruikerCode || row.UserCode),
    naam: safeText(row.Naam || row.Name),
    rol: safeText(row.Rol || row.Role),
    actief: row.Actief === undefined ? true : isTrue(row.Actief),
    email: normalizeLoginEmail(row.Email || row.Mail || ''),
    loginEmail: normalizeLoginEmail(
      row.LoginEmail || row.LoginMail || row.Login || row.Email || row.Mail || ''
    ),
    loginCode: safeText(row.LoginCode || row.CodeLogin || row.Pin || ''),
    laatsteLoginOp: safeText(row.LaatsteLoginOp || row.LastLoginAt || ''),
    mobileWarehouseCode: safeText(row.MobileWarehouseCode || ''),
    technicianCode: safeText(row.TechniekerCode || row.TechnicianCode || row.Code || ''),
  };
}

function mapLoginFailureAdmin_(row) {
  return {
    loginFailureId: safeText(row.LoginFailureID || row.LoginFailureId || row.ID),
    loginEmail: normalizeLoginEmail(row.LoginEmail || ''),
    reason: safeText(row.Reason || row.Reden),
    payloadInfo: safeText(row.PayloadInfo || row.ExtraJson || ''),
    aangemaaktOp: safeText(row.AangemaaktOp || row.CreatedAt || ''),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.CreatedAtRaw || row.AangemaaktOp || row.CreatedAt || ''),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllUsers() {
  return readObjectsSafe(getUserAdminUsersTab_())
    .map(mapUserAdminUser_)
    .sort(function (a, b) {
      return (
        safeText(a.naam).localeCompare(safeText(b.naam)) ||
        safeText(a.code).localeCompare(safeText(b.code))
      );
    });
}

function getActiveUsers_() {
  return getAllUsers().filter(function (item) {
    return item.actief;
  });
}

function getAllNonAdminUsers() {
  return getActiveUsers_().filter(function (item) {
    return safeText(item.rol) !== safeText(ROLE.ADMIN || 'Admin');
  });
}

function getUserByCode(userCode) {
  var code = safeText(userCode);
  if (!code) return null;

  return getAllUsers().find(function (item) {
    return safeText(item.code) === code;
  }) || null;
}

function getUserByEffectiveLoginEmail(loginEmail) {
  var target = normalizeLoginEmail(loginEmail);
  if (!target) return null;

  return getActiveUsers_().find(function (item) {
    return normalizeLoginEmail(item.loginEmail || item.email) === target;
  }) || null;
}

function getUserByContactEmail(email) {
  var target = normalizeLoginEmail(email);
  if (!target) return null;

  return getActiveUsers_().find(function (item) {
    return normalizeLoginEmail(item.email) === target;
  }) || null;
}

function getRecentLoginFailures(limit) {
  var max = safeNumber(limit, 50);
  if (max <= 0) max = 50;

  return readObjectsSafe(getUserAdminLoginFailuresTab_())
    .map(mapLoginFailureAdmin_)
    .sort(function (a, b) {
      return (
        safeText(b.aangemaaktOpRaw).localeCompare(safeText(a.aangemaaktOpRaw)) ||
        safeText(b.loginFailureId).localeCompare(safeText(a.loginFailureId))
      );
    })
    .slice(0, max);
}

/* ---------------------------------------------------------
   User row editing helpers
   --------------------------------------------------------- */

function getEditableUserAdminRowByCode_(userCode) {
  var table = getAllValues(getUserAdminUsersTab_());
  if (!table.length) {
    throw new Error('Gebruikerstab is leeg of ongeldig.');
  }

  var headerRow = table[0];
  var dataRows = table.slice(1);

  for (var i = 0; i < dataRows.length; i += 1) {
    var obj = rowToObject(headerRow, dataRows[i]);
    var mapped = mapUserAdminUser_(obj);

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

function writeBackEditedUserRow_(editable) {
  var rebuiltRows = editable.rows.map(function (row, index) {
    if (index !== editable.rowIndex) {
      return row;
    }
    return buildRowFromHeaders(editable.headerRow, editable.obj);
  });

  writeFullTable(getUserAdminUsersTab_(), editable.headerRow, rebuiltRows);
}

function assertLoginEmailUniqueForOtherUsersAdmin_(loginEmail, currentUserCode) {
  var target = normalizeLoginEmail(loginEmail);

  var duplicate = getActiveUsers_().find(function (item) {
    return safeText(item.code) !== safeText(currentUserCode) &&
      normalizeLoginEmail(item.loginEmail || item.email) === target;
  });

  if (duplicate) {
    throw new Error('Deze login e-mail is al in gebruik.');
  }
}

function assertUserExistsAndIsNotAdmin_(userCode) {
  var user = getUserByCode(userCode);
  if (!user) {
    throw new Error('Gebruiker niet gevonden.');
  }
  if (safeText(user.rol) === safeText(ROLE.ADMIN || 'Admin')) {
    throw new Error('Admin loginbeheer verloopt via Google-auth.');
  }
  return user;
}

/* ---------------------------------------------------------
   Last login
   --------------------------------------------------------- */

function updateUserLastLogin(userCode) {
  var editable = getEditableUserAdminRowByCode_(userCode);
  if (!editable) {
    return {
      updated: false,
      userCode: safeText(userCode),
    };
  }

  editable.obj.LaatsteLoginOp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  writeBackEditedUserRow_(editable);

  return {
    updated: true,
    userCode: safeText(userCode),
  };
}

/* ---------------------------------------------------------
   Self-service wrapper
   --------------------------------------------------------- */

function updateOwnLoginSettings(payload) {
  if (typeof changeOwnLoginAccess !== 'function') {
    throw new Error('Auth service ontbreekt. Werk eerst het authblok in.');
  }
  return changeOwnLoginAccess(payload);
}

/* ---------------------------------------------------------
   Admin read model
   --------------------------------------------------------- */

function adminGetLoginAdminData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.MANAGER, ROLE.ADMIN])) {
    throw new Error('Geen rechten voor loginbeheer.');
  }

  return {
    users: getAllNonAdminUsers(),
    loginFailures: getRecentLoginFailures(payload.limit || 100),
  };
}

/* ---------------------------------------------------------
   Admin update login settings
   --------------------------------------------------------- */

function adminUpdateUserLoginSettings(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.MANAGER, ROLE.ADMIN])) {
    throw new Error('Geen rechten voor loginbeheer.');
  }

  var userCode = safeText(payload.userCode || payload.targetUserCode);
  if (!userCode) {
    throw new Error('UserCode ontbreekt.');
  }

  var targetUser = assertUserExistsAndIsNotAdmin_(userCode);
  var nextLoginEmail = assertValidLoginEmail_(payload.newLoginEmail || payload.loginEmail);
  var nextLoginCode = assertValidLoginCode_(payload.newLoginCode || payload.loginCode);

  assertLoginEmailUniqueForOtherUsersAdmin_(nextLoginEmail, targetUser.code);

  var editable = getEditableUserAdminRowByCode_(targetUser.code);
  if (!editable) {
    throw new Error('Gebruiker niet gevonden.');
  }

  editable.obj.LoginEmail = nextLoginEmail;
  editable.obj.LoginCode = nextLoginCode;
  writeBackEditedUserRow_(editable);

  if (typeof syncOpenSessionsToNewLoginEmail === 'function') {
    syncOpenSessionsToNewLoginEmail(targetUser.code, nextLoginEmail);
  }

  writeAudit({
    actie: 'ADMIN_UPDATE_USER_LOGIN_SETTINGS',
    actor: actor,
    documentType: 'Gebruiker',
    documentId: targetUser.code,
    details: {
      loginEmail: nextLoginEmail,
    },
  });

  return {
    ok: true,
    userCode: targetUser.code,
    naam: targetUser.naam,
    loginEmail: nextLoginEmail,
  };
}

/* ---------------------------------------------------------
   Convenience exports for admin screens
   --------------------------------------------------------- */

function adminSearchUsers(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.MANAGER, ROLE.ADMIN])) {
    throw new Error('Geen rechten voor gebruikersbeheer.');
  }

  var query = safeText(payload.query).toLowerCase();

  return getAllNonAdminUsers().filter(function (item) {
    if (!query) return true;

    return [
      item.code,
      item.naam,
      item.rol,
      item.email,
      item.loginEmail,
      item.mobileWarehouseCode,
      item.technicianCode,
    ].some(function (value) {
      return safeText(value).toLowerCase().indexOf(query) >= 0;
    });
  });
}
