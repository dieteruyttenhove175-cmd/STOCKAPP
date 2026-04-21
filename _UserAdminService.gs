/* =========================================================
   13_UserAdminService.gs — gebruikersbeheer / admin loginbeheer
   ========================================================= */

function getAllUsers() {
  return readObjectsSafe(TABS.USERS)
    .map(mapUser)
    .filter(user => user.active && user.email);
}

function getAllNonAdminUsers() {
  return getAllUsers().filter(user => user.rol !== ROLE.ADMIN);
}

function getUsersByRole(role) {
  return getAllUsers().filter(user => user.rol === role);
}

function findUserByEffectiveLoginEmail(loginEmail) {
  const normalized = normalizeLoginEmail(loginEmail);
  if (!normalized) return null;

  return getAllUsers().find(user => getEffectiveLoginEmail(user) === normalized) || null;
}

function findUserByContactEmail(email) {
  const normalized = normalizeLoginEmail(email);
  if (!normalized) return null;

  return getAllUsers().find(user => normalizeLoginEmail(user.email) === normalized) || null;
}

function updateUserLastLogin(user) {
  if (!user) return;

  const sheet = getSheetOrThrow(TABS.USERS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) return;

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    const rowUser = mapUser({
      Email: values[i][col['Email']],
      LoginEmail: col['LoginEmail'] !== undefined ? values[i][col['LoginEmail']] : ''
    });

    if (getEffectiveLoginEmail(rowUser) !== getEffectiveLoginEmail(user)) continue;

    if (col['LaatsteLoginOp'] !== undefined) {
      values[i][col['LaatsteLoginOp']] = nowStamp();
    }

    updated = true;
    break;
  }

  if (updated && values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }
}

function updateOwnLoginSettings(user, newLoginEmail, newLoginCode) {
  const oldLoginEmail = getEffectiveLoginEmail(user);

  const sheet = getSheetOrThrow(TABS.USERS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab Gebruikers is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    const rowUser = mapUser({
      Email: values[i][col['Email']],
      LoginEmail: col['LoginEmail'] !== undefined ? values[i][col['LoginEmail']] : ''
    });

    if (getEffectiveLoginEmail(rowUser) !== oldLoginEmail) continue;

    if (col['LoginEmail'] !== undefined) values[i][col['LoginEmail']] = newLoginEmail;
    if (col['LoginCode'] !== undefined) values[i][col['LoginCode']] = newLoginCode;
    if (col['CodeGewijzigdOp'] !== undefined) values[i][col['CodeGewijzigdOp']] = nowStamp();

    updated = true;
    break;
  }

  if (!updated) {
    throw new Error('Gebruiker niet gevonden in tab Gebruikers.');
  }

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  syncOpenSessionsToNewLoginEmail(oldLoginEmail, newLoginEmail);
}

function adminGetLoginAdminData(payload) {
  const sessionId = getPayloadSessionId(payload);
  assertRoleAllowed([ROLE.ADMIN], sessionId);

  const users = getAllNonAdminUsers()
    .map(user => ({
      naam: user.naam,
      rol: user.rol,
      techniekerCode: user.techniekerCode,
      contactEmail: user.email,
      loginEmail: getEffectiveLoginEmail(user),
      loginCode: user.loginCode || '',
      codeGewijzigdOp: user.codeGewijzigdOp || '',
      laatsteLoginOp: user.laatsteLoginOp || ''
    }))
    .sort((a, b) => String(a.naam || '').localeCompare(String(b.naam || '')));

  const failures = readObjectsSafe(TABS.LOGIN_FAILURES)
    .map(mapLoginFailure)
    .sort((a, b) => String(b.tijdstipRaw || '').localeCompare(String(a.tijdstipRaw || '')));

  return { users, failures };
}

function adminUpdateUserLoginSettings(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const admin = assertRoleAllowed([ROLE.ADMIN], sessionId);

  const targetEmail = normalizeLoginEmail(payload.targetEmail);
  const currentLoginEmail = normalizeLoginEmail(payload.currentLoginEmail);
  const newLoginEmail = normalizeLoginEmail(payload.newLoginEmail);
  const newLoginCode = safeText(payload.newLoginCode);

  if (!targetEmail && !currentLoginEmail) {
    throw new Error('Doelgebruiker ontbreekt.');
  }

  if (!newLoginEmail) {
    throw new Error('Nieuwe login-email ontbreekt.');
  }

  if (!isLikelyEmail(newLoginEmail)) {
    throw new Error('Nieuwe login-email is ongeldig.');
  }

  if (!newLoginCode || newLoginCode.length < 4) {
    throw new Error('Nieuwe code moet minstens 4 tekens hebben.');
  }

  const users = getAllUsers();

  const duplicate = users.find(user =>
    user.active &&
    user.rol !== ROLE.ADMIN &&
    getEffectiveLoginEmail(user) === newLoginEmail &&
    !(
      (targetEmail && normalizeLoginEmail(user.email) === targetEmail) ||
      (currentLoginEmail && getEffectiveLoginEmail(user) === currentLoginEmail)
    )
  );

  if (duplicate) {
    throw new Error('Deze login-email is al in gebruik.');
  }

  const sheet = getSheetOrThrow(TABS.USERS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab Gebruikers is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;
  let oldLoginEmail = '';
  let targetName = '';

  for (let i = 1; i < values.length; i++) {
    const rowContactEmail = normalizeLoginEmail(values[i][col['Email']]);
    const rowLoginEmail = normalizeLoginEmail(
      col['LoginEmail'] !== undefined ? values[i][col['LoginEmail']] : rowContactEmail
    );
    const rowRole = safeText(values[i][col['Rol']]);

    const matches =
      (targetEmail && rowContactEmail === targetEmail) ||
      (currentLoginEmail && rowLoginEmail === currentLoginEmail);

    if (!matches) continue;

    if (rowRole === ROLE.ADMIN) {
      throw new Error('Admin gebruikt Google-authenticatie en heeft hier geen login-code.');
    }

    oldLoginEmail = rowLoginEmail;
    targetName = safeText(values[i][col['Naam']]);

    if (col['LoginEmail'] !== undefined) values[i][col['LoginEmail']] = newLoginEmail;
    if (col['LoginCode'] !== undefined) values[i][col['LoginCode']] = newLoginCode;
    if (col['CodeGewijzigdOp'] !== undefined) values[i][col['CodeGewijzigdOp']] = nowStamp();

    updated = true;
    break;
  }

  if (!updated) {
    throw new Error('Gebruiker niet gevonden.');
  }

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  syncOpenSessionsToNewLoginEmail(oldLoginEmail, newLoginEmail);

  safeWriteAuthAudit(
    'Logingegevens aangepast door admin',
    admin.rol,
    admin.naam || admin.email,
    'Gebruiker',
    targetEmail || currentLoginEmail,
    {
      targetName: targetName,
      newLoginEmail: newLoginEmail
    }
  );

  return {
    success: true,
    message: 'Logingegevens aangepast.'
  };
}