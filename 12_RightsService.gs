/* =========================================================
   12_RightsService.gs — rechten / rolchecks / techniekerchecks
   ========================================================= */

function roleAllowed(user, allowedRoles) {
  if (!user) return false;
  if (user.rol === ROLE.ADMIN) return true;
  return (allowedRoles || []).includes(user.rol);
}

function assertRoleAllowed(allowedRoles, sessionId) {
  const user = requireLoggedInUser(sessionId);
  if (!roleAllowed(user, allowedRoles)) {
    throw new Error('Geen rechten voor deze actie.');
  }
  return user;
}

function assertWarehouseAccess(sessionId) {
  return assertRoleAllowed([ROLE.WAREHOUSE, ROLE.MANAGER], sessionId);
}

function assertMobileWarehouseAccess(sessionId) {
  return assertRoleAllowed([ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER], sessionId);
}

function assertWarehouseOrMobileAccess(sessionId) {
  return assertRoleAllowed([ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER], sessionId);
}

function assertManagerAccess(sessionId) {
  return assertRoleAllowed([ROLE.MANAGER], sessionId);
}

function assertManagerOrAnalysisAccess(sessionId) {
  return assertRoleAllowed([ROLE.MANAGER, ROLE.ANALYSIS], sessionId);
}

function getActiveTechnicians() {
  return readObjectsSafe(TABS.TECHNICIANS)
    .map(mapTechnician)
    .filter(x => x.active);
}

function getTechnicianByCode(code) {
  return getActiveTechnicians().find(x => normalizeRef(x.code) === normalizeRef(code)) || null;
}

function getTechnicianNameByCode(code) {
  const found = getTechnicianByCode(code);
  return found ? found.naam : safeText(code);
}

function findTechnicianByRef(technicians, techRef) {
  const ref = normalizeRef(techRef);
  if (!ref) return null;

  return (technicians || []).find(t => {
    const code = normalizeRef(t.code);
    const name = normalizeRef(t.naam);
    return [
      code,
      name,
      normalizeRef(code + '-' + name),
      normalizeRef(name + '-' + code)
    ].includes(ref);
  }) || null;
}

function resolveTechnicianByRef(techRef) {
  const technicians = readObjectsSafe(TABS.TECHNICIANS).map(mapTechnician);
  const technician = findTechnicianByRef(technicians, techRef);

  if (!technician || !technician.active) {
    throw new Error('Technieker niet gevonden of niet actief.');
  }

  return technician;
}

function assertTechnicianAccessToRef(techRef, sessionId) {
  const user = requireLoggedInUser(sessionId);
  const technician = resolveTechnicianByRef(techRef);

  if (user.rol === ROLE.ADMIN) {
    return { user, technician, readOnly: false };
  }

  if (
    user.rol === ROLE.TECHNICIAN &&
    normalizeRef(user.techniekerCode) === normalizeRef(technician.code)
  ) {
    return { user, technician, readOnly: false };
  }

  throw new Error('Geen rechten voor deze techniekerdata.');
}

function deliveryMatchesTechnician(delivery, technician) {
  const deliveryCode = normalizeRef(delivery.techniekerCode);
  const deliveryName = normalizeRef(delivery.technieker);
  const techCode = normalizeRef(technician.code);
  const techName = normalizeRef(technician.naam);

  return (deliveryCode && deliveryCode === techCode) || (deliveryName && deliveryName === techName);
}