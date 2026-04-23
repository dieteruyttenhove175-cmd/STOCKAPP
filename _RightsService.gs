/* =========================================================
   12_RightsService.gs
   Refactor: rights + access policy service
   Doel:
   - centrale rolchecks
   - duidelijke page/domain access helpers
   - technieker-toegang apart en herbruikbaar
   ========================================================= */

/* ---------------------------------------------------------
   Generic role helpers
   --------------------------------------------------------- */

function getUserRole_(user) {
  return safeText(user && (user.rol || user.userRole || user.Role));
}

function isRole_(user, roleName) {
  return getUserRole_(user) === safeText(roleName);
}

function isAdmin_(user) {
  return isRole_(user, ROLE.ADMIN || 'Admin');
}

function roleAllowed(user, allowedRoles) {
  if (!user) return false;
  if (isAdmin_(user)) return true;

  var roles = Array.isArray(allowedRoles) ? allowedRoles : [];
  var currentRole = getUserRole_(user);

  return roles.some(function (roleName) {
    return currentRole === safeText(roleName);
  });
}

function assertRoleAllowed(user, allowedRoles, message) {
  if (!roleAllowed(user, allowedRoles)) {
    throw new Error(message || 'Geen rechten voor deze actie.');
  }
  return true;
}

/* ---------------------------------------------------------
   Generic access assertions by domain
   --------------------------------------------------------- */

function assertWarehouseAccess(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.WAREHOUSE, ROLE.MANAGER],
    'Geen rechten voor magazijn.'
  );
  return user;
}

function assertMobileWarehouseRoleAccess(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER],
    'Geen rechten voor mobiel magazijn.'
  );
  return user;
}

function assertManagerAccess(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.MANAGER],
    'Geen rechten voor manager.'
  );
  return user;
}

function assertWarehouseOrMobileWarehouseAccess(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER],
    'Geen rechten voor magazijn of mobiel magazijn.'
  );
  return user;
}

function assertAuthenticatedAccess(sessionId) {
  return requireLoggedInUser(sessionId);
}

/* ---------------------------------------------------------
   Technician directory
   --------------------------------------------------------- */

function mapRightsTechnician_(row) {
  return {
    code: safeText(row.Code || row.TechniekerCode || row.Ref),
    naam: safeText(row.Naam || row.Name),
    ref: safeText(row.Ref || row.Code || row.TechniekerCode),
    actief: row.Actief === undefined ? true : isTrue(row.Actief),
    email: normalizeLoginEmail(row.Email || row.Mail || ''),
    mobileWarehouseCode: safeText(row.MobileWarehouseCode || ''),
  };
}

function getAllActiveTechniciansForRights_() {
  return readObjectsSafe(TABS.TECHNICIANS || 'Techniekers')
    .map(mapRightsTechnician_)
    .filter(function (item) {
      return item.actief;
    });
}

function resolveTechnicianByRef(techRef) {
  var target = safeText(techRef);
  if (!target) return null;

  var all = getAllActiveTechniciansForRights_();
  return all.find(function (item) {
    return (
      safeText(item.ref) === target ||
      safeText(item.code) === target ||
      safeText(item.naam) === target
    );
  }) || null;
}

function resolveActorTechnicianRef_(user) {
  return safeText(
    user &&
    (
      user.techniekerCode ||
      user.technicianCode ||
      user.code ||
      user.ref
    )
  );
}

/* ---------------------------------------------------------
   Technician access policy
   --------------------------------------------------------- */

function canUserAccessTechnicianRef(user, techRef) {
  var target = resolveTechnicianByRef(techRef);
  if (!target) return false;

  if (isAdmin_(user)) return true;
  if (roleAllowed(user, [ROLE.MANAGER, ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE])) return true;

  if (roleAllowed(user, [ROLE.TECHNICIAN])) {
    var ownRef = resolveActorTechnicianRef_(user);
    return !!ownRef && ownRef === safeText(target.ref);
  }

  return false;
}

function assertTechnicianAccessToRef(sessionId, techRef) {
  var user = requireLoggedInUser(sessionId);
  var target = resolveTechnicianByRef(techRef);

  if (!target) {
    throw new Error('Technieker niet gevonden.');
  }

  if (!canUserAccessTechnicianRef(user, techRef)) {
    throw new Error('Geen toegang tot deze technieker.');
  }

  return {
    user: user,
    technician: target,
  };
}

function assertTechnicianSelfAccess(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.TECHNICIAN],
    'Geen rechten voor technieker.'
  );

  var ownRef = resolveActorTechnicianRef_(user);
  if (!ownRef) {
    throw new Error('Geen gekoppelde techniekerref gevonden.');
  }

  var technician = resolveTechnicianByRef(ownRef);
  if (!technician) {
    throw new Error('Gekoppelde technieker niet gevonden.');
  }

  return {
    user: user,
    technician: technician,
  };
}

/* ---------------------------------------------------------
   Page access policy
   --------------------------------------------------------- */

function canAccessPage(user, pageName, techRef) {
  var page = safeText(pageName);

  if (!user) return false;
  if (isAdmin_(user)) return true;

  if (page === 'Warehouse') {
    return roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER]);
  }

  if (page === 'Manager') {
    return roleAllowed(user, [ROLE.MANAGER]);
  }

  if (page === 'MobileWarehouse') {
    return roleAllowed(user, [ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER]);
  }

  if (page === 'Index') {
    if (roleAllowed(user, [ROLE.MANAGER, ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE])) {
      return true;
    }
    if (roleAllowed(user, [ROLE.TECHNICIAN])) {
      return canUserAccessTechnicianRef(user, techRef || resolveActorTechnicianRef_(user));
    }
    return false;
  }

  if (page === 'Login') {
    return true;
  }

  return false;
}

function assertPageAccess(user, pageName, techRef) {
  if (!canAccessPage(user, pageName, techRef)) {
    throw new Error('Geen toegang tot pagina ' + safeText(pageName) + '.');
  }

  if (safeText(pageName) === 'Index' && roleAllowed(user, [ROLE.TECHNICIAN])) {
    var targetRef = safeText(techRef || resolveActorTechnicianRef_(user));
    if (!targetRef) {
      throw new Error('Geen techniekerreferentie beschikbaar.');
    }
    assertTechnicianAccessToRef(user.sessionId || '', targetRef);
  }

  return true;
}

/* ---------------------------------------------------------
   Convenience policy helpers for services
   --------------------------------------------------------- */

function canApproveReceipts(user) {
  return roleAllowed(user, [ROLE.MANAGER]);
}

function canApproveReturns(user) {
  return roleAllowed(user, [ROLE.MANAGER]);
}

function canBookNeedIssues(user) {
  return roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.MOBILE_WAREHOUSE]);
}

function canCreateTransfers(user) {
  return roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER]);
}

function canApproveTransfers(user) {
  return roleAllowed(user, [ROLE.MANAGER]);
}

function canViewCentralStock(user) {
  return roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.MOBILE_WAREHOUSE]);
}
