/* =========================================================
   15_WebApp.gs — webapp routing / template bootstrap
   ========================================================= */

function getForbiddenOutput(message) {
  return HtmlService
    .createHtmlOutput(`
## Geen toegang

${String(message || 'Je hebt geen toegang tot deze pagina.')}
`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function evaluateTemplateWithData(fileName, templateData, pageTitle) {
  const template = HtmlService.createTemplateFromFile(fileName);

  Object.keys(templateData || {}).forEach(function (key) {
    template[key] = templateData[key];
  });

  return template
    .evaluate()
    .setTitle(pageTitle || APP_CONFIG.DEFAULT_PAGE_TITLE || 'DigiQS Warehouse')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function evaluateTemplateWithFallback(primaryFileName, fallbackFileName, templateData, pageTitle) {
  try {
    return evaluateTemplateWithData(primaryFileName, templateData, pageTitle);
  } catch (e) {
    if (!fallbackFileName) throw e;
    return evaluateTemplateWithData(fallbackFileName, templateData, pageTitle);
  }
}

function buildLoginTemplateContext(view, techRef, sid) {
  return {
    targetView: safeText(view || ''),
    targetTechRef: safeText(techRef || ''),
    loginError: sid ? 'Sessie verlopen. Log opnieuw in.' : ''
  };
}

function getRequestedView(e) {
  return safeText(e && e.parameter && e.parameter.view).toLowerCase();
}

function getRequestedTechRef(e) {
  return safeText(e && e.parameter && e.parameter.tech);
}

function getRequestedSessionId(e) {
  return safeText(e && e.parameter && e.parameter.sid);
}

function resolveLoggedInUserFromRequest(sessionId) {
  const adminUser = getCurrentAdminUserRecord();
  const sessionUser = sessionId ? getUserBySessionId(sessionId) : null;
  return adminUser || sessionUser || null;
}

function resolveTargetFileName(view, user) {
  const requestedView = safeText(view).toLowerCase();

  if (requestedView === 'warehouse') return 'Warehouse';
  if (requestedView === 'manager') return 'Manager';
  if (requestedView === 'mobilewarehouse') return 'MobileWarehouse';
  if (requestedView === 'mobile') return 'MobileWarehouse';

  if (!user) return 'Index';
  if (user.rol === ROLE.WAREHOUSE) return 'Warehouse';
  if (user.rol === ROLE.MOBILE_WAREHOUSE) return 'MobileWarehouse';
  if (user.rol === ROLE.MANAGER || user.rol === ROLE.ANALYSIS) return 'Manager';

  return 'Index';
}

function assertPageAccess(fileName, user, techRef) {
  if (!user) {
    throw new Error('Geen gebruiker gevonden.');
  }

  if (fileName === 'Index') {
    const technician = resolveTechnicianByRef(techRef);

    if (user.rol === ROLE.ADMIN) {
      return { technician: technician };
    }

    if (!roleAllowed(user, [ROLE.TECHNICIAN])) {
      throw new Error('Deze pagina is enkel voor techniekers.');
    }

    if (normalizeRef(user.techniekerCode) !== normalizeRef(technician.code)) {
      throw new Error('Je kan enkel je eigen techniekerpagina openen.');
    }

    return { technician: technician };
  }

  if (fileName === 'Warehouse') {
    if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER])) {
      throw new Error('Deze pagina is enkel voor magazijn of manager.');
    }
    return {};
  }

  if (fileName === 'MobileWarehouse') {
    if (!roleAllowed(user, [ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
      throw new Error('Deze pagina is enkel voor mobiel magazijn of manager.');
    }
    return {};
  }

  if (fileName === 'Manager') {
    if (!roleAllowed(user, [ROLE.MANAGER, ROLE.ANALYSIS])) {
      throw new Error('Deze pagina is enkel voor manager of analyse.');
    }
    return {};
  }

  return {};
}

function buildPageTemplateContext(fileName, user, techRef, sessionId) {
  const authContext = buildAuthTemplateContext(user, sessionId);

  const context = {
    techRef: safeText(techRef),
    ...authContext,

    // nieuw: alleen voor eerste bootstrap in de browser
    bootstrapSessionId: safeText(sessionId || '')
  };

  if (fileName === 'Index') {
    context.pageTitle = 'DigiQS Grabbelstock';
  } else if (fileName === 'Warehouse') {
    context.pageTitle = 'DigiQS Magazijn';
  } else if (fileName === 'MobileWarehouse') {
    context.pageTitle = 'DigiQS Mobiel Magazijn';
  } else if (fileName === 'Manager') {
    context.pageTitle = 'DigiQS Manager';
  } else {
    context.pageTitle = APP_CONFIG.DEFAULT_PAGE_TITLE || 'DigiQS Warehouse';
  }

  return context;
}

function renderLoginPage(view, techRef, sid) {
  const context = buildLoginTemplateContext(view, techRef, sid);
  return evaluateTemplateWithData('Login', context, 'DigiQS Login');
}

function renderAppPage(fileName, user, techRef, sessionId) {
  const templateContext = buildPageTemplateContext(fileName, user, techRef, sessionId);
  const pageTitle = templateContext.pageTitle || APP_CONFIG.DEFAULT_PAGE_TITLE || 'DigiQS Warehouse';

  if (fileName === 'MobileWarehouse') {
    return evaluateTemplateWithFallback('MobileWarehouse', 'Warehouse', templateContext, pageTitle);
  }

  return evaluateTemplateWithData(fileName, templateContext, pageTitle);
}

function doGet(e) {
  try {
    const techRef = getRequestedTechRef(e);
    const view = getRequestedView(e);
    const sid = getRequestedSessionId(e);

    const user = resolveLoggedInUserFromRequest(sid);

    if (!user) {
      return renderLoginPage(view, techRef, sid);
    }

    const fileName = resolveTargetFileName(view, user);
    assertPageAccess(fileName, user, techRef);

    return renderAppPage(fileName, user, techRef, sid);
  } catch (err) {
    return getForbiddenOutput(
      err && err.message ? err.message : 'Onbekende fout bij openen van de pagina.'
    );
  }
}
