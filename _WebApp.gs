/* =========================================================
   15_WebApp.gs
   Refactor: webapp routing + shell rendering
   Doel:
   - één centrale doGet-router
   - auth/sessioncontext koppelen aan views
   - consistente modelopbouw voor Login / Index / Warehouse / Manager / MobileWarehouse
   - veilige fallback voor access denied en fouten
   ========================================================= */

/* ---------------------------------------------------------
   View helpers
   --------------------------------------------------------- */

function getWebAppDefaultView_() {
  return 'Login';
}

function normalizeWebAppViewName_(value) {
  var raw = safeText(value).toLowerCase();

  if (!raw) return getWebAppDefaultView_();
  if (raw === 'login') return 'Login';
  if (raw === 'index') return 'Index';
  if (raw === 'warehouse') return 'Warehouse';
  if (raw === 'manager') return 'Manager';
  if (raw === 'mobilewarehouse' || raw === 'mobile_warehouse' || raw === 'mobile-magazijn') {
    return 'MobileWarehouse';
  }

  return getWebAppDefaultView_();
}

function getRequestParam_(e, key) {
  var params = (e && e.parameter) || {};
  return safeText(params[key]);
}

function getRequestedSessionId_(e) {
  return safeText(
    getRequestParam_(e, 'sid') ||
    getRequestParam_(e, 'sessionId')
  );
}

function getRequestedTechnicianRef_(e) {
  return safeText(
    getRequestParam_(e, 'tech') ||
    getRequestParam_(e, 'techniekerCode') ||
    getRequestParam_(e, 'technicianCode')
  );
}

function getRequestedMobileWarehouseCode_(e) {
  return safeText(
    getRequestParam_(e, 'mw') ||
    getRequestParam_(e, 'mobileWarehouseCode')
  );
}

function resolveRequestedView_(e) {
  return normalizeWebAppViewName_(getRequestParam_(e, 'view'));
}

/* ---------------------------------------------------------
   Render helpers
   --------------------------------------------------------- */

function createTemplateFromFile_(fileName) {
  return HtmlService.createTemplateFromFile(fileName);
}

function renderHtmlFromTemplate_(fileName, model) {
  var template = createTemplateFromFile_(fileName);
  template.model = model || {};

  return template
    .evaluate()
    .setTitle(safeText(model && model.pageTitle) || 'STOCKAPP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderPlainHtmlPage_(title, heading, bodyHtml) {
  var html =
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
      '<meta charset="utf-8" />' +
      '<meta name="viewport" content="width=device-width, initial-scale=1" />' +
      '<title>' + sanitizeHtmlText_(title || 'STOCKAPP') + '</title>' +
      '<style>' +
        'body{font-family:Arial,sans-serif;background:#f6f7fb;color:#1f2937;margin:0;padding:24px;}' +
        '.card{max-width:900px;margin:32px auto;background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:24px;box-shadow:0 10px 30px rgba(0,0,0,.06);}' +
        'h1{margin:0 0 12px 0;font-size:28px;}' +
        'p{line-height:1.5;margin:0 0 10px 0;}' +
        '.muted{color:#6b7280;}' +
      '</style>' +
    '</head>' +
    '<body>' +
      '<div class="card">' +
        '<h1>' + sanitizeHtmlText_(heading || 'STOCKAPP') + '</h1>' +
        '<div>' + safeText(bodyHtml) + '</div>' +
      '</div>' +
    '</body>' +
    '</html>';

  return HtmlService
    .createHtmlOutput(html)
    .setTitle(title || 'STOCKAPP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderAccessDeniedPage_(message) {
  return renderPlainHtmlPage_(
    'Geen toegang',
    'Geen toegang',
    '<p>' + sanitizeHtmlText_(message || 'Je hebt geen toegang tot deze pagina.') + '</p>'
  );
}

function renderErrorPage_(message) {
  return renderPlainHtmlPage_(
    'Fout',
    'Er liep iets mis',
    '<p>' + sanitizeHtmlText_(message || 'Onbekende fout.') + '</p>'
  );
}

function sanitizeHtmlText_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/* ---------------------------------------------------------
   Session / route context
   --------------------------------------------------------- */

function buildBaseRouteModel_(e, sessionContext) {
  var user = (sessionContext && sessionContext.user) || null;
  var requestedView = resolveRequestedView_(e);
  var requestedTechRef = getRequestedTechnicianRef_(e);
  var requestedMobileWarehouseCode = getRequestedMobileWarehouseCode_(e);
  var sessionId = safeText((sessionContext && sessionContext.sessionId) || getRequestedSessionId_(e));

  return {
    appName: 'STOCKAPP',
    requestedView: requestedView,
    requestedTechRef: requestedTechRef,
    requestedMobileWarehouseCode: requestedMobileWarehouseCode,

    sessionId: sessionId,
    isAuthenticated: !!(sessionContext && sessionContext.authenticated),
    authMode: safeText(sessionContext && sessionContext.authMode),

    currentUserCode: safeText(user && user.code),
    currentUserName: safeText(user && user.naam),
    currentUserRole: safeText(user && user.rol),
    currentUserEmail: safeText(user && user.email),

    bootstrapSessionId: sessionId,
    currentSessionId: sessionId,
  };
}

function resolveIndexTechRef_(model) {
  var requestedTechRef = safeText(model.requestedTechRef);
  if (requestedTechRef) return requestedTechRef;

  if (safeText(model.currentUserRole) === safeText(ROLE.TECHNICIAN)) {
    return safeText(model.currentUserCode);
  }

  return '';
}

function buildPageSpecificModel_(viewName, baseModel) {
  var model = Object.assign({}, baseModel || {});
  var page = safeText(viewName);

  if (page === 'Login') {
    model.pageTitle = 'Aanmelden';
    model.pageSubtitle = 'Meld aan om STOCKAPP te openen.';
    model.pageRole = 'login';
    model.bodyClass = 'page-login';
    return model;
  }

  if (page === 'Index') {
    model.pageTitle = 'Technieker dashboard';
    model.pageSubtitle = 'Overzicht van beleveringen, grabbelbestellingen en meldingen.';
    model.pageRole = 'technician';
    model.bodyClass = 'page-index';
    model.selectedTechnicianCode = resolveIndexTechRef_(model);
    model.techRef = model.selectedTechnicianCode;
    return model;
  }

  if (page === 'Warehouse') {
    model.pageTitle = 'Magazijn dashboard';
    model.pageSubtitle = 'Grabbelstock, ontvangsten, retouren, behoefte-uitgiftes, tellingen en meldingen.';
    model.pageRole = 'warehouse';
    model.bodyClass = 'page-warehouse';
    return model;
  }

  if (page === 'Manager') {
    model.pageTitle = 'Manager dashboard';
    model.pageSubtitle = 'Goedkeuringen, stock, tellingen, audits en meldingen.';
    model.pageRole = 'manager';
    model.bodyClass = 'page-manager';
    return model;
  }

  if (page === 'MobileWarehouse') {
    model.pageTitle = 'Mobiel magazijn dashboard';
    model.pageSubtitle = 'Overzicht voor mobiel magazijn.';
    model.pageRole = 'mobilewarehouse';
    model.bodyClass = 'page-mobilewarehouse';
    model.selectedMobileWarehouseCode = safeText(model.requestedMobileWarehouseCode);
    return model;
  }

  model.pageTitle = 'STOCKAPP';
  model.pageSubtitle = '';
  model.pageRole = '';
  model.bodyClass = '';
  return model;
}

function getViewFileForPage_(pageName) {
  var page = safeText(pageName);

  if (page === 'Login') return 'Login';
  if (page === 'Index') return 'Index';
  if (page === 'Warehouse') return 'Warehouse';
  if (page === 'Manager') return 'Manager';
  if (page === 'MobileWarehouse') return 'MobileWarehouse';

  return 'Login';
}

/* ---------------------------------------------------------
   Route policy
   --------------------------------------------------------- */

function isPublicRoute_(pageName) {
  return safeText(pageName) === 'Login';
}

function resolveSessionContextForRoute_(e) {
  if (typeof getSessionContext !== 'function') {
    throw new Error('Session service ontbreekt. Werk eerst het sessionblok in.');
  }

  return getSessionContext({
    sessionId: getRequestedSessionId_(e),
  });
}

function assertRouteAccessOrReturnPage_(pageName, model) {
  var page = safeText(pageName);

  if (isPublicRoute_(page)) {
    return null;
  }

  if (!isTrue(model.isAuthenticated)) {
    return renderHtmlFromTemplate_(
      'Login',
      buildPageSpecificModel_('Login', Object.assign({}, model, {
        pageTitle: 'Aanmelden',
        pageSubtitle: 'Je sessie is verlopen of je bent nog niet aangemeld.',
      }))
    );
  }

  if (typeof canAccessPage !== 'function') {
    throw new Error('Rights service ontbreekt. Werk eerst het rightsblok in.');
  }

  var user = {
    code: model.currentUserCode,
    naam: model.currentUserName,
    rol: model.currentUserRole,
    email: model.currentUserEmail,
    sessionId: model.sessionId,
  };

  var techRef = page === 'Index' ? resolveIndexTechRef_(model) : '';
  if (!canAccessPage(user, page, techRef)) {
    return renderAccessDeniedPage_('Je hebt geen rechten om deze pagina te openen.');
  }

  return null;
}

/* ---------------------------------------------------------
   Main routing
   --------------------------------------------------------- */

function buildRouteModel_(e) {
  var sessionContext = resolveSessionContextForRoute_(e);
  var baseModel = buildBaseRouteModel_(e, sessionContext);
  return buildPageSpecificModel_(baseModel.requestedView, baseModel);
}

function doGet(e) {
  try {
    var model = buildRouteModel_(e);
    var requestedView = safeText(model.requestedView || getWebAppDefaultView_());

    var guardPage = assertRouteAccessOrReturnPage_(requestedView, model);
    if (guardPage) {
      return guardPage;
    }

    var fileName = getViewFileForPage_(requestedView);
    return renderHtmlFromTemplate_(fileName, model);
  } catch (err) {
    return renderErrorPage_(extractErrorMessage(err, 'Onbekende fout tijdens openen van de webapp.'));
  }
}
