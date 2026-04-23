/* =========================================================
   90_HtmlIncludes.gs
   Refactor: html include helpers
   Doel:
   - één vaste includeconventie voor HTML-partials
   - model veilig doorgeven aan partials
   - helpers voor script/style/head rendering
   - basis voor dunne shells met herbruikbare blocks
   ========================================================= */

/* ---------------------------------------------------------
   Basic include helpers
   --------------------------------------------------------- */

function include(fileName, model) {
  var template = HtmlService.createTemplateFromFile(fileName);
  template.model = model || {};
  return template.evaluate().getContent();
}

function includeRaw(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function includeIf(condition, fileName, model) {
  return condition ? include(fileName, model) : '';
}

/* ---------------------------------------------------------
   Safe bootstrap helpers
   --------------------------------------------------------- */

function safeTemplateJson_(value) {
  try {
    return JSON.stringify(value == null ? {} : value);
  } catch (err) {
    return '{}';
  }
}

function buildTemplateBootstrapScript(model, globalName) {
  var name = safeText(globalName || 'BOOTSTRAP_MODEL');
  return [
    '<script>',
    'window.' + name + ' = ' + safeTemplateJson_(model || {}) + ';',
    '</script>'
  ].join('');
}

function buildSessionBootstrapScript(model) {
  var sessionId = safeText(model && (model.bootstrapSessionId || model.currentSessionId || model.sessionId));
  return [
    '<script>',
    'window.BOOTSTRAP_SESSION_ID = ' + JSON.stringify(sessionId) + ';',
    'window.CURRENT_SESSION_ID = ' + JSON.stringify(sessionId) + ';',
    '</script>'
  ].join('');
}

/* ---------------------------------------------------------
   Partial wrappers
   --------------------------------------------------------- */

function renderScriptPartial(fileName, model) {
  return [
    '<script>',
    include(fileName, model),
    '</script>'
  ].join('\n');
}

function renderStylePartial(fileName, model) {
  return [
    '<style>',
    include(fileName, model),
    '</style>'
  ].join('\n');
}

function renderHtmlPartial(fileName, model) {
  return include(fileName, model);
}

/* ---------------------------------------------------------
   Shared shell composition
   --------------------------------------------------------- */

function buildSharedShellHead(model) {
  return include('_shared_shell_head', model);
}

function buildSharedCommonScripts(model) {
  return [
    renderScriptPartial('_shared_common_js', model),
    renderScriptPartial('_shared_loader_js', model)
  ].join('\n');
}

function buildTechnicianScripts(model) {
  return [
    renderScriptPartial('_shared_common_js', model),
    renderScriptPartial('_tech_loader_js', model)
  ].join('\n');
}

function buildMobileWarehouseScripts(model) {
  return [
    renderScriptPartial('_shared_common_js', model),
    renderScriptPartial('_mobilewarehouse_js', model)
  ].join('\n');
}

/* ---------------------------------------------------------
   View model helpers
   --------------------------------------------------------- */

function buildShellModel(baseModel, overrides) {
  return Object.assign({}, baseModel || {}, overrides || {});
}

function buildLoginViewModel(baseModel) {
  return buildShellModel(baseModel, {
    pageTitle: safeText(baseModel && baseModel.pageTitle) || 'Aanmelden',
    pageRole: 'login',
    bodyClass: 'page-login',
  });
}

function buildTechnicianViewModel(baseModel) {
  return buildShellModel(baseModel, {
    pageTitle: safeText(baseModel && baseModel.pageTitle) || 'Technieker dashboard',
    pageRole: 'technician',
    bodyClass: 'page-index',
    techRef: safeText(baseModel && (baseModel.techRef || baseModel.selectedTechnicianCode || '')),
  });
}

function buildWarehouseViewModel(baseModel) {
  return buildShellModel(baseModel, {
    pageTitle: safeText(baseModel && baseModel.pageTitle) || 'Magazijn dashboard',
    pageRole: 'warehouse',
    bodyClass: 'page-warehouse',
  });
}

function buildManagerViewModel(baseModel) {
  return buildShellModel(baseModel, {
    pageTitle: safeText(baseModel && baseModel.pageTitle) || 'Manager dashboard',
    pageRole: 'manager',
    bodyClass: 'page-manager',
  });
}

function buildMobileWarehouseViewModel(baseModel) {
  return buildShellModel(baseModel, {
    pageTitle: safeText(baseModel && baseModel.pageTitle) || 'Mobiel magazijn dashboard',
    pageRole: 'mobilewarehouse',
    bodyClass: 'page-mobilewarehouse',
    selectedMobileWarehouseCode: safeText(
      baseModel && (baseModel.selectedMobileWarehouseCode || baseModel.requestedMobileWarehouseCode || '')
    ),
  });
}

/* ---------------------------------------------------------
   Shell block helpers
   --------------------------------------------------------- */

function buildLoginShellBlocks(model) {
  var m = buildLoginViewModel(model);
  return {
    head: buildSharedShellHead(m),
    bootstrap: buildTemplateBootstrapScript(m, 'BOOTSTRAP_MODEL') + '\n' + buildSessionBootstrapScript(m),
    bodyStart: '<div id="loginApp" class="app-shell app-shell-login"></div>',
    scripts: buildSharedCommonScripts(m),
  };
}

function buildTechnicianShellBlocks(model) {
  var m = buildTechnicianViewModel(model);
  return {
    head: buildSharedShellHead(m),
    bootstrap: buildTemplateBootstrapScript(m, 'BOOTSTRAP_MODEL') + '\n' + buildSessionBootstrapScript(m),
    bodyStart: '<div id="techApp" class="app-shell app-shell-tech"></div>',
    scripts: buildTechnicianScripts(m),
  };
}

function buildWarehouseShellBlocks(model) {
  var m = buildWarehouseViewModel(model);
  return {
    head: buildSharedShellHead(m),
    bootstrap: buildTemplateBootstrapScript(m, 'BOOTSTRAP_MODEL') + '\n' + buildSessionBootstrapScript(m),
    bodyStart: '<div id="warehouseApp" class="app-shell app-shell-warehouse"></div>',
    scripts: buildSharedCommonScripts(m),
  };
}

function buildManagerShellBlocks(model) {
  var m = buildManagerViewModel(model);
  return {
    head: buildSharedShellHead(m),
    bootstrap: buildTemplateBootstrapScript(m, 'BOOTSTRAP_MODEL') + '\n' + buildSessionBootstrapScript(m),
    bodyStart: '<div id="managerApp" class="app-shell app-shell-manager"></div>',
    scripts: buildSharedCommonScripts(m),
  };
}

function buildMobileWarehouseShellBlocks(model) {
  var m = buildMobileWarehouseViewModel(model);
  return {
    head: buildSharedShellHead(m),
    bootstrap: buildTemplateBootstrapScript(m, 'BOOTSTRAP_MODEL') + '\n' + buildSessionBootstrapScript(m),
    bodyStart: '<div id="mobileWarehouseApp" class="app-shell app-shell-mobilewarehouse"></div>',
    scripts: buildMobileWarehouseScripts(m),
  };
}

/* ---------------------------------------------------------
   Full shell render helpers
   --------------------------------------------------------- */

function renderShellDocument(blocks) {
  var b = blocks || {};

  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
    safeText(b.head),
    safeText(b.bootstrap),
    '</head>',
    '<body>',
    safeText(b.bodyStart),
    safeText(b.scripts),
    '</body>',
    '</html>'
  ].join('\n');
}

function renderLoginShell(model) {
  return renderShellDocument(buildLoginShellBlocks(model));
}

function renderTechnicianShell(model) {
  return renderShellDocument(buildTechnicianShellBlocks(model));
}

function renderWarehouseShell(model) {
  return renderShellDocument(buildWarehouseShellBlocks(model));
}

function renderManagerShell(model) {
  return renderShellDocument(buildManagerShellBlocks(model));
}

function renderMobileWarehouseShell(model) {
  return renderShellDocument(buildMobileWarehouseShellBlocks(model));
}

/* ---------------------------------------------------------
   Developer guardrails
   --------------------------------------------------------- */

function getIncludeConventionsReadme_() {
  return [
    'Belangrijke include-regels:',
    '1. JS-partials bevatten pure JavaScript, zonder eigen <script>-tags.',
    '2. Gebruik renderScriptPartial(...) of <script><?!= include(...) ?></script> in shells.',
    '3. CSS-partials bevatten pure CSS, zonder eigen <style>-tags.',
    '4. Pagina-shells bevatten geen businesslogica.',
    '5. Model wordt altijd doorgegeven als template.model.',
  ].join('\n');
}
