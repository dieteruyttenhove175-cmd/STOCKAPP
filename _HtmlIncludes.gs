/* =========================================================
   _HtmlIncludes.gs — shared html includes + view models
   ========================================================= */

function include(filename, model) {
  const template = HtmlService.createTemplateFromFile(filename);
  template.model = model || {};
  return template.evaluate().getContent();
}

function buildWarehouseViewModel(session) {
  session = session || {};

  return {
    pageTitle: 'Magazijn dashboard',
    pageSubtitle: 'Grabbelstock, ontvangsten, retouren, behoefte-uitgiftes, tellingen en meldingen.',
    actorLabel: 'Naam magazijnier',
    actorPlaceholder: 'Bijv. Dieter',
    pageRole: 'warehouse',
    currentUserName: safeText(session.userName || ''),
    currentUserRole: safeText(session.userRole || ''),
    currentAuthMode: safeText(session.authMode || ''),
    currentSessionId: safeText(session.sessionId || ''),
    bootstrapSessionId: safeText(session.sessionId || ''),
    permissions: {
      canCreateReceipt: true,
      canUploadReceipt: true,
      canEditReceipt: true,
      canSubmitReceipt: true,
      canApproveReceipt: false,
      canCreateReturn: true,
      canApproveReturn: false,
      canCreateNeedIssue: true,
      canBookNeedIssue: true,
      canSeeCentralStock: true
    }
  };
}

function buildManagerViewModel(session) {
  session = session || {};

  return {
    pageTitle: 'Manager dashboard',
    pageSubtitle: 'Manageroverzicht van grabbelstock, ontvangsten, retouren, behoefte-uitgiftes, tellingen en meldingen.',
    actorLabel: 'Naam manager',
    actorPlaceholder: 'Bijv. Dieter',
    pageRole: 'manager',
    currentUserName: safeText(session.userName || ''),
    currentUserRole: safeText(session.userRole || ''),
    currentAuthMode: safeText(session.authMode || ''),
    currentSessionId: safeText(session.sessionId || ''),
    bootstrapSessionId: safeText(session.sessionId || ''),
    permissions: {
      canCreateReceipt: true,
      canUploadReceipt: true,
      canEditReceipt: true,
      canSubmitReceipt: true,
      canApproveReceipt: true,
      canCreateReturn: true,
      canApproveReturn: true,
      canCreateNeedIssue: true,
      canBookNeedIssue: true,
      canSeeCentralStock: true
    }
  };
}

function buildMobileWarehouseViewModel(session) {
  session = session || {};

  return {
    pageTitle: 'Mobiel magazijn dashboard',
    pageSubtitle: 'Overzicht voor mobiel magazijn.',
    actorLabel: 'Naam mobiel magazijn',
    actorPlaceholder: 'Bijv. Dieter',
    pageRole: 'mobilewarehouse',
    currentUserName: safeText(session.userName || ''),
    currentUserRole: safeText(session.userRole || ''),
    currentAuthMode: safeText(session.authMode || ''),
    currentSessionId: safeText(session.sessionId || ''),
    bootstrapSessionId: safeText(session.sessionId || ''),
    permissions: {
      canCreateReceipt: false,
      canUploadReceipt: false,
      canEditReceipt: false,
      canSubmitReceipt: false,
      canApproveReceipt: false,
      canCreateReturn: true,
      canApproveReturn: false,
      canCreateNeedIssue: true,
      canBookNeedIssue: true,
      canSeeCentralStock: true
    }
  };
}

function buildTechnicianViewModel(session) {
  session = session || {};

  return {
    pageTitle: 'Techniekerscherm',
    pageSubtitle: 'Bestellingen, leveringen, meldingen, busstocktellingen en aanvragen naar mobiel magazijn.',
    actorLabel: 'Naam technieker',
    actorPlaceholder: 'Bijv. Dieter',
    pageRole: 'technician',
    currentUserName: safeText(session.userName || ''),
    currentUserRole: safeText(session.userRole || ''),
    currentAuthMode: safeText(session.authMode || ''),
    currentSessionId: safeText(session.sessionId || ''),
    bootstrapSessionId: safeText(session.sessionId || ''),
    permissions: {
      canCreateOrder: true,
      canEditOrder: true,
      canReceiveDelivery: true,
      canViewNotifications: true,
      canCountBusStock: true,
      canCreateMobileRequest: true,
      canEditMobileRequest: true,
      silentAutoRefresh: true
    }
  };
}