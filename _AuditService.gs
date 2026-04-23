/* =========================================================
   20_AuditService.gs
   Refactor: audit core service
   Doel:
   - centrale auditlaag voor alle write-acties
   - uniforme logging van actor, document en details
   - read/querylaag voor manager en magazijn
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getAuditTab_() {
  return TABS.AUDIT_LOG || 'AuditLog';
}

function getAuditDefaultDocumentType_() {
  return 'Onbekend';
}

function getAuditDefaultAction_() {
  return 'UNKNOWN_ACTION';
}

function makeAuditId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'AUD-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapAuditRow(row) {
  return {
    auditId: safeText(row.AuditID || row.AuditId || row.ID),
    actie: safeText(row.Actie || row.Action),
    actorCode: safeText(row.ActorCode || row.UserCode),
    actorNaam: safeText(row.ActorNaam || row.ActorName),
    actorRol: safeText(row.ActorRol || row.ActorRole),
    actorEmail: safeText(row.ActorEmail || row.Email),
    documentType: safeText(row.DocumentType || row.DocType || getAuditDefaultDocumentType_()),
    documentId: safeText(row.DocumentID || row.DocumentId || row.DocId),
    detailsJson: safeText(row.DetailsJson || row.DetailsJSON || row.Details || ''),
    aangemaaktOp: safeText(row.AangemaaktOp || row.CreatedAt),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.CreatedAtRaw || row.AangemaaktOp || row.CreatedAt),
  };
}

/* ---------------------------------------------------------
   Helpers
   --------------------------------------------------------- */

function safeAuditJson_(value) {
  try {
    return safeJson(value || {});
  } catch (err) {
    return '{}';
  }
}

function normalizeAuditActor_(actor) {
  actor = actor || {};

  return {
    code: safeText(actor.code || actor.userCode || actor.actorCode),
    naam: safeText(actor.naam || actor.userName || actor.actorName),
    rol: safeText(actor.rol || actor.userRole || actor.actorRole),
    email: safeText(actor.email || actor.actorEmail),
  };
}

function normalizeAuditPayload_(payload) {
  payload = payload || {};

  var actor = normalizeAuditActor_(payload.actor);

  return {
    actie: safeText(payload.actie || payload.action || getAuditDefaultAction_()),
    actor: actor,
    documentType: safeText(payload.documentType || payload.docType || getAuditDefaultDocumentType_()),
    documentId: safeText(payload.documentId || payload.docId),
    details: payload.details || {},
  };
}

/* ---------------------------------------------------------
   Write
   --------------------------------------------------------- */

function writeAudit(payload) {
  var normalized = normalizeAuditPayload_(payload);
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var obj = {
    AuditID: makeAuditId_(),
    Actie: normalized.actie || getAuditDefaultAction_(),
    ActorCode: normalized.actor.code,
    ActorNaam: normalized.actor.naam,
    ActorRol: normalized.actor.rol,
    ActorEmail: normalized.actor.email,
    DocumentType: normalized.documentType || getAuditDefaultDocumentType_(),
    DocumentID: normalized.documentId,
    DetailsJson: safeAuditJson_(normalized.details),
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
  };

  appendObjects(getAuditTab_(), [obj]);

  return mapAuditRow(obj);
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllAuditEntries() {
  return readObjectsSafe(getAuditTab_())
    .map(mapAuditRow)
    .sort(function (a, b) {
      return (
        safeText(b.aangemaaktOpRaw).localeCompare(safeText(a.aangemaaktOpRaw)) ||
        safeText(b.auditId).localeCompare(safeText(a.auditId))
      );
    });
}

function getAuditEntryById(auditId) {
  var id = safeText(auditId);
  if (!id) return null;

  return getAllAuditEntries().find(function (item) {
    return safeText(item.auditId) === id;
  }) || null;
}

/* ---------------------------------------------------------
   Filters
   --------------------------------------------------------- */

function matchesAuditQuery_(item, query) {
  var q = safeText(query).toLowerCase();
  if (!q) return true;

  return [
    item.auditId,
    item.actie,
    item.actorCode,
    item.actorNaam,
    item.actorRol,
    item.actorEmail,
    item.documentType,
    item.documentId,
    item.detailsJson,
  ].some(function (value) {
    return safeText(value).toLowerCase().indexOf(q) >= 0;
  });
}

function filterAuditEntries_(rows, filters) {
  var f = filters || {};
  var query = safeText(f.query || f.search);
  var actie = safeText(f.actie || f.action);
  var actorRol = safeText(f.actorRol || f.actorRole);
  var documentType = safeText(f.documentType || f.docType);
  var documentId = safeText(f.documentId || f.docId);
  var startDate = safeText(f.startDate);
  var endDate = safeText(f.endDate);

  return (rows || []).filter(function (item) {
    if (!matchesAuditQuery_(item, query)) {
      return false;
    }
    if (actie && safeText(item.actie) !== actie) {
      return false;
    }
    if (actorRol && safeText(item.actorRol) !== actorRol) {
      return false;
    }
    if (documentType && safeText(item.documentType) !== documentType) {
      return false;
    }
    if (documentId && safeText(item.documentId) !== documentId) {
      return false;
    }
    if (startDate && safeText(item.aangemaaktOpRaw).slice(0, 10) < startDate) {
      return false;
    }
    if (endDate && safeText(item.aangemaaktOpRaw).slice(0, 10) > endDate) {
      return false;
    }
    return true;
  });
}

/* ---------------------------------------------------------
   Access policy
   --------------------------------------------------------- */

function assertAuditReadAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.MANAGER, ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE],
    'Geen rechten om auditlog te bekijken.'
  );
  return user;
}

/* ---------------------------------------------------------
   Summary / query
   --------------------------------------------------------- */

function summarizeAuditEntries_(rows) {
  var items = Array.isArray(rows) ? rows : [];

  return {
    totaal: items.length,
    uniekeActies: Object.keys(
      items.reduce(function (acc, item) {
        acc[safeText(item.actie)] = true;
        return acc;
      }, {})
    ).length,
    uniekeActoren: Object.keys(
      items.reduce(function (acc, item) {
        acc[safeText(item.actorNaam || item.actorCode || item.actorEmail)] = true;
        return acc;
      }, {})
    ).length,
    uniekeDocumenten: Object.keys(
      items.reduce(function (acc, item) {
        acc[safeText(item.documentType) + ':' + safeText(item.documentId)] = true;
        return acc;
      }, {})
    ).length,
  };
}

function getAuditData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertAuditReadAccess_(sessionId);
  var rows = filterAuditEntries_(getAllAuditEntries(), payload.filters || payload);

  return {
    items: rows,
    audit: rows,
    summary: Object.assign(
      {},
      summarizeAuditEntries_(rows),
      { actorRol: safeText(actor.rol) }
    ),
  };
}
