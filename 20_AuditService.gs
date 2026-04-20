/* =========================================================
   20_AuditService.gs — audittrail
   ========================================================= */

function makeAuditId() {
  return makeStampedId('A');
}

function buildAuditObject(action, role, actor, documentType, documentId, details) {
  return {
    AuditID: makeAuditId(),
    Tijdstip: nowStamp(),
    Actie: safeText(action),
    Rol: safeText(role),
    Actor: safeText(actor),
    DocumentType: safeText(documentType),
    DocumentID: safeText(documentId),
    Details: safeJson(details)
  };
}

function writeAudit(action, role, actor, documentType, documentId, details) {
  appendObjects(TABS.AUDIT_LOG, [
    buildAuditObject(action, role, actor, documentType, documentId, details)
  ]);
}

function mapAuditRow(row) {
  return {
    auditId: safeText(row.AuditID),
    tijdstip: toDisplayDateTime(row.Tijdstip),
    tijdstipRaw: safeText(row.Tijdstip),
    actie: safeText(row.Actie),
    rol: safeText(row.Rol),
    actor: safeText(row.Actor),
    documentType: safeText(row.DocumentType),
    documentId: safeText(row.DocumentID),
    details: safeText(row.Details)
  };
}

function getAuditLog(limit) {
  const rows = readObjectsSafe(TABS.AUDIT_LOG)
    .map(mapAuditRow)
    .sort((a, b) => String(b.tijdstipRaw || '').localeCompare(String(a.tijdstipRaw || '')));

  if (typeof limit === 'number' && limit > 0) {
    return rows.slice(0, limit);
  }

  return rows;
}

function getLatestAuditLog() {
  return getAuditLog(APP_CONFIG.MAX_AUDIT_ROWS || 10000);
}

function writeSystemAudit(action, documentType, documentId, details) {
  writeAudit(action, 'Systeem', 'Automatisch', documentType, documentId, details);
}