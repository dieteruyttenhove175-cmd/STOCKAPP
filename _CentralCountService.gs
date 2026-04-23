/* =========================================================
   50_CentralCountService.gs
   Refactor: central count core service
   Doel:
   - centrale documentlaag voor magazijntellingen
   - volledige of gerichte tellingen
   - lijnen bewaren
   - indienen
   - goedkeuren en correctiemutaties boeken
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getCentralCountHeaderTab_() {
  return TABS.CENTRAL_COUNTS || 'CentraleTellingen';
}

function getCentralCountLineTab_() {
  return TABS.CENTRAL_COUNT_LINES || 'CentraleTelLijnen';
}

function getCentralCountStatusOpen_() {
  if (typeof CENTRAL_COUNT_STATUS !== 'undefined' && CENTRAL_COUNT_STATUS && CENTRAL_COUNT_STATUS.OPEN) {
    return CENTRAL_COUNT_STATUS.OPEN;
  }
  return 'Open';
}

function getCentralCountStatusSubmitted_() {
  if (typeof CENTRAL_COUNT_STATUS !== 'undefined' && CENTRAL_COUNT_STATUS && CENTRAL_COUNT_STATUS.SUBMITTED) {
    return CENTRAL_COUNT_STATUS.SUBMITTED;
  }
  return 'Ingediend';
}

function getCentralCountStatusApproved_() {
  if (typeof CENTRAL_COUNT_STATUS !== 'undefined' && CENTRAL_COUNT_STATUS && CENTRAL_COUNT_STATUS.APPROVED) {
    return CENTRAL_COUNT_STATUS.APPROVED;
  }
  return 'Goedgekeurd';
}

function getCentralCountStatusClosed_() {
  if (typeof CENTRAL_COUNT_STATUS !== 'undefined' && CENTRAL_COUNT_STATUS && CENTRAL_COUNT_STATUS.CLOSED) {
    return CENTRAL_COUNT_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function getCentralCountScopeFull_() {
  return 'FULL';
}

function getCentralCountScopeTargeted_() {
  return 'TARGETED';
}

function getCentralCountMovementIn_() {
  if (typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE && MOVEMENT_TYPE.CENTRAL_COUNT_IN) {
    return MOVEMENT_TYPE.CENTRAL_COUNT_IN;
  }
  return 'CentralCountIn';
}

function getCentralCountMovementOut_() {
  if (typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE && MOVEMENT_TYPE.CENTRAL_COUNT_OUT) {
    return MOVEMENT_TYPE.CENTRAL_COUNT_OUT;
  }
  return 'CentralCountOut';
}

function isCentralCountEditable_(status) {
  var value = safeText(status);
  return !value || value === getCentralCountStatusOpen_();
}

function makeCentralCountId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'CCT-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function makeCentralCountLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'CCL-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapCentralCountHeader(row) {
  return {
    centralCountId: safeText(row.CentralCountID || row.CentralCountId || row.CentraleTellingID || row.ID),
    scopeType: safeText(row.ScopeType || getCentralCountScopeFull_()),
    requestedBy: safeText(row.RequestedBy || row.AangevraagdDoor),
    reason: safeText(row.Reason || row.Reden),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    status: safeText(row.Status),
    actor: safeText(row.Actor),
    opmerking: safeText(row.Opmerking),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    ingediendOp: safeText(row.IngediendOp),
    goedgekeurdOp: safeText(row.GoedgekeurdOp),
  };
}

function mapCentralCountLine(row) {
  var systemQty = safeNumber(row.SystemAantal || row.SystemQty, 0);
  var countedQty = safeNumber(row.GeteldAantal || row.CountedQty, 0);

  return {
    centralCountLineId: safeText(row.CentralCountLineID || row.CentralCountLineId || row.CentraleTellingLijnID || row.ID),
    centralCountId: safeText(row.CentralCountID || row.CentralCountId || row.CentraleTellingID),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode || row.ArtikelNr))),
    eenheid: safeText(row.Eenheid || row.Unit),
    systemAantal: systemQty,
    geteldAantal: countedQty,
    deltaAantal: safeNumber(row.DeltaAantal, countedQty - systemQty),
    opmerking: safeText(row.Opmerking),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllCentralCountHeaders() {
  return readObjectsSafe(getCentralCountHeaderTab_())
    .map(mapCentralCountHeader)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.centralCountId).localeCompare(safeText(a.centralCountId))
      );
    });
}

function getAllCentralCountLines() {
  return readObjectsSafe(getCentralCountLineTab_())
    .map(mapCentralCountLine)
    .sort(function (a, b) {
      return (
        safeText(a.centralCountId).localeCompare(safeText(b.centralCountId)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
      );
    });
}

function getCentralCountHeaderById(centralCountId) {
  var id = safeText(centralCountId);
  if (!id) return null;

  return getAllCentralCountHeaders().find(function (item) {
    return safeText(item.centralCountId) === id;
  }) || null;
}

function getCentralCountLinesById(centralCountId) {
  var id = safeText(centralCountId);
  return getAllCentralCountLines().filter(function (item) {
    return safeText(item.centralCountId) === id;
  });
}

function buildCentralCountsWithLines(headers, lines) {
  var lineMap = {};

  (lines || []).forEach(function (line) {
    var id = safeText(line.centralCountId);
    if (!lineMap[id]) lineMap[id] = [];
    lineMap[id].push(line);
  });

  return (headers || []).map(function (header) {
    var countLines = lineMap[safeText(header.centralCountId)] || [];
    return Object.assign({}, header, {
      lines: countLines,
      lineCount: countLines.length,
      afwijkingen: countLines.filter(function (line) {
        return safeNumber(line.deltaAantal, 0) !== 0;
      }).length,
    });
  });
}

function getCentralCountsWithLines() {
  return buildCentralCountsWithLines(getAllCentralCountHeaders(), getAllCentralCountLines());
}

function getCentralCountWithLines(centralCountId) {
  var header = getCentralCountHeaderById(centralCountId);
  if (!header) return null;
  return buildCentralCountsWithLines([header], getCentralCountLinesById(centralCountId))[0] || null;
}

/* ---------------------------------------------------------
   Snapshot helpers
   --------------------------------------------------------- */

function buildCurrentCentralSnapshotRows_() {
  if (typeof buildCentralWarehouseStockRows === 'function') {
    return buildCentralWarehouseStockRows();
  }
  return [];
}

function buildCurrentCentralSnapshotMap_() {
  if (typeof buildCentralStockMap === 'function') {
    return buildCentralStockMap();
  }

  var map = {};
  buildCurrentCentralSnapshotRows_().forEach(function (row) {
    map[row.artikelCode] = row;
  });
  return map;
}

function deriveCentralCountScopeRows_(scopeType, articleCodes) {
  var currentRows = buildCurrentCentralSnapshotRows_();

  if (safeText(scopeType) !== getCentralCountScopeTargeted_()) {
    return currentRows;
  }

  var allowed = {};
  (Array.isArray(articleCodes) ? articleCodes : []).forEach(function (code) {
    var key = safeText(code);
    if (key) {
      allowed[key] = true;
    }
  });

  return currentRows.filter(function (row) {
    return !!allowed[safeText(row.artikelCode)];
  });
}

/* ---------------------------------------------------------
   Normalization / validation
   --------------------------------------------------------- */

function normalizeCentralCountHeaderPayload_(payload) {
  payload = payload || {};

  return {
    sessionId: getPayloadSessionId(payload),
    scopeType: safeText(payload.scopeType || payload.scope || getCentralCountScopeFull_()),
    requestedBy: safeText(payload.requestedBy),
    reason: safeText(payload.reason || payload.reden),
    documentDatum: safeText(payload.documentDatum || payload.documentDate),
    opmerking: safeText(payload.opmerking || payload.remark),
    actor: safeText(payload.actor),
  };
}

function normalizeCentralCountLines_(lines) {
  return (Array.isArray(lines) ? lines : []).map(function (line) {
    var systemQty = safeNumber(line.systemAantal || line.systemQty, 0);
    var countedQty = safeNumber(line.geteldAantal || line.countedQty, 0);

    return {
      artikelCode: safeText(line.artikelCode || line.artikelNr),
      artikelOmschrijving: safeText(line.artikelOmschrijving || line.artikel),
      typeMateriaal: safeText(line.typeMateriaal || determineMaterialTypeFromArticle(safeText(line.artikelCode || line.artikelNr))),
      eenheid: safeText(line.eenheid || line.unit || 'Stuk'),
      systemAantal: systemQty,
      geteldAantal: countedQty,
      deltaAantal: safeNumber(line.deltaAantal, countedQty - systemQty),
      opmerking: safeText(line.opmerking),
    };
  });
}

function validateCentralCountHeader_(payload) {
  if (!payload.sessionId) throw new Error('Sessie ontbreekt.');
  if (!payload.documentDatum) throw new Error('Documentdatum is verplicht.');
  return true;
}

function validateCentralCountLines_(lines) {
  if (!lines.length) {
    throw new Error('Geen tellijnen ontvangen.');
  }

  lines.forEach(function (line, index) {
    var rowNr = index + 1;

    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.geteldAantal, 0) < 0) {
      throw new Error('Geteld aantal mag niet negatief zijn op lijn ' + rowNr + '.');
    }
  });

  return true;
}

/* ---------------------------------------------------------
   Access policy
   --------------------------------------------------------- */

function assertCentralCountReadAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.WAREHOUSE, ROLE.MANAGER],
    'Geen rechten voor centrale tellingen.'
  );
  return user;
}

function assertCentralCountWriteAccess_(sessionId) {
  return assertCentralCountReadAccess_(sessionId);
}

function assertCentralCountApproveAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.MANAGER],
    'Geen rechten om centrale tellingen goed te keuren.'
  );
  return user;
}

/* ---------------------------------------------------------
   Create / save
   --------------------------------------------------------- */

function createCentralCountRequest(payload) {
  var normalized = normalizeCentralCountHeaderPayload_(payload);
  var actor = assertCentralCountWriteAccess_(normalized.sessionId);

  validateCentralCountHeader_(normalized);

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var centralCountId = makeCentralCountId_();

  var obj = {
    CentralCountID: centralCountId,
    ScopeType: normalized.scopeType || getCentralCountScopeFull_(),
    RequestedBy: normalized.requestedBy || safeText(actor.naam || actor.email),
    Reason: normalized.reason,
    DocumentDatum: normalized.documentDatum,
    DocumentDatumIso: normalized.documentDatum,
    Status: getCentralCountStatusOpen_(),
    Actor: normalized.actor || safeText(actor.naam || actor.email),
    Opmerking: normalized.opmerking,
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    IngediendOp: '',
    GoedgekeurdOp: '',
  };

  appendObjects(getCentralCountHeaderTab_(), [obj]);

  writeAudit({
    actie: 'CREATE_CENTRAL_COUNT',
    actor: actor,
    documentType: 'CentraleTelling',
    documentId: centralCountId,
    details: {
      scopeType: normalized.scopeType,
      documentDatum: normalized.documentDatum,
    },
  });

  return mapCentralCountHeader(obj);
}

function createCentralCountDraftFromSnapshot(payload) {
  payload = payload || {};

  var normalized = normalizeCentralCountHeaderPayload_(payload);
  var actor = assertCentralCountWriteAccess_(normalized.sessionId);
  var scopeType = safeText(normalized.scopeType || getCentralCountScopeFull_());
  var requestedArticleCodes = Array.isArray(payload.articleCodes) ? payload.articleCodes : [];

  var header = createCentralCountRequest({
    sessionId: normalized.sessionId,
    scopeType: scopeType,
    requestedBy: normalized.requestedBy || safeText(actor.naam || actor.email),
    reason: normalized.reason || 'Centrale telling',
    documentDatum: normalized.documentDatum,
    opmerking: normalized.opmerking,
    actor: normalized.actor || safeText(actor.naam || actor.email),
  });

  var snapshotRows = deriveCentralCountScopeRows_(scopeType, requestedArticleCodes);

  var lines = snapshotRows.map(function (row) {
    return {
      artikelCode: row.artikelCode,
      artikelOmschrijving: row.artikelOmschrijving,
      typeMateriaal: row.typeMateriaal,
      eenheid: row.eenheid,
      systemAantal: safeNumber(row.voorraadCentraal, 0),
      geteldAantal: safeNumber(row.voorraadCentraal, 0),
      deltaAantal: 0,
      opmerking: '',
    };
  });

  return saveCentralCountLines({
    sessionId: normalized.sessionId,
    centralCountId: header.centralCountId,
    lines: lines,
  });
}

function saveCentralCountLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var centralCountId = safeText(payload.centralCountId);
  if (!centralCountId) throw new Error('CentralCountId ontbreekt.');

  var header = getCentralCountHeaderById(centralCountId);
  if (!header) throw new Error('Centrale telling niet gevonden.');

  var actor = assertCentralCountWriteAccess_(sessionId);

  if (!isCentralCountEditable_(header.status)) {
    throw new Error('Centrale telling is niet meer bewerkbaar.');
  }

  var lines = normalizeCentralCountLines_(payload.lines);
  validateCentralCountLines_(lines);

  var table = getAllValues(getCentralCountLineTab_());
  var headerRow = table.length ? table[0] : null;
  var currentRows = readObjectsSafe(getCentralCountLineTab_());

  var kept = currentRows.filter(function (row) {
    return safeText(row.CentralCountID || row.CentralCountId || row.CentraleTellingID) !== centralCountId;
  });

  var newRows = lines.map(function (line) {
    return {
      CentralCountLineID: makeCentralCountLineId_(),
      CentralCountID: centralCountId,
      ArtikelCode: line.artikelCode,
      ArtikelOmschrijving: line.artikelOmschrijving,
      TypeMateriaal: line.typeMateriaal,
      Eenheid: line.eenheid,
      SystemAantal: line.systemAantal,
      GeteldAantal: line.geteldAantal,
      DeltaAantal: line.deltaAantal,
      Opmerking: line.opmerking,
    };
  });

  var finalObjects = kept.concat(newRows);

  if (!headerRow && finalObjects.length) {
    appendObjects(getCentralCountLineTab_(), newRows);
  } else if (headerRow) {
    writeFullTable(
      getCentralCountLineTab_(),
      headerRow,
      finalObjects.map(function (obj) {
        return buildRowFromHeaders(headerRow, obj);
      })
    );
  } else if (newRows.length) {
    appendObjects(getCentralCountLineTab_(), newRows);
  }

  writeAudit({
    actie: 'SAVE_CENTRAL_COUNT_LINES',
    actor: actor,
    documentType: 'CentraleTelling',
    documentId: centralCountId,
    details: {
      lineCount: lines.length,
    },
  });

  return getCentralCountWithLines(centralCountId);
}

/* ---------------------------------------------------------
   Submit / approve
   --------------------------------------------------------- */

function submitCentralCount(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var centralCountId = safeText(payload.centralCountId);
  if (!centralCountId) throw new Error('CentralCountId ontbreekt.');

  var count = getCentralCountWithLines(centralCountId);
  if (!count) throw new Error('Centrale telling niet gevonden.');

  var actor = assertCentralCountWriteAccess_(sessionId);

  if (!isCentralCountEditable_(count.status)) {
    throw new Error('Centrale telling kan niet meer ingediend worden.');
  }

  validateCentralCountLines_(count.lines || []);

  var table = getAllValues(getCentralCountHeaderTab_());
  if (!table.length) throw new Error('Centrale tellingtab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getCentralCountHeaderTab_()).map(function (row) {
    var current = mapCentralCountHeader(row);
    if (safeText(current.centralCountId) !== centralCountId) {
      return row;
    }

    row.Status = getCentralCountStatusSubmitted_();
    row.IngediendOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getCentralCountHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  if (typeof pushManagerNotification === 'function') {
    pushManagerNotification(
      'CentraleTelling',
      'Centrale telling wacht op goedkeuring',
      'Er werd een centrale telling ingediend die goedkeuring vraagt.',
      'CentraleTelling',
      centralCountId,
      'MANAGER'
    );
  }

  writeAudit({
    actie: 'SUBMIT_CENTRAL_COUNT',
    actor: actor,
    documentType: 'CentraleTelling',
    documentId: centralCountId,
    details: {
      afwijkingen: (count.lines || []).filter(function (line) {
        return safeNumber(line.deltaAantal, 0) !== 0;
      }).length,
    },
  });

  return getCentralCountWithLines(centralCountId);
}

function buildCentralCountCorrectionMovements_(count) {
  var header = count || {};
  var lines = header.lines || [];
  var movements = [];

  lines.forEach(function (line) {
    var delta = safeNumber(line.deltaAantal, 0);
    if (delta === 0) {
      return;
    }

    if (delta > 0) {
      movements.push(buildMovementObject({
        movementType: getCentralCountMovementIn_(),
        bronType: 'CentraleTelling',
        bronId: header.centralCountId,
        datumBoeking: header.documentDatum,
        artikelCode: line.artikelCode,
        artikelOmschrijving: line.artikelOmschrijving,
        typeMateriaal: line.typeMateriaal,
        eenheid: line.eenheid,
        aantalIn: delta,
        aantalUit: 0,
        nettoAantal: delta,
        locatieVan: '',
        locatieNaar: LOCATION.CENTRAL,
        reden: 'Centrale telling correctie',
        opmerking: line.opmerking || header.opmerking,
        actor: header.actor,
      }));
      return;
    }

    movements.push(buildMovementObject({
      movementType: getCentralCountMovementOut_(),
      bronType: 'CentraleTelling',
      bronId: header.centralCountId,
      datumBoeking: header.documentDatum,
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      typeMateriaal: line.typeMateriaal,
      eenheid: line.eenheid,
      aantalIn: 0,
      aantalUit: Math.abs(delta),
      nettoAantal: delta,
      locatieVan: LOCATION.CENTRAL,
      locatieNaar: '',
      reden: 'Centrale telling correctie',
      opmerking: line.opmerking || header.opmerking,
      actor: header.actor,
    }));
  });

  return movements;
}

function approveCentralCount(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertCentralCountApproveAccess_(sessionId);
  var centralCountId = safeText(payload.centralCountId);
  if (!centralCountId) throw new Error('CentralCountId ontbreekt.');

  if (typeof replaceSourceMovements !== 'function') {
    throw new Error('Movement service ontbreekt. Werk eerst het movementblok in.');
  }

  var count = getCentralCountWithLines(centralCountId);
  if (!count) throw new Error('Centrale telling niet gevonden.');
  if (safeText(count.status) !== getCentralCountStatusSubmitted_()) {
    throw new Error('Centrale telling staat niet in ingediende status.');
  }

  validateCentralCountLines_(count.lines || []);

  var movements = buildCentralCountCorrectionMovements_(count);
  replaceSourceMovements('CentraleTelling', count.centralCountId, movements);

  var table = getAllValues(getCentralCountHeaderTab_());
  if (!table.length) throw new Error('Centrale tellingtab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getCentralCountHeaderTab_()).map(function (row) {
    var current = mapCentralCountHeader(row);
    if (safeText(current.centralCountId) !== centralCountId) {
      return row;
    }

    row.Status = getCentralCountStatusApproved_();
    row.GoedgekeurdOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getCentralCountHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  if (typeof rebuildCentralWarehouseOverview === 'function') {
    rebuildCentralWarehouseOverview();
  }

  if (typeof markNotificationsProcessedBySource === 'function') {
    markNotificationsProcessedBySource({
      sessionId: sessionId,
      bronType: 'CentraleTelling',
      bronId: centralCountId,
    });
  }

  writeAudit({
    actie: 'APPROVE_CENTRAL_COUNT',
    actor: actor,
    documentType: 'CentraleTelling',
    documentId: centralCountId,
    details: {
      movementCount: movements.length,
      afwijkingen: (count.lines || []).filter(function (line) {
        return safeNumber(line.deltaAantal, 0) !== 0;
      }).length,
    },
  });

  return getCentralCountWithLines(centralCountId);
}

/* ---------------------------------------------------------
   Queries for screens
   --------------------------------------------------------- */

function getCentralCountsData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertCentralCountReadAccess_(sessionId);
  var rows = getCentralCountsWithLines();

  return {
    items: rows,
    centralCounts: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getCentralCountStatusOpen_(); }).length,
      ingediend: rows.filter(function (x) { return safeText(x.status) === getCentralCountStatusSubmitted_(); }).length,
      goedgekeurd: rows.filter(function (x) { return safeText(x.status) === getCentralCountStatusApproved_(); }).length,
      gesloten: rows.filter(function (x) { return safeText(x.status) === getCentralCountStatusClosed_(); }).length,
      afwijkingen: rows.reduce(function (sum, item) {
        return sum + safeNumber(item.afwijkingen, 0);
      }, 0),
      actorRol: safeText(actor.rol),
    }
  };
}