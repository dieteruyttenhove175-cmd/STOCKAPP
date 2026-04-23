/* =========================================================
   43_BusCountService.gs
   Refactor: bus count core service
   Doel:
   - centrale documentlaag voor bustellingen
   - volledige of gerichte tellingen per technieker
   - lijnen bewaren
   - indienen
   - goedkeuren en correctiemutaties boeken
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getBusCountHeaderTab_() {
  return TABS.BUS_COUNTS || 'BusTellingen';
}

function getBusCountLineTab_() {
  return TABS.BUS_COUNT_LINES || 'BusTelLijnen';
}

function getBusCountStatusOpen_() {
  if (typeof BUS_COUNT_STATUS !== 'undefined' && BUS_COUNT_STATUS && BUS_COUNT_STATUS.OPEN) {
    return BUS_COUNT_STATUS.OPEN;
  }
  return 'Open';
}

function getBusCountStatusSubmitted_() {
  if (typeof BUS_COUNT_STATUS !== 'undefined' && BUS_COUNT_STATUS && BUS_COUNT_STATUS.SUBMITTED) {
    return BUS_COUNT_STATUS.SUBMITTED;
  }
  return 'Ingediend';
}

function getBusCountStatusApproved_() {
  if (typeof BUS_COUNT_STATUS !== 'undefined' && BUS_COUNT_STATUS && BUS_COUNT_STATUS.APPROVED) {
    return BUS_COUNT_STATUS.APPROVED;
  }
  return 'Goedgekeurd';
}

function getBusCountStatusClosed_() {
  if (typeof BUS_COUNT_STATUS !== 'undefined' && BUS_COUNT_STATUS && BUS_COUNT_STATUS.CLOSED) {
    return BUS_COUNT_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function getBusCountScopeFull_() {
  return 'FULL';
}

function getBusCountScopeTargeted_() {
  return 'TARGETED';
}

function getBusCountMovementIn_() {
  if (typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE && MOVEMENT_TYPE.BUS_COUNT_IN) {
    return MOVEMENT_TYPE.BUS_COUNT_IN;
  }
  return 'BusCorrectieIn';
}

function getBusCountMovementOut_() {
  if (typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE && MOVEMENT_TYPE.BUS_COUNT_OUT) {
    return MOVEMENT_TYPE.BUS_COUNT_OUT;
  }
  return 'BusCorrectieUit';
}

function isBusCountEditable_(status) {
  var value = safeText(status);
  return !value || value === getBusCountStatusOpen_();
}

function makeBusCountId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'BCT-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function makeBusCountLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'BCL-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapBusCountHeader(row) {
  return {
    busCountId: safeText(row.BusCountID || row.BusCountId || row.BusTellingID || row.ID),
    techniekerCode: safeText(row.TechniekerCode || row.TechnicianCode || row.TechCode),
    techniekerNaam: safeText(row.TechniekerNaam || row.TechnicianName),
    scopeType: safeText(row.ScopeType || getBusCountScopeFull_()),
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

function mapBusCountLine(row) {
  var systemQty = safeNumber(row.SystemAantal || row.SystemQty, 0);
  var countedQty = safeNumber(row.GeteldAantal || row.CountedQty, 0);

  return {
    busCountLineId: safeText(row.BusCountLineID || row.BusCountLineId || row.BusTellingLijnID || row.ID),
    busCountId: safeText(row.BusCountID || row.BusCountId || row.BusTellingID),
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

function getAllBusCountHeaders() {
  return readObjectsSafe(getBusCountHeaderTab_())
    .map(mapBusCountHeader)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.busCountId).localeCompare(safeText(a.busCountId))
      );
    });
}

function getAllBusCountLines() {
  return readObjectsSafe(getBusCountLineTab_())
    .map(mapBusCountLine)
    .sort(function (a, b) {
      return (
        safeText(a.busCountId).localeCompare(safeText(b.busCountId)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
      );
    });
}

function getBusCountHeaderById(busCountId) {
  var id = safeText(busCountId);
  if (!id) return null;

  return getAllBusCountHeaders().find(function (item) {
    return safeText(item.busCountId) === id;
  }) || null;
}

function getBusCountLinesById(busCountId) {
  var id = safeText(busCountId);
  return getAllBusCountLines().filter(function (item) {
    return safeText(item.busCountId) === id;
  });
}

function buildBusCountsWithLines(headers, lines) {
  var lineMap = {};

  (lines || []).forEach(function (line) {
    var id = safeText(line.busCountId);
    if (!lineMap[id]) lineMap[id] = [];
    lineMap[id].push(line);
  });

  return (headers || []).map(function (header) {
    var countLines = lineMap[safeText(header.busCountId)] || [];
    return Object.assign({}, header, {
      lines: countLines,
      lineCount: countLines.length,
      afwijkingen: countLines.filter(function (line) {
        return safeNumber(line.deltaAantal, 0) !== 0;
      }).length,
    });
  });
}

function getBusCountsWithLines() {
  return buildBusCountsWithLines(getAllBusCountHeaders(), getAllBusCountLines());
}

function getBusCountWithLines(busCountId) {
  var header = getBusCountHeaderById(busCountId);
  if (!header) return null;
  return buildBusCountsWithLines([header], getBusCountLinesById(busCountId))[0] || null;
}

/* ---------------------------------------------------------
   Helpers / snapshots
   --------------------------------------------------------- */

function getBusCountTechnicianCodeFromActor_(actor, payloadTechCode) {
  var explicitCode = safeText(payloadTechCode);
  if (explicitCode) return explicitCode;

  return safeText(
    actor &&
    (
      actor.techniekerCode ||
      actor.technicianCode ||
      actor.code
    )
  );
}

function buildCurrentBusSnapshotRows_(techniekerCode) {
  if (typeof buildBusStockRowsForTechnician === 'function') {
    return buildBusStockRowsForTechnician(techniekerCode);
  }
  return [];
}

function buildCurrentBusSnapshotMap_(techniekerCode) {
  if (typeof buildBusStockMapForTechnician === 'function') {
    return buildBusStockMapForTechnician(techniekerCode);
  }

  var map = {};
  buildCurrentBusSnapshotRows_(techniekerCode).forEach(function (row) {
    map[row.artikelCode] = row;
  });
  return map;
}

function deriveBusCountScopeRows_(techniekerCode, scopeType, articleCodes) {
  var currentRows = buildCurrentBusSnapshotRows_(techniekerCode);

  if (safeText(scopeType) !== getBusCountScopeTargeted_()) {
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

function normalizeBusCountHeaderPayload_(payload) {
  payload = payload || {};

  return {
    sessionId: getPayloadSessionId(payload),
    techniekerCode: safeText(payload.techniekerCode || payload.technicianCode || payload.techCode),
    techniekerNaam: safeText(payload.techniekerNaam || payload.technicianName),
    scopeType: safeText(payload.scopeType || payload.scope || getBusCountScopeFull_()),
    requestedBy: safeText(payload.requestedBy),
    reason: safeText(payload.reason || payload.reden),
    documentDatum: safeText(payload.documentDatum || payload.documentDate),
    opmerking: safeText(payload.opmerking || payload.remark),
    actor: safeText(payload.actor),
  };
}

function normalizeBusCountLines_(lines) {
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

function validateBusCountHeader_(payload) {
  if (!payload.sessionId) throw new Error('Sessie ontbreekt.');
  if (!payload.techniekerCode) throw new Error('TechniekerCode is verplicht.');
  if (!payload.documentDatum) throw new Error('Documentdatum is verplicht.');
  return true;
}

function validateBusCountLines_(lines) {
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

function assertBusCountReadAccess_(sessionId, techniekerCode) {
  var user = requireLoggedInUser(sessionId);

  if (roleAllowed(user, [ROLE.MANAGER, ROLE.WAREHOUSE])) {
    return user;
  }

  assertRoleAllowed(user, [ROLE.TECHNICIAN], 'Geen rechten voor bustellingen.');

  var ownCode = getBusCountTechnicianCodeFromActor_(user, '');
  if (safeText(ownCode) !== safeText(techniekerCode)) {
    throw new Error('Geen toegang tot deze bustelling.');
  }

  return user;
}

function assertBusCountWriteAccess_(sessionId, techniekerCode) {
  return assertBusCountReadAccess_(sessionId, techniekerCode);
}

function assertBusCountApproveAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(user, [ROLE.MANAGER, ROLE.WAREHOUSE], 'Geen rechten om bustellingen goed te keuren.');
  return user;
}

/* ---------------------------------------------------------
   Create / save
   --------------------------------------------------------- */

function createBusCountRequest(payload) {
  var normalized = normalizeBusCountHeaderPayload_(payload);
  var actor = assertBusCountWriteAccess_(normalized.sessionId, normalized.techniekerCode);

  validateBusCountHeader_(normalized);

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var busCountId = makeBusCountId_();

  var obj = {
    BusCountID: busCountId,
    TechniekerCode: normalized.techniekerCode,
    TechniekerNaam: normalized.techniekerNaam,
    ScopeType: normalized.scopeType || getBusCountScopeFull_(),
    RequestedBy: normalized.requestedBy || safeText(actor.naam || actor.email),
    Reason: normalized.reason,
    DocumentDatum: normalized.documentDatum,
    DocumentDatumIso: normalized.documentDatum,
    Status: getBusCountStatusOpen_(),
    Actor: normalized.actor || safeText(actor.naam || actor.email),
    Opmerking: normalized.opmerking,
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    IngediendOp: '',
    GoedgekeurdOp: '',
  };

  appendObjects(getBusCountHeaderTab_(), [obj]);

  writeAudit({
    actie: 'CREATE_BUS_COUNT',
    actor: actor,
    documentType: 'BusTelling',
    documentId: busCountId,
    details: {
      techniekerCode: normalized.techniekerCode,
      scopeType: normalized.scopeType,
      documentDatum: normalized.documentDatum,
    },
  });

  return mapBusCountHeader(obj);
}

function createBusCountDraftFromSnapshot(payload) {
  payload = payload || {};

  var normalized = normalizeBusCountHeaderPayload_(payload);
  var actor = assertBusCountWriteAccess_(normalized.sessionId, normalized.techniekerCode);
  var scopeType = safeText(normalized.scopeType || getBusCountScopeFull_());
  var requestedArticleCodes = Array.isArray(payload.articleCodes) ? payload.articleCodes : [];

  var header = createBusCountRequest({
    sessionId: normalized.sessionId,
    techniekerCode: normalized.techniekerCode,
    techniekerNaam: normalized.techniekerNaam,
    scopeType: scopeType,
    requestedBy: normalized.requestedBy || safeText(actor.naam || actor.email),
    reason: normalized.reason || 'Bustelling',
    documentDatum: normalized.documentDatum,
    opmerking: normalized.opmerking,
    actor: normalized.actor || safeText(actor.naam || actor.email),
  });

  var snapshotRows = deriveBusCountScopeRows_(
    normalized.techniekerCode,
    scopeType,
    requestedArticleCodes
  );

  var lines = snapshotRows.map(function (row) {
    return {
      artikelCode: row.artikelCode,
      artikelOmschrijving: row.artikelOmschrijving,
      typeMateriaal: row.typeMateriaal,
      eenheid: row.eenheid,
      systemAantal: safeNumber(row.voorraadBus, 0),
      geteldAantal: safeNumber(row.voorraadBus, 0),
      deltaAantal: 0,
      opmerking: '',
    };
  });

  return saveBusCountLines({
    sessionId: normalized.sessionId,
    busCountId: header.busCountId,
    lines: lines,
  });
}

function saveBusCountLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var busCountId = safeText(payload.busCountId);
  if (!busCountId) throw new Error('BusCountId ontbreekt.');

  var header = getBusCountHeaderById(busCountId);
  if (!header) throw new Error('Bustelling niet gevonden.');

  var actor = assertBusCountWriteAccess_(sessionId, header.techniekerCode);

  if (!isBusCountEditable_(header.status)) {
    throw new Error('Bustelling is niet meer bewerkbaar.');
  }

  var lines = normalizeBusCountLines_(payload.lines);
  validateBusCountLines_(lines);

  var table = getAllValues(getBusCountLineTab_());
  var headerRow = table.length ? table[0] : null;
  var currentRows = readObjectsSafe(getBusCountLineTab_());

  var kept = currentRows.filter(function (row) {
    return safeText(row.BusCountID || row.BusCountId || row.BusTellingID) !== busCountId;
  });

  var newRows = lines.map(function (line) {
    return {
      BusCountLineID: makeBusCountLineId_(),
      BusCountID: busCountId,
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
    appendObjects(getBusCountLineTab_(), newRows);
  } else if (headerRow) {
    writeFullTable(
      getBusCountLineTab_(),
      headerRow,
      finalObjects.map(function (obj) {
        return buildRowFromHeaders(headerRow, obj);
      })
    );
  } else if (newRows.length) {
    appendObjects(getBusCountLineTab_(), newRows);
  }

  writeAudit({
    actie: 'SAVE_BUS_COUNT_LINES',
    actor: actor,
    documentType: 'BusTelling',
    documentId: busCountId,
    details: {
      lineCount: lines.length,
    },
  });

  return getBusCountWithLines(busCountId);
}

/* ---------------------------------------------------------
   Submit / approve
   --------------------------------------------------------- */

function submitBusCount(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var busCountId = safeText(payload.busCountId);
  if (!busCountId) throw new Error('BusCountId ontbreekt.');

  var count = getBusCountWithLines(busCountId);
  if (!count) throw new Error('Bustelling niet gevonden.');

  var actor = assertBusCountWriteAccess_(sessionId, count.techniekerCode);

  if (!isBusCountEditable_(count.status)) {
    throw new Error('Bustelling kan niet meer ingediend worden.');
  }

  validateBusCountLines_(count.lines || []);

  var table = getAllValues(getBusCountHeaderTab_());
  if (!table.length) throw new Error('Bustellingtab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getBusCountHeaderTab_()).map(function (row) {
    var current = mapBusCountHeader(row);
    if (safeText(current.busCountId) !== busCountId) {
      return row;
    }

    row.Status = getBusCountStatusSubmitted_();
    row.IngediendOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getBusCountHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  if (typeof pushWarehouseNotification === 'function') {
    pushWarehouseNotification(
      'BusTelling',
      'Bustelling wacht op goedkeuring',
      'Er werd een bustelling ingediend die goedkeuring vraagt.',
      'BusTelling',
      busCountId,
      'WAREHOUSE'
    );
  }

  writeAudit({
    actie: 'SUBMIT_BUS_COUNT',
    actor: actor,
    documentType: 'BusTelling',
    documentId: busCountId,
    details: {
      techniekerCode: count.techniekerCode,
      afwijkingen: (count.lines || []).filter(function (line) {
        return safeNumber(line.deltaAantal, 0) !== 0;
      }).length,
    },
  });

  return getBusCountWithLines(busCountId);
}

function buildBusCountCorrectionMovements_(count) {
  var header = count || {};
  var lines = header.lines || [];
  var busLocation = typeof getBusLocationCode === 'function'
    ? getBusLocationCode(header.techniekerCode)
    : ('Bus:' + safeText(header.techniekerCode));

  var movements = [];

  lines.forEach(function (line) {
    var delta = safeNumber(line.deltaAantal, 0);
    if (delta === 0) {
      return;
    }

    if (delta > 0) {
      movements.push(buildMovementObject({
        movementType: getBusCountMovementIn_(),
        bronType: 'BusTelling',
        bronId: header.busCountId,
        datumBoeking: header.documentDatum,
        artikelCode: line.artikelCode,
        artikelOmschrijving: line.artikelOmschrijving,
        typeMateriaal: line.typeMateriaal,
        eenheid: line.eenheid,
        aantalIn: delta,
        aantalUit: 0,
        nettoAantal: delta,
        locatieVan: '',
        locatieNaar: busLocation,
        reden: 'Bustelling correctie',
        opmerking: line.opmerking || header.opmerking,
        actor: header.actor,
      }));
      return;
    }

    movements.push(buildMovementObject({
      movementType: getBusCountMovementOut_(),
      bronType: 'BusTelling',
      bronId: header.busCountId,
      datumBoeking: header.documentDatum,
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      typeMateriaal: line.typeMateriaal,
      eenheid: line.eenheid,
      aantalIn: 0,
      aantalUit: Math.abs(delta),
      nettoAantal: delta,
      locatieVan: busLocation,
      locatieNaar: '',
      reden: 'Bustelling correctie',
      opmerking: line.opmerking || header.opmerking,
      actor: header.actor,
    }));
  });

  return movements;
}

function approveBusCount(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertBusCountApproveAccess_(sessionId);
  var busCountId = safeText(payload.busCountId);
  if (!busCountId) throw new Error('BusCountId ontbreekt.');

  if (typeof replaceSourceMovements !== 'function') {
    throw new Error('Movement service ontbreekt. Werk eerst het movementblok in.');
  }

  var count = getBusCountWithLines(busCountId);
  if (!count) throw new Error('Bustelling niet gevonden.');
  if (safeText(count.status) !== getBusCountStatusSubmitted_()) {
    throw new Error('Bustelling staat niet in ingediende status.');
  }

  validateBusCountLines_(count.lines || []);

  var movements = buildBusCountCorrectionMovements_(count);
  replaceSourceMovements('BusTelling', count.busCountId, movements);

  var table = getAllValues(getBusCountHeaderTab_());
  if (!table.length) throw new Error('Bustellingtab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getBusCountHeaderTab_()).map(function (row) {
    var current = mapBusCountHeader(row);
    if (safeText(current.busCountId) !== busCountId) {
      return row;
    }

    row.Status = getBusCountStatusApproved_();
    row.GoedgekeurdOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getBusCountHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  if (typeof markNotificationsProcessedBySource === 'function') {
    markNotificationsProcessedBySource({
      sessionId: sessionId,
      bronType: 'BusTelling',
      bronId: busCountId,
    });
  }

  writeAudit({
    actie: 'APPROVE_BUS_COUNT',
    actor: actor,
    documentType: 'BusTelling',
    documentId: busCountId,
    details: {
      movementCount: movements.length,
      techniekerCode: count.techniekerCode,
      afwijkingen: (count.lines || []).filter(function (line) {
        return safeNumber(line.deltaAantal, 0) !== 0;
      }).length,
    },
  });

  return getBusCountWithLines(busCountId);
}

/* ---------------------------------------------------------
   Queries for screens
   --------------------------------------------------------- */

function filterBusCountsForUser_(rows, user) {
  if (roleAllowed(user, [ROLE.MANAGER, ROLE.WAREHOUSE])) {
    return rows;
  }

  if (roleAllowed(user, [ROLE.TECHNICIAN])) {
    var ownCode = getBusCountTechnicianCodeFromActor_(user, '');
    return rows.filter(function (item) {
      return safeText(item.techniekerCode) === safeText(ownCode);
    });
  }

  return [];
}

function getBusCountsData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  assertRoleAllowed(
    actor,
    [ROLE.TECHNICIAN, ROLE.WAREHOUSE, ROLE.MANAGER],
    'Geen rechten om bustellingen te bekijken.'
  );

  var rows = filterBusCountsForUser_(getBusCountsWithLines(), actor);

  return {
    items: rows,
    busCounts: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getBusCountStatusOpen_(); }).length,
      ingediend: rows.filter(function (x) { return safeText(x.status) === getBusCountStatusSubmitted_(); }).length,
      goedgekeurd: rows.filter(function (x) { return safeText(x.status) === getBusCountStatusApproved_(); }).length,
      gesloten: rows.filter(function (x) { return safeText(x.status) === getBusCountStatusClosed_(); }).length,
      afwijkingen: rows.reduce(function (sum, item) {
        return sum + safeNumber(item.afwijkingen, 0);
      }, 0),
    }
  };
}