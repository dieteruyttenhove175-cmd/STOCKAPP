/* =========================================================
   41_ReturnService.gs
   Refactor: return core service
   Doel:
   - centrale documentlaag voor retouren
   - technieker -> centraal
   - centraal -> Fluvius
   - lijnen bewaren
   - indienen
   - goedkeuren en stockmutaties boeken
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getReturnHeaderTab_() {
  return TABS.RETURNS || 'Retouren';
}

function getReturnLineTab_() {
  return TABS.RETURN_LINES || 'RetourLijnen';
}

function getReturnStatusOpen_() {
  if (typeof RETURN_STATUS !== 'undefined' && RETURN_STATUS && RETURN_STATUS.OPEN) {
    return RETURN_STATUS.OPEN;
  }
  return 'Open';
}

function getReturnStatusSubmitted_() {
  if (typeof RETURN_STATUS !== 'undefined' && RETURN_STATUS && RETURN_STATUS.SUBMITTED) {
    return RETURN_STATUS.SUBMITTED;
  }
  return 'Ingediend';
}

function getReturnStatusApproved_() {
  if (typeof RETURN_STATUS !== 'undefined' && RETURN_STATUS && RETURN_STATUS.APPROVED) {
    return RETURN_STATUS.APPROVED;
  }
  return 'Goedgekeurd';
}

function getReturnStatusClosed_() {
  if (typeof RETURN_STATUS !== 'undefined' && RETURN_STATUS && RETURN_STATUS.CLOSED) {
    return RETURN_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function getReturnFlowBusToCentral_() {
  return 'BUS_TO_CENTRAL';
}

function getReturnFlowCentralToFluvius_() {
  return 'CENTRAL_TO_FLUVIUS';
}

function isReturnEditable_(status) {
  var value = safeText(status);
  return !value || value === getReturnStatusOpen_();
}

function makeReturnId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'RTN-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function makeReturnLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'RTL-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapReturnHeader(row) {
  return {
    returnId: safeText(row.ReturnID || row.ReturnId || row.RetourID || row.ID),
    flowType: safeText(row.FlowType || row.ReturnFlowType || getReturnFlowBusToCentral_()),
    leverancier: safeText(row.Leverancier || 'Fluvius'),
    vanLocatie: safeText(row.VanLocatie),
    naarLocatie: safeText(row.NaarLocatie),
    techniekerCode: safeText(row.TechniekerCode || row.TechnicianCode),
    typeMateriaal: safeText(row.TypeMateriaal),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    status: safeText(row.Status),
    reden: safeText(row.Reden),
    actor: safeText(row.Actor),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    ingediendOp: safeText(row.IngediendOp),
    goedgekeurdOp: safeText(row.GoedgekeurdOp),
    opmerking: safeText(row.Opmerking),
    deltaLijnen: safeNumber(row.DeltaLijnen, 0),
  };
}

function mapReturnLine(row) {
  var qty = safeNumber(row.Aantal || row.Quantity, 0);

  return {
    returnLineId: safeText(row.ReturnLineID || row.ReturnLineId || row.RetourLijnID || row.ID),
    returnId: safeText(row.ReturnID || row.ReturnId || row.RetourID),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode || row.ArtikelNr))),
    eenheid: safeText(row.Eenheid || row.Unit),
    aantal: qty,
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllReturnHeaders() {
  return readObjectsSafe(getReturnHeaderTab_())
    .map(mapReturnHeader)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.returnId).localeCompare(safeText(a.returnId))
      );
    });
}

function getAllReturnLines() {
  return readObjectsSafe(getReturnLineTab_())
    .map(mapReturnLine)
    .sort(function (a, b) {
      return (
        safeText(a.returnId).localeCompare(safeText(b.returnId)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
      );
    });
}

function getReturnHeaderById(returnId) {
  var id = safeText(returnId);
  if (!id) return null;

  return getAllReturnHeaders().find(function (item) {
    return safeText(item.returnId) === id;
  }) || null;
}

function getReturnLinesByReturnId(returnId) {
  var id = safeText(returnId);
  return getAllReturnLines().filter(function (item) {
    return safeText(item.returnId) === id;
  });
}

function buildReturnsWithLines(headers, lines) {
  var lineMap = {};

  (lines || []).forEach(function (line) {
    var id = safeText(line.returnId);
    if (!lineMap[id]) lineMap[id] = [];
    lineMap[id].push(line);
  });

  return (headers || []).map(function (header) {
    var returnLines = lineMap[safeText(header.returnId)] || [];

    return Object.assign({}, header, {
      lines: returnLines,
      lineCount: returnLines.length,
      totaalAantal: returnLines.reduce(function (sum, line) {
        return sum + safeNumber(line.aantal, 0);
      }, 0),
    });
  });
}

function getReturnsWithLines() {
  return buildReturnsWithLines(getAllReturnHeaders(), getAllReturnLines());
}

function getReturnWithLines(returnId) {
  var header = getReturnHeaderById(returnId);
  if (!header) return null;
  return buildReturnsWithLines([header], getReturnLinesByReturnId(returnId))[0] || null;
}

/* ---------------------------------------------------------
   Helpers flow / location
   --------------------------------------------------------- */

function buildBusLocationFromTechnician_(techCode) {
  if (typeof getBusLocationCode === 'function') {
    return getBusLocationCode(techCode);
  }
  return 'Bus:' + safeText(techCode);
}

function getFluviusReturnLocation_() {
  return 'Fluvius';
}

function normalizeReturnHeaderPayload_(payload) {
  payload = payload || {};

  return {
    sessionId: getPayloadSessionId(payload),
    flowType: safeText(payload.flowType || payload.returnFlowType || getReturnFlowBusToCentral_()),
    leverancier: safeText(payload.leverancier || 'Fluvius'),
    vanLocatie: safeText(payload.vanLocatie || payload.fromLocation),
    naarLocatie: safeText(payload.naarLocatie || payload.toLocation),
    techniekerCode: safeText(payload.techniekerCode || payload.technicianCode || payload.techCode),
    typeMateriaal: safeText(payload.typeMateriaal || payload.materialType),
    documentDatum: safeText(payload.documentDatum || payload.documentDate),
    reden: safeText(payload.reden || payload.reason),
    opmerking: safeText(payload.opmerking || payload.remark),
    actor: safeText(payload.actor),
  };
}

function normalizeReturnLines_(lines) {
  return (Array.isArray(lines) ? lines : []).map(function (line) {
    return {
      artikelCode: safeText(line.artikelCode || line.artikelNr),
      artikelOmschrijving: safeText(line.artikelOmschrijving || line.artikel),
      typeMateriaal: safeText(line.typeMateriaal || determineMaterialTypeFromArticle(safeText(line.artikelCode || line.artikelNr))),
      eenheid: safeText(line.eenheid || line.unit || 'Stuk'),
      aantal: safeNumber(line.aantal || line.quantity, 0),
      reden: safeText(line.reden || line.reason),
      opmerking: safeText(line.opmerking),
    };
  });
}

function deriveReturnLocations_(normalized, actor) {
  var flowType = safeText(normalized.flowType);

  if (flowType === getReturnFlowBusToCentral_()) {
    var techCode = safeText(normalized.techniekerCode || actor.techniekerCode || actor.technicianCode || actor.code);
    if (!techCode) {
      throw new Error('TechniekerCode is verplicht voor bus -> centraal retour.');
    }

    return {
      flowType: flowType,
      techniekerCode: techCode,
      vanLocatie: normalized.vanLocatie || buildBusLocationFromTechnician_(techCode),
      naarLocatie: normalized.naarLocatie || LOCATION.CENTRAL,
    };
  }

  if (flowType === getReturnFlowCentralToFluvius_()) {
    return {
      flowType: flowType,
      techniekerCode: '',
      vanLocatie: normalized.vanLocatie || LOCATION.CENTRAL,
      naarLocatie: normalized.naarLocatie || getFluviusReturnLocation_(),
    };
  }

  throw new Error('Onbekend retourflowtype.');
}

function validateReturnHeader_(payload) {
  if (!payload.sessionId) throw new Error('Sessie ontbreekt.');
  if (!payload.flowType) throw new Error('FlowType is verplicht.');
  if (!payload.documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!payload.reden) throw new Error('Reden is verplicht.');
}

function validateReturnLines_(lines) {
  if (!lines.length) {
    throw new Error('Geen retourlijnen ontvangen.');
  }

  lines.forEach(function (line, index) {
    var rowNr = index + 1;

    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.aantal, 0) <= 0) {
      throw new Error('Aantal moet groter zijn dan 0 op lijn ' + rowNr + '.');
    }
    if (!safeText(line.reden)) {
      throw new Error('Reden is verplicht op lijn ' + rowNr + '.');
    }
  });

  return true;
}

function deriveReturnMaterialType_(lines) {
  var types = {};
  (lines || []).forEach(function (line) {
    var t = safeText(line.typeMateriaal);
    if (t) {
      types[t] = true;
    }
  });

  var keys = Object.keys(types);
  if (!keys.length) return '';
  if (keys.length === 1) return keys[0];
  return 'Gemengd';
}

/* ---------------------------------------------------------
   Stock validation
   --------------------------------------------------------- */

function getReturnAvailableQty_(locationCode, artikelCode) {
  var code = safeText(locationCode);
  var article = safeText(artikelCode);
  if (!code || !article) return 0;

  if (safeText(code) === safeText(LOCATION.CENTRAL) && typeof buildCentralStockMap === 'function') {
    var centralMap = buildCentralStockMap();
    return centralMap[article] ? safeNumber(centralMap[article].voorraadCentraal, 0) : 0;
  }

  if (typeof isBusLocation === 'function' && isBusLocation(code) &&
      typeof buildBusStockMapForTechnician === 'function' && typeof parseBusLocation === 'function') {
    var techCode = parseBusLocation(code);
    var busMap = buildBusStockMapForTechnician(techCode);
    return busMap[article] ? safeNumber(busMap[article].voorraadBus, 0) : 0;
  }

  return 0;
}

function validateReturnSourceStock_(header, lines) {
  var sourceLocation = safeText(header.vanLocatie);

  lines.forEach(function (line) {
    var code = safeText(line.artikelCode);
    var requested = safeNumber(line.aantal, 0);
    var available = getReturnAvailableQty_(sourceLocation, code);

    if (requested > available) {
      throw new Error(
        'Onvoldoende voorraad op bronlocatie voor artikel ' + code +
        '. Beschikbaar: ' + available + ', gevraagd: ' + requested + '.'
      );
    }
  });
}

/* ---------------------------------------------------------
   Access policy
   --------------------------------------------------------- */

function assertReturnCreateAccess_(sessionId, flowType) {
  var user = requireLoggedInUser(sessionId);
  var flow = safeText(flowType);

  if (flow === getReturnFlowBusToCentral_()) {
    assertRoleAllowed(
      user,
      [ROLE.TECHNICIAN, ROLE.WAREHOUSE, ROLE.MANAGER],
      'Geen rechten voor bus -> centraal retour.'
    );
    return user;
  }

  if (flow === getReturnFlowCentralToFluvius_()) {
    assertRoleAllowed(
      user,
      [ROLE.WAREHOUSE, ROLE.MANAGER],
      'Geen rechten voor centraal -> Fluvius retour.'
    );
    return user;
  }

  throw new Error('Onbekend retourflowtype.');
}

function assertReturnReadAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.TECHNICIAN, ROLE.WAREHOUSE, ROLE.MANAGER],
    'Geen rechten voor retouren.'
  );
  return user;
}

/* ---------------------------------------------------------
   Create / save
   --------------------------------------------------------- */

function createReturn(payload) {
  var normalized = normalizeReturnHeaderPayload_(payload);
  validateReturnHeader_(normalized);

  var actor = assertReturnCreateAccess_(normalized.sessionId, normalized.flowType);
  var flow = deriveReturnLocations_(normalized, actor);

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var returnId = makeReturnId_();

  var obj = {
    ReturnID: returnId,
    FlowType: flow.flowType,
    Leverancier: normalized.leverancier || 'Fluvius',
    VanLocatie: flow.vanLocatie,
    NaarLocatie: flow.naarLocatie,
    TechniekerCode: flow.techniekerCode,
    TypeMateriaal: normalized.typeMateriaal,
    DocumentDatum: normalized.documentDatum,
    DocumentDatumIso: normalized.documentDatum,
    Status: getReturnStatusOpen_(),
    Reden: normalized.reden,
    Actor: normalized.actor || safeText(actor.naam || actor.email),
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    IngediendOp: '',
    GoedgekeurdOp: '',
    Opmerking: normalized.opmerking,
    DeltaLijnen: 0,
  };

  appendObjects(getReturnHeaderTab_(), [obj]);

  writeAudit({
    actie: 'CREATE_RETURN',
    actor: actor,
    documentType: 'Retour',
    documentId: returnId,
    details: {
      flowType: obj.FlowType,
      vanLocatie: obj.VanLocatie,
      naarLocatie: obj.NaarLocatie,
      documentDatum: obj.DocumentDatum,
    },
  });

  return mapReturnHeader(obj);
}

function saveReturnLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertReturnReadAccess_(sessionId);
  var returnId = safeText(payload.returnId);
  if (!returnId) throw new Error('ReturnId ontbreekt.');

  var header = getReturnHeaderById(returnId);
  if (!header) throw new Error('Retour niet gevonden.');
  if (!isReturnEditable_(header.status)) {
    throw new Error('Retour is niet meer bewerkbaar.');
  }

  validateReturnHeader_({
    sessionId: sessionId,
    flowType: header.flowType,
    documentDatum: header.documentDatum,
    reden: header.reden,
  });

  var lines = normalizeReturnLines_(payload.lines);
  validateReturnLines_(lines);

  var table = getAllValues(getReturnLineTab_());
  var headerRow = table.length ? table[0] : null;
  var currentRows = readObjectsSafe(getReturnLineTab_());

  var kept = currentRows.filter(function (row) {
    return safeText(row.ReturnID || row.ReturnId || row.RetourID) !== returnId;
  });

  var newRows = lines.map(function (line) {
    return {
      ReturnLineID: makeReturnLineId_(),
      ReturnID: returnId,
      ArtikelCode: line.artikelCode,
      ArtikelOmschrijving: line.artikelOmschrijving,
      TypeMateriaal: line.typeMateriaal,
      Eenheid: line.eenheid,
      Aantal: line.aantal,
      Reden: line.reden,
      Opmerking: line.opmerking,
    };
  });

  var finalObjects = kept.concat(newRows);

  if (!headerRow && finalObjects.length) {
    appendObjects(getReturnLineTab_(), newRows);
  } else if (headerRow) {
    writeFullTable(
      getReturnLineTab_(),
      headerRow,
      finalObjects.map(function (obj) {
        return buildRowFromHeaders(headerRow, obj);
      })
    );
  } else {
    appendObjects(getReturnLineTab_(), newRows);
  }

  updateReturnDerivedFields_(returnId, {
    typeMateriaal: deriveReturnMaterialType_(lines),
    deltaLijnen: 0,
  });

  writeAudit({
    actie: 'SAVE_RETURN_LINES',
    actor: actor,
    documentType: 'Retour',
    documentId: returnId,
    details: {
      lineCount: lines.length,
    },
  });

  return getReturnWithLines(returnId);
}

function updateReturnDerivedFields_(returnId, values) {
  var table = getAllValues(getReturnHeaderTab_());
  if (!table.length) throw new Error('Retourtab is leeg of ongeldig.');

  var headerRow = table[0];
  var rows = readObjectsSafe(getReturnHeaderTab_()).map(function (row) {
    var current = mapReturnHeader(row);
    if (safeText(current.returnId) !== safeText(returnId)) {
      return row;
    }

    row.TypeMateriaal = safeText(values.typeMateriaal || row.TypeMateriaal);
    row.DeltaLijnen = safeNumber(values.deltaLijnen, 0);
    return row;
  });

  writeFullTable(
    getReturnHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );
}

/* ---------------------------------------------------------
   Submit / approve
   --------------------------------------------------------- */

function submitReturn(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertReturnReadAccess_(sessionId);
  var returnId = safeText(payload.returnId);
  if (!returnId) throw new Error('ReturnId ontbreekt.');

  var header = getReturnHeaderById(returnId);
  if (!header) throw new Error('Retour niet gevonden.');
  if (!isReturnEditable_(header.status)) {
    throw new Error('Retour kan niet meer ingediend worden.');
  }

  var lines = getReturnLinesByReturnId(returnId);
  validateReturnLines_(lines);
  validateReturnSourceStock_(header, lines);

  var table = getAllValues(getReturnHeaderTab_());
  if (!table.length) throw new Error('Retourtab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getReturnHeaderTab_()).map(function (row) {
    var current = mapReturnHeader(row);
    if (safeText(current.returnId) !== returnId) {
      return row;
    }

    row.Status = getReturnStatusSubmitted_();
    row.IngediendOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getReturnHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  if (typeof pushManagerNotification === 'function') {
    pushManagerNotification(
      'Retour',
      'Retour wacht op goedkeuring',
      'Er werd een retour ingediend die goedkeuring vraagt.',
      'Retour',
      returnId,
      'MANAGER'
    );
  }

  writeAudit({
    actie: 'SUBMIT_RETURN',
    actor: actor,
    documentType: 'Retour',
    documentId: returnId,
    details: {
      lineCount: lines.length,
      flowType: header.flowType,
    },
  });

  return getReturnWithLines(returnId);
}

/* ---------------------------------------------------------
   Movement mapping
   --------------------------------------------------------- */

function getReturnMovementTypes_(flowType) {
  var flow = safeText(flowType);

  if (typeof MOVEMENT_TYPE === 'undefined') {
    return {
      outType: 'ReturnOut',
      inType: 'ReturnIn',
    };
  }

  if (flow === getReturnFlowBusToCentral_()) {
    return {
      outType: MOVEMENT_TYPE.RETURN_OUT || 'ReturnOut',
      inType: MOVEMENT_TYPE.RETURN_IN || 'ReturnIn',
    };
  }

  if (flow === getReturnFlowCentralToFluvius_()) {
    return {
      outType: MOVEMENT_TYPE.RETURN_OUT || 'ReturnOut',
      inType: '',
    };
  }

  return {
    outType: MOVEMENT_TYPE.RETURN_OUT || 'ReturnOut',
    inType: MOVEMENT_TYPE.RETURN_IN || 'ReturnIn',
  };
}

function buildReturnMovements_(ret) {
  var header = ret || {};
  var lines = header.lines || [];
  var movementTypes = getReturnMovementTypes_(header.flowType);

  var outMovements = lines.map(function (line) {
    return buildMovementObject({
      movementType: movementTypes.outType,
      bronType: 'Retour',
      bronId: header.returnId,
      datumBoeking: header.documentDatum,
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      typeMateriaal: line.typeMateriaal,
      eenheid: line.eenheid,
      aantalUit: line.aantal,
      aantalIn: 0,
      nettoAantal: -safeNumber(line.aantal, 0),
      locatieVan: header.vanLocatie,
      locatieNaar: header.naarLocatie,
      reden: line.reden || header.reden,
      opmerking: line.opmerking || header.opmerking,
      actor: header.actor,
    });
  });

  if (!movementTypes.inType) {
    return outMovements;
  }

  var inMovements = lines.map(function (line) {
    return buildMovementObject({
      movementType: movementTypes.inType,
      bronType: 'Retour',
      bronId: header.returnId,
      datumBoeking: header.documentDatum,
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      typeMateriaal: line.typeMateriaal,
      eenheid: line.eenheid,
      aantalUit: 0,
      aantalIn: line.aantal,
      nettoAantal: safeNumber(line.aantal, 0),
      locatieVan: header.vanLocatie,
      locatieNaar: header.naarLocatie,
      reden: line.reden || header.reden,
      opmerking: line.opmerking || header.opmerking,
      actor: header.actor,
    });
  });

  return outMovements.concat(inMovements);
}

function approveReturn(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertManagerAccess(sessionId);
  var returnId = safeText(payload.returnId);
  if (!returnId) throw new Error('ReturnId ontbreekt.');

  if (typeof replaceSourceMovements !== 'function') {
    throw new Error('Movement service ontbreekt. Werk eerst het movementblok in.');
  }

  var ret = getReturnWithLines(returnId);
  if (!ret) throw new Error('Retour niet gevonden.');
  if (safeText(ret.status) !== getReturnStatusSubmitted_()) {
    throw new Error('Retour staat niet in ingediende status.');
  }

  validateReturnLines_(ret.lines || []);
  validateReturnSourceStock_(ret, ret.lines || []);

  var movements = buildReturnMovements_(ret);
  replaceSourceMovements('Retour', ret.returnId, movements);

  var table = getAllValues(getReturnHeaderTab_());
  if (!table.length) throw new Error('Retourtab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getReturnHeaderTab_()).map(function (row) {
    var current = mapReturnHeader(row);
    if (safeText(current.returnId) !== returnId) {
      return row;
    }

    row.Status = getReturnStatusApproved_();
    row.GoedgekeurdOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getReturnHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  if (safeText(ret.vanLocatie) === safeText(LOCATION.CENTRAL) &&
      typeof rebuildCentralWarehouseOverview === 'function') {
    rebuildCentralWarehouseOverview();
  }
  if (safeText(ret.naarLocatie) === safeText(LOCATION.CENTRAL) &&
      typeof rebuildCentralWarehouseOverview === 'function') {
    rebuildCentralWarehouseOverview();
  }

  if (typeof markNotificationsProcessedBySource === 'function') {
    markNotificationsProcessedBySource({
      sessionId: sessionId,
      bronType: 'Retour',
      bronId: returnId,
    });
  }

  writeAudit({
    actie: 'APPROVE_RETURN',
    actor: actor,
    documentType: 'Retour',
    documentId: returnId,
    details: {
      movementCount: movements.length,
      flowType: ret.flowType,
      totalReturned: (ret.lines || []).reduce(function (sum, line) {
        return sum + safeNumber(line.aantal, 0);
      }, 0),
    },
  });

  return getReturnWithLines(returnId);
}

/* ---------------------------------------------------------
   Queries for screens
   --------------------------------------------------------- */

function filterReturnsForUser_(rows, user) {
  if (roleAllowed(user, [ROLE.MANAGER, ROLE.WAREHOUSE])) {
    return rows;
  }

  if (roleAllowed(user, [ROLE.TECHNICIAN])) {
    var ownTechCode = safeText(user.techniekerCode || user.technicianCode || user.code);
    return rows.filter(function (item) {
      return safeText(item.techniekerCode) === ownTechCode;
    });
  }

  return [];
}

function getReturnsData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertReturnReadAccess_(sessionId);
  var rows = filterReturnsForUser_(getReturnsWithLines(), actor);

  return {
    items: rows,
    returns: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getReturnStatusOpen_(); }).length,
      ingediend: rows.filter(function (x) { return safeText(x.status) === getReturnStatusSubmitted_(); }).length,
      goedgekeurd: rows.filter(function (x) { return safeText(x.status) === getReturnStatusApproved_(); }).length,
      gesloten: rows.filter(function (x) { return safeText(x.status) === getReturnStatusClosed_(); }).length,
      busNaarCentraal: rows.filter(function (x) { return safeText(x.flowType) === getReturnFlowBusToCentral_(); }).length,
      centraalNaarFluvius: rows.filter(function (x) { return safeText(x.flowType) === getReturnFlowCentralToFluvius_(); }).length,
    }
  };
}
