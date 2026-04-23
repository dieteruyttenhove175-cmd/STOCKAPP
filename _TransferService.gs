/* =========================================================
   47_TransferService.gs
   Refactor: transfer core service
   Doel:
   - één duidelijke documentlaag voor transfers
   - leeslaag + lijnlaag + submitlaag
   - basis waarop mobiel magazijn, centraal magazijn en busflows kunnen steunen
   ========================================================= */

/* ---------------------------------------------------------
   Constants / fallbacks
   --------------------------------------------------------- */

var TRANSFER_FLOW = {
  CENTRAL_TO_MOBILE: 'CENTRAL_TO_MOBILE',
  MOBILE_TO_CENTRAL: 'MOBILE_TO_CENTRAL',
  MOBILE_TO_BUS: 'MOBILE_TO_BUS',
  CENTRAL_TO_BUS: 'CENTRAL_TO_BUS',
  BUS_TO_CENTRAL: 'BUS_TO_CENTRAL',
  BUS_TO_MOBILE: 'BUS_TO_MOBILE'
};

function getTransferHeaderTab_() {
  return TABS.TRANSFERS || 'Transfers';
}

function getTransferLineTab_() {
  return TABS.TRANSFER_LINES || 'TransferLijnen';
}

/* ---------------------------------------------------------
   ID / mapping
   --------------------------------------------------------- */

function makeTransferId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'TRF-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function makeTransferLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'TRFL-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function mapTransferHeader(row) {
  return {
    transferId: safeText(row.TransferID || row.TransferId || row.ID),
    flowType: safeText(row.FlowType),
    vanLocatie: safeText(row.VanLocatie),
    naarLocatie: safeText(row.NaarLocatie),
    doelTechniekerCode: safeText(row.DoelTechniekerCode),
    bronReferentie: safeText(row.BronReferentie),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    status: safeText(row.Status),
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),
    actor: safeText(row.Actor),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    ingediendOp: safeText(row.IngediendOp),
    goedgekeurdOp: safeText(row.GoedgekeurdOp),
    geboektOp: safeText(row.GeboektOp),
    mobileWarehouseCode: safeText(row.MobileWarehouseCode),
    createdByRole: safeText(row.CreatedByRole),
  };
}

function mapTransferLine(row) {
  return {
    transferLineId: safeText(row.TransferLineID || row.TransferLineId || row.ID),
    transferId: safeText(row.TransferID || row.TransferId),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode))),
    eenheid: safeText(row.Eenheid),
    aantal: safeNumber(row.Aantal, 0),
    opmerking: safeText(row.Opmerking),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllTransfers() {
  return readObjectsSafe(getTransferHeaderTab_())
    .map(mapTransferHeader)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.transferId).localeCompare(safeText(a.transferId))
      );
    });
}

function getAllTransferLines() {
  return readObjectsSafe(getTransferLineTab_())
    .map(mapTransferLine)
    .sort(function (a, b) {
      return (
        safeText(a.transferId).localeCompare(safeText(b.transferId)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
      );
    });
}

function getTransferById(transferId) {
  var id = safeText(transferId);
  if (!id) return null;
  return getAllTransfers().find(function (item) {
    return safeText(item.transferId) === id;
  }) || null;
}

function getTransferLinesByTransferId(transferId) {
  var id = safeText(transferId);
  return getAllTransferLines().filter(function (item) {
    return safeText(item.transferId) === id;
  });
}

function buildTransfersWithLines(headers, lines) {
  var lineMap = {};

  (lines || []).forEach(function (line) {
    var id = safeText(line.transferId);
    if (!lineMap[id]) lineMap[id] = [];
    lineMap[id].push(line);
  });

  return (headers || []).map(function (header) {
    var transferLines = lineMap[safeText(header.transferId)] || [];
    return Object.assign({}, header, {
      lines: transferLines,
      lineCount: transferLines.length,
      totaalAantal: transferLines.reduce(function (sum, line) {
        return sum + safeNumber(line.aantal, 0);
      }, 0),
    });
  });
}

function getTransfersWithLines() {
  return buildTransfersWithLines(getAllTransfers(), getAllTransferLines());
}

function getTransferWithLines(transferId) {
  var header = getTransferById(transferId);
  if (!header) return null;
  var lines = getTransferLinesByTransferId(transferId);
  return buildTransfersWithLines([header], lines)[0] || null;
}

/* ---------------------------------------------------------
   Flow / status helpers
   --------------------------------------------------------- */

function normalizeTransferPayload_(payload) {
  payload = payload || {};
  return {
    sessionId: getPayloadSessionId(payload),
    flowType: safeText(payload.flowType || payload.FlowType),
    vanLocatie: safeText(payload.vanLocatie || payload.fromLocation),
    naarLocatie: safeText(payload.naarLocatie || payload.toLocation),
    doelTechniekerCode: safeText(payload.doelTechniekerCode || payload.technicianCode || payload.techCode),
    bronReferentie: safeText(payload.bronReferentie || payload.sourceReference),
    documentDatum: safeText(payload.documentDatum || payload.documentDate),
    reden: safeText(payload.reden || payload.reason),
    opmerking: safeText(payload.opmerking || payload.remark),
    actor: safeText(payload.actor),
    mobileWarehouseCode: safeText(payload.mobileWarehouseCode),
  };
}

function normalizeTransferLines_(lines) {
  return (Array.isArray(lines) ? lines : []).map(function (line) {
    return {
      artikelCode: safeText(line.artikelCode),
      artikelOmschrijving: safeText(line.artikelOmschrijving),
      typeMateriaal: safeText(line.typeMateriaal || determineMaterialTypeFromArticle(safeText(line.artikelCode))),
      eenheid: safeText(line.eenheid || 'Stuk'),
      aantal: safeNumber(line.aantal, 0),
      opmerking: safeText(line.opmerking),
    };
  });
}

function getDefaultTransferStatus_() {
  if (typeof TRANSFER_STATUS !== 'undefined' && TRANSFER_STATUS && TRANSFER_STATUS.OPEN) {
    return TRANSFER_STATUS.OPEN;
  }
  return 'Open';
}

function getSubmittedTransferStatus_() {
  if (typeof TRANSFER_STATUS !== 'undefined' && TRANSFER_STATUS && TRANSFER_STATUS.SUBMITTED) {
    return TRANSFER_STATUS.SUBMITTED;
  }
  return 'Ingediend';
}

function getApprovedTransferStatus_() {
  if (typeof TRANSFER_STATUS !== 'undefined' && TRANSFER_STATUS && TRANSFER_STATUS.APPROVED) {
    return TRANSFER_STATUS.APPROVED;
  }
  return 'Goedgekeurd';
}

function getBookedTransferStatus_() {
  if (typeof TRANSFER_STATUS !== 'undefined' && TRANSFER_STATUS && TRANSFER_STATUS.BOOKED) {
    return TRANSFER_STATUS.BOOKED;
  }
  return 'Geboekt';
}

function isTransferEditable_(status) {
  var value = safeText(status);
  return !value || value === getDefaultTransferStatus_();
}

/* ---------------------------------------------------------
   Validatie
   --------------------------------------------------------- */

function validateTransferHeader_(payload) {
  if (!payload.sessionId) throw new Error('Sessie ontbreekt.');
  if (!payload.flowType) throw new Error('FlowType is verplicht.');
  if (!payload.vanLocatie) throw new Error('VanLocatie is verplicht.');
  if (!payload.naarLocatie) throw new Error('NaarLocatie is verplicht.');
  if (!payload.documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!payload.reden) throw new Error('Reden is verplicht.');
  if (payload.vanLocatie === payload.naarLocatie) {
    throw new Error('VanLocatie en NaarLocatie mogen niet gelijk zijn.');
  }
}

function validateTransferLines_(lines) {
  if (!lines.length) throw new Error('Geen transferlijnen ontvangen.');

  lines.forEach(function (line, index) {
    var rowNr = index + 1;
    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.aantal, 0) <= 0) {
      throw new Error('Aantal moet groter zijn dan 0 op lijn ' + rowNr + '.');
    }
  });
}

function getLocationAvailableQty_(locationCode, artikelCode) {
  var code = safeText(locationCode);
  var article = safeText(artikelCode);
  if (!code || !article) return 0;

  if (safeText(code) === safeText(LOCATION.CENTRAL) && typeof buildCentralStockMap === 'function') {
    var centralMap = buildCentralStockMap();
    return centralMap[article] ? safeNumber(centralMap[article].voorraadCentraal, 0) : 0;
  }

  if (typeof isMobileWarehouseLocation === 'function' && isMobileWarehouseLocation(code) &&
      typeof buildMobileWarehouseStockMap === 'function' && typeof parseMobileWarehouseLocation === 'function') {
    var mwCode = parseMobileWarehouseLocation(code);
    var mobileMap = buildMobileWarehouseStockMap(mwCode);
    return mobileMap[article] ? safeNumber(mobileMap[article].voorraadMobiel, 0) : 0;
  }

  if (typeof isBusLocation === 'function' && isBusLocation(code) &&
      typeof buildBusStockMapForTechnician === 'function' && typeof parseBusLocation === 'function') {
    var techCode = parseBusLocation(code);
    var busMap = buildBusStockMapForTechnician(techCode);
    return busMap[article] ? safeNumber(busMap[article].voorraadBus, 0) : 0;
  }

  return 0;
}

function validateTransferSourceStock_(header, lines) {
  var fromLocation = safeText(header.vanLocatie);

  lines.forEach(function (line) {
    var code = safeText(line.artikelCode);
    var requested = safeNumber(line.aantal, 0);
    var available = getLocationAvailableQty_(fromLocation, code);

    if (requested > available) {
      throw new Error(
        'Onvoldoende voorraad op bronlocatie voor artikel ' + code +
        '. Beschikbaar: ' + available + ', gevraagd: ' + requested + '.'
      );
    }
  });
}

/* ---------------------------------------------------------
   Create / save
   --------------------------------------------------------- */

function createTransfer(payload) {
  var normalized = normalizeTransferPayload_(payload);
  var actor = requireLoggedInUser(normalized.sessionId);

  if (!roleAllowed(actor, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om transfers aan te maken.');
  }

  validateTransferHeader_(normalized);

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var transferId = makeTransferId_();

  var row = {
    TransferID: transferId,
    FlowType: normalized.flowType,
    VanLocatie: normalized.vanLocatie,
    NaarLocatie: normalized.naarLocatie,
    DoelTechniekerCode: normalized.doelTechniekerCode,
    BronReferentie: normalized.bronReferentie,
    DocumentDatum: normalized.documentDatum,
    DocumentDatumIso: normalized.documentDatum,
    Status: getDefaultTransferStatus_(),
    Reden: normalized.reden,
    Opmerking: normalized.opmerking,
    Actor: normalized.actor || safeText(actor.naam || actor.email),
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    IngediendOp: '',
    GoedgekeurdOp: '',
    GeboektOp: '',
    MobileWarehouseCode: normalized.mobileWarehouseCode,
    CreatedByRole: safeText(actor.rol),
  };

  appendObjects(getTransferHeaderTab_(), [row]);

  writeAudit({
    actie: 'CREATE_TRANSFER',
    actor: actor,
    documentType: 'Transfer',
    documentId: transferId,
    details: {
      flowType: normalized.flowType,
      vanLocatie: normalized.vanLocatie,
      naarLocatie: normalized.naarLocatie,
      documentDatum: normalized.documentDatum,
    },
  });

  return mapTransferHeader(row);
}

function saveTransferLines(payload) {
  payload = payload || {};
  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om transferlijnen te bewaren.');
  }

  var transferId = safeText(payload.transferId);
  if (!transferId) throw new Error('TransferId ontbreekt.');

  var header = getTransferById(transferId);
  if (!header) throw new Error('Transfer niet gevonden.');
  if (!isTransferEditable_(header.status)) {
    throw new Error('Transfer is niet meer bewerkbaar.');
  }

  var lines = normalizeTransferLines_(payload.lines);
  validateTransferLines_(lines);

  var existingTable = getAllValues(getTransferLineTab_());
  var headerRow = existingTable.length ? existingTable[0] : null;
  var currentRows = readObjectsSafe(getTransferLineTab_());

  var kept = currentRows.filter(function (row) {
    return safeText(row.TransferID || row.TransferId) !== transferId;
  });

  var newRows = lines.map(function (line) {
    return {
      TransferLineID: makeTransferLineId_(),
      TransferID: transferId,
      ArtikelCode: line.artikelCode,
      ArtikelOmschrijving: line.artikelOmschrijving,
      TypeMateriaal: line.typeMateriaal,
      Eenheid: line.eenheid,
      Aantal: line.aantal,
      Opmerking: line.opmerking,
    };
  });

  var finalObjects = kept.concat(newRows);

  if (!headerRow && finalObjects.length) {
    appendObjects(getTransferLineTab_(), newRows);
  } else if (headerRow) {
    writeFullTable(
      getTransferLineTab_(),
      headerRow,
      finalObjects.map(function (obj) {
        return buildRowFromHeaders(headerRow, obj);
      })
    );
  } else {
    appendObjects(getTransferLineTab_(), newRows);
  }

  writeAudit({
    actie: 'SAVE_TRANSFER_LINES',
    actor: actor,
    documentType: 'Transfer',
    documentId: transferId,
    details: {
      lineCount: lines.length,
    },
  });

  return getTransferWithLines(transferId);
}

/* ---------------------------------------------------------
   Submit / approve / book
   --------------------------------------------------------- */

function submitTransfer(payload) {
  payload = payload || {};
  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om transfer in te dienen.');
  }

  var transferId = safeText(payload.transferId);
  var header = getTransferById(transferId);
  if (!header) throw new Error('Transfer niet gevonden.');
  if (!isTransferEditable_(header.status)) {
    throw new Error('Transfer kan niet meer ingediend worden.');
  }

  var lines = getTransferLinesByTransferId(transferId);
  validateTransferLines_(lines);
  validateTransferSourceStock_(header, lines);

  var table = getAllValues(getTransferHeaderTab_());
  if (!table.length) throw new Error('Transfertab is leeg of ongeldig.');

  var headerRow = table[0];
  var rows = readObjectsSafe(getTransferHeaderTab_()).map(function (row) {
    var current = mapTransferHeader(row);
    if (safeText(current.transferId) !== transferId) {
      return row;
    }

    row.Status = getSubmittedTransferStatus_();
    row.IngediendOp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
    return row;
  });

  writeFullTable(
    getTransferHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  writeAudit({
    actie: 'SUBMIT_TRANSFER',
    actor: actor,
    documentType: 'Transfer',
    documentId: transferId,
    details: {
      flowType: header.flowType,
      vanLocatie: header.vanLocatie,
      naarLocatie: header.naarLocatie,
      lineCount: lines.length,
    },
  });

  return getTransferWithLines(transferId);
}

function approveTransfer(payload) {
  payload = payload || {};
  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.MANAGER])) {
    throw new Error('Alleen manager kan transfers goedkeuren.');
  }

  var transferId = safeText(payload.transferId);
  var header = getTransferById(transferId);
  if (!header) throw new Error('Transfer niet gevonden.');
  if (safeText(header.status) !== getSubmittedTransferStatus_()) {
    throw new Error('Transfer staat niet in ingediende status.');
  }

  var table = getAllValues(getTransferHeaderTab_());
  if (!table.length) throw new Error('Transfertab is leeg of ongeldig.');

  var headerRow = table[0];
  var rows = readObjectsSafe(getTransferHeaderTab_()).map(function (row) {
    var current = mapTransferHeader(row);
    if (safeText(current.transferId) !== transferId) {
      return row;
    }

    row.Status = getApprovedTransferStatus_();
    row.GoedgekeurdOp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
    return row;
  });

  writeFullTable(
    getTransferHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  writeAudit({
    actie: 'APPROVE_TRANSFER',
    actor: actor,
    documentType: 'Transfer',
    documentId: transferId,
    details: {
      flowType: header.flowType,
    },
  });

  return getTransferWithLines(transferId);
}

/* ---------------------------------------------------------
   Movement mapping
   --------------------------------------------------------- */

function getTransferMovementTypePair_(flowType) {
  if (typeof MOVEMENT_TYPE === 'undefined') {
    return { outType: 'TransferUit', inType: 'TransferIn' };
  }

  return {
    outType: MOVEMENT_TYPE.TRANSFER_OUT || 'TransferUit',
    inType: MOVEMENT_TYPE.TRANSFER_IN || 'TransferIn',
  };
}

function buildTransferMovements_(transferWithLines) {
  var header = transferWithLines || {};
  var lines = header.lines || [];
  var pair = getTransferMovementTypePair_(header.flowType);

  var outMovements = lines.map(function (line) {
    return buildMovementObject({
      movementType: pair.outType,
      bronType: 'Transfer',
      bronId: header.transferId,
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
      reden: header.reden,
      opmerking: line.opmerking || header.opmerking,
      actor: header.actor,
    });
  });

  var inMovements = lines.map(function (line) {
    return buildMovementObject({
      movementType: pair.inType,
      bronType: 'Transfer',
      bronId: header.transferId,
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
      reden: header.reden,
      opmerking: line.opmerking || header.opmerking,
      actor: header.actor,
    });
  });

  return outMovements.concat(inMovements);
}

function bookTransfer(payload) {
  payload = payload || {};
  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om transfer te boeken.');
  }

  if (typeof replaceSourceMovements !== 'function') {
    throw new Error('Movement service ontbreekt. Werk eerst het movementblok in.');
  }

  var transferId = safeText(payload.transferId);
  var transfer = getTransferWithLines(transferId);
  if (!transfer) throw new Error('Transfer niet gevonden.');

  var allowedStatuses = [getApprovedTransferStatus_(), getSubmittedTransferStatus_()];
  if (allowedStatuses.indexOf(safeText(transfer.status)) < 0) {
    throw new Error('Transfer kan niet geboekt worden vanuit status "' + safeText(transfer.status) + '".');
  }

  validateTransferLines_(transfer.lines || []);
  validateTransferSourceStock_(transfer, transfer.lines || []);

  var movements = buildTransferMovements_(transfer);
  replaceSourceMovements('Transfer', transfer.transferId, movements);

  var table = getAllValues(getTransferHeaderTab_());
  if (!table.length) throw new Error('Transfertab is leeg of ongeldig.');

  var headerRow = table[0];
  var rows = readObjectsSafe(getTransferHeaderTab_()).map(function (row) {
    var current = mapTransferHeader(row);
    if (safeText(current.transferId) !== transferId) {
      return row;
    }

    row.Status = getBookedTransferStatus_();
    row.GeboektOp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
    return row;
  });

  writeFullTable(
    getTransferHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  if (safeText(transfer.vanLocatie) === safeText(LOCATION.CENTRAL) &&
      typeof rebuildCentralWarehouseOverview === 'function') {
    rebuildCentralWarehouseOverview();
  }
  if (safeText(transfer.naarLocatie) === safeText(LOCATION.CENTRAL) &&
      typeof rebuildCentralWarehouseOverview === 'function') {
    rebuildCentralWarehouseOverview();
  }

  writeAudit({
    actie: 'BOOK_TRANSFER',
    actor: actor,
    documentType: 'Transfer',
    documentId: transferId,
    details: {
      flowType: transfer.flowType,
      movementCount: movements.length,
    },
  });

  return getTransferWithLines(transferId);
}

/* ---------------------------------------------------------
   Queries voor schermen
   --------------------------------------------------------- */

function filterTransfersForUser_(rows, user) {
  if (roleAllowed(user, [ROLE.MANAGER])) {
    return rows;
  }

  if (roleAllowed(user, [ROLE.MOBILE_WAREHOUSE])) {
    var ownCode = safeText(user.mobileWarehouseCode || '');
    if (!ownCode) return [];
    var ownLocation = typeof getMobileWarehouseLocationCode === 'function'
      ? getMobileWarehouseLocationCode(ownCode)
      : ('Mobiel:' + ownCode);

    return rows.filter(function (row) {
      return safeText(row.vanLocatie) === ownLocation || safeText(row.naarLocatie) === ownLocation;
    });
  }

  if (roleAllowed(user, [ROLE.WAREHOUSE])) {
    return rows.filter(function (row) {
      return safeText(row.vanLocatie) === safeText(LOCATION.CENTRAL) ||
             safeText(row.naarLocatie) === safeText(LOCATION.CENTRAL);
    });
  }

  return [];
}

function getTransfersData(payload) {
  payload = payload || {};
  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om transfers te bekijken.');
  }

  var rows = filterTransfersForUser_(getTransfersWithLines(), actor);

  return {
    items: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getDefaultTransferStatus_(); }).length,
      ingediend: rows.filter(function (x) { return safeText(x.status) === getSubmittedTransferStatus_(); }).length,
      goedgekeurd: rows.filter(function (x) { return safeText(x.status) === getApprovedTransferStatus_(); }).length,
      geboekt: rows.filter(function (x) { return safeText(x.status) === getBookedTransferStatus_(); }).length,
    }
  };
}
