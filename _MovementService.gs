/* =========================================================
   23_MovementService.gs
   Refactor: movement core service
   Doel:
   - voorraadmutaties als centrale bron van waarheid
   - generieke movement builder
   - source-based replace/query helpers
   - duidelijke location delta helpers
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getMovementTab_() {
  return TABS.WAREHOUSE_MOVEMENTS || TABS.MOVEMENTS || 'MagazijnMutaties';
}

function getMovementTypeFallback_(value, fallbackValue) {
  var txt = safeText(value);
  return txt || fallbackValue;
}

function getMovementTypeTransferOut_() {
  return getMovementTypeFallback_(
    typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE ? MOVEMENT_TYPE.TRANSFER_OUT : '',
    'TransferUit'
  );
}

function getMovementTypeTransferIn_() {
  return getMovementTypeFallback_(
    typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE ? MOVEMENT_TYPE.TRANSFER_IN : '',
    'TransferIn'
  );
}

function getMovementTypeReceiptIn_() {
  return getMovementTypeFallback_(
    typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE ? MOVEMENT_TYPE.RECEIPT_IN : '',
    'ReceiptIn'
  );
}

function getMovementTypeReturnOut_() {
  return getMovementTypeFallback_(
    typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE ? MOVEMENT_TYPE.RETURN_OUT : '',
    'ReturnOut'
  );
}

function getMovementTypeReturnIn_() {
  return getMovementTypeFallback_(
    typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE ? MOVEMENT_TYPE.RETURN_IN : '',
    'ReturnIn'
  );
}

function getMovementTypeNeedIssueOut_() {
  return getMovementTypeFallback_(
    typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE ? MOVEMENT_TYPE.NEED_ISSUE_OUT : '',
    'NeedIssueOut'
  );
}

function getMovementTypeNeedIssueIn_() {
  return getMovementTypeFallback_(
    typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE ? MOVEMENT_TYPE.NEED_ISSUE_IN : '',
    'NeedIssueIn'
  );
}

/* ---------------------------------------------------------
   IDs / mapping
   --------------------------------------------------------- */

function makeMovementId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'MOV-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function mapWarehouseMovement(row) {
  var qtyIn = safeNumber(row.AantalIn || row.InQty, 0);
  var qtyOut = safeNumber(row.AantalUit || row.OutQty, 0);

  return {
    movementId: safeText(row.MovementID || row.MovementId || row.ID),
    movementType: safeText(row.MovementType || row.Type),
    bronType: safeText(row.BronType),
    bronId: safeText(row.BronID || row.BronId),
    datumBoeking: safeText(row.DatumBoeking || row.DocumentDatum || row.Datum),
    datumBoekingRaw: safeText(row.DatumBoekingRaw || row.DatumBoeking || row.DocumentDatum || row.Datum),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode || row.ArtikelNr))),
    eenheid: safeText(row.Eenheid || row.Unit),
    aantalIn: qtyIn,
    aantalUit: qtyOut,
    nettoAantal: safeNumber(row.NettoAantal, qtyIn - qtyOut),
    locatieVan: safeText(row.LocatieVan),
    locatieNaar: safeText(row.LocatieNaar),
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),
    actor: safeText(row.Actor),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
  };
}

/* ---------------------------------------------------------
   Builder
   --------------------------------------------------------- */

function normalizeMovementPayload_(payload) {
  payload = payload || {};

  var aantalIn = safeNumber(payload.aantalIn, 0);
  var aantalUit = safeNumber(payload.aantalUit, 0);
  var nettoAantal =
    payload.nettoAantal === undefined || payload.nettoAantal === null || payload.nettoAantal === ''
      ? (aantalIn - aantalUit)
      : safeNumber(payload.nettoAantal, 0);

  return {
    movementType: safeText(payload.movementType || payload.type),
    bronType: safeText(payload.bronType),
    bronId: safeText(payload.bronId),
    datumBoeking: safeText(payload.datumBoeking || payload.documentDatum || payload.datum),
    artikelCode: safeText(payload.artikelCode || payload.artikelNr),
    artikelOmschrijving: safeText(payload.artikelOmschrijving || payload.artikel),
    typeMateriaal: safeText(payload.typeMateriaal || determineMaterialTypeFromArticle(safeText(payload.artikelCode || payload.artikelNr))),
    eenheid: safeText(payload.eenheid || payload.unit || 'Stuk'),
    aantalIn: aantalIn,
    aantalUit: aantalUit,
    nettoAantal: nettoAantal,
    locatieVan: safeText(payload.locatieVan || payload.fromLocation),
    locatieNaar: safeText(payload.locatieNaar || payload.toLocation),
    reden: safeText(payload.reden || payload.reason),
    opmerking: safeText(payload.opmerking || payload.remark),
    actor: safeText(payload.actor),
  };
}

function validateMovementPayload_(payload) {
  if (!safeText(payload.movementType)) throw new Error('MovementType is verplicht.');
  if (!safeText(payload.bronType)) throw new Error('BronType is verplicht.');
  if (!safeText(payload.bronId)) throw new Error('BronId is verplicht.');
  if (!safeText(payload.datumBoeking)) throw new Error('DatumBoeking is verplicht.');
  if (!safeText(payload.artikelCode)) throw new Error('ArtikelCode is verplicht.');

  if (safeNumber(payload.aantalIn, 0) < 0) {
    throw new Error('AantalIn mag niet negatief zijn.');
  }
  if (safeNumber(payload.aantalUit, 0) < 0) {
    throw new Error('AantalUit mag niet negatief zijn.');
  }

  if (
    safeNumber(payload.aantalIn, 0) === 0 &&
    safeNumber(payload.aantalUit, 0) === 0 &&
    safeNumber(payload.nettoAantal, 0) === 0
  ) {
    throw new Error('Een mutatie moet een niet-nul hoeveelheid hebben.');
  }

  return true;
}

function buildMovementObject(payload) {
  var normalized = normalizeMovementPayload_(payload);
  validateMovementPayload_(normalized);

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  return {
    MovementID: makeMovementId_(),
    MovementType: normalized.movementType,
    BronType: normalized.bronType,
    BronID: normalized.bronId,
    DatumBoeking: normalized.datumBoeking,
    DatumBoekingRaw: normalized.datumBoeking,
    ArtikelCode: normalized.artikelCode,
    ArtikelOmschrijving: normalized.artikelOmschrijving,
    TypeMateriaal: normalized.typeMateriaal,
    Eenheid: normalized.eenheid,
    AantalIn: normalized.aantalIn,
    AantalUit: normalized.aantalUit,
    NettoAantal: normalized.nettoAantal,
    LocatieVan: normalized.locatieVan,
    LocatieNaar: normalized.locatieNaar,
    Reden: normalized.reden,
    Opmerking: normalized.opmerking,
    Actor: normalized.actor,
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllMovements() {
  return readObjectsSafe(getMovementTab_())
    .map(mapWarehouseMovement)
    .sort(function (a, b) {
      return (
        safeText(a.datumBoekingRaw).localeCompare(safeText(b.datumBoekingRaw)) ||
        safeText(a.movementId).localeCompare(safeText(b.movementId))
      );
    });
}

function getMovementsBySource(bronType, bronId) {
  var type = safeText(bronType);
  var id = safeText(bronId);

  return getAllMovements().filter(function (item) {
    return safeText(item.bronType) === type &&
      safeText(item.bronId) === id;
  });
}

function getMovementsByArticleCode(artikelCode) {
  var code = safeText(artikelCode);
  return getAllMovements().filter(function (item) {
    return safeText(item.artikelCode) === code;
  });
}

function getMovementsByLocation(locationCode) {
  var code = safeText(locationCode);
  return getAllMovements().filter(function (item) {
    return safeText(item.locatieVan) === code ||
      safeText(item.locatieNaar) === code;
  });
}

/* ---------------------------------------------------------
   Write layer
   --------------------------------------------------------- */

function appendMovementObjects(movements) {
  var rows = Array.isArray(movements) ? movements : [];
  if (!rows.length) {
    return {
      insertedCount: 0,
    };
  }

  appendObjects(getMovementTab_(), rows);

  return {
    insertedCount: rows.length,
  };
}

function appendMovementBatch(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.MOBILE_WAREHOUSE])) {
    throw new Error('Geen rechten om mutaties te schrijven.');
  }

  var inputRows = Array.isArray(payload.movements) ? payload.movements : [];
  var rows = inputRows.map(buildMovementObject);

  var result = appendMovementObjects(rows);

  writeAudit({
    actie: 'APPEND_MOVEMENT_BATCH',
    actor: actor,
    documentType: 'Mutaties',
    documentId: safeText(payload.batchId || 'BATCH'),
    details: {
      insertedCount: result.insertedCount,
    },
  });

  return result;
}

function replaceSourceMovements(bronType, bronId, movementObjects) {
  var type = safeText(bronType);
  var id = safeText(bronId);
  var newRows = Array.isArray(movementObjects) ? movementObjects : [];

  if (!type) throw new Error('BronType ontbreekt.');
  if (!id) throw new Error('BronId ontbreekt.');

  var table = getAllValues(getMovementTab_());
  var headerRow = table.length ? table[0] : null;
  var currentRows = readObjectsSafe(getMovementTab_());

  var kept = currentRows.filter(function (row) {
    var mapped = mapWarehouseMovement(row);
    return !(
      safeText(mapped.bronType) === type &&
      safeText(mapped.bronId) === id
    );
  });

  var finalObjects = kept.concat(newRows);

  if (!headerRow && finalObjects.length) {
    appendObjects(getMovementTab_(), newRows);
  } else if (headerRow) {
    writeFullTable(
      getMovementTab_(),
      headerRow,
      finalObjects.map(function (obj) {
        return buildRowFromHeaders(headerRow, obj);
      })
    );
  } else if (newRows.length) {
    appendObjects(getMovementTab_(), newRows);
  }

  return {
    bronType: type,
    bronId: id,
    insertedCount: newRows.length,
    remainingCount: finalObjects.length,
  };
}

/* ---------------------------------------------------------
   Quantitative helpers
   --------------------------------------------------------- */

function getInboundQtyFromMovement(move, locationCode) {
  var movement = move || {};
  var code = safeText(locationCode);

  if (code && safeText(movement.locatieNaar) !== code) {
    return 0;
  }

  if (safeNumber(movement.aantalIn, 0) > 0) {
    return safeNumber(movement.aantalIn, 0);
  }

  if (!code || safeText(movement.locatieNaar) === code) {
    return Math.max(safeNumber(movement.nettoAantal, 0), 0);
  }

  return 0;
}

function getOutboundQtyFromMovement(move, locationCode) {
  var movement = move || {};
  var code = safeText(locationCode);

  if (code && safeText(movement.locatieVan) !== code) {
    return 0;
  }

  if (safeNumber(movement.aantalUit, 0) > 0) {
    return safeNumber(movement.aantalUit, 0);
  }

  if (!code || safeText(movement.locatieVan) === code) {
    return Math.max(-safeNumber(movement.nettoAantal, 0), 0);
  }

  return 0;
}

function getNetQtyFromMovement(move, locationCode) {
  var movement = move || {};
  var code = safeText(locationCode);

  if (!code) {
    return safeNumber(movement.nettoAantal, 0);
  }

  var net = 0;
  if (safeText(movement.locatieNaar) === code) {
    net += getInboundQtyFromMovement(movement, code);
  }
  if (safeText(movement.locatieVan) === code) {
    net -= getOutboundQtyFromMovement(movement, code);
  }
  return net;
}

/* ---------------------------------------------------------
   Location classification helpers
   --------------------------------------------------------- */

function isCentralMovementLocation(location) {
  return safeText(location) === safeText(LOCATION.CENTRAL);
}

function isBusMovementLocation(location) {
  if (typeof isBusLocation === 'function') {
    return isBusLocation(location);
  }
  return /^Bus:/i.test(safeText(location));
}

function isMobileMovementLocation(location) {
  if (typeof isMobileWarehouseLocation === 'function') {
    return isMobileWarehouseLocation(location);
  }
  return /^Mobiel:/i.test(safeText(location));
}

function getLocationDeltaFromMovement(move) {
  var movement = move || {};

  return {
    artikelCode: safeText(movement.artikelCode),
    vanLocatie: safeText(movement.locatieVan),
    naarLocatie: safeText(movement.locatieNaar),
    outboundQty: safeNumber(movement.aantalUit, 0),
    inboundQty: safeNumber(movement.aantalIn, 0),
    nettoAantal: safeNumber(movement.nettoAantal, 0),
    affectsCentral:
      isCentralMovementLocation(movement.locatieVan) ||
      isCentralMovementLocation(movement.locatieNaar),
    affectsBus:
      isBusMovementLocation(movement.locatieVan) ||
      isBusMovementLocation(movement.locatieNaar),
    affectsMobileWarehouse:
      isMobileMovementLocation(movement.locatieVan) ||
      isMobileMovementLocation(movement.locatieNaar),
  };
}

/* ---------------------------------------------------------
   Query models for screens
   --------------------------------------------------------- */

function summarizeMovementSet_(rows) {
  var items = Array.isArray(rows) ? rows : [];

  return {
    totaal: items.length,
    totaalIn: items.reduce(function (sum, row) {
      return sum + safeNumber(row.aantalIn, 0);
    }, 0),
    totaalUit: items.reduce(function (sum, row) {
      return sum + safeNumber(row.aantalUit, 0);
    }, 0),
    netto: items.reduce(function (sum, row) {
      return sum + safeNumber(row.nettoAantal, 0);
    }, 0),
  };
}

function getMovementData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.MOBILE_WAREHOUSE])) {
    throw new Error('Geen rechten om mutaties te bekijken.');
  }

  var bronType = safeText(payload.bronType);
  var bronId = safeText(payload.bronId);
  var artikelCode = safeText(payload.artikelCode);
  var locatie = safeText(payload.locatie);

  var rows = getAllMovements();

  if (bronType) {
    rows = rows.filter(function (item) {
      return safeText(item.bronType) === bronType;
    });
  }
  if (bronId) {
    rows = rows.filter(function (item) {
      return safeText(item.bronId) === bronId;
    });
  }
  if (artikelCode) {
    rows = rows.filter(function (item) {
      return safeText(item.artikelCode) === artikelCode;
    });
  }
  if (locatie) {
    rows = rows.filter(function (item) {
      return safeText(item.locatieVan) === locatie ||
        safeText(item.locatieNaar) === locatie;
    });
  }

  return {
    items: rows,
    summary: summarizeMovementSet_(rows),
  };
}
