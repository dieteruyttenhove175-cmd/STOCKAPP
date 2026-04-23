/* =========================================================
   40_ReceiptService.gs
   Refactor: receipt core service
   Doel:
   - centrale documentlaag voor ontvangsten
   - manuele ontvangst aanmaken
   - lijnen bewaren
   - indienen
   - goedkeuren en stockmutaties boeken
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getReceiptHeaderTab_() {
  return TABS.RECEIPTS || 'Ontvangsten';
}

function getReceiptLineTab_() {
  return TABS.RECEIPT_LINES || 'OntvangstLijnen';
}

function getReceiptDefaultSupplier_() {
  return 'Fluvius';
}

function getReceiptStatusOpen_() {
  if (typeof RECEIPT_STATUS !== 'undefined' && RECEIPT_STATUS && RECEIPT_STATUS.OPEN) {
    return RECEIPT_STATUS.OPEN;
  }
  return 'Open';
}

function getReceiptStatusSubmitted_() {
  if (typeof RECEIPT_STATUS !== 'undefined' && RECEIPT_STATUS && RECEIPT_STATUS.SUBMITTED) {
    return RECEIPT_STATUS.SUBMITTED;
  }
  return 'Ingediend';
}

function getReceiptStatusApproved_() {
  if (typeof RECEIPT_STATUS !== 'undefined' && RECEIPT_STATUS && RECEIPT_STATUS.APPROVED) {
    return RECEIPT_STATUS.APPROVED;
  }
  return 'Goedgekeurd';
}

function getReceiptStatusClosed_() {
  if (typeof RECEIPT_STATUS !== 'undefined' && RECEIPT_STATUS && RECEIPT_STATUS.CLOSED) {
    return RECEIPT_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function isReceiptEditable_(status) {
  var value = safeText(status);
  return !value || value === getReceiptStatusOpen_();
}

function makeReceiptId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'RCT-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function makeReceiptLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'RCL-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapReceiptHeader(row) {
  return {
    receiptId: safeText(row.ReceiptID || row.ReceiptId || row.OntvangstID || row.ID),
    leverancier: safeText(row.Leverancier || row.Supplier || getReceiptDefaultSupplier_()),
    typeMateriaal: safeText(row.TypeMateriaal || row.MaterialType),
    bestelbonNr: safeText(row.BestelbonNr || row.Bestelbon || row.PONumber),
    externeReferentie: safeText(row.ExterneReferentie || row.Reference || row.FluviusReferentie),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    ontvangstDatum: safeText(row.OntvangstDatum || row.ReceiveDate),
    ontvangstDatumIso: safeText(row.OntvangstDatumIso || row.ReceiveDateIso || row.OntvangstDatum || row.ReceiveDate),
    bronType: safeText(row.BronType || 'Manueel'),
    bronReferentie: safeText(row.BronReferentie || row.SourceReference),
    status: safeText(row.Status),
    actor: safeText(row.Actor),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    ingediendOp: safeText(row.IngediendOp),
    goedgekeurdOp: safeText(row.GoedgekeurdOp),
    opmerking: safeText(row.Opmerking),
    deltaLijnen: safeNumber(row.DeltaLijnen, 0),
  };
}

function mapReceiptLine(row) {
  var besteld = safeNumber(row.BesteldAantal || row.Besteld || row.ExpectedQty, 0);
  var ontvangen = safeNumber(row.OntvangenAantal || row.Ontvangen || row.ReceivedQty, 0);

  return {
    receiptLineId: safeText(row.ReceiptLineID || row.ReceiptLineId || row.OntvangstLijnID || row.ID),
    receiptId: safeText(row.ReceiptID || row.ReceiptId || row.OntvangstID),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode || row.ArtikelNr))),
    eenheid: safeText(row.Eenheid || row.Unit),
    besteldAantal: besteld,
    ontvangenAantal: ontvangen,
    deltaAantal: safeNumber(row.DeltaAantal, ontvangen - besteld),
    redenDelta: safeText(row.RedenDelta || row.DeltaReden),
    opmerking: safeText(row.Opmerking),
    palletNr: safeText(row.PalletNr || row.Pallet || ''),
    label: safeText(row.Label || ''),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllReceiptHeaders() {
  return readObjectsSafe(getReceiptHeaderTab_())
    .map(mapReceiptHeader)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.receiptId).localeCompare(safeText(a.receiptId))
      );
    });
}

function getAllReceiptLines() {
  return readObjectsSafe(getReceiptLineTab_())
    .map(mapReceiptLine)
    .sort(function (a, b) {
      return (
        safeText(a.receiptId).localeCompare(safeText(b.receiptId)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
      );
    });
}

function getReceiptHeaderById(receiptId) {
  var id = safeText(receiptId);
  if (!id) return null;

  return getAllReceiptHeaders().find(function (item) {
    return safeText(item.receiptId) === id;
  }) || null;
}

function getReceiptLinesByReceiptId(receiptId) {
  var id = safeText(receiptId);
  return getAllReceiptLines().filter(function (item) {
    return safeText(item.receiptId) === id;
  });
}

function buildReceiptsWithLines(headers, lines) {
  var lineMap = {};

  (lines || []).forEach(function (line) {
    var id = safeText(line.receiptId);
    if (!lineMap[id]) lineMap[id] = [];
    lineMap[id].push(line);
  });

  return (headers || []).map(function (header) {
    var receiptLines = lineMap[safeText(header.receiptId)] || [];
    var deltaLijnen = receiptLines.filter(function (line) {
      return safeNumber(line.deltaAantal, 0) !== 0;
    }).length;

    return Object.assign({}, header, {
      lines: receiptLines,
      lineCount: receiptLines.length,
      totaalOntvangen: receiptLines.reduce(function (sum, line) {
        return sum + safeNumber(line.ontvangenAantal, 0);
      }, 0),
      deltaLijnen: deltaLijnen,
    });
  });
}

function getReceiptsWithLines() {
  return buildReceiptsWithLines(getAllReceiptHeaders(), getAllReceiptLines());
}

function getReceiptWithLines(receiptId) {
  var header = getReceiptHeaderById(receiptId);
  if (!header) return null;
  return buildReceiptsWithLines([header], getReceiptLinesByReceiptId(receiptId))[0] || null;
}

/* ---------------------------------------------------------
   Payload normalization
   --------------------------------------------------------- */

function normalizeReceiptHeaderPayload_(payload) {
  payload = payload || {};

  return {
    sessionId: getPayloadSessionId(payload),
    leverancier: safeText(payload.leverancier || payload.supplier || getReceiptDefaultSupplier_()),
    typeMateriaal: safeText(payload.typeMateriaal || payload.materialType),
    bestelbonNr: safeText(payload.bestelbonNr || payload.purchaseOrderNumber || payload.poNumber),
    externeReferentie: safeText(payload.externeReferentie || payload.reference || payload.fluviusReferentie),
    documentDatum: safeText(payload.documentDatum || payload.documentDate),
    ontvangstDatum: safeText(payload.ontvangstDatum || payload.receiveDate || payload.documentDatum || payload.documentDate),
    bronType: safeText(payload.bronType || payload.sourceType || 'Manueel'),
    bronReferentie: safeText(payload.bronReferentie || payload.sourceReference),
    opmerking: safeText(payload.opmerking || payload.remark),
    actor: safeText(payload.actor),
  };
}

function normalizeReceiptLines_(lines) {
  return (Array.isArray(lines) ? lines : []).map(function (line) {
    var artikelCode = safeText(line.artikelCode || line.artikelNr);
    var artikelOmschrijving = safeText(line.artikelOmschrijving || line.artikel);
    var eenheid = safeText(line.eenheid || line.unit || '');
    var besteldAantal = safeNumber(line.besteldAantal || line.besteld || line.expectedQty, 0);
    var ontvangenAantal = safeNumber(line.ontvangenAantal || line.ontvangen || line.receivedQty, 0);
    var typeMateriaal = safeText(line.typeMateriaal || determineMaterialTypeFromArticle(artikelCode));
    var deltaAantal = ontvangenAantal - besteldAantal;

    return {
      artikelCode: artikelCode,
      artikelOmschrijving: artikelOmschrijving,
      typeMateriaal: typeMateriaal,
      eenheid: eenheid,
      besteldAantal: besteldAantal,
      ontvangenAantal: ontvangenAantal,
      deltaAantal: safeNumber(line.deltaAantal, deltaAantal),
      redenDelta: safeText(line.redenDelta || line.deltaReason),
      opmerking: safeText(line.opmerking),
      palletNr: safeText(line.palletNr || line.pallet),
      label: safeText(line.label),
    };
  });
}

/* ---------------------------------------------------------
   Validation
   --------------------------------------------------------- */

function validateReceiptHeader_(payload) {
  if (!payload.sessionId) throw new Error('Sessie ontbreekt.');
  if (!payload.documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!payload.ontvangstDatum) throw new Error('Ontvangstdatum is verplicht.');
  return true;
}

function validateReceiptLines_(lines) {
  if (!lines.length) {
    throw new Error('Geen ontvangstlijnen ontvangen.');
  }

  lines.forEach(function (line, index) {
    var rowNr = index + 1;

    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.ontvangenAantal, 0) < 0) {
      throw new Error('Ontvangen aantal mag niet negatief zijn op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.deltaAantal, 0) !== 0 && !safeText(line.redenDelta)) {
      throw new Error('Reden delta is verplicht op lijn ' + rowNr + '.');
    }
  });

  return true;
}

function deriveReceiptMaterialType_(lines) {
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
   Header create / lines save
   --------------------------------------------------------- */

function createManualReceipt(payload) {
  var normalized = normalizeReceiptHeaderPayload_(payload);
  var actor = assertWarehouseAccess(normalized.sessionId);

  validateReceiptHeader_(normalized);

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var receiptId = makeReceiptId_();

  var obj = {
    ReceiptID: receiptId,
    Leverancier: normalized.leverancier || getReceiptDefaultSupplier_(),
    TypeMateriaal: normalized.typeMateriaal,
    BestelbonNr: normalized.bestelbonNr,
    ExterneReferentie: normalized.externeReferentie,
    DocumentDatum: normalized.documentDatum,
    DocumentDatumIso: normalized.documentDatum,
    OntvangstDatum: normalized.ontvangstDatum,
    OntvangstDatumIso: normalized.ontvangstDatum,
    BronType: normalized.bronType || 'Manueel',
    BronReferentie: normalized.bronReferentie,
    Status: getReceiptStatusOpen_(),
    Actor: normalized.actor || safeText(actor.naam || actor.email),
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    IngediendOp: '',
    GoedgekeurdOp: '',
    Opmerking: normalized.opmerking,
    DeltaLijnen: 0,
  };

  appendObjects(getReceiptHeaderTab_(), [obj]);

  writeAudit({
    actie: 'CREATE_RECEIPT',
    actor: actor,
    documentType: 'Ontvangst',
    documentId: receiptId,
    details: {
      leverancier: obj.Leverancier,
      documentDatum: obj.DocumentDatum,
      bronType: obj.BronType,
    },
  });

  return mapReceiptHeader(obj);
}

function saveReceiptLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertWarehouseAccess(sessionId);
  var receiptId = safeText(payload.receiptId);
  if (!receiptId) throw new Error('ReceiptId ontbreekt.');

  var header = getReceiptHeaderById(receiptId);
  if (!header) throw new Error('Ontvangst niet gevonden.');
  if (!isReceiptEditable_(header.status)) {
    throw new Error('Ontvangst is niet meer bewerkbaar.');
  }

  var lines = normalizeReceiptLines_(payload.lines);
  validateReceiptLines_(lines);

  var table = getAllValues(getReceiptLineTab_());
  var headerRow = table.length ? table[0] : null;
  var currentRows = readObjectsSafe(getReceiptLineTab_());

  var kept = currentRows.filter(function (row) {
    return safeText(row.ReceiptID || row.ReceiptId || row.OntvangstID) !== receiptId;
  });

  var newRows = lines.map(function (line) {
    return {
      ReceiptLineID: makeReceiptLineId_(),
      ReceiptID: receiptId,
      ArtikelCode: line.artikelCode,
      ArtikelOmschrijving: line.artikelOmschrijving,
      TypeMateriaal: line.typeMateriaal,
      Eenheid: line.eenheid,
      BesteldAantal: line.besteldAantal,
      OntvangenAantal: line.ontvangenAantal,
      DeltaAantal: line.deltaAantal,
      RedenDelta: line.redenDelta,
      Opmerking: line.opmerking,
      PalletNr: line.palletNr,
      Label: line.label,
    };
  });

  var finalObjects = kept.concat(newRows);

  if (!headerRow && finalObjects.length) {
    appendObjects(getReceiptLineTab_(), newRows);
  } else if (headerRow) {
    writeFullTable(
      getReceiptLineTab_(),
      headerRow,
      finalObjects.map(function (obj) {
        return buildRowFromHeaders(headerRow, obj);
      })
    );
  } else {
    appendObjects(getReceiptLineTab_(), newRows);
  }

  updateReceiptDerivedFields_(receiptId, {
    typeMateriaal: deriveReceiptMaterialType_(lines),
    deltaLijnen: lines.filter(function (line) {
      return safeNumber(line.deltaAantal, 0) !== 0;
    }).length,
  });

  writeAudit({
    actie: 'SAVE_RECEIPT_LINES',
    actor: actor,
    documentType: 'Ontvangst',
    documentId: receiptId,
    details: {
      lineCount: lines.length,
    },
  });

  return getReceiptWithLines(receiptId);
}

function updateReceiptDerivedFields_(receiptId, values) {
  var table = getAllValues(getReceiptHeaderTab_());
  if (!table.length) throw new Error('Ontvangsttab is leeg of ongeldig.');

  var headerRow = table[0];
  var rows = readObjectsSafe(getReceiptHeaderTab_()).map(function (row) {
    var current = mapReceiptHeader(row);
    if (safeText(current.receiptId) !== safeText(receiptId)) {
      return row;
    }

    row.TypeMateriaal = safeText(values.typeMateriaal || row.TypeMateriaal);
    row.DeltaLijnen = safeNumber(values.deltaLijnen, 0);
    return row;
  });

  writeFullTable(
    getReceiptHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );
}

/* ---------------------------------------------------------
   Submit / approve
   --------------------------------------------------------- */

function submitReceipt(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertWarehouseAccess(sessionId);
  var receiptId = safeText(payload.receiptId);
  if (!receiptId) throw new Error('ReceiptId ontbreekt.');

  var header = getReceiptHeaderById(receiptId);
  if (!header) throw new Error('Ontvangst niet gevonden.');
  if (!isReceiptEditable_(header.status)) {
    throw new Error('Ontvangst kan niet meer ingediend worden.');
  }

  var lines = getReceiptLinesByReceiptId(receiptId);
  validateReceiptLines_(lines);

  var table = getAllValues(getReceiptHeaderTab_());
  if (!table.length) throw new Error('Ontvangsttab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getReceiptHeaderTab_()).map(function (row) {
    var current = mapReceiptHeader(row);
    if (safeText(current.receiptId) !== receiptId) {
      return row;
    }

    row.Status = getReceiptStatusSubmitted_();
    row.IngediendOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getReceiptHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  if (typeof pushManagerNotification === 'function') {
    pushManagerNotification(
      'Ontvangst',
      'Ontvangst wacht op goedkeuring',
      'Er werd een ontvangst ingediend die goedkeuring vraagt.',
      'Ontvangst',
      receiptId,
      'MANAGER'
    );
  }

  writeAudit({
    actie: 'SUBMIT_RECEIPT',
    actor: actor,
    documentType: 'Ontvangst',
    documentId: receiptId,
    details: {
      lineCount: lines.length,
    },
  });

  return getReceiptWithLines(receiptId);
}

function buildReceiptMovements_(receipt) {
  var header = receipt || {};
  var lines = header.lines || [];
  var movementType =
    typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE && MOVEMENT_TYPE.RECEIPT_IN
      ? MOVEMENT_TYPE.RECEIPT_IN
      : 'ReceiptIn';

  return lines
    .filter(function (line) {
      return safeNumber(line.ontvangenAantal, 0) > 0;
    })
    .map(function (line) {
      return buildMovementObject({
        movementType: movementType,
        bronType: 'Ontvangst',
        bronId: header.receiptId,
        datumBoeking: header.ontvangstDatum || header.documentDatum,
        artikelCode: line.artikelCode,
        artikelOmschrijving: line.artikelOmschrijving,
        typeMateriaal: line.typeMateriaal,
        eenheid: line.eenheid,
        aantalIn: line.ontvangenAantal,
        aantalUit: 0,
        nettoAantal: safeNumber(line.ontvangenAantal, 0),
        locatieVan: '',
        locatieNaar: LOCATION.CENTRAL,
        reden: 'Ontvangst',
        opmerking: line.opmerking || header.opmerking,
        actor: header.actor,
      });
    });
}

function approveReceipt(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertManagerAccess(sessionId);
  var receiptId = safeText(payload.receiptId);
  if (!receiptId) throw new Error('ReceiptId ontbreekt.');

  if (typeof replaceSourceMovements !== 'function') {
    throw new Error('Movement service ontbreekt. Werk eerst het movementblok in.');
  }

  var receipt = getReceiptWithLines(receiptId);
  if (!receipt) throw new Error('Ontvangst niet gevonden.');
  if (safeText(receipt.status) !== getReceiptStatusSubmitted_()) {
    throw new Error('Ontvangst staat niet in ingediende status.');
  }

  validateReceiptLines_(receipt.lines || []);

  var movements = buildReceiptMovements_(receipt);
  replaceSourceMovements('Ontvangst', receipt.receiptId, movements);

  var table = getAllValues(getReceiptHeaderTab_());
  if (!table.length) throw new Error('Ontvangsttab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getReceiptHeaderTab_()).map(function (row) {
    var current = mapReceiptHeader(row);
    if (safeText(current.receiptId) !== receiptId) {
      return row;
    }

    row.Status = getReceiptStatusApproved_();
    row.GoedgekeurdOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getReceiptHeaderTab_(),
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
      bronType: 'Ontvangst',
      bronId: receiptId,
    });
  }

  writeAudit({
    actie: 'APPROVE_RECEIPT',
    actor: actor,
    documentType: 'Ontvangst',
    documentId: receiptId,
    details: {
      movementCount: movements.length,
      totalReceived: (receipt.lines || []).reduce(function (sum, line) {
        return sum + safeNumber(line.ontvangenAantal, 0);
      }, 0),
    },
  });

  return getReceiptWithLines(receiptId);
}

/* ---------------------------------------------------------
   Queries for screens
   --------------------------------------------------------- */

function filterReceiptsForUser_(rows, user) {
  if (roleAllowed(user, [ROLE.MANAGER])) {
    return rows;
  }

  if (roleAllowed(user, [ROLE.WAREHOUSE])) {
    return rows;
  }

  return [];
}

function getReceiptsData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertWarehouseAccess(sessionId);
  var rows = filterReceiptsForUser_(getReceiptsWithLines(), actor);

  return {
    items: rows,
    receipts: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getReceiptStatusOpen_(); }).length,
      ingediend: rows.filter(function (x) { return safeText(x.status) === getReceiptStatusSubmitted_(); }).length,
      goedgekeurd: rows.filter(function (x) { return safeText(x.status) === getReceiptStatusApproved_(); }).length,
      gesloten: rows.filter(function (x) { return safeText(x.status) === getReceiptStatusClosed_(); }).length,
      deltaLijnen: rows.reduce(function (sum, item) {
        return sum + safeNumber(item.deltaLijnen, 0);
      }, 0),
    }
  };
}
