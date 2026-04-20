/* =========================================================
   40_ReceiptService.gs — ontvangsten
   ========================================================= */

function makeReceiptId() {
  return makeStampedId('R');
}

function getAllReceipts() {
  return readObjectsSafe(TABS.RECEIPTS)
    .map(mapReceipt)
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.ontvangstId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.ontvangstId)}`
      )
    );
}

function getAllReceiptLines() {
  return readObjectsSafe(TABS.RECEIPT_LINES)
    .map(mapReceiptLine)
    .filter(line => line.actief);
}

function getReceiptById(ontvangstId) {
  const id = safeText(ontvangstId);
  if (!id) return null;

  return getAllReceipts().find(item => item.ontvangstId === id) || null;
}

function getReceiptLinesByReceiptId(ontvangstId) {
  const id = safeText(ontvangstId);
  if (!id) return [];

  return getAllReceiptLines().filter(line => line.ontvangstId === id);
}

function buildReceiptLineObject(ontvangstId, line) {
  const artikelCode = safeText(line.artikelCode);
  const article = getArticleMaster(artikelCode);

  const besteldAantal = safeNumber(line.besteldAantal, 0);
  const ontvangenAantal = safeNumber(line.ontvangenAantal, 0);
  const deltaAantal = besteldAantal - ontvangenAantal;

  return {
    OntvangstID: safeText(ontvangstId),
    ArtikelCode: artikelCode,
    ArtikelOmschrijving: safeText(line.artikelOmschrijving) || safeText(article && article.artikelOmschrijving),
    Eenheid: safeText(line.eenheid) || safeText(article && article.eenheid),
    BesteldAantal: besteldAantal,
    OntvangenAantal: ontvangenAantal,
    DeltaAantal: deltaAantal,
    RedenDelta: safeText(line.redenDelta),
    Opmerking: safeText(line.opmerking),
    Actief: 'Ja'
  };
}

function replaceReceiptLinesForReceipt(ontvangstId, lineObjects) {
  const receiptId = safeText(ontvangstId);
  const sheet = getSheetOrThrow(TABS.RECEIPT_LINES);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.RECEIPT_LINES);
  const col = getColMap(headers);
  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => safeText(row[col['OntvangstID']]) !== receiptId);
  const newRows = (lineObjects || []).map(obj => buildRowFromHeaders(headers, obj));

  writeFullTable(TABS.RECEIPT_LINES, headers, keptRows.concat(newRows));

  return {
    success: true,
    lines: newRows.length
  };
}

function updateReceiptHeader(ontvangstId, updates) {
  const receiptId = safeText(ontvangstId);
  const sheet = getSheetOrThrow(TABS.RECEIPTS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab Ontvangsten is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['OntvangstID']]) !== receiptId) continue;

    Object.keys(updates || {}).forEach(field => {
      if (col[field] !== undefined) {
        values[i][col[field]] = updates[field];
      }
    });

    updated = true;
    break;
  }

  if (!updated) {
    throw new Error('Ontvangst niet gevonden.');
  }

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return { success: true };
}

function buildReceiptsWithLines(receipts, lines) {
  const linesByReceipt = {};

  (lines || []).forEach(line => {
    const id = safeText(line.ontvangstId);
    if (!id) return;

    if (!linesByReceipt[id]) linesByReceipt[id] = [];
    linesByReceipt[id].push(line);
  });

  return (receipts || []).map(receipt => {
    const receiptLines = linesByReceipt[receipt.ontvangstId] || [];
    return {
      ...receipt,
      lines: receiptLines,
      lineCount: receiptLines.length,
      deltaCount: receiptLines.filter(line => safeNumber(line.deltaAantal, 0) !== 0).length
    };
  });
}

function getReceiptsData(sessionId) {
  assertWarehouseAccess(sessionId);

  const receipts = getAllReceipts();
  const lines = getAllReceiptLines();

  return {
    receipts: buildReceiptsWithLines(receipts, lines),
    centralWarehouse: getCentralWarehouseOverview()
  };
}

function createManualReceipt(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertWarehouseAccess(sessionId);

  const bestelbonNr = safeText(payload.bestelbonNr);
  const externeReferentie = safeText(payload.externeReferentie);
  const bronType = safeText(payload.bronType || 'Manueel');
  const documentDatum = safeText(payload.documentDatum);
  const ontvangstdatum = safeText(payload.ontvangstdatum);

  if (!documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!ontvangstdatum) throw new Error('Ontvangstdatum is verplicht.');

  const ontvangstId = makeReceiptId();

  appendObjects(TABS.RECEIPTS, [{
    OntvangstID: ontvangstId,
    TypeMateriaal: '',
    Leverancier: 'Fluvius',
    BestelbonNr: bestelbonNr,
    ExterneReferentie: externeReferentie,
    BronType: bronType,
    BronBestand: '',
    DocumentDatum: documentDatum,
    Ontvangstdatum: ontvangstdatum,
    Status: RECEIPT_STATUS.IN_PROGRESS,
    IngediendDoor: '',
    IngediendOp: '',
    GoedgekeurdDoor: '',
    GoedgekeurdOp: '',
    ManagerOpmerking: ''
  }]);

  writeAudit(
    'Ontvangst aangemaakt',
    user.rol,
    user.naam || user.email || 'Magazijn',
    'Ontvangst',
    ontvangstId,
    {
      bestelbonNr: bestelbonNr,
      documentDatum: documentDatum,
      ontvangstdatum: ontvangstdatum,
      bronType: bronType
    }
  );

  return {
    success: true,
    ontvangstId,
    message: 'Ontvangst aangemaakt.'
  };
}

function validateReceiptLines(cleanedLines) {
  if (!(cleanedLines || []).length) {
    throw new Error('Geen geldige ontvangstlijnen om te bewaren.');
  }

  cleanedLines.forEach(line => {
    const artikelCode = safeText(line.artikelCode);
    const artikelOmschrijving = safeText(line.artikelOmschrijving);

    if (!artikelCode || !artikelOmschrijving) {
      throw new Error('Artikelcode en omschrijving zijn verplicht op elke lijn.');
    }

    const besteldAantal = safeNumber(line.besteldAantal, 0);
    const ontvangenAantal = safeNumber(line.ontvangenAantal, 0);
    const deltaAantal = besteldAantal - ontvangenAantal;
    const redenDelta = safeText(line.redenDelta);
    const opmerking = safeText(line.opmerking);

    if (besteldAantal < 0 || ontvangenAantal < 0) {
      throw new Error('Besteld en ontvangen aantal mogen niet negatief zijn.');
    }

    if (deltaAantal !== 0) {
      if (!RECEIPT_DELTA_REASONS.includes(redenDelta)) {
        throw new Error('Ongeldige reden delta voor artikel ' + artikelCode);
      }

      if (!redenDelta) {
        throw new Error('Reden delta is verplicht voor artikel ' + artikelCode);
      }

      if (redenDelta === 'Andere reden' && !opmerking) {
        throw new Error('Opmerking is verplicht bij Andere reden voor artikel ' + artikelCode);
      }
    }
  });
}

function saveReceiptLines(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertWarehouseAccess(sessionId);

  const ontvangstId = safeText(payload.ontvangstId);
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!ontvangstId) throw new Error('OntvangstID ontbreekt.');
  if (!lines.length) throw new Error('Geen ontvangstlijnen ontvangen.');

  const receipt = getReceiptById(ontvangstId);
  if (!receipt) throw new Error('Ontvangst niet gevonden.');

  if ([RECEIPT_STATUS.SUBMITTED, RECEIPT_STATUS.APPROVED].includes(receipt.status)) {
    throw new Error('Deze ontvangst kan niet meer aangepast worden.');
  }

  const cleaned = lines
    .map(line => buildReceiptLineObject(ontvangstId, {
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      eenheid: line.eenheid,
      besteldAantal: line.besteldAantal,
      ontvangenAantal: line.ontvangenAantal,
      redenDelta: line.redenDelta,
      opmerking: line.opmerking
    }))
    .filter(line => safeText(line.ArtikelCode) && safeText(line.ArtikelOmschrijving));

  const cleanedForValidation = cleaned.map(line => ({
    artikelCode: line.ArtikelCode,
    artikelOmschrijving: line.ArtikelOmschrijving,
    eenheid: line.Eenheid,
    besteldAantal: line.BesteldAantal,
    ontvangenAantal: line.OntvangenAantal,
    deltaAantal: line.DeltaAantal,
    redenDelta: line.RedenDelta,
    opmerking: line.Opmerking
  }));

  validateReceiptLines(cleanedForValidation);

  replaceReceiptLinesForReceipt(ontvangstId, cleaned);

  const receiptType = determineReceiptMaterialType(
    cleaned.map(line => ({ artikelCode: line.ArtikelCode }))
  );

  updateReceiptHeader(ontvangstId, {
    TypeMateriaal: receiptType,
    Leverancier: 'Fluvius'
  });

  writeAudit(
    'Ontvangstlijnen opgeslagen',
    user.rol,
    user.naam || user.email || 'Magazijn',
    'Ontvangst',
    ontvangstId,
    {
      lijnen: cleaned.length,
      typeMateriaal: receiptType
    }
  );

  return {
    success: true,
    lines: cleaned.length,
    message: 'Ontvangstlijnen opgeslagen.'
  };
}

function submitReceipt(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertWarehouseAccess(sessionId);

  const ontvangstId = safeText(payload.ontvangstId);
  const actor = safeText(payload.actor || user.naam || 'Magazijn');

  if (!ontvangstId) throw new Error('OntvangstID ontbreekt.');

  const receipt = getReceiptById(ontvangstId);
  if (!receipt) throw new Error('Ontvangst niet gevonden.');

  const lineRows = getReceiptLinesByReceiptId(ontvangstId);
  if (!lineRows.length) {
    throw new Error('Deze ontvangst bevat nog geen actieve lijnen.');
  }

  updateReceiptHeader(ontvangstId, {
    Status: RECEIPT_STATUS.SUBMITTED,
    IngediendDoor: actor,
    IngediendOp: nowStamp()
  });

  pushManagerNotification(
    'OntvangstGoedTeKeuren',
    'Ontvangst ingediend',
    `Ontvangst ${ontvangstId} is ingediend en wacht op goedkeuring.`,
    'Ontvangst',
    ontvangstId
  );

  writeAudit(
    'Ontvangst ingediend',
    user.rol,
    actor,
    'Ontvangst',
    ontvangstId,
    {
      status: RECEIPT_STATUS.SUBMITTED,
      lijnen: lineRows.length
    }
  );

  return {
    success: true,
    message: 'Ontvangst ingediend voor goedkeuring.'
  };
}

function buildReceiptApprovalMovementPayloads(receipt, receiptLines, actor, note) {
  const documentDatum = safeText(receipt.documentDatumIso || receipt.documentDatum);
  const typeMateriaal = safeText(receipt.typeMateriaal);

  return (receiptLines || [])
    .filter(line => safeNumber(line.ontvangenAantal, 0) > 0)
    .map(line => ({
      datumDocument: documentDatum,
      typeMutatie: 'Ontvangst',
      bronId: receipt.ontvangstId,
      typeMateriaal: typeMateriaal || determineMaterialTypeFromArticle(line.artikelCode),
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      eenheid: line.eenheid,
      aantalIn: safeNumber(line.ontvangenAantal, 0),
      aantalUit: 0,
      nettoAantal: safeNumber(line.ontvangenAantal, 0),
      locatieVan: '',
      locatieNaar: LOCATION.CENTRAL,
      reden: safeText(line.redenDelta),
      opmerking: safeText(line.opmerking || note),
      goedgekeurdDoor: actor,
      goedgekeurdOp: nowStamp()
    }));
}

function approveReceipt(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertManagerAccess(sessionId);

  const ontvangstId = safeText(payload.ontvangstId);
  const actor = safeText(payload.actor || user.naam || 'Manager');
  const note = safeText(payload.note);

  if (!ontvangstId) throw new Error('OntvangstID ontbreekt.');

  const receipt = getReceiptById(ontvangstId);
  if (!receipt) throw new Error('Ontvangst niet gevonden.');

  const receiptLines = getReceiptLinesByReceiptId(ontvangstId);
  if (!receiptLines.length) {
    throw new Error('Deze ontvangst bevat geen actieve lijnen.');
  }

  updateReceiptHeader(ontvangstId, {
    Status: RECEIPT_STATUS.APPROVED,
    GoedgekeurdDoor: actor,
    GoedgekeurdOp: nowStamp(),
    ManagerOpmerking: note
  });

  const movementPayloads = buildReceiptApprovalMovementPayloads(receipt, receiptLines, actor, note);

  replaceSourceMovements('Ontvangst', ontvangstId, movementPayloads);
  rebuildCentralWarehouseOverview();

  markManagerNotificationsBySource('Ontvangst', ontvangstId);

  pushWarehouseNotification(
    'OntvangstGoedgekeurd',
    'Ontvangst goedgekeurd',
    `Ontvangst ${ontvangstId} is goedgekeurd en geboekt in centraal magazijn.`,
    'Ontvangst',
    ontvangstId
  );

  writeAudit(
    'Ontvangst goedgekeurd',
    user.rol,
    actor,
    'Ontvangst',
    ontvangstId,
    {
      note: note,
      typeMateriaal: receipt.typeMateriaal,
      lijnen: movementPayloads.length
    }
  );

  return {
    success: true,
    message: 'Ontvangst goedgekeurd en centraal magazijn bijgewerkt.'
  };
}