/* =========================================================
   40_ReceiptImportService.gs
   Refactor: receipt import service
   Doel:
   - browser-geparste ontvangstlijnen normaliseren
   - import bundelen per artikel
   - manuele ontvangst automatisch aanmaken
   - lijnen opslaan via ReceiptService
   ========================================================= */

/* ---------------------------------------------------------
   Helpers
   --------------------------------------------------------- */

function normalizeReceiptImportLine_(line) {
  line = line || {};

  var artikelCode = safeText(
    line.artikelCode ||
    line.artikelNr ||
    line.code ||
    line.articleCode
  );

  var artikelOmschrijving = safeText(
    line.artikelOmschrijving ||
    line.artikel ||
    line.omschrijving ||
    line.description
  );

  var eenheid = safeText(
    line.eenheid ||
    line.unit ||
    'Stuk'
  );

  var besteldAantal = safeNumber(
    line.besteldAantal ||
    line.besteld ||
    line.expectedQty,
    0
  );

  var ontvangenAantal = safeNumber(
    line.ontvangenAantal ||
    line.ontvangen ||
    line.receivedQty ||
    line.aantal,
    0
  );

  return {
    artikelCode: artikelCode,
    artikelOmschrijving: artikelOmschrijving,
    typeMateriaal: safeText(
      line.typeMateriaal ||
      determineMaterialTypeFromArticle(artikelCode)
    ),
    eenheid: eenheid,
    besteldAantal: besteldAantal,
    ontvangenAantal: ontvangenAantal,
    deltaAantal: safeNumber(
      line.deltaAantal,
      ontvangenAantal - besteldAantal
    ),
    redenDelta: safeText(line.redenDelta || line.deltaReason),
    opmerking: safeText(line.opmerking || line.remark || line.note),
    palletNr: safeText(line.palletNr || line.pallet || ''),
    label: safeText(line.label || ''),
  };
}

function validateReceiptImportLines_(lines) {
  var rows = Array.isArray(lines) ? lines : [];
  if (!rows.length) {
    throw new Error('Geen importlijnen ontvangen.');
  }

  rows.forEach(function (line, index) {
    var rowNr = index + 1;

    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op importlijn ' + rowNr + '.');
    }
    if (safeNumber(line.ontvangenAantal, 0) < 0) {
      throw new Error('Ontvangen aantal mag niet negatief zijn op importlijn ' + rowNr + '.');
    }
    if (
      safeNumber(line.deltaAantal, 0) !== 0 &&
      !safeText(line.redenDelta)
    ) {
      throw new Error('Reden delta is verplicht op importlijn ' + rowNr + '.');
    }
  });

  return true;
}

function buildReceiptImportGroupKey_(line) {
  return [
    safeText(line.artikelCode),
    safeText(line.eenheid),
    safeText(line.typeMateriaal)
  ].join('|');
}

function groupImportedReceiptLines_(lines) {
  var grouped = {};

  (Array.isArray(lines) ? lines : []).forEach(function (line) {
    var key = buildReceiptImportGroupKey_(line);

    if (!grouped[key]) {
      grouped[key] = {
        artikelCode: safeText(line.artikelCode),
        artikelOmschrijving: safeText(line.artikelOmschrijving),
        typeMateriaal: safeText(line.typeMateriaal),
        eenheid: safeText(line.eenheid),
        besteldAantal: 0,
        ontvangenAantal: 0,
        deltaAantal: 0,
        redenDelta: '',
        opmerking: '',
        palletNr: '',
        label: '',
      };
    }

    grouped[key].besteldAantal += safeNumber(line.besteldAantal, 0);
    grouped[key].ontvangenAantal += safeNumber(line.ontvangenAantal, 0);
    grouped[key].deltaAantal =
      safeNumber(grouped[key].ontvangenAantal, 0) -
      safeNumber(grouped[key].besteldAantal, 0);

    if (!grouped[key].artikelOmschrijving) {
      grouped[key].artikelOmschrijving = safeText(line.artikelOmschrijving);
    }
    if (!grouped[key].opmerking && safeText(line.opmerking)) {
      grouped[key].opmerking = safeText(line.opmerking);
    }
    if (!grouped[key].palletNr && safeText(line.palletNr)) {
      grouped[key].palletNr = safeText(line.palletNr);
    }
    if (!grouped[key].label && safeText(line.label)) {
      grouped[key].label = safeText(line.label);
    }

    if (safeText(line.redenDelta)) {
      grouped[key].redenDelta = safeText(line.redenDelta);
    }
  });

  return Object.keys(grouped)
    .map(function (key) {
      var row = grouped[key];
      if (safeNumber(row.deltaAantal, 0) === 0) {
        row.redenDelta = '';
      }
      return row;
    })
    .sort(function (a, b) {
      return (
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function deriveReceiptImportMaterialType_(lines) {
  var types = {};
  (Array.isArray(lines) ? lines : []).forEach(function (line) {
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

function summarizeImportedReceiptLines_(lines) {
  var rows = Array.isArray(lines) ? lines : [];

  return {
    lineCount: rows.length,
    articleCount: rows.length,
    totaalBesteld: rows.reduce(function (sum, row) {
      return sum + safeNumber(row.besteldAantal, 0);
    }, 0),
    totaalOntvangen: rows.reduce(function (sum, row) {
      return sum + safeNumber(row.ontvangenAantal, 0);
    }, 0),
    deltaLijnen: rows.filter(function (row) {
      return safeNumber(row.deltaAantal, 0) !== 0;
    }).length,
  };
}

/* ---------------------------------------------------------
   Main import flow
   --------------------------------------------------------- */

function importParsedReceiptLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertWarehouseAccess(sessionId);

  if (typeof createManualReceipt !== 'function' || typeof saveReceiptLines !== 'function') {
    throw new Error('Receipt service ontbreekt. Werk eerst het receiptblok in.');
  }

  var rawLines = Array.isArray(payload.lines) ? payload.lines : [];
  var normalizedLines = rawLines.map(normalizeReceiptImportLine_);
  validateReceiptImportLines_(normalizedLines);

  var groupedLines = groupImportedReceiptLines_(normalizedLines);
  var materialType = deriveReceiptImportMaterialType_(groupedLines);

  var headerPayload = {
    sessionId: sessionId,
    leverancier: safeText(payload.leverancier || payload.supplier || 'Fluvius'),
    typeMateriaal: safeText(payload.typeMateriaal || materialType),
    bestelbonNr: safeText(payload.bestelbonNr || payload.purchaseOrderNumber || payload.poNumber),
    externeReferentie: safeText(payload.externeReferentie || payload.reference || payload.fluviusReferentie),
    documentDatum: safeText(payload.documentDatum || payload.documentDate),
    ontvangstDatum: safeText(payload.ontvangstDatum || payload.receiveDate || payload.documentDatum || payload.documentDate),
    bronType: safeText(payload.bronType || payload.sourceType || 'Upload'),
    bronReferentie: safeText(payload.bronReferentie || payload.sourceReference || payload.bestandsNaam || payload.fileName),
    opmerking: safeText(payload.opmerking || payload.remark),
    actor: safeText(actor.naam || actor.email),
  };

  var receipt = createManualReceipt(headerPayload);

  var saved = saveReceiptLines({
    sessionId: sessionId,
    receiptId: receipt.receiptId,
    lines: groupedLines,
  });

  writeAudit({
    actie: 'IMPORT_RECEIPT_LINES',
    actor: actor,
    documentType: 'Ontvangst',
    documentId: receipt.receiptId,
    details: {
      sourceType: safeText(headerPayload.bronType),
      sourceReference: safeText(headerPayload.bronReferentie),
      rawLineCount: rawLines.length,
      groupedLineCount: groupedLines.length,
    },
  });

  return {
    receipt: saved,
    summary: summarizeImportedReceiptLines_(groupedLines),
    rawLineCount: rawLines.length,
    groupedLineCount: groupedLines.length,
  };
}

/* ---------------------------------------------------------
   Preview helper
   --------------------------------------------------------- */

function previewParsedReceiptLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  assertWarehouseAccess(sessionId);

  var rawLines = Array.isArray(payload.lines) ? payload.lines : [];
  var normalizedLines = rawLines.map(normalizeReceiptImportLine_);
  validateReceiptImportLines_(normalizedLines);

  var groupedLines = groupImportedReceiptLines_(normalizedLines);

  return {
    items: groupedLines,
    summary: summarizeImportedReceiptLines_(groupedLines),
    rawLineCount: rawLines.length,
  };
}
