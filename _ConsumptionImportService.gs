/* =========================================================
   44_ConsumptionImportService.gs
   Refactor: consumption import service
   Doel:
   - ruwe verbruiksregels uit Excel/CSV normaliseren
   - lijnen groeperen per technieker, datum en artikel
   - previewlaag voor importcontrole
   - importbatch bewaren voor verdere verwerking
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getConsumptionImportBatchTab_() {
  return TABS.CONSUMPTION_IMPORT_BATCHES || 'VerbruiksImportBatches';
}

function getConsumptionImportLineTab_() {
  return TABS.CONSUMPTION_IMPORT_LINES || 'VerbruiksImportLijnen';
}

function getConsumptionImportStatusOpen_() {
  if (typeof CONSUMPTION_IMPORT_STATUS !== 'undefined' && CONSUMPTION_IMPORT_STATUS && CONSUMPTION_IMPORT_STATUS.OPEN) {
    return CONSUMPTION_IMPORT_STATUS.OPEN;
  }
  return 'Open';
}

function getConsumptionImportStatusProcessed_() {
  if (typeof CONSUMPTION_IMPORT_STATUS !== 'undefined' && CONSUMPTION_IMPORT_STATUS && CONSUMPTION_IMPORT_STATUS.PROCESSED) {
    return CONSUMPTION_IMPORT_STATUS.PROCESSED;
  }
  return 'Verwerkt';
}

function getConsumptionImportStatusClosed_() {
  if (typeof CONSUMPTION_IMPORT_STATUS !== 'undefined' && CONSUMPTION_IMPORT_STATUS && CONSUMPTION_IMPORT_STATUS.CLOSED) {
    return CONSUMPTION_IMPORT_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function makeConsumptionImportBatchId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'CIB-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function makeConsumptionImportLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'CIL-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapConsumptionImportBatch(row) {
  return {
    importBatchId: safeText(row.ImportBatchID || row.ImportBatchId || row.ID),
    sourceType: safeText(row.SourceType || 'Upload'),
    sourceName: safeText(row.SourceName || row.BestandsNaam || row.FileName),
    sourceReference: safeText(row.SourceReference || row.Reference),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    status: safeText(row.Status || getConsumptionImportStatusOpen_()),
    actor: safeText(row.Actor),
    opmerking: safeText(row.Opmerking),
    rawLineCount: safeNumber(row.RawLineCount, 0),
    groupedLineCount: safeNumber(row.GroupedLineCount, 0),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    verwerktOp: safeText(row.VerwerktOp),
  };
}

function mapConsumptionImportLine(row) {
  return {
    importLineId: safeText(row.ImportLineID || row.ImportLineId || row.ID),
    importBatchId: safeText(row.ImportBatchID || row.ImportBatchId),
    techniekerCode: safeText(row.TechniekerCode || row.TechnicianCode || row.TechCode),
    techniekerNaam: safeText(row.TechniekerNaam || row.TechnicianName),
    werkorderNr: safeText(row.WerkorderNr || row.WorkOrderNr || row.OrderNr),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode || row.ArtikelNr))),
    eenheid: safeText(row.Eenheid || row.Unit),
    aantal: safeNumber(row.Aantal || row.Quantity, 0),
    bronRij: safeText(row.BronRij || row.SourceRow),
    opmerking: safeText(row.Opmerking),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllConsumptionImportBatches() {
  return readObjectsSafe(getConsumptionImportBatchTab_())
    .map(mapConsumptionImportBatch)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.aangemaaktOpRaw).localeCompare(safeText(a.aangemaaktOpRaw)) ||
        safeText(b.importBatchId).localeCompare(safeText(a.importBatchId))
      );
    });
}

function getAllConsumptionImportLines() {
  return readObjectsSafe(getConsumptionImportLineTab_())
    .map(mapConsumptionImportLine)
    .sort(function (a, b) {
      return (
        safeText(a.importBatchId).localeCompare(safeText(b.importBatchId)) ||
        safeText(a.techniekerCode).localeCompare(safeText(b.techniekerCode)) ||
        safeText(a.documentDatumIso).localeCompare(safeText(b.documentDatumIso)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
      );
    });
}

function getConsumptionImportBatchById(importBatchId) {
  var id = safeText(importBatchId);
  if (!id) return null;

  return getAllConsumptionImportBatches().find(function (item) {
    return safeText(item.importBatchId) === id;
  }) || null;
}

function getConsumptionImportLinesByBatchId(importBatchId) {
  var id = safeText(importBatchId);
  return getAllConsumptionImportLines().filter(function (item) {
    return safeText(item.importBatchId) === id;
  });
}

function buildConsumptionImportBatchesWithLines(headers, lines) {
  var lineMap = {};

  (lines || []).forEach(function (line) {
    var id = safeText(line.importBatchId);
    if (!lineMap[id]) lineMap[id] = [];
    lineMap[id].push(line);
  });

  return (headers || []).map(function (header) {
    var importLines = lineMap[safeText(header.importBatchId)] || [];

    return Object.assign({}, header, {
      lines: importLines,
      lineCount: importLines.length,
      totaalAantal: importLines.reduce(function (sum, line) {
        return sum + safeNumber(line.aantal, 0);
      }, 0),
      techniekerCount: Object.keys(
        importLines.reduce(function (acc, line) {
          acc[safeText(line.techniekerCode)] = true;
          return acc;
        }, {})
      ).length,
    });
  });
}

function getConsumptionImportBatchWithLines(importBatchId) {
  var header = getConsumptionImportBatchById(importBatchId);
  if (!header) return null;
  return buildConsumptionImportBatchesWithLines(
    [header],
    getConsumptionImportLinesByBatchId(importBatchId)
  )[0] || null;
}

/* ---------------------------------------------------------
   Normalization helpers
   --------------------------------------------------------- */

function normalizeConsumptionImportLine_(line, index) {
  line = line || {};

  var techniekerCode = safeText(
    line.techniekerCode ||
    line.technicianCode ||
    line.techCode ||
    line.ref ||
    line.busCode
  );

  var documentDatum = safeText(
    line.documentDatum ||
    line.documentDate ||
    line.datum ||
    line.date
  );

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

  var aantal = safeNumber(
    line.aantal ||
    line.quantity ||
    line.verbruik ||
    line.qty,
    0
  );

  return {
    techniekerCode: techniekerCode,
    techniekerNaam: safeText(line.techniekerNaam || line.technicianName),
    werkorderNr: safeText(line.werkorderNr || line.workOrderNr || line.orderNr),
    documentDatum: documentDatum,
    documentDatumIso: documentDatum,
    artikelCode: artikelCode,
    artikelOmschrijving: artikelOmschrijving,
    typeMateriaal: safeText(
      line.typeMateriaal ||
      determineMaterialTypeFromArticle(artikelCode)
    ),
    eenheid: safeText(line.eenheid || line.unit || 'Stuk'),
    aantal: aantal,
    bronRij: safeText(line.bronRij || line.sourceRow || (index + 1)),
    opmerking: safeText(line.opmerking || line.remark),
  };
}

function validateConsumptionImportLines_(lines) {
  var rows = Array.isArray(lines) ? lines : [];
  if (!rows.length) {
    throw new Error('Geen verbruikslijnen ontvangen.');
  }

  rows.forEach(function (line, index) {
    var rowNr = index + 1;

    if (!safeText(line.techniekerCode)) {
      throw new Error('TechniekerCode ontbreekt op lijn ' + rowNr + '.');
    }
    if (!safeText(line.documentDatumIso)) {
      throw new Error('Documentdatum ontbreekt op lijn ' + rowNr + '.');
    }
    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.aantal, 0) <= 0) {
      throw new Error('Aantal moet groter zijn dan 0 op lijn ' + rowNr + '.');
    }
  });

  return true;
}

function buildConsumptionImportGroupKey_(line) {
  return [
    safeText(line.techniekerCode),
    safeText(line.documentDatumIso),
    safeText(line.artikelCode),
    safeText(line.eenheid),
    safeText(line.typeMateriaal),
    safeText(line.werkorderNr)
  ].join('|');
}

function groupConsumptionImportLines_(lines) {
  var grouped = {};

  (Array.isArray(lines) ? lines : []).forEach(function (line) {
    var key = buildConsumptionImportGroupKey_(line);

    if (!grouped[key]) {
      grouped[key] = {
        techniekerCode: safeText(line.techniekerCode),
        techniekerNaam: safeText(line.techniekerNaam),
        werkorderNr: safeText(line.werkorderNr),
        documentDatum: safeText(line.documentDatum),
        documentDatumIso: safeText(line.documentDatumIso),
        artikelCode: safeText(line.artikelCode),
        artikelOmschrijving: safeText(line.artikelOmschrijving),
        typeMateriaal: safeText(line.typeMateriaal),
        eenheid: safeText(line.eenheid),
        aantal: 0,
        bronRij: safeText(line.bronRij),
        opmerking: safeText(line.opmerking),
      };
    }

    grouped[key].aantal += safeNumber(line.aantal, 0);

    if (!grouped[key].techniekerNaam) {
      grouped[key].techniekerNaam = safeText(line.techniekerNaam);
    }
    if (!grouped[key].artikelOmschrijving) {
      grouped[key].artikelOmschrijving = safeText(line.artikelOmschrijving);
    }
    if (!grouped[key].opmerking && safeText(line.opmerking)) {
      grouped[key].opmerking = safeText(line.opmerking);
    }
  });

  return Object.keys(grouped)
    .map(function (key) { return grouped[key]; })
    .sort(function (a, b) {
      return (
        safeText(a.techniekerCode).localeCompare(safeText(b.techniekerCode)) ||
        safeText(a.documentDatumIso).localeCompare(safeText(b.documentDatumIso)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function summarizeConsumptionImportLines_(lines) {
  var rows = Array.isArray(lines) ? lines : [];

  return {
    lineCount: rows.length,
    totaalAantal: rows.reduce(function (sum, row) {
      return sum + safeNumber(row.aantal, 0);
    }, 0),
    techniekerCount: Object.keys(
      rows.reduce(function (acc, row) {
        acc[safeText(row.techniekerCode)] = true;
        return acc;
      }, {})
    ).length,
    artikelCount: Object.keys(
      rows.reduce(function (acc, row) {
        acc[safeText(row.artikelCode)] = true;
        return acc;
      }, {})
    ).length,
  };
}

/* ---------------------------------------------------------
   Access policy
   --------------------------------------------------------- */

function assertConsumptionImportAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.MOBILE_WAREHOUSE],
    'Geen rechten voor verbruiksimport.'
  );
  return user;
}

/* ---------------------------------------------------------
   Preview
   --------------------------------------------------------- */

function previewConsumptionImportLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  assertConsumptionImportAccess_(sessionId);

  var rawLines = Array.isArray(payload.lines) ? payload.lines : [];
  var normalizedLines = rawLines.map(normalizeConsumptionImportLine_);
  validateConsumptionImportLines_(normalizedLines);

  var groupedLines = groupConsumptionImportLines_(normalizedLines);

  return {
    items: groupedLines,
    summary: summarizeConsumptionImportLines_(groupedLines),
    rawLineCount: rawLines.length,
  };
}

/* ---------------------------------------------------------
   Import save
   --------------------------------------------------------- */

function importConsumptionLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionImportAccess_(sessionId);

  var rawLines = Array.isArray(payload.lines) ? payload.lines : [];
  var normalizedLines = rawLines.map(normalizeConsumptionImportLine_);
  validateConsumptionImportLines_(normalizedLines);

  var groupedLines = groupConsumptionImportLines_(normalizedLines);

  var sourceType = safeText(payload.sourceType || 'Upload');
  var sourceName = safeText(payload.sourceName || payload.bestandsNaam || payload.fileName);
  var sourceReference = safeText(payload.sourceReference || payload.reference);
  var documentDatum = safeText(
    payload.documentDatum ||
    payload.documentDate ||
    (groupedLines.length ? groupedLines[0].documentDatumIso : Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd'))
  );
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var importBatchId = makeConsumptionImportBatchId_();

  var batchObj = {
    ImportBatchID: importBatchId,
    SourceType: sourceType,
    SourceName: sourceName,
    SourceReference: sourceReference,
    DocumentDatum: documentDatum,
    DocumentDatumIso: documentDatum,
    Status: getConsumptionImportStatusOpen_(),
    Actor: safeText(payload.actor || actor.naam || actor.email),
    Opmerking: safeText(payload.opmerking || payload.remark),
    RawLineCount: rawLines.length,
    GroupedLineCount: groupedLines.length,
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    VerwerktOp: '',
  };

  appendObjects(getConsumptionImportBatchTab_(), [batchObj]);

  var lineObjects = groupedLines.map(function (line) {
    return {
      ImportLineID: makeConsumptionImportLineId_(),
      ImportBatchID: importBatchId,
      TechniekerCode: line.techniekerCode,
      TechniekerNaam: line.techniekerNaam,
      WerkorderNr: line.werkorderNr,
      DocumentDatum: line.documentDatum,
      DocumentDatumIso: line.documentDatumIso,
      ArtikelCode: line.artikelCode,
      ArtikelOmschrijving: line.artikelOmschrijving,
      TypeMateriaal: line.typeMateriaal,
      Eenheid: line.eenheid,
      Aantal: line.aantal,
      BronRij: line.bronRij,
      Opmerking: line.opmerking,
    };
  });

  if (lineObjects.length) {
    appendObjects(getConsumptionImportLineTab_(), lineObjects);
  }

  writeAudit({
    actie: 'IMPORT_CONSUMPTION_LINES',
    actor: actor,
    documentType: 'VerbruiksImport',
    documentId: importBatchId,
    details: {
      sourceType: sourceType,
      rawLineCount: rawLines.length,
      groupedLineCount: groupedLines.length,
    },
  });

  return {
    batch: getConsumptionImportBatchWithLines(importBatchId),
    summary: summarizeConsumptionImportLines_(groupedLines),
    rawLineCount: rawLines.length,
    groupedLineCount: groupedLines.length,
  };
}

/* ---------------------------------------------------------
   Status updates
   --------------------------------------------------------- */

function markConsumptionImportProcessed(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionImportAccess_(sessionId);
  var importBatchId = safeText(payload.importBatchId);
  if (!importBatchId) throw new Error('ImportBatchId ontbreekt.');

  var batch = getConsumptionImportBatchById(importBatchId);
  if (!batch) throw new Error('Importbatch niet gevonden.');

  var table = getAllValues(getConsumptionImportBatchTab_());
  if (!table.length) throw new Error('Importbatchtab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getConsumptionImportBatchTab_()).map(function (row) {
    var current = mapConsumptionImportBatch(row);
    if (safeText(current.importBatchId) !== importBatchId) {
      return row;
    }

    row.Status = getConsumptionImportStatusProcessed_();
    row.VerwerktOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getConsumptionImportBatchTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  writeAudit({
    actie: 'MARK_CONSUMPTION_IMPORT_PROCESSED',
    actor: actor,
    documentType: 'VerbruiksImport',
    documentId: importBatchId,
    details: {},
  });

  return getConsumptionImportBatchWithLines(importBatchId);
}

/* ---------------------------------------------------------
   Queries
   --------------------------------------------------------- */

function getConsumptionImportData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionImportAccess_(sessionId);
  var rows = buildConsumptionImportBatchesWithLines(
    getAllConsumptionImportBatches(),
    getAllConsumptionImportLines()
  );

  return {
    items: rows,
    batches: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getConsumptionImportStatusOpen_(); }).length,
      verwerkt: rows.filter(function (x) { return safeText(x.status) === getConsumptionImportStatusProcessed_(); }).length,
      gesloten: rows.filter(function (x) { return safeText(x.status) === getConsumptionImportStatusClosed_(); }).length,
      actorRol: safeText(actor.rol),
    }
  };
}
