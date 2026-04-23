/* =========================================================
   45_ConsumptionProcessingService.gs
   Refactor: consumption processing service
   Doel:
   - geïmporteerde verbruiksregels verwerken tot boekbare voorstellen
   - validatie per technieker, artikel en busstock
   - bundelen per technieker en datum
   - voorbereiding op boeking naar mutaties
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getConsumptionProcessBatchTab_() {
  return TABS.CONSUMPTION_PROCESS_BATCHES || 'VerbruiksProcessBatches';
}

function getConsumptionProcessLineTab_() {
  return TABS.CONSUMPTION_PROCESS_LINES || 'VerbruiksProcessLijnen';
}

function getConsumptionProcessStatusOpen_() {
  if (typeof CONSUMPTION_PROCESS_STATUS !== 'undefined' && CONSUMPTION_PROCESS_STATUS && CONSUMPTION_PROCESS_STATUS.OPEN) {
    return CONSUMPTION_PROCESS_STATUS.OPEN;
  }
  return 'Open';
}

function getConsumptionProcessStatusReady_() {
  if (typeof CONSUMPTION_PROCESS_STATUS !== 'undefined' && CONSUMPTION_PROCESS_STATUS && CONSUMPTION_PROCESS_STATUS.READY) {
    return CONSUMPTION_PROCESS_STATUS.READY;
  }
  return 'Klaar';
}

function getConsumptionProcessStatusBooked_() {
  if (typeof CONSUMPTION_PROCESS_STATUS !== 'undefined' && CONSUMPTION_PROCESS_STATUS && CONSUMPTION_PROCESS_STATUS.BOOKED) {
    return CONSUMPTION_PROCESS_STATUS.BOOKED;
  }
  return 'Geboekt';
}

function getConsumptionProcessStatusClosed_() {
  if (typeof CONSUMPTION_PROCESS_STATUS !== 'undefined' && CONSUMPTION_PROCESS_STATUS && CONSUMPTION_PROCESS_STATUS.CLOSED) {
    return CONSUMPTION_PROCESS_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function getConsumptionValidationOk_() {
  return 'OK';
}

function getConsumptionValidationWarning_() {
  return 'WAARSCHUWING';
}

function getConsumptionValidationError_() {
  return 'FOUT';
}

function makeConsumptionProcessBatchId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'CPB-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function makeConsumptionProcessLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'CPL-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapConsumptionProcessBatch(row) {
  return {
    processBatchId: safeText(row.ProcessBatchID || row.ProcessBatchId || row.ID),
    importBatchId: safeText(row.ImportBatchID || row.ImportBatchId),
    sourceType: safeText(row.SourceType || 'Import'),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    status: safeText(row.Status || getConsumptionProcessStatusOpen_()),
    actor: safeText(row.Actor),
    opmerking: safeText(row.Opmerking),
    rawLineCount: safeNumber(row.RawLineCount, 0),
    processedLineCount: safeNumber(row.ProcessedLineCount, 0),
    warningCount: safeNumber(row.WarningCount, 0),
    errorCount: safeNumber(row.ErrorCount, 0),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    geboektOp: safeText(row.GeboektOp),
  };
}

function mapConsumptionProcessLine(row) {
  return {
    processLineId: safeText(row.ProcessLineID || row.ProcessLineId || row.ID),
    processBatchId: safeText(row.ProcessBatchID || row.ProcessBatchId),
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
    systemBusAantal: safeNumber(row.SystemBusAantal || row.SystemQty, 0),
    projectedBusAantal: safeNumber(row.ProjectedBusAantal || row.ProjectedQty, 0),
    validationStatus: safeText(row.ValidationStatus || getConsumptionValidationOk_()),
    validationMessage: safeText(row.ValidationMessage || ''),
    bookingGroupKey: safeText(row.BookingGroupKey || ''),
    opmerking: safeText(row.Opmerking),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllConsumptionProcessBatches() {
  return readObjectsSafe(getConsumptionProcessBatchTab_())
    .map(mapConsumptionProcessBatch)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.aangemaaktOpRaw).localeCompare(safeText(a.aangemaaktOpRaw)) ||
        safeText(b.processBatchId).localeCompare(safeText(a.processBatchId))
      );
    });
}

function getAllConsumptionProcessLines() {
  return readObjectsSafe(getConsumptionProcessLineTab_())
    .map(mapConsumptionProcessLine)
    .sort(function (a, b) {
      return (
        safeText(a.processBatchId).localeCompare(safeText(b.processBatchId)) ||
        safeText(a.techniekerCode).localeCompare(safeText(b.techniekerCode)) ||
        safeText(a.documentDatumIso).localeCompare(safeText(b.documentDatumIso)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function getConsumptionProcessBatchById(processBatchId) {
  var id = safeText(processBatchId);
  if (!id) return null;

  return getAllConsumptionProcessBatches().find(function (item) {
    return safeText(item.processBatchId) === id;
  }) || null;
}

function getConsumptionProcessLinesByBatchId(processBatchId) {
  var id = safeText(processBatchId);
  return getAllConsumptionProcessLines().filter(function (item) {
    return safeText(item.processBatchId) === id;
  });
}

function buildConsumptionProcessBatchesWithLines(headers, lines) {
  var lineMap = {};

  (lines || []).forEach(function (line) {
    var id = safeText(line.processBatchId);
    if (!lineMap[id]) lineMap[id] = [];
    lineMap[id].push(line);
  });

  return (headers || []).map(function (header) {
    var processLines = lineMap[safeText(header.processBatchId)] || [];

    return Object.assign({}, header, {
      lines: processLines,
      lineCount: processLines.length,
      totaalAantal: processLines.reduce(function (sum, line) {
        return sum + safeNumber(line.aantal, 0);
      }, 0),
      techniekerCount: Object.keys(
        processLines.reduce(function (acc, line) {
          acc[safeText(line.techniekerCode)] = true;
          return acc;
        }, {})
      ).length,
    });
  });
}

function getConsumptionProcessBatchWithLines(processBatchId) {
  var header = getConsumptionProcessBatchById(processBatchId);
  if (!header) return null;
  return buildConsumptionProcessBatchesWithLines(
    [header],
    getConsumptionProcessLinesByBatchId(processBatchId)
  )[0] || null;
}

/* ---------------------------------------------------------
   Helpers
   --------------------------------------------------------- */

function assertConsumptionProcessingAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.MOBILE_WAREHOUSE],
    'Geen rechten voor verbruiksverwerking.'
  );
  return user;
}

function resolveConsumptionImportBatchWithLines_(importBatchId) {
  var id = safeText(importBatchId);
  if (!id) throw new Error('ImportBatchId ontbreekt.');

  if (typeof getConsumptionImportBatchWithLines !== 'function') {
    throw new Error('Consumption import service ontbreekt. Werk eerst het importblok in.');
  }

  var batch = getConsumptionImportBatchWithLines(id);
  if (!batch) {
    throw new Error('Importbatch niet gevonden.');
  }

  return batch;
}

function getBusSnapshotForConsumption_(techniekerCode, artikelCode) {
  var code = safeText(techniekerCode);
  var article = safeText(artikelCode);
  if (!code || !article) return 0;

  if (typeof buildBusStockMapForTechnician !== 'function') {
    return 0;
  }

  var stockMap = buildBusStockMapForTechnician(code);
  return stockMap[article] ? safeNumber(stockMap[article].voorraadBus, 0) : 0;
}

function buildConsumptionBookingGroupKey_(line) {
  return [
    safeText(line.techniekerCode),
    safeText(line.documentDatumIso),
    safeText(line.werkorderNr)
  ].join('|');
}

function buildConsumptionProjectionLines_(importLines) {
  var runningBus = {};
  var projected = [];

  (importLines || []).forEach(function (line) {
    var techCode = safeText(line.techniekerCode);
    var articleCode = safeText(line.artikelCode);
    var stateKey = techCode + '|' + articleCode;

    if (runningBus[stateKey] === undefined) {
      runningBus[stateKey] = getBusSnapshotForConsumption_(techCode, articleCode);
    }

    var currentQty = safeNumber(runningBus[stateKey], 0);
    var consumeQty = safeNumber(line.aantal, 0);
    var projectedQty = currentQty - consumeQty;

    var validationStatus = getConsumptionValidationOk_();
    var validationMessage = '';

    if (consumeQty <= 0) {
      validationStatus = getConsumptionValidationError_();
      validationMessage = 'Verbruik moet groter zijn dan 0.';
    } else if (currentQty <= 0) {
      validationStatus = getConsumptionValidationWarning_();
      validationMessage = 'Geen bekende busvoorraad voor dit artikel.';
    } else if (projectedQty < 0) {
      validationStatus = getConsumptionValidationWarning_();
      validationMessage = 'Verbruik overschrijdt huidige busvoorraad.';
    }

    projected.push(Object.assign({}, line, {
      systemBusAantal: currentQty,
      projectedBusAantal: projectedQty,
      validationStatus: validationStatus,
      validationMessage: validationMessage,
      bookingGroupKey: buildConsumptionBookingGroupKey_(line),
    }));

    runningBus[stateKey] = projectedQty;
  });

  return projected;
}

function summarizeConsumptionProcessLines_(lines) {
  var rows = Array.isArray(lines) ? lines : [];

  return {
    lineCount: rows.length,
    totaalAantal: rows.reduce(function (sum, row) {
      return sum + safeNumber(row.aantal, 0);
    }, 0),
    warningCount: rows.filter(function (row) {
      return safeText(row.validationStatus) === getConsumptionValidationWarning_();
    }).length,
    errorCount: rows.filter(function (row) {
      return safeText(row.validationStatus) === getConsumptionValidationError_();
    }).length,
    bookingGroupCount: Object.keys(
      rows.reduce(function (acc, row) {
        acc[safeText(row.bookingGroupKey)] = true;
        return acc;
      }, {})
    ).length,
  };
}

/* ---------------------------------------------------------
   Process batch creation
   --------------------------------------------------------- */

function createConsumptionProcessBatch(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionProcessingAccess_(sessionId);
  var importBatch = resolveConsumptionImportBatchWithLines_(payload.importBatchId);

  var importLines = Array.isArray(importBatch.lines) ? importBatch.lines : [];
  if (!importLines.length) {
    throw new Error('Importbatch bevat geen lijnen.');
  }

  var projectedLines = buildConsumptionProjectionLines_(importLines);
  var summary = summarizeConsumptionProcessLines_(projectedLines);

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var processBatchId = makeConsumptionProcessBatchId_();

  var batchObj = {
    ProcessBatchID: processBatchId,
    ImportBatchID: safeText(importBatch.importBatchId),
    SourceType: 'Import',
    DocumentDatum: safeText(importBatch.documentDatum),
    DocumentDatumIso: safeText(importBatch.documentDatumIso),
    Status: summary.errorCount ? getConsumptionProcessStatusOpen_() : getConsumptionProcessStatusReady_(),
    Actor: safeText(payload.actor || actor.naam || actor.email),
    Opmerking: safeText(payload.opmerking || payload.remark),
    RawLineCount: safeNumber(importBatch.rawLineCount, importLines.length),
    ProcessedLineCount: projectedLines.length,
    WarningCount: summary.warningCount,
    ErrorCount: summary.errorCount,
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    GeboektOp: '',
  };

  appendObjects(getConsumptionProcessBatchTab_(), [batchObj]);

  var lineObjects = projectedLines.map(function (line) {
    return {
      ProcessLineID: makeConsumptionProcessLineId_(),
      ProcessBatchID: processBatchId,
      ImportBatchID: safeText(importBatch.importBatchId),
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
      SystemBusAantal: line.systemBusAantal,
      ProjectedBusAantal: line.projectedBusAantal,
      ValidationStatus: line.validationStatus,
      ValidationMessage: line.validationMessage,
      BookingGroupKey: line.bookingGroupKey,
      Opmerking: line.opmerking,
    };
  });

  if (lineObjects.length) {
    appendObjects(getConsumptionProcessLineTab_(), lineObjects);
  }

  writeAudit({
    actie: 'CREATE_CONSUMPTION_PROCESS_BATCH',
    actor: actor,
    documentType: 'VerbruiksProcess',
    documentId: processBatchId,
    details: {
      importBatchId: importBatch.importBatchId,
      processedLineCount: projectedLines.length,
      warningCount: summary.warningCount,
      errorCount: summary.errorCount,
    },
  });

  return {
    batch: getConsumptionProcessBatchWithLines(processBatchId),
    summary: summary,
  };
}

/* ---------------------------------------------------------
   Recalculate / status update
   --------------------------------------------------------- */

function recalculateConsumptionProcessBatch(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionProcessingAccess_(sessionId);
  var processBatchId = safeText(payload.processBatchId);
  if (!processBatchId) throw new Error('ProcessBatchId ontbreekt.');

  var existing = getConsumptionProcessBatchById(processBatchId);
  if (!existing) throw new Error('Processbatch niet gevonden.');

  return createConsumptionProcessBatch({
    sessionId: sessionId,
    importBatchId: existing.importBatchId,
    actor: safeText(actor.naam || actor.email),
    opmerking: safeText(payload.opmerking || payload.remark || existing.opmerking),
  });
}

function markConsumptionProcessBatchBooked(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionProcessingAccess_(sessionId);
  var processBatchId = safeText(payload.processBatchId);
  if (!processBatchId) throw new Error('ProcessBatchId ontbreekt.');

  var batch = getConsumptionProcessBatchById(processBatchId);
  if (!batch) throw new Error('Processbatch niet gevonden.');

  var table = getAllValues(getConsumptionProcessBatchTab_());
  if (!table.length) throw new Error('Processbatchtab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getConsumptionProcessBatchTab_()).map(function (row) {
    var current = mapConsumptionProcessBatch(row);
    if (safeText(current.processBatchId) !== processBatchId) {
      return row;
    }

    row.Status = getConsumptionProcessStatusBooked_();
    row.GeboektOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getConsumptionProcessBatchTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  writeAudit({
    actie: 'MARK_CONSUMPTION_PROCESS_BATCH_BOOKED',
    actor: actor,
    documentType: 'VerbruiksProcess',
    documentId: processBatchId,
    details: {},
  });

  return getConsumptionProcessBatchWithLines(processBatchId);
}

/* ---------------------------------------------------------
   Booking preparation queries
   --------------------------------------------------------- */

function buildConsumptionBookingGroups(processLines) {
  var grouped = {};

  (processLines || []).forEach(function (line) {
    var key = safeText(line.bookingGroupKey);
    if (!key) return;

    if (!grouped[key]) {
      grouped[key] = {
        bookingGroupKey: key,
        techniekerCode: safeText(line.techniekerCode),
        techniekerNaam: safeText(line.techniekerNaam),
        werkorderNr: safeText(line.werkorderNr),
        documentDatum: safeText(line.documentDatum),
        documentDatumIso: safeText(line.documentDatumIso),
        lines: [],
      };
    }

    grouped[key].lines.push(line);
  });

  return Object.keys(grouped)
    .map(function (key) {
      var group = grouped[key];
      var statuses = group.lines.map(function (line) {
        return safeText(line.validationStatus);
      });

      return Object.assign({}, group, {
        validationStatus:
          statuses.indexOf(getConsumptionValidationError_()) >= 0
            ? getConsumptionValidationError_()
            : statuses.indexOf(getConsumptionValidationWarning_()) >= 0
              ? getConsumptionValidationWarning_()
              : getConsumptionValidationOk_(),
        totaalAantal: group.lines.reduce(function (sum, line) {
          return sum + safeNumber(line.aantal, 0);
        }, 0),
      });
    })
    .sort(function (a, b) {
      return (
        safeText(a.techniekerCode).localeCompare(safeText(b.techniekerCode)) ||
        safeText(a.documentDatumIso).localeCompare(safeText(b.documentDatumIso)) ||
        safeText(a.werkorderNr).localeCompare(safeText(b.werkorderNr))
      );
    });
}

function getConsumptionProcessBookingData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  assertConsumptionProcessingAccess_(sessionId);

  var processBatchId = safeText(payload.processBatchId);
  if (!processBatchId) throw new Error('ProcessBatchId ontbreekt.');

  var batch = getConsumptionProcessBatchWithLines(processBatchId);
  if (!batch) throw new Error('Processbatch niet gevonden.');

  var groups = buildConsumptionBookingGroups(batch.lines || []);

  return {
    batch: batch,
    bookingGroups: groups,
    summary: summarizeConsumptionProcessLines_(batch.lines || []),
  };
}

/* ---------------------------------------------------------
   Queries
   --------------------------------------------------------- */

function getConsumptionProcessData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionProcessingAccess_(sessionId);
  var rows = buildConsumptionProcessBatchesWithLines(
    getAllConsumptionProcessBatches(),
    getAllConsumptionProcessLines()
  );

  return {
    items: rows,
    batches: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getConsumptionProcessStatusOpen_(); }).length,
      klaar: rows.filter(function (x) { return safeText(x.status) === getConsumptionProcessStatusReady_(); }).length,
      geboekt: rows.filter(function (x) { return safeText(x.status) === getConsumptionProcessStatusBooked_(); }).length,
      gesloten: rows.filter(function (x) { return safeText(x.status) === getConsumptionProcessStatusClosed_(); }).length,
      actorRol: safeText(actor.rol),
    }
  };
}
