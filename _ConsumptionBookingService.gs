/* =========================================================
   46_ConsumptionBookingService.gs
   Refactor: consumption booking service
   Doel:
   - boekbare verbruiksgroepen omzetten naar echte mutaties
   - documentlaag voor verbruiksboekingen
   - draft aanmaken vanuit processing batch
   - lijnen bewaren
   - boeken naar movementlaag
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getConsumptionBookingHeaderTab_() {
  return TABS.CONSUMPTION_BOOKINGS || 'VerbruiksBoekingen';
}

function getConsumptionBookingLineTab_() {
  return TABS.CONSUMPTION_BOOKING_LINES || 'VerbruiksBoekingLijnen';
}

function getConsumptionBookingStatusOpen_() {
  if (typeof CONSUMPTION_BOOKING_STATUS !== 'undefined' && CONSUMPTION_BOOKING_STATUS && CONSUMPTION_BOOKING_STATUS.OPEN) {
    return CONSUMPTION_BOOKING_STATUS.OPEN;
  }
  return 'Open';
}

function getConsumptionBookingStatusBooked_() {
  if (typeof CONSUMPTION_BOOKING_STATUS !== 'undefined' && CONSUMPTION_BOOKING_STATUS && CONSUMPTION_BOOKING_STATUS.BOOKED) {
    return CONSUMPTION_BOOKING_STATUS.BOOKED;
  }
  return 'Geboekt';
}

function getConsumptionBookingStatusClosed_() {
  if (typeof CONSUMPTION_BOOKING_STATUS !== 'undefined' && CONSUMPTION_BOOKING_STATUS && CONSUMPTION_BOOKING_STATUS.CLOSED) {
    return CONSUMPTION_BOOKING_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function getConsumptionMovementOut_() {
  if (typeof MOVEMENT_TYPE !== 'undefined' && MOVEMENT_TYPE && MOVEMENT_TYPE.CONSUMPTION_OUT) {
    return MOVEMENT_TYPE.CONSUMPTION_OUT;
  }
  return 'ConsumptionOut';
}

function getConsumptionSinkLocation_() {
  if (typeof LOCATION !== 'undefined' && LOCATION && LOCATION.CONSUMED) {
    return LOCATION.CONSUMED;
  }
  return 'Verbruik';
}

function makeConsumptionBookingId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'CBK-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function makeConsumptionBookingLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'CBL-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function isConsumptionBookingEditable_(status) {
  var value = safeText(status);
  return !value || value === getConsumptionBookingStatusOpen_();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapConsumptionBookingHeader(row) {
  return {
    bookingId: safeText(row.BookingID || row.ConsumptionBookingID || row.ConsumptionBookingId || row.ID),
    processBatchId: safeText(row.ProcessBatchID || row.ProcessBatchId),
    importBatchId: safeText(row.ImportBatchID || row.ImportBatchId),
    bookingGroupKey: safeText(row.BookingGroupKey),
    techniekerCode: safeText(row.TechniekerCode || row.TechnicianCode || row.TechCode),
    techniekerNaam: safeText(row.TechniekerNaam || row.TechnicianName),
    werkorderNr: safeText(row.WerkorderNr || row.WorkOrderNr || row.OrderNr),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    status: safeText(row.Status || getConsumptionBookingStatusOpen_()),
    validationStatus: safeText(row.ValidationStatus || getConsumptionValidationOk_()),
    reason: safeText(row.Reason || row.Reden),
    actor: safeText(row.Actor),
    opmerking: safeText(row.Opmerking),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    geboektOp: safeText(row.GeboektOp),
    geslotenOp: safeText(row.GeslotenOp),
  };
}

function mapConsumptionBookingLine(row) {
  return {
    bookingLineId: safeText(row.BookingLineID || row.ConsumptionBookingLineID || row.ConsumptionBookingLineId || row.ID),
    bookingId: safeText(row.BookingID || row.ConsumptionBookingID || row.ConsumptionBookingId),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode || row.ArtikelNr))),
    eenheid: safeText(row.Eenheid || row.Unit),
    aantal: safeNumber(row.Aantal || row.Quantity, 0),
    systemBusAantal: safeNumber(row.SystemBusAantal || row.SystemQty, 0),
    projectedBusAantal: safeNumber(row.ProjectedBusAantal || row.ProjectedQty, 0),
    validationStatus: safeText(row.ValidationStatus || getConsumptionValidationOk_()),
    validationMessage: safeText(row.ValidationMessage || ''),
    opmerking: safeText(row.Opmerking),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllConsumptionBookings() {
  return readObjectsSafe(getConsumptionBookingHeaderTab_())
    .map(mapConsumptionBookingHeader)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.aangemaaktOpRaw).localeCompare(safeText(a.aangemaaktOpRaw)) ||
        safeText(b.bookingId).localeCompare(safeText(a.bookingId))
      );
    });
}

function getAllConsumptionBookingLines() {
  return readObjectsSafe(getConsumptionBookingLineTab_())
    .map(mapConsumptionBookingLine)
    .sort(function (a, b) {
      return (
        safeText(a.bookingId).localeCompare(safeText(b.bookingId)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function getConsumptionBookingById(bookingId) {
  var id = safeText(bookingId);
  if (!id) return null;

  return getAllConsumptionBookings().find(function (item) {
    return safeText(item.bookingId) === id;
  }) || null;
}

function getConsumptionBookingLinesById(bookingId) {
  var id = safeText(bookingId);
  return getAllConsumptionBookingLines().filter(function (item) {
    return safeText(item.bookingId) === id;
  });
}

function buildConsumptionBookingsWithLines(headers, lines) {
  var lineMap = {};

  (lines || []).forEach(function (line) {
    var id = safeText(line.bookingId);
    if (!lineMap[id]) lineMap[id] = [];
    lineMap[id].push(line);
  });

  return (headers || []).map(function (header) {
    var bookingLines = lineMap[safeText(header.bookingId)] || [];
    return Object.assign({}, header, {
      lines: bookingLines,
      lineCount: bookingLines.length,
      totaalAantal: bookingLines.reduce(function (sum, line) {
        return sum + safeNumber(line.aantal, 0);
      }, 0),
      warningCount: bookingLines.filter(function (line) {
        return safeText(line.validationStatus) === getConsumptionValidationWarning_();
      }).length,
      errorCount: bookingLines.filter(function (line) {
        return safeText(line.validationStatus) === getConsumptionValidationError_();
      }).length,
    });
  });
}

function getConsumptionBookingWithLines(bookingId) {
  var header = getConsumptionBookingById(bookingId);
  if (!header) return null;
  return buildConsumptionBookingsWithLines(
    [header],
    getConsumptionBookingLinesById(bookingId)
  )[0] || null;
}

function hasConsumptionBookingForGroup_(processBatchId, bookingGroupKey) {
  var pid = safeText(processBatchId);
  var key = safeText(bookingGroupKey);

  return getAllConsumptionBookings().some(function (item) {
    return safeText(item.processBatchId) === pid &&
      safeText(item.bookingGroupKey) === key &&
      [getConsumptionBookingStatusOpen_(), getConsumptionBookingStatusBooked_()].indexOf(safeText(item.status)) >= 0;
  });
}

/* ---------------------------------------------------------
   Access policy
   --------------------------------------------------------- */

function assertConsumptionBookingAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.MOBILE_WAREHOUSE],
    'Geen rechten voor verbruiksboekingen.'
  );
  return user;
}

/* ---------------------------------------------------------
   Helpers from processing layer
   --------------------------------------------------------- */

function resolveConsumptionProcessBatchForBooking_(processBatchId) {
  var id = safeText(processBatchId);
  if (!id) throw new Error('ProcessBatchId ontbreekt.');

  if (typeof getConsumptionProcessBatchWithLines !== 'function') {
    throw new Error('Consumption processing service ontbreekt. Werk eerst het processingblok in.');
  }

  var batch = getConsumptionProcessBatchWithLines(id);
  if (!batch) {
    throw new Error('Processbatch niet gevonden.');
  }

  return batch;
}

function resolveConsumptionBookingGroups_(processBatch) {
  if (typeof buildConsumptionBookingGroups === 'function') {
    return buildConsumptionBookingGroups(processBatch.lines || []);
  }

  var grouped = {};
  (processBatch.lines || []).forEach(function (line) {
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

  return Object.keys(grouped).map(function (key) { return grouped[key]; });
}

function getConsumptionBookingGroupByKey_(processBatch, bookingGroupKey) {
  var key = safeText(bookingGroupKey);
  return resolveConsumptionBookingGroups_(processBatch).find(function (item) {
    return safeText(item.bookingGroupKey) === key;
  }) || null;
}

function calculateConsumptionBookingValidationStatus_(lines) {
  var statuses = (lines || []).map(function (line) {
    return safeText(line.validationStatus);
  });

  if (statuses.indexOf(getConsumptionValidationError_()) >= 0) {
    return getConsumptionValidationError_();
  }
  if (statuses.indexOf(getConsumptionValidationWarning_()) >= 0) {
    return getConsumptionValidationWarning_();
  }
  return getConsumptionValidationOk_();
}

/* ---------------------------------------------------------
   Draft creation from processing group
   --------------------------------------------------------- */

function createConsumptionBookingFromProcessGroup(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionBookingAccess_(sessionId);
  var processBatchId = safeText(payload.processBatchId);
  var bookingGroupKey = safeText(payload.bookingGroupKey);

  if (!processBatchId) throw new Error('ProcessBatchId ontbreekt.');
  if (!bookingGroupKey) throw new Error('BookingGroupKey ontbreekt.');

  if (hasConsumptionBookingForGroup_(processBatchId, bookingGroupKey)) {
    throw new Error('Er bestaat al een verbruiksboeking voor deze boekingsgroep.');
  }

  var processBatch = resolveConsumptionProcessBatchForBooking_(processBatchId);
  var group = getConsumptionBookingGroupByKey_(processBatch, bookingGroupKey);
  if (!group) throw new Error('Boekingsgroep niet gevonden.');

  var validationStatus = calculateConsumptionBookingValidationStatus_(group.lines || []);
  if (validationStatus === getConsumptionValidationError_()) {
    throw new Error('Boekingsgroep bevat fouten en kan nog niet geboekt worden.');
  }

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var bookingId = makeConsumptionBookingId_();

  var headerObj = {
    BookingID: bookingId,
    ProcessBatchID: safeText(processBatch.processBatchId),
    ImportBatchID: safeText(processBatch.importBatchId),
    BookingGroupKey: safeText(group.bookingGroupKey),
    TechniekerCode: safeText(group.techniekerCode),
    TechniekerNaam: safeText(group.techniekerNaam),
    WerkorderNr: safeText(group.werkorderNr),
    DocumentDatum: safeText(group.documentDatum),
    DocumentDatumIso: safeText(group.documentDatumIso),
    Status: getConsumptionBookingStatusOpen_(),
    ValidationStatus: validationStatus,
    Reason: safeText(payload.reason || payload.reden || 'Verbruiksboeking'),
    Actor: safeText(payload.actor || actor.naam || actor.email),
    Opmerking: safeText(payload.opmerking || payload.remark),
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    GeboektOp: '',
    GeslotenOp: '',
  };

  appendObjects(getConsumptionBookingHeaderTab_(), [headerObj]);

  var lineObjects = (group.lines || []).map(function (line) {
    return {
      BookingLineID: makeConsumptionBookingLineId_(),
      BookingID: bookingId,
      ArtikelCode: safeText(line.artikelCode),
      ArtikelOmschrijving: safeText(line.artikelOmschrijving),
      TypeMateriaal: safeText(line.typeMateriaal),
      Eenheid: safeText(line.eenheid),
      Aantal: safeNumber(line.aantal, 0),
      SystemBusAantal: safeNumber(line.systemBusAantal, 0),
      ProjectedBusAantal: safeNumber(line.projectedBusAantal, 0),
      ValidationStatus: safeText(line.validationStatus),
      ValidationMessage: safeText(line.validationMessage),
      Opmerking: safeText(line.opmerking),
    };
  });

  if (lineObjects.length) {
    appendObjects(getConsumptionBookingLineTab_(), lineObjects);
  }

  writeAudit({
    actie: 'CREATE_CONSUMPTION_BOOKING',
    actor: actor,
    documentType: 'VerbruiksBoeking',
    documentId: bookingId,
    details: {
      processBatchId: processBatchId,
      bookingGroupKey: bookingGroupKey,
      lineCount: lineObjects.length,
      validationStatus: validationStatus,
    },
  });

  return getConsumptionBookingWithLines(bookingId);
}

/* ---------------------------------------------------------
   Save lines
   --------------------------------------------------------- */

function normalizeConsumptionBookingLines_(lines) {
  return (Array.isArray(lines) ? lines : []).map(function (line) {
    return {
      artikelCode: safeText(line.artikelCode || line.artikelNr),
      artikelOmschrijving: safeText(line.artikelOmschrijving || line.artikel),
      typeMateriaal: safeText(line.typeMateriaal || determineMaterialTypeFromArticle(safeText(line.artikelCode || line.artikelNr))),
      eenheid: safeText(line.eenheid || line.unit || 'Stuk'),
      aantal: safeNumber(line.aantal || line.quantity, 0),
      systemBusAantal: safeNumber(line.systemBusAantal || line.systemQty, 0),
      projectedBusAantal: safeNumber(line.projectedBusAantal || line.projectedQty, 0),
      validationStatus: safeText(line.validationStatus || getConsumptionValidationOk_()),
      validationMessage: safeText(line.validationMessage || ''),
      opmerking: safeText(line.opmerking),
    };
  });
}

function validateConsumptionBookingLines_(lines) {
  if (!lines.length) {
    throw new Error('Geen boekingslijnen ontvangen.');
  }

  lines.forEach(function (line, index) {
    var rowNr = index + 1;
    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.aantal, 0) <= 0) {
      throw new Error('Aantal moet groter zijn dan 0 op lijn ' + rowNr + '.');
    }
    if (safeText(line.validationStatus) === getConsumptionValidationError_()) {
      throw new Error('Boekingslijn ' + rowNr + ' bevat nog een fout.');
    }
  });

  return true;
}

function saveConsumptionBookingLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionBookingAccess_(sessionId);
  var bookingId = safeText(payload.bookingId);
  if (!bookingId) throw new Error('BookingId ontbreekt.');

  var header = getConsumptionBookingById(bookingId);
  if (!header) throw new Error('Verbruiksboeking niet gevonden.');
  if (!isConsumptionBookingEditable_(header.status)) {
    throw new Error('Verbruiksboeking is niet meer bewerkbaar.');
  }

  var lines = normalizeConsumptionBookingLines_(payload.lines);
  validateConsumptionBookingLines_(lines);

  var table = getAllValues(getConsumptionBookingLineTab_());
  var headerRow = table.length ? table[0] : null;
  var currentRows = readObjectsSafe(getConsumptionBookingLineTab_());

  var kept = currentRows.filter(function (row) {
    return safeText(row.BookingID || row.ConsumptionBookingID || row.ConsumptionBookingId) !== bookingId;
  });

  var newRows = lines.map(function (line) {
    return {
      BookingLineID: makeConsumptionBookingLineId_(),
      BookingID: bookingId,
      ArtikelCode: line.artikelCode,
      ArtikelOmschrijving: line.artikelOmschrijving,
      TypeMateriaal: line.typeMateriaal,
      Eenheid: line.eenheid,
      Aantal: line.aantal,
      SystemBusAantal: line.systemBusAantal,
      ProjectedBusAantal: line.projectedBusAantal,
      ValidationStatus: line.validationStatus,
      ValidationMessage: line.validationMessage,
      Opmerking: line.opmerking,
    };
  });

  var finalObjects = kept.concat(newRows);

  if (!headerRow && finalObjects.length) {
    appendObjects(getConsumptionBookingLineTab_(), newRows);
  } else if (headerRow) {
    writeFullTable(
      getConsumptionBookingLineTab_(),
      headerRow,
      finalObjects.map(function (obj) {
        return buildRowFromHeaders(headerRow, obj);
      })
    );
  } else if (newRows.length) {
    appendObjects(getConsumptionBookingLineTab_(), newRows);
  }

  var headerTable = getAllValues(getConsumptionBookingHeaderTab_());
  if (headerTable.length) {
    var bookingHeaderRow = headerTable[0];
    var newValidationStatus = calculateConsumptionBookingValidationStatus_(lines);

    var headerRows = readObjectsSafe(getConsumptionBookingHeaderTab_()).map(function (row) {
      var current = mapConsumptionBookingHeader(row);
      if (safeText(current.bookingId) !== bookingId) {
        return row;
      }

      row.ValidationStatus = newValidationStatus;
      return row;
    });

    writeFullTable(
      getConsumptionBookingHeaderTab_(),
      bookingHeaderRow,
      headerRows.map(function (row) {
        return buildRowFromHeaders(bookingHeaderRow, row);
      })
    );
  }

  writeAudit({
    actie: 'SAVE_CONSUMPTION_BOOKING_LINES',
    actor: actor,
    documentType: 'VerbruiksBoeking',
    documentId: bookingId,
    details: {
      lineCount: lines.length,
    },
  });

  return getConsumptionBookingWithLines(bookingId);
}

/* ---------------------------------------------------------
   Movement build / booking
   --------------------------------------------------------- */

function buildConsumptionSourceLocation_(techniekerCode) {
  if (typeof getBusLocationCode === 'function') {
    return getBusLocationCode(techniekerCode);
  }
  return 'Bus:' + safeText(techniekerCode);
}

function buildConsumptionBookingMovements_(booking) {
  var header = booking || {};
  var lines = header.lines || [];
  var sourceLocation = buildConsumptionSourceLocation_(header.techniekerCode);
  var sinkLocation = getConsumptionSinkLocation_();

  return lines.map(function (line) {
    return buildMovementObject({
      movementType: getConsumptionMovementOut_(),
      bronType: 'VerbruiksBoeking',
      bronId: header.bookingId,
      datumBoeking: header.documentDatum,
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      typeMateriaal: line.typeMateriaal,
      eenheid: line.eenheid,
      aantalIn: 0,
      aantalUit: safeNumber(line.aantal, 0),
      nettoAantal: -safeNumber(line.aantal, 0),
      locatieVan: sourceLocation,
      locatieNaar: sinkLocation,
      reden: header.reason || 'Verbruik',
      opmerking: line.opmerking || header.opmerking,
      actor: header.actor,
    });
  });
}

function validateConsumptionBookingStock_(booking) {
  var lines = booking.lines || [];

  lines.forEach(function (line) {
    var available = safeNumber(line.systemBusAantal, 0);
    var requested = safeNumber(line.aantal, 0);

    if (requested > available) {
      throw new Error(
        'Onvoldoende busvoorraad voor artikel ' + safeText(line.artikelCode) +
        '. Beschikbaar: ' + available + ', gevraagd: ' + requested + '.'
      );
    }
  });

  return true;
}

function bookConsumptionBooking(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionBookingAccess_(sessionId);
  var bookingId = safeText(payload.bookingId);
  if (!bookingId) throw new Error('BookingId ontbreekt.');

  if (typeof replaceSourceMovements !== 'function') {
    throw new Error('Movement service ontbreekt. Werk eerst het movementblok in.');
  }

  var booking = getConsumptionBookingWithLines(bookingId);
  if (!booking) throw new Error('Verbruiksboeking niet gevonden.');
  if (safeText(booking.status) !== getConsumptionBookingStatusOpen_()) {
    throw new Error('Verbruiksboeking kan niet geboekt worden vanuit status "' + safeText(booking.status) + '".');
  }
  if (safeText(booking.validationStatus) === getConsumptionValidationError_()) {
    throw new Error('Verbruiksboeking bevat nog fouten.');
  }

  validateConsumptionBookingLines_(booking.lines || []);
  validateConsumptionBookingStock_(booking);

  var movements = buildConsumptionBookingMovements_(booking);
  replaceSourceMovements('VerbruiksBoeking', booking.bookingId, movements);

  var table = getAllValues(getConsumptionBookingHeaderTab_());
  if (!table.length) throw new Error('Verbruiksboekingtab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getConsumptionBookingHeaderTab_()).map(function (row) {
    var current = mapConsumptionBookingHeader(row);
    if (safeText(current.bookingId) !== bookingId) {
      return row;
    }

    row.Status = getConsumptionBookingStatusBooked_();
    row.GeboektOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getConsumptionBookingHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  if (typeof markConsumptionProcessBatchBooked === 'function') {
    markConsumptionProcessBatchBooked({
      sessionId: sessionId,
      processBatchId: booking.processBatchId,
    });
  }

  writeAudit({
    actie: 'BOOK_CONSUMPTION_BOOKING',
    actor: actor,
    documentType: 'VerbruiksBoeking',
    documentId: bookingId,
    details: {
      processBatchId: booking.processBatchId,
      movementCount: movements.length,
      techniekerCode: booking.techniekerCode,
      totaalAantal: (booking.lines || []).reduce(function (sum, line) {
        return sum + safeNumber(line.aantal, 0);
      }, 0),
    },
  });

  return getConsumptionBookingWithLines(bookingId);
}

/* ---------------------------------------------------------
   Queries
   --------------------------------------------------------- */

function getConsumptionBookingData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionBookingAccess_(sessionId);
  var rows = buildConsumptionBookingsWithLines(
    getAllConsumptionBookings(),
    getAllConsumptionBookingLines()
  );

  return {
    items: rows,
    bookings: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getConsumptionBookingStatusOpen_(); }).length,
      geboekt: rows.filter(function (x) { return safeText(x.status) === getConsumptionBookingStatusBooked_(); }).length,
      gesloten: rows.filter(function (x) { return safeText(x.status) === getConsumptionBookingStatusClosed_(); }).length,
      warnings: rows.reduce(function (sum, item) {
        return sum + safeNumber(item.warningCount, 0);
      }, 0),
      errors: rows.reduce(function (sum, item) {
        return sum + safeNumber(item.errorCount, 0);
      }, 0),
      actorRol: safeText(actor.rol),
    }
  };
}
