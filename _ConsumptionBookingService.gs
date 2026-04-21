/* =========================================================
   46_ConsumptionBookingService.gs — booking verbruik
   Verbruiken / VerbruikLijnen -> MagazijnMutaties
   ========================================================= */

/* ---------------------------------------------------------
   Basis readers
   --------------------------------------------------------- */

function getAllConsumptions() {
  return readObjectsSafe(TABS.CONSUMPTIONS)
    .map(mapConsumption)
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.verbruikId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.verbruikId)}`
      )
    );
}

function getAllConsumptionLines() {
  return readObjectsSafe(TABS.CONSUMPTION_LINES)
    .map(mapConsumptionLine)
    .filter(line => line.actief);
}

function getConsumptionById(verbruikId) {
  const id = safeText(verbruikId);
  if (!id) return null;
  return getAllConsumptions().find(item => item.verbruikId === id) || null;
}

function getConsumptionLinesById(verbruikId) {
  const id = safeText(verbruikId);
  if (!id) return [];
  return getAllConsumptionLines().filter(line => line.verbruikId === id);
}

function getConsumptionsByRunId(runId) {
  const id = safeText(runId);
  if (!id) return [];

  return getAllConsumptions().filter(item => {
    if (safeText(item.bronRunId) === id) return true;
    return safeText(item.verbruikId).startsWith(`VIMP-${id}-`);
  });
}

/* ---------------------------------------------------------
   Historiek
   --------------------------------------------------------- */

function buildConsumptionHistory() {
  const consumptions = getAllConsumptions();
  const lines = getAllConsumptionLines();

  const linesByConsumption = {};
  lines.forEach(line => {
    const id = safeText(line.verbruikId);
    if (!id) return;
    if (!linesByConsumption[id]) linesByConsumption[id] = [];
    linesByConsumption[id].push(line);
  });

  return consumptions
    .map(consumption => ({
      ...consumption,
      bronRunId: safeText(consumption.bronRunId || ''),
      bronType: safeText(consumption.bronType || ''),
      lines: linesByConsumption[consumption.verbruikId] || [],
      lineCount: (linesByConsumption[consumption.verbruikId] || []).length
    }))
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.verbruikId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.verbruikId)}`
      )
    );
}

/* ---------------------------------------------------------
   Helpers
   --------------------------------------------------------- */

function updateConsumptionHeader(verbruikId, updates) {
  const id = safeText(verbruikId);

  const sheet = getSheetOrThrow(TABS.CONSUMPTIONS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab Verbruiken is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['VerbruikID']]) !== id) continue;

    Object.keys(updates || {}).forEach(field => {
      if (col[field] !== undefined) {
        values[i][col[field]] = updates[field];
      }
    });

    updated = true;
    break;
  }

  if (!updated) throw new Error('Verbruik niet gevonden.');

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return { success: true };
}

function getConsumptionBusStockMapForTechnician(techniekerCode) {
  const map = {};
  buildBusStockRows()
    .filter(item => normalizeRef(item.techniekerCode) === normalizeRef(techniekerCode))
    .forEach(item => {
      map[item.artikelCode] = item;
    });
  return map;
}

function replaceConsumptionBookingMovements(verbruikId, movementPayloads) {
  const id = safeText(verbruikId);

  const sheet = getSheetOrThrow(TABS.WAREHOUSE_MOVEMENTS);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.WAREHOUSE_MOVEMENTS);
  const col = getColMap(headers);
  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => {
    const bronId = safeText(row[col['BronID']]);
    const typeMutatie = safeText(row[col['TypeMutatie']]);
    if (bronId !== id) return true;
    if (typeMutatie !== MOVEMENT_TYPE.CONSUMPTION && typeMutatie !== 'Verbruik') return true;
    return false;
  });

  const newRows = (movementPayloads || []).map(payload =>
    buildRowFromHeaders(headers, {
      MutatieID: makeMovementId(),
      DatumBoeking: nowStamp(),
      DatumDocument: safeText(payload.datumDocument),
      TypeMutatie: safeText(payload.typeMutatie),
      BronID: id,
      TypeMateriaal: safeText(payload.typeMateriaal),
      ArtikelCode: safeText(payload.artikelCode),
      ArtikelOmschrijving: safeText(payload.artikelOmschrijving),
      Eenheid: safeText(payload.eenheid),
      AantalIn: safeNumber(payload.aantalIn, 0),
      AantalUit: safeNumber(payload.aantalUit, 0),
      NettoAantal: safeNumber(payload.nettoAantal, 0),
      LocatieVan: safeText(payload.locatieVan),
      LocatieNaar: safeText(payload.locatieNaar),
      Reden: safeText(payload.reden),
      Opmerking: safeText(payload.opmerking),
      GoedgekeurdDoor: safeText(payload.goedgekeurdDoor),
      GoedgekeurdOp: safeText(payload.goedgekeurdOp)
    })
  );

  writeFullTable(TABS.WAREHOUSE_MOVEMENTS, headers, keptRows.concat(newRows));

  return {
    success: true,
    lines: newRows.length
  };
}

function updateRawBookStatusByConsumptionId(verbruikId, bookStatus, bookError) {
  const id = safeText(verbruikId);
  if (!id) return { success: true, updated: 0 };

  const sheet = getSheetOrThrow(TABS.CONSUMPTION_IMPORT_RAW);
  const values = sheet.getDataRange().getValues();
  if (!values.length) return { success: true, updated: 0 };

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = 0;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['VerbruikID']]) !== id) continue;

    if (col['BoekStatus'] !== undefined) values[i][col['BoekStatus']] = safeText(bookStatus);
    if (col['BoekFout'] !== undefined) values[i][col['BoekFout']] = safeText(bookError);
    if (col['VerwerktOp'] !== undefined) values[i][col['VerwerktOp']] = nowStamp();

    updated++;
  }

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return {
    success: true,
    updated: updated
  };
}

function buildConsumptionMovementPayloads(consumption, lines, actor) {
  const busLocation = getBusLocationCode(consumption.techniekerCode);
  const targetLocation = safeText(consumption.werfRef) || LOCATION.SITE;
  const documentDatum = safeText(consumption.documentDatumIso || consumption.documentDatum);

  return (lines || [])
    .filter(line => safeNumber(line.aantal, 0) > 0)
    .map(line => ({
      datumDocument: documentDatum,
      typeMutatie: MOVEMENT_TYPE.CONSUMPTION,
      typeMateriaal: determineMaterialTypeFromArticle(line.artikelCode),
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      eenheid: line.eenheid,
      aantalIn: 0,
      aantalUit: safeNumber(line.aantal, 0),
      nettoAantal: -safeNumber(line.aantal, 0),
      locatieVan: busLocation,
      locatieNaar: targetLocation,
      reden: safeText(consumption.reden) || 'Import',
      opmerking: safeText(consumption.opmerking),
      goedgekeurdDoor: safeText(actor),
      goedgekeurdOp: nowStamp()
    }));
}

/* ---------------------------------------------------------
   Validatie voor boeking
   --------------------------------------------------------- */

function validateConsumptionBeforeBooking(consumption, lines) {
  if (!consumption) throw new Error('Verbruik niet gevonden.');
  if (!(lines || []).length) throw new Error('Dit verbruik bevat nog geen actieve lijnen.');

  if (safeText(consumption.status) === CONSUMPTION_STATUS.BOOKED) {
    throw new Error('Dit verbruik is al geboekt.');
  }

  const busStockMap = getConsumptionBusStockMapForTechnician(consumption.techniekerCode);

  (lines || []).forEach(line => {
    const code = safeText(line.artikelCode);
    const aantal = safeNumber(line.aantal, 0);
    const currentBus = busStockMap[code] ? safeNumber(busStockMap[code].voorraadBus, 0) : 0;

    if (aantal > currentBus) {
      throw new Error(`Onvoldoende voorraad in bus voor artikel ${code}. Beschikbaar: ${currentBus}`);
    }
  });

  return { success: true };
}

/* ---------------------------------------------------------
   1 document boeken
   --------------------------------------------------------- */

function bookSingleConsumptionDocument(verbruikId, actor) {
  const id = safeText(verbruikId);
  const consumption = getConsumptionById(id);
  const lines = getConsumptionLinesById(id);

  try {
    validateConsumptionBeforeBooking(consumption, lines);

    const movementPayloads = buildConsumptionMovementPayloads(consumption, lines, actor);
    replaceConsumptionBookingMovements(id, movementPayloads);

    updateConsumptionHeader(id, {
      Status: CONSUMPTION_STATUS.BOOKED,
      GeboektDoor: safeText(actor),
      GeboektOp: nowStamp()
    });

    updateRawBookStatusByConsumptionId(id, IMPORT_BOOK_STATUS.BOOKED, '');
    rebuildCentralWarehouseOverview();

    pushManagerNotification(
      'VerbruikGeboekt',
      'Verbruik geboekt',
      `Verbruik ${id} voor ${consumption.techniekerNaam || consumption.techniekerCode} is geboekt.`,
      'Verbruik',
      id
    );

    writeAudit(
      'Verbruik geboekt',
      'Systeem',
      safeText(actor),
      'Verbruik',
      id,
      {
        techniekerCode: consumption.techniekerCode,
        techniekerNaam: consumption.techniekerNaam,
        werfRef: consumption.werfRef,
        lijnen: movementPayloads.length
      }
    );

    return {
      success: true,
      verbruikId: id,
      booked: true,
      lines: movementPayloads.length
    };

  } catch (err) {
    updateRawBookStatusByConsumptionId(id, IMPORT_BOOK_STATUS.BOOK_ERROR, safeText(err && err.message ? err.message : err));

    pushManagerNotification(
      'VerbruikBoekfout',
      'Verbruik kon niet geboekt worden',
      `Verbruik ${id} kon niet geboekt worden: ${safeText(err && err.message ? err.message : err)}`,
      'Verbruik',
      id
    );

    writeAudit(
      'Verbruik boekfout',
      'Systeem',
      safeText(actor),
      'Verbruik',
      id,
      {
        error: safeText(err && err.message ? err.message : err)
      }
    );

    return {
      success: false,
      verbruikId: id,
      booked: false,
      error: safeText(err && err.message ? err.message : err)
    };
  }
}

/* ---------------------------------------------------------
   Public booking functies
   --------------------------------------------------------- */

function bookConsumption(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om verbruik te boeken.');
  }

  const verbruikId = safeText(payload.verbruikId);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');

  if (!verbruikId) throw new Error('VerbruikID ontbreekt.');

  const result = bookSingleConsumptionDocument(verbruikId, actor);
  if (!result.success) {
    throw new Error(result.error || 'Verbruik kon niet geboekt worden.');
  }

  return {
    success: true,
    message: 'Verbruik geboekt en busvoorraad verminderd.'
  };
}

function bookConsumptionImportRun(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om importverbruik te boeken.');
  }

  const runId = safeText(payload.runId);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');

  if (!runId) throw new Error('RunID ontbreekt.');

  const consumptions = getConsumptionsByRunId(runId)
    .filter(item => safeText(item.status) !== CONSUMPTION_STATUS.BOOKED);

  if (!consumptions.length) {
    return {
      success: true,
      runId: runId,
      bookedCount: 0,
      errorCount: 0,
      message: 'Geen open verbruiken gevonden voor deze run.'
    };
  }

  const results = consumptions.map(item => bookSingleConsumptionDocument(item.verbruikId, actor));
  const bookedCount = results.filter(x => x.success).length;
  const errorCount = results.filter(x => !x.success).length;

  writeConsumptionImportLog(runId, 'INFO', 'Boeking importrun afgerond', {
    bookedCount: bookedCount,
    errorCount: errorCount
  });

  return {
    success: true,
    runId: runId,
    bookedCount: bookedCount,
    errorCount: errorCount,
    results: results,
    message: `Importrun verwerkt: ${bookedCount} geboekt, ${errorCount} fout.`
  };
}

function bookLatestProcessedConsumptionImportRun(payload) {
  const latestRunId = getLatestConsumptionImportRunId();
  if (!latestRunId) {
    return {
      success: true,
      bookedCount: 0,
      errorCount: 0,
      message: 'Geen import run gevonden.'
    };
  }

  return bookConsumptionImportRun({
    sessionId: getPayloadSessionId(payload || {}),
    actor: payload && payload.actor ? payload.actor : '',
    runId: latestRunId
  });
}

/* ---------------------------------------------------------
   Optionele automatisering
   --------------------------------------------------------- */

function autoBookLatestConsumptionImportRunIfEnabled() {
  const autoBook = safeText(getConsumptionImportConfigValue('AutoBoeken', 'Nee')).toLowerCase();
  if (!['ja', 'true', '1', 'yes'].includes(autoBook)) {
    return {
      success: true,
      skipped: true,
      message: 'AutoBoeken staat uit.'
    };
  }

  return bookLatestProcessedConsumptionImportRun({
    actor: 'Automatisch'
  });
}