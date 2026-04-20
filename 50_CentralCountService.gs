/* =========================================================
   50_CentralCountService.gs — centrale stocktelling
   ========================================================= */

/* ---------------------------------------------------------
   ID
   --------------------------------------------------------- */

function makeCentralCountId() {
  return makeStampedId('CCT');
}

/* ---------------------------------------------------------
   Mappers
   --------------------------------------------------------- */

function mapCentralCount(row) {
  return {
    tellingId: safeText(row.TellingID),

    documentDatumIso: toIsoDate(row.DocumentDatum),
    documentDatum: toDisplayDate(row.DocumentDatum),

    status: safeText(row.Status),
    scopeType: safeText(row.ScopeType) || 'Volledig',
    scopeArtikelCodes: safeText(row.ScopeArtikelCodes),
    reden: safeText(row.Reden),

    aangemaaktDoor: safeText(row.AangemaaktDoor),
    aangemaaktOp: toDisplayDateTime(row.AangemaaktOp),

    ingediendDoor: safeText(row.IngediendDoor),
    ingediendOp: toDisplayDateTime(row.IngediendOp),

    goedgekeurdDoor: safeText(row.GoedgekeurdDoor),
    goedgekeurdOp: toDisplayDateTime(row.GoedgekeurdOp),

    managerOpmerking: safeText(row.ManagerOpmerking)
  };
}

function mapCentralCountLine(row) {
  return {
    tellingId: safeText(row.TellingID),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    systeemAantal: safeNumber(row.SysteemAantal, 0),
    geteldAantal: row.GeteldAantal === '' || row.GeteldAantal === null || row.GeteldAantal === undefined
      ? ''
      : safeNumber(row.GeteldAantal, 0),
    deltaAantal: row.DeltaAantal === '' || row.DeltaAantal === null || row.DeltaAantal === undefined
      ? ''
      : safeNumber(row.DeltaAantal, 0),
    actief: safeText(row.Actief) !== 'Nee'
  };
}

/* ---------------------------------------------------------
   Readers
   --------------------------------------------------------- */

function getAllCentralCounts() {
  return readObjectsSafe(TABS.CENTRAL_COUNTS)
    .map(mapCentralCount)
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.tellingId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.tellingId)}`
      )
    );
}

function getAllCentralCountLines() {
  return readObjectsSafe(TABS.CENTRAL_COUNT_LINES)
    .map(mapCentralCountLine)
    .filter(line => line.actief);
}

function getCentralCountById(tellingId) {
  const id = safeText(tellingId);
  if (!id) return null;
  return getAllCentralCounts().find(item => item.tellingId === id) || null;
}

function getCentralCountLinesById(tellingId) {
  const id = safeText(tellingId);
  if (!id) return [];
  return getAllCentralCountLines().filter(line => line.tellingId === id);
}

function buildCentralCountsWithLines(counts, lines) {
  const linesByCount = {};

  (lines || []).forEach(line => {
    const id = safeText(line.tellingId);
    if (!id) return;
    if (!linesByCount[id]) linesByCount[id] = [];
    linesByCount[id].push(line);
  });

  return (counts || []).map(count => ({
    ...count,
    lines: linesByCount[count.tellingId] || [],
    lineCount: (linesByCount[count.tellingId] || []).length,
    deltaLineCount: (linesByCount[count.tellingId] || []).filter(line => safeNumber(line.deltaAantal, 0) !== 0).length
  }));
}

/* ---------------------------------------------------------
   Data
   --------------------------------------------------------- */

function getCentralCountData(sessionId) {
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om centrale tellingen te bekijken.');
  }

  return {
    counts: buildCentralCountsWithLines(getAllCentralCounts(), getAllCentralCountLines()),
    centralWarehouse: getCentralWarehouseOverview(),
    stockScoring: typeof getStockScoringSnapshot === 'function'
      ? getStockScoringSnapshot(sessionId)
      : null
  };
}

/* ---------------------------------------------------------
   Helpers
   --------------------------------------------------------- */

function buildCentralCountLineObject(tellingId, line) {
  return {
    TellingID: safeText(tellingId),
    ArtikelCode: safeText(line.artikelCode),
    ArtikelOmschrijving: safeText(line.artikelOmschrijving),
    Eenheid: safeText(line.eenheid),
    SysteemAantal: safeNumber(line.systeemAantal, 0),
    GeteldAantal: line.geteldAantal === '' || line.geteldAantal === null || line.geteldAantal === undefined
      ? ''
      : safeNumber(line.geteldAantal, 0),
    DeltaAantal: line.deltaAantal === '' || line.deltaAantal === null || line.deltaAantal === undefined
      ? ''
      : safeNumber(line.deltaAantal, 0),
    Actief: 'Ja'
  };
}

function replaceCentralCountLines(tellingId, lineObjects) {
  const id = safeText(tellingId);

  const sheet = getSheetOrThrow(TABS.CENTRAL_COUNT_LINES);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.CENTRAL_COUNT_LINES);
  const col = getColMap(headers);
  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => safeText(row[col['TellingID']]) !== id);
  const newRows = (lineObjects || []).map(obj => buildRowFromHeaders(headers, obj));

  writeFullTable(TABS.CENTRAL_COUNT_LINES, headers, keptRows.concat(newRows));

  return {
    success: true,
    lines: newRows.length
  };
}

function updateCentralCountHeader(tellingId, updates) {
  const id = safeText(tellingId);

  const sheet = getSheetOrThrow(TABS.CENTRAL_COUNTS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab CentralStockTellingen is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['TellingID']]) !== id) continue;

    Object.keys(updates || {}).forEach(field => {
      if (col[field] !== undefined) {
        values[i][col[field]] = updates[field];
      }
    });

    updated = true;
    break;
  }

  if (!updated) throw new Error('Centrale telling niet gevonden.');

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return { success: true };
}

function buildCentralCountLinesFromStock(scopeArtikelCodes) {
  const centralRows = buildCentralWarehouseRows();

  let filtered = centralRows;

  if (scopeArtikelCodes && scopeArtikelCodes.length) {
    const set = {};
    scopeArtikelCodes.forEach(code => {
      set[safeText(code)] = true;
    });

    filtered = centralRows.filter(item => set[safeText(item.artikelCode)]);
  }

  return filtered
    .map(item => ({
      artikelCode: safeText(item.artikelCode),
      artikelOmschrijving: safeText(item.artikelOmschrijving),
      eenheid: safeText(item.eenheid),
      systeemAantal: safeNumber(item.voorraadCentraal, 0),
      geteldAantal: '',
      deltaAantal: '',
      actief: true
    }))
    .sort((a, b) =>
      safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
    );
}

function validateCentralCountLines(lines) {
  if (!(lines || []).length) throw new Error('Geen tellinglijnen ontvangen.');

  (lines || []).forEach(line => {
    const code = safeText(line.artikelCode);
    if (!code) throw new Error('ArtikelCode ontbreekt op een tellinglijn.');

    const systeemAantal = safeNumber(line.systeemAantal, 0);
    const geteldAantal = line.geteldAantal === '' || line.geteldAantal === null || line.geteldAantal === undefined
      ? ''
      : safeNumber(line.geteldAantal, 0);

    if (systeemAantal < 0) throw new Error(`Systeemaantal mag niet negatief zijn voor artikel ${code}.`);
    if (geteldAantal !== '' && geteldAantal < 0) throw new Error(`Geteld aantal mag niet negatief zijn voor artikel ${code}.`);
  });

  return { success: true };
}

/* ---------------------------------------------------------
   Create
   --------------------------------------------------------- */

function createCentralCount(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om centrale telling aan te maken.');
  }

  const documentDatum = safeText(payload.documentDatum);
  const reden = safeText(payload.reden);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');
  const requestedArticles = Array.isArray(payload.requestedArticles) ? payload.requestedArticles : [];

  if (!documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!reden) throw new Error('Reden is verplicht.');

  const uniqueCodes = [...new Set(
    requestedArticles.map(code => safeText(code)).filter(Boolean)
  )];

  const scopeType = uniqueCodes.length ? 'Gericht' : 'Volledig';
  const lines = buildCentralCountLinesFromStock(uniqueCodes);

  if (!lines.length) {
    throw new Error('Geen artikels gevonden voor centrale telling.');
  }

  const tellingId = makeCentralCountId();

  appendObjects(TABS.CENTRAL_COUNTS, [{
    TellingID: tellingId,
    DocumentDatum: documentDatum,
    Status: CENTRAL_COUNT_STATUS.OPEN,
    ScopeType: scopeType,
    ScopeArtikelCodes: uniqueCodes.join(', '),
    Reden: reden,
    AangemaaktDoor: actor,
    AangemaaktOp: nowStamp(),
    IngediendDoor: '',
    IngediendOp: '',
    GoedgekeurdDoor: '',
    GoedgekeurdOp: '',
    ManagerOpmerking: ''
  }]);

  replaceCentralCountLines(
    tellingId,
    lines.map(line => buildCentralCountLineObject(tellingId, line))
  );

  writeAudit(
    scopeType === 'Gericht' ? 'Gerichte centrale telling aangemaakt' : 'Centrale stocktelling aangemaakt',
    user.rol,
    actor,
    'CentraleStocktelling',
    tellingId,
    {
      scopeType: scopeType,
      scopeArtikelCodes: uniqueCodes.join(', '),
      reden: reden,
      lijnen: lines.length
    }
  );

  return {
    success: true,
    tellingId: tellingId,
    message: 'Centrale telling aangemaakt.'
  };
}

function createCentralCountFromRiskAlerts(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om gerichte centrale telling te maken.');
  }

  const alerts = typeof buildArticleRiskScores === 'function'
    ? buildArticleRiskScores().filter(item => item.shouldTriggerCentralCount)
    : [];

  const requestedArticles = alerts.map(item => item.artikelCode);

  if (!requestedArticles.length) {
    throw new Error('Er zijn momenteel geen artikels met centrale teltrigger.');
  }

  return createCentralCount({
    sessionId: sessionId,
    documentDatum: safeText(payload.documentDatum),
    reden: safeText(payload.reden || 'Gerichte centrale telling op basis van risicoscore'),
    actor: safeText(payload.actor || user.naam || user.email || 'Magazijn'),
    requestedArticles: requestedArticles
  });
}

/* ---------------------------------------------------------
   Save lines
   --------------------------------------------------------- */

function saveCentralCountLines(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om centrale telling op te slaan.');
  }

  const tellingId = safeText(payload.tellingId);
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!tellingId) throw new Error('TellingID ontbreekt.');
  if (!lines.length) throw new Error('Geen tellinglijnen ontvangen.');

  const count = getCentralCountById(tellingId);
  if (!count) throw new Error('Centrale telling niet gevonden.');

  if ([CENTRAL_COUNT_STATUS.SUBMITTED, CENTRAL_COUNT_STATUS.APPROVED].includes(count.status)) {
    throw new Error('Deze centrale telling kan niet meer aangepast worden.');
  }

  validateCentralCountLines(lines);

  const cleaned = lines.map(line => {
    const geteldAantal = line.geteldAantal === '' || line.geteldAantal === null || line.geteldAantal === undefined
      ? ''
      : safeNumber(line.geteldAantal, 0);

    const systeemAantal = safeNumber(line.systeemAantal, 0);

    return buildCentralCountLineObject(tellingId, {
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      eenheid: line.eenheid,
      systeemAantal: systeemAantal,
      geteldAantal: geteldAantal,
      deltaAantal: geteldAantal === '' ? '' : (geteldAantal - systeemAantal)
    });
  });

  replaceCentralCountLines(tellingId, cleaned);

  writeAudit(
    'Centrale telling opgeslagen',
    user.rol,
    user.naam || user.email,
    'CentraleStocktelling',
    tellingId,
    {
      lijnen: cleaned.length
    }
  );

  return {
    success: true,
    lines: cleaned.length,
    message: 'Centrale telling opgeslagen.'
  };
}

/* ---------------------------------------------------------
   Submit
   --------------------------------------------------------- */

function submitCentralCount(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om centrale telling in te dienen.');
  }

  const tellingId = safeText(payload.tellingId);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');

  if (!tellingId) throw new Error('TellingID ontbreekt.');

  const count = getCentralCountById(tellingId);
  if (!count) throw new Error('Centrale telling niet gevonden.');

  if (count.status !== CENTRAL_COUNT_STATUS.OPEN) {
    throw new Error('Deze centrale telling kan niet meer ingediend worden.');
  }

  const lines = getCentralCountLinesById(tellingId);
  if (!lines.length) throw new Error('Deze telling bevat nog geen actieve lijnen.');

  const incomplete = lines.some(line => line.geteldAantal === '');
  if (incomplete) throw new Error('Vul eerst alle getelde aantallen in.');

  updateCentralCountHeader(tellingId, {
    Status: CENTRAL_COUNT_STATUS.SUBMITTED,
    IngediendDoor: actor,
    IngediendOp: nowStamp()
  });

  pushManagerNotification(
    'CentraleTellingGoedTeKeuren',
    'Centrale telling ingediend',
    `Centrale telling ${tellingId} is ingediend en wacht op goedkeuring.`,
    'CentraleStocktelling',
    tellingId
  );

  writeAudit(
    'Centrale telling ingediend',
    user.rol,
    actor,
    'CentraleStocktelling',
    tellingId,
    {
      status: CENTRAL_COUNT_STATUS.SUBMITTED
    }
  );

  return {
    success: true,
    message: 'Centrale telling ingediend.'
  };
}

/* ---------------------------------------------------------
   Booking payloads
   --------------------------------------------------------- */

function buildCentralCountMovementPayloads(count, lines, actor, note) {
  const documentDatum = safeText(count.documentDatumIso || count.documentDatum);

  return (lines || [])
    .filter(line => safeNumber(line.deltaAantal, 0) !== 0)
    .map(line => {
      const delta = safeNumber(line.deltaAantal, 0);

      if (delta > 0) {
        return {
          datumDocument: documentDatum,
          typeMutatie: MOVEMENT_TYPE.CENTRAL_COUNT_IN,
          typeMateriaal: determineMaterialTypeFromArticle(line.artikelCode),
          artikelCode: line.artikelCode,
          artikelOmschrijving: line.artikelOmschrijving,
          eenheid: line.eenheid,
          aantalIn: delta,
          aantalUit: 0,
          nettoAantal: delta,
          locatieVan: '',
          locatieNaar: LOCATION.CENTRAL,
          reden: 'Centrale stocktelling',
          opmerking: note || '',
          goedgekeurdDoor: actor,
          goedgekeurdOp: nowStamp()
        };
      }

      return {
        datumDocument: documentDatum,
        typeMutatie: MOVEMENT_TYPE.CENTRAL_COUNT_OUT,
        typeMateriaal: determineMaterialTypeFromArticle(line.artikelCode),
        artikelCode: line.artikelCode,
        artikelOmschrijving: line.artikelOmschrijving,
        eenheid: line.eenheid,
        aantalIn: 0,
        aantalUit: Math.abs(delta),
        nettoAantal: delta,
        locatieVan: LOCATION.CENTRAL,
        locatieNaar: '',
        reden: 'Centrale stocktelling',
        opmerking: note || '',
        goedgekeurdDoor: actor,
        goedgekeurdOp: nowStamp()
      };
    });
}

function replaceCentralCountBookingMovements(tellingId, movementPayloads) {
  const id = safeText(tellingId);

  const sheet = getSheetOrThrow(TABS.WAREHOUSE_MOVEMENTS);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.WAREHOUSE_MOVEMENTS);
  const col = getColMap(headers);
  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => {
    if (safeText(row[col['BronID']]) !== id) return true;

    const typeMutatie = safeText(row[col['TypeMutatie']]);
    return ![
      MOVEMENT_TYPE.CENTRAL_COUNT_IN,
      MOVEMENT_TYPE.CENTRAL_COUNT_OUT,
      'CentralCountIn',
      'CentralCountOut'
    ].includes(typeMutatie);
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

/* ---------------------------------------------------------
   Approve
   --------------------------------------------------------- */

function approveCentralCount(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertManagerAccess(sessionId);

  const tellingId = safeText(payload.tellingId);
  const actor = safeText(payload.actor || user.naam || 'Manager');
  const note = safeText(payload.note);

  if (!tellingId) throw new Error('TellingID ontbreekt.');

  const count = getCentralCountById(tellingId);
  if (!count) throw new Error('Centrale telling niet gevonden.');

  if (count.status !== CENTRAL_COUNT_STATUS.SUBMITTED) {
    throw new Error('Enkel ingediende centrale tellingen kunnen goedgekeurd worden.');
  }

  const lines = getCentralCountLinesById(tellingId);
  if (!lines.length) throw new Error('Geen tellinglijnen gevonden.');

  const incomplete = lines.some(line => line.geteldAantal === '');
  if (incomplete) throw new Error('Niet alle tellinglijnen bevatten een geteld aantal.');

  const movements = buildCentralCountMovementPayloads(count, lines, actor, note);
  replaceCentralCountBookingMovements(tellingId, movements);

  updateCentralCountHeader(tellingId, {
    Status: CENTRAL_COUNT_STATUS.APPROVED,
    GoedgekeurdDoor: actor,
    GoedgekeurdOp: nowStamp(),
    ManagerOpmerking: note
  });

  rebuildCentralWarehouseOverview();

  pushWarehouseNotification(
    'CentraleTellingGoedgekeurd',
    'Centrale telling goedgekeurd',
    `Centrale telling ${tellingId} is goedgekeurd en verwerkt.`,
    'CentraleStocktelling',
    tellingId
  );

  writeAudit(
    count.scopeType === 'Gericht'
      ? 'Gerichte centrale telling goedgekeurd'
      : 'Centrale stocktelling goedgekeurd',
    user.rol,
    actor,
    'CentraleStocktelling',
    tellingId,
    {
      note: note,
      mutaties: movements.length,
      deltaLijnen: lines.filter(line => safeNumber(line.deltaAantal, 0) !== 0).length
    }
  );

  return {
    success: true,
    message: 'Centrale telling goedgekeurd.'
  };
}

/* ---------------------------------------------------------
   Historiek
   --------------------------------------------------------- */

function buildCentralCountHistory() {
  return buildCentralCountsWithLines(getAllCentralCounts(), getAllCentralCountLines());
}