/* =========================================================
   43_BusCountService.gs — busstocktellingen
   ========================================================= */

function makeBusCountId() {
  return makeStampedId('T');
}

function getAllBusCounts() {
  return readObjectsSafe(TABS.BUS_COUNTS)
    .map(mapBusCount)
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.tellingId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.tellingId)}`
      )
    );
}

function getAllBusCountLines() {
  return readObjectsSafe(TABS.BUS_COUNT_LINES)
    .map(mapBusCountLine)
    .filter(line => line.actief);
}

function getBusCountById(tellingId) {
  const id = safeText(tellingId);
  if (!id) return null;
  return getAllBusCounts().find(item => item.tellingId === id) || null;
}

function getBusCountLinesById(tellingId) {
  const id = safeText(tellingId);
  if (!id) return [];
  return getAllBusCountLines().filter(line => line.tellingId === id);
}

function buildBusCountsWithLines(counts, lines) {
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
    lineCount: (linesByCount[count.tellingId] || []).length
  }));
}

function buildBusCountLineObject(tellingId, line) {
  const artikelCode = safeText(line.artikelCode);
  const article = getArticleMaster(artikelCode);

  const systeemAantal = safeNumber(line.systeemAantal, 0);
  const geteldAantal = line.geteldAantal === '' || line.geteldAantal == null
    ? ''
    : safeNumber(line.geteldAantal, 0);

  const deltaAantal = geteldAantal === ''
    ? ''
    : safeNumber(geteldAantal, 0) - systeemAantal;

  return {
    TellingID: safeText(tellingId),
    ArtikelCode: artikelCode,
    ArtikelOmschrijving: safeText(line.artikelOmschrijving) || safeText(article && article.artikelOmschrijving),
    Eenheid: safeText(line.eenheid) || safeText(article && article.eenheid),
    SysteemAantal: systeemAantal,
    GeteldAantal: geteldAantal,
    DeltaAantal: deltaAantal,
    Actief: 'Ja'
  };
}

function replaceBusCountLinesForCount(tellingId, lineObjects) {
  const id = safeText(tellingId);

  const sheet = getSheetOrThrow(TABS.BUS_COUNT_LINES);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.BUS_COUNT_LINES);
  const col = getColMap(headers);
  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => safeText(row[col['TellingID']]) !== id);
  const newRows = (lineObjects || []).map(obj => buildRowFromHeaders(headers, obj));

  writeFullTable(TABS.BUS_COUNT_LINES, headers, keptRows.concat(newRows));

  return {
    success: true,
    lines: newRows.length
  };
}

function updateBusCountHeader(tellingId, updates) {
  const id = safeText(tellingId);

  const sheet = getSheetOrThrow(TABS.BUS_COUNTS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab BusStockTellingen is leeg.');

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

  if (!updated) {
    throw new Error('Busstocktelling niet gevonden.');
  }

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return { success: true };
}

function replaceBusCountMovements(tellingId, movementPayloads) {
  const id = safeText(tellingId);

  const sheet = getSheetOrThrow(TABS.WAREHOUSE_MOVEMENTS);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.WAREHOUSE_MOVEMENTS);
  const col = getColMap(headers);
  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => {
    const bronId = safeText(row[col['BronID']]);
    const typeMutatie = safeText(row[col['TypeMutatie']]);

    if (bronId !== id) return true;
    if (typeMutatie !== 'BusCorrectieIn' && typeMutatie !== 'BusCorrectieUit') return true;
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

function canUserCreateBusCount(user) {
  return roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER]);
}

function canUserEditBusCount(count, user) {
  if (!count || !user) return false;
  if (user.rol === ROLE.ADMIN) return true;
  if (user.rol === ROLE.MANAGER) return true;

  if (user.rol === ROLE.TECHNICIAN) {
    return normalizeRef(count.techniekerCode) === normalizeRef(user.techniekerCode);
  }

  return false;
}

function assertBusCountTechnicianAccess(count, user) {
  if (!count) throw new Error('Busstocktelling niet gevonden.');
  if (!canUserEditBusCount(count, user)) {
    throw new Error('Je kan enkel je eigen busstocktelling aanpassen.');
  }
}

function buildBusCountRequestLineObjects(techniekerCode, requestedArticleCodes, tellingId) {
  const busStock = buildBusStockRows()
    .filter(item => normalizeRef(item.techniekerCode) === normalizeRef(techniekerCode));

  const busStockMap = {};
  busStock.forEach(item => {
    busStockMap[item.artikelCode] = item;
  });

  const articleMaster = getArticleMasterMap();

  const codes = (requestedArticleCodes || []).map(safeText).filter(Boolean);
  const uniqueCodes = [...new Set(codes)];

  return uniqueCodes.map(code => {
    const busItem = busStockMap[code] || null;
    const article = articleMaster[code] || null;

    return buildBusCountLineObject(tellingId, {
      artikelCode: code,
      artikelOmschrijving: busItem ? busItem.artikelOmschrijving : safeText(article && article.artikelOmschrijving),
      eenheid: busItem ? busItem.eenheid : safeText(article && article.eenheid),
      systeemAantal: busItem ? safeNumber(busItem.voorraadBus, 0) : 0,
      geteldAantal: ''
    });
  });
}

function createBusCountRequest(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!canUserCreateBusCount(user)) {
    throw new Error('Geen rechten om een busstocktelling aan te maken.');
  }

  const techniekerCode = safeText(payload.techniekerCode);
  const documentDatum = safeText(payload.documentDatum);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');
  const reden = safeText(payload.reden);
  const requestedArticles = Array.isArray(payload.requestedArticles) ? payload.requestedArticles : [];

  if (!techniekerCode) throw new Error('Technieker is verplicht.');
  if (!documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!reden) throw new Error('Reden is verplicht.');

  const techniekerNaam = getTechnicianNameByCode(techniekerCode);
  const busStockRows = buildBusStockRows()
    .filter(item => normalizeRef(item.techniekerCode) === normalizeRef(techniekerCode));

  const articleCodes = requestedArticles.length
    ? requestedArticles.map(safeText).filter(Boolean)
    : busStockRows.map(item => safeText(item.artikelCode));

  const uniqueCodes = [...new Set(articleCodes)];
  if (!uniqueCodes.length) {
    throw new Error('Geen artikels gevonden voor busstocktelling.');
  }

  const tellingId = makeBusCountId();
  const scopeType = requestedArticles.length ? 'Gericht' : 'Volledig';
  const scopeArtikelCodes = requestedArticles.length ? uniqueCodes.join(', ') : '';

  appendObjects(TABS.BUS_COUNTS, [{
    TellingID: tellingId,
    TechniekerCode: techniekerCode,
    TechniekerNaam: techniekerNaam,
    DocumentDatum: documentDatum,
    Status: BUS_COUNT_STATUS.OPEN,
    ScopeType: scopeType,
    ScopeArtikelCodes: scopeArtikelCodes,
    Reden: reden,
    AangemaaktDoor: actor,
    AangemaaktOp: nowStamp(),
    IngediendDoor: '',
    IngediendOp: '',
    GoedgekeurdDoor: '',
    GoedgekeurdOp: '',
    ManagerOpmerking: ''
  }]);

  const lineObjects = buildBusCountRequestLineObjects(techniekerCode, uniqueCodes, tellingId);
  appendObjects(TABS.BUS_COUNT_LINES, lineObjects);

  pushTechnicianNotification(
    techniekerCode,
    techniekerNaam,
    'Busstocktelling',
    'Nieuwe busstocktelling',
    `Er staat een ${scopeType.toLowerCase()} busstocktelling klaar (${tellingId}). Vul deze in en dien ze in.`,
    'Busstocktelling',
    tellingId
  );

  writeAudit(
    'Busstocktelling aangemaakt',
    user.rol,
    actor,
    'Busstocktelling',
    tellingId,
    {
      techniekerCode: techniekerCode,
      techniekerNaam: techniekerNaam,
      scopeType: scopeType,
      scopeArtikelCodes: scopeArtikelCodes,
      lijnen: lineObjects.length,
      reden: reden
    }
  );

  return {
    success: true,
    tellingId,
    message: 'Busstocktelling aangemaakt.'
  };
}

function createBusCountRequestFromAlert(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!canUserCreateBusCount(user)) {
    throw new Error('Geen rechten om een gerichte telling aan te maken.');
  }

  const techniekerCode = safeText(payload.techniekerCode);
  const documentDatum = safeText(payload.documentDatum);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');
  const artikelCode = safeText(payload.artikelCode);
  const reden = safeText(payload.reden);

  if (!artikelCode) throw new Error('Artikelcode ontbreekt.');

  return createBusCountRequest({
    sessionId: sessionId,
    techniekerCode,
    documentDatum,
    actor,
    reden,
    requestedArticles: [artikelCode]
  });
}

function getBusCountDataForTechnician(techRef, sessionId) {
  const access = assertTechnicianAccessToRef(techRef, sessionId);

  const counts = getAllBusCounts()
    .filter(item => normalizeRef(item.techniekerCode) === normalizeRef(access.technician.code))
    .filter(item => item.status !== BUS_COUNT_STATUS.APPROVED)
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.tellingId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.tellingId)}`
      )
    );

  const lines = getAllBusCountLines();

  return {
    counts: buildBusCountsWithLines(counts, lines)
  };
}

function validateSavedBusCountLines(lines) {
  if (!(lines || []).length) {
    throw new Error('Geen tellinglijnen ontvangen.');
  }

  (lines || []).forEach(line => {
    const code = safeText(line.artikelCode);
    const qty = safeNumber(line.geteldAantal, 0);

    if (!code) {
      throw new Error('Artikelcode ontbreekt op een tellinglijn.');
    }

    if (qty < 0) {
      throw new Error('Geteld aantal mag niet negatief zijn voor artikel ' + code);
    }
  });
}

function saveBusCountLines(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  const tellingId = safeText(payload.tellingId);
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!tellingId) throw new Error('TellingID ontbreekt.');
  if (!lines.length) throw new Error('Geen tellinglijnen ontvangen.');

  const count = getBusCountById(tellingId);
  if (!count) throw new Error('Busstocktelling niet gevonden.');

  assertBusCountTechnicianAccess(count, user);

  if ([BUS_COUNT_STATUS.SUBMITTED, BUS_COUNT_STATUS.APPROVED].includes(count.status)) {
    throw new Error('Deze busstocktelling kan niet meer aangepast worden.');
  }

  const existingLines = getBusCountLinesById(tellingId);
  if (!existingLines.length) {
    throw new Error('Geen lijnstructuur gevonden voor deze busstocktelling.');
  }

  const inputMap = {};
  lines.forEach(line => {
    const code = safeText(line.artikelCode);
    if (!code) return;
    inputMap[code] = safeNumber(line.geteldAantal, 0);
  });

  const rewrittenObjects = existingLines.map(line => {
    const code = safeText(line.artikelCode);
    const systeemAantal = safeNumber(line.systeemAantal, 0);
    const geteldAantal = Object.prototype.hasOwnProperty.call(inputMap, code)
      ? inputMap[code]
      : (line.geteldAantal === '' || line.geteldAantal == null ? '' : safeNumber(line.geteldAantal, 0));

    return buildBusCountLineObject(tellingId, {
      artikelCode: code,
      artikelOmschrijving: line.artikelOmschrijving,
      eenheid: line.eenheid,
      systeemAantal: systeemAantal,
      geteldAantal: geteldAantal
    });
  });

  const validateLines = rewrittenObjects.map(obj => ({
    artikelCode: obj.ArtikelCode,
    geteldAantal: obj.GeteldAantal === '' ? 0 : obj.GeteldAantal
  }));

  validateSavedBusCountLines(validateLines);

  replaceBusCountLinesForCount(tellingId, rewrittenObjects);

  writeAudit(
    'Busstocktelling opgeslagen',
    user.rol,
    user.naam || user.techniekerCode || user.email,
    'Busstocktelling',
    tellingId,
    {
      lijnen: rewrittenObjects.length
    }
  );

  return {
    success: true,
    message: 'Busstocktelling opgeslagen.'
  };
}

function submitBusCount(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  const tellingId = safeText(payload.tellingId);
  const actor = safeText(payload.actor || user.naam || user.techniekerCode || 'Technieker');

  if (!tellingId) throw new Error('TellingID ontbreekt.');

  const count = getBusCountById(tellingId);
  if (!count) throw new Error('Busstocktelling niet gevonden.');

  assertBusCountTechnicianAccess(count, user);

  const lineRows = getBusCountLinesById(tellingId);
  if (!lineRows.length) {
    throw new Error('Deze busstocktelling bevat nog geen actieve lijnen.');
  }

  if (lineRows.some(row => row.geteldAantal === '' || row.geteldAantal == null)) {
    throw new Error('Vul eerst alle getelde aantallen in.');
  }

  updateBusCountHeader(tellingId, {
    Status: BUS_COUNT_STATUS.SUBMITTED,
    IngediendDoor: actor,
    IngediendOp: nowStamp()
  });

  pushManagerNotification(
    'BusstocktellingGoedTeKeuren',
    'Busstocktelling ingediend',
    `Busstocktelling ${tellingId} is ingediend en wacht op goedkeuring.`,
    'Busstocktelling',
    tellingId
  );

  writeAudit(
    'Busstocktelling ingediend',
    user.rol,
    actor,
    'Busstocktelling',
    tellingId,
    {
      status: BUS_COUNT_STATUS.SUBMITTED,
      scopeType: count.scopeType
    }
  );

  return {
    success: true,
    message: 'Busstocktelling ingediend voor goedkeuring.'
  };
}

function buildBusCountApprovalMovements(count, lineRows, actor, note) {
  const busLocation = getBusLocationCode(count.techniekerCode);
  const documentDatum = safeText(count.documentDatumIso || count.documentDatum);

  return (lineRows || [])
    .map(line => {
      const delta = safeNumber(line.deltaAantal, 0);
      if (!delta) return null;

      if (delta > 0) {
        return {
          datumDocument: documentDatum,
          typeMutatie: 'BusCorrectieIn',
          typeMateriaal: determineMaterialTypeFromArticle(line.artikelCode),
          artikelCode: line.artikelCode,
          artikelOmschrijving: line.artikelOmschrijving,
          eenheid: line.eenheid,
          aantalIn: delta,
          aantalUit: 0,
          nettoAantal: delta,
          locatieVan: '',
          locatieNaar: busLocation,
          reden: 'Busstocktelling',
          opmerking: safeText(note),
          goedgekeurdDoor: actor,
          goedgekeurdOp: nowStamp()
        };
      }

      return {
        datumDocument: documentDatum,
        typeMutatie: 'BusCorrectieUit',
        typeMateriaal: determineMaterialTypeFromArticle(line.artikelCode),
        artikelCode: line.artikelCode,
        artikelOmschrijving: line.artikelOmschrijving,
        eenheid: line.eenheid,
        aantalIn: 0,
        aantalUit: Math.abs(delta),
        nettoAantal: delta,
        locatieVan: busLocation,
        locatieNaar: '',
        reden: 'Busstocktelling',
        opmerking: safeText(note),
        goedgekeurdDoor: actor,
        goedgekeurdOp: nowStamp()
      };
    })
    .filter(Boolean);
}

function approveBusCount(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertManagerAccess(sessionId);

  const tellingId = safeText(payload.tellingId);
  const actor = safeText(payload.actor || user.naam || 'Manager');
  const note = safeText(payload.note);

  if (!tellingId) throw new Error('TellingID ontbreekt.');

  const count = getBusCountById(tellingId);
  if (!count) throw new Error('Busstocktelling niet gevonden.');

  const lineRows = getBusCountLinesById(tellingId);
  if (!lineRows.length) {
    throw new Error('Deze busstocktelling bevat geen actieve lijnen.');
  }

  updateBusCountHeader(tellingId, {
    Status: BUS_COUNT_STATUS.APPROVED,
    GoedgekeurdDoor: actor,
    GoedgekeurdOp: nowStamp(),
    ManagerOpmerking: note
  });

  const movementPayloads = buildBusCountApprovalMovements(count, lineRows, actor, note);
  replaceBusCountMovements(tellingId, movementPayloads);

  rebuildCentralWarehouseOverview();

  markManagerNotificationsBySource('Busstocktelling', tellingId);

  pushWarehouseNotification(
    'BusstocktellingGoedgekeurd',
    'Busstocktelling goedgekeurd',
    `Busstocktelling ${tellingId} van ${count.techniekerNaam || count.techniekerCode} is goedgekeurd en verwerkt.`,
    'Busstocktelling',
    tellingId
  );

  pushTechnicianNotification(
    count.techniekerCode,
    count.techniekerNaam,
    'BusstocktellingGoedgekeurd',
    'Je busstocktelling is verwerkt',
    `Busstocktelling ${tellingId} werd goedgekeurd door de manager.`,
    'Busstocktelling',
    tellingId
  );

  writeAudit(
    'Busstocktelling goedgekeurd',
    user.rol,
    actor,
    'Busstocktelling',
    tellingId,
    {
      techniekerCode: count.techniekerCode,
      techniekerNaam: count.techniekerNaam,
      scopeType: count.scopeType,
      note: note,
      mutaties: movementPayloads.length
    }
  );

  return {
    success: true,
    message: 'Busstocktelling goedgekeurd.'
  };
}