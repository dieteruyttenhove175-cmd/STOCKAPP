/* =========================================================
   41_ReturnService.gs — retouren
   Volledig herschreven:
   - bus -> centraal magazijn
   - centraal magazijn -> Fluvius
   - beschikbaar aantal op read
   - zoekfunctie per retourtype
   - extra validatie op save en approval
   ========================================================= */

function makeReturnId() {
  return makeStampedId('RET');
}

function getAllReturns() {
  return readObjectsSafe(TABS.RETURNS)
    .map(mapReturn)
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.retourId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.retourId)}`
      )
    );
}

function getAllReturnLines() {
  return readObjectsSafe(TABS.RETURN_LINES)
    .map(mapReturnLine)
    .filter(line => line.actief);
}

function getReturnById(retourId) {
  const id = safeText(retourId);
  if (!id) return null;
  return getAllReturns().find(item => item.retourId === id) || null;
}

function getReturnLinesByReturnId(retourId) {
  const id = safeText(retourId);
  if (!id) return [];
  return getAllReturnLines().filter(line => line.retourId === id);
}

function isBusReturnDocument(retour) {
  return !!safeText(retour && retour.techniekerCode);
}

function isCentralReturnToFluvius(retour) {
  return !isBusReturnDocument(retour);
}

function getReturnMode(retour) {
  return isBusReturnDocument(retour) ? 'BUS_TO_CENTRAL' : 'CENTRAL_TO_FLUVIUS';
}

function normalizeReturnReason_(reason) {
  const value = safeText(reason);

  if (value === 'NCP') return 'NPC / beschadigde artikelen';
  if (value === 'NPC') return 'NPC / beschadigde artikelen';

  return value;
}

function matchesReturnCatalogQuery_(query, artikelCode, artikelOmschrijving) {
  const q = safeText(query).toLowerCase();
  if (!q) return true;

  return (
    safeText(artikelCode).toLowerCase().includes(q) ||
    safeText(artikelOmschrijving).toLowerCase().includes(q)
  );
}

function getAvailableReturnQty_(retour, artikelCode) {
  const code = safeText(artikelCode);
  if (!code) return 0;

  if (isBusReturnDocument(retour)) {
    const busStockMap = getBusStockMapForTechnician(retour.techniekerCode);
    return busStockMap[code] ? safeNumber(busStockMap[code].voorraadBus, 0) : 0;
  }

  const centralMap = buildCentralWarehouseMap();
  return centralMap[code] ? safeNumber(centralMap[code].voorraadCentraal, 0) : 0;
}

function addAvailabilityToReturnLines_(retour, lines) {
  return (lines || []).map(line => ({
    ...line,
    beschikbaarAantal: getAvailableReturnQty_(retour, line.artikelCode)
  }));
}

function buildReturnsWithLines(returns, lines) {
  const linesByReturn = {};

  (lines || []).forEach(line => {
    const id = safeText(line.retourId);
    if (!id) return;

    if (!linesByReturn[id]) linesByReturn[id] = [];
    linesByReturn[id].push(line);
  });

  return (returns || []).map(retour => {
    const currentLines = addAvailabilityToReturnLines_(retour, linesByReturn[retour.retourId] || []);

    return {
      ...retour,
      returnMode: getReturnMode(retour),
      lines: currentLines,
      lineCount: currentLines.length
    };
  });
}

function getReturnsData(sessionId) {
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten voor retourdata.');
  }

  const returns = getAllReturns();
  const lines = getAllReturnLines();

  return {
    technicians: getActiveTechnicians().map(t => ({ code: t.code, naam: t.naam })),
    returns: buildReturnsWithLines(returns, lines)
  };
}

function getReturnsDataForTechnician(techRef, sessionId) {
  const access = assertTechnicianAccessToRef(techRef, sessionId);

  const allReturns = buildReturnsWithLines(getAllReturns(), getAllReturnLines())
    .filter(retour => normalizeRef(retour.techniekerCode) === normalizeRef(access.technician.code))
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.retourId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.retourId)}`
      )
    );

  return {
    technician: {
      code: access.technician.code,
      naam: access.technician.naam
    },
    returns: allReturns
  };
}

function canUserEditReturn(retour, user) {
  if (!retour || !user) return false;
  if (user.rol === ROLE.ADMIN) return true;
  if (user.rol === ROLE.MANAGER) return true;
  if (user.rol === ROLE.WAREHOUSE || user.rol === ROLE.MOBILE_WAREHOUSE) return true;

  if (user.rol === ROLE.TECHNICIAN) {
    return (
      isBusReturnDocument(retour) &&
      normalizeRef(retour.techniekerCode) === normalizeRef(user.techniekerCode)
    );
  }

  return false;
}

function assertReturnEditAccess(retour, user) {
  if (!canUserEditReturn(retour, user)) {
    throw new Error('Geen rechten voor deze retour.');
  }
}

function getAllowedReasonsForReturn(retour, user) {
  if (user && user.rol === ROLE.TECHNICIAN) {
    return getAllowedReturnReasonsByActor(ROLE.TECHNICIAN);
  }

  if (isCentralReturnToFluvius(retour)) {
    return getAllowedReturnReasonsByActor(ROLE.WAREHOUSE);
  }

  return getAllowedReturnReasonsByActor(ROLE.WAREHOUSE);
}

function buildReturnLineObject(retourId, line) {
  const artikelCode = safeText(line.artikelCode);
  const article = getArticleMaster(artikelCode);

  return {
    RetourID: safeText(retourId),
    ArtikelCode: artikelCode,
    ArtikelOmschrijving: safeText(line.artikelOmschrijving) || safeText(article && article.artikelOmschrijving),
    Eenheid: safeText(line.eenheid) || safeText(article && article.eenheid),
    Aantal: safeNumber(line.aantal, 0),
    Reden: normalizeReturnReason_(line.reden),
    Opmerking: safeText(line.opmerking),
    Actief: 'Ja'
  };
}

function replaceReturnLinesForReturn(retourId, lineObjects) {
  const id = safeText(retourId);

  const sheet = getSheetOrThrow(TABS.RETURN_LINES);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.RETURN_LINES);
  const col = getColMap(headers);
  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => safeText(row[col['RetourID']]) !== id);
  const newRows = (lineObjects || []).map(obj => buildRowFromHeaders(headers, obj));

  writeFullTable(TABS.RETURN_LINES, headers, keptRows.concat(newRows));

  return {
    success: true,
    lines: newRows.length
  };
}

function updateReturnHeader(retourId, updates) {
  const id = safeText(retourId);

  const sheet = getSheetOrThrow(TABS.RETURNS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab Retouren is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['RetourID']]) !== id) continue;

    Object.keys(updates || {}).forEach(field => {
      if (col[field] !== undefined) {
        values[i][col[field]] = updates[field];
      }
    });

    updated = true;
    break;
  }

  if (!updated) {
    throw new Error('Retour niet gevonden.');
  }

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return { success: true };
}

function createReturn(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  const documentDatum = safeText(payload.documentDatum);
  const requestedMode = safeText(payload.returnMode || payload.retourType || 'BUS_TO_CENTRAL');

  if (!documentDatum) throw new Error('Documentdatum is verplicht.');

  if (
    !roleAllowed(user, [
      ROLE.TECHNICIAN,
      ROLE.WAREHOUSE,
      ROLE.MOBILE_WAREHOUSE,
      ROLE.MANAGER
    ])
  ) {
    throw new Error('Geen rechten om een retour aan te maken.');
  }

  let techniekerCode = '';
  let techniekerNaam = '';
  let returnMode = requestedMode === 'CENTRAL_TO_FLUVIUS' ? 'CENTRAL_TO_FLUVIUS' : 'BUS_TO_CENTRAL';

  if (user.rol === ROLE.TECHNICIAN) {
    returnMode = 'BUS_TO_CENTRAL';
    techniekerCode = safeText(user.techniekerCode);

    if (!techniekerCode) {
      throw new Error('Techniekerkoppeling ontbreekt op gebruiker.');
    }

    techniekerNaam = getTechnicianNameByCode(techniekerCode);
  } else if (returnMode === 'BUS_TO_CENTRAL') {
    techniekerCode = safeText(payload.techniekerCode);
    if (!techniekerCode) {
      throw new Error('Technieker / bus is verplicht voor busretour.');
    }
    techniekerNaam = getTechnicianNameByCode(techniekerCode);
  }

  if (returnMode === 'CENTRAL_TO_FLUVIUS' && user.rol === ROLE.TECHNICIAN) {
    throw new Error('Techniekers kunnen geen centrale retour naar Fluvius aanmaken.');
  }

  const retourId = makeReturnId();

  appendObjects(TABS.RETURNS, [{
    RetourID: retourId,
    TypeMateriaal: '',
    Leverancier: 'Fluvius',
    TechniekerCode: techniekerCode,
    TechniekerNaam: techniekerNaam,
    DocumentDatum: documentDatum,
    Status: RETURN_STATUS.IN_PROGRESS,
    IngediendDoor: '',
    IngediendOp: '',
    GoedgekeurdDoor: '',
    GoedgekeurdOp: '',
    ManagerOpmerking: '',
    RetourFlow: returnMode,
    LocatieVan: returnMode === 'CENTRAL_TO_FLUVIUS' ? LOCATION.CENTRAL : getBusLocationCode(techniekerCode),
    LocatieNaar: returnMode === 'CENTRAL_TO_FLUVIUS' ? 'Fluvius' : LOCATION.CENTRAL,
    InitiatiefRol: user.rol
  }]);

  writeAudit(
    'Retour aangemaakt',
    user.rol,
    user.naam || user.email || user.techniekerCode,
    'Retour',
    retourId,
    {
      returnMode: returnMode,
      techniekerCode: techniekerCode,
      techniekerNaam: techniekerNaam,
      documentDatum: documentDatum
    }
  );

  return {
    success: true,
    retourId: retourId,
    message: 'Retour aangemaakt.'
  };
}

function searchBusReturnCatalog_(techniekerCode, query) {
  const busStockMap = getBusStockMapForTechnician(techniekerCode);
  const codes = Object.keys(busStockMap || {});

  return codes
    .map(code => {
      const row = busStockMap[code] || {};
      const article = getArticleMaster(code) || {};

      return {
        artikelCode: code,
        artikelOmschrijving: safeText(row.artikelOmschrijving || article.artikelOmschrijving),
        eenheid: safeText(row.eenheid || article.eenheid),
        beschikbaarAantal: safeNumber(row.voorraadBus, 0)
      };
    })
    .filter(item =>
      item.beschikbaarAantal > 0 &&
      matchesReturnCatalogQuery_(query, item.artikelCode, item.artikelOmschrijving)
    )
    .sort((a, b) => safeText(a.artikelCode).localeCompare(safeText(b.artikelCode)));
}

function searchCentralReturnCatalog_(query) {
  const centralMap = buildCentralWarehouseMap();
  const codes = Object.keys(centralMap || {});

  return codes
    .map(code => {
      const row = centralMap[code] || {};
      const article = getArticleMaster(code) || {};

      return {
        artikelCode: code,
        artikelOmschrijving: safeText(row.artikelOmschrijving || article.artikelOmschrijving),
        eenheid: safeText(row.eenheid || article.eenheid),
        beschikbaarAantal: safeNumber(row.voorraadCentraal, 0)
      };
    })
    .filter(item =>
      item.beschikbaarAantal > 0 &&
      matchesReturnCatalogQuery_(query, item.artikelCode, item.artikelOmschrijving)
    )
    .sort((a, b) => safeText(a.artikelCode).localeCompare(safeText(b.artikelCode)));
}

function searchReturnCatalog(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  const retourId = safeText(payload.retourId);
  const query = safeText(payload.query);

  if (!retourId) throw new Error('RetourID ontbreekt.');

  const retour = getReturnById(retourId);
  if (!retour) throw new Error('Retour niet gevonden.');

  assertReturnEditAccess(retour, user);

  if (isBusReturnDocument(retour)) {
    return searchBusReturnCatalog_(retour.techniekerCode, query);
  }

  return searchCentralReturnCatalog_(query);
}

function validateReturnLinesForContext(retour, cleanedLines, user) {
  if (!(cleanedLines || []).length) {
    throw new Error('Geen geldige retourlijnen om te bewaren.');
  }

  const allowedReasons = getAllowedReasonsForReturn(retour, user);

  cleanedLines.forEach(line => {
    const artikelCode = safeText(line.artikelCode);
    const aantal = safeNumber(line.aantal, 0);
    const reden = normalizeReturnReason_(line.reden);
    const opmerking = safeText(line.opmerking);

    if (!artikelCode) {
      throw new Error('Artikelcode ontbreekt op een retourlijn.');
    }

    if (aantal <= 0) {
      throw new Error('Aantal moet groter zijn dan 0 voor artikel ' + artikelCode);
    }

    if (!allowedReasons.includes(reden)) {
      throw new Error('Ongeldige retourreden voor artikel ' + artikelCode);
    }

    if (reden === 'Andere reden' && !opmerking) {
      throw new Error('Opmerking is verplicht bij Andere reden voor artikel ' + artikelCode);
    }

    if (isBusReturnDocument(retour) && !isBehoefteArticle(artikelCode)) {
      throw new Error('Techniekerretour is enkel toegestaan voor behoeftemateriaal. Artikel: ' + artikelCode);
    }

    const beschikbaar = getAvailableReturnQty_(retour, artikelCode);

    if (beschikbaar <= 0) {
      throw new Error('Artikel ' + artikelCode + ' is niet beschikbaar voor retour.');
    }

    if (aantal > beschikbaar) {
      throw new Error(
        'Retouraantal overschrijdt beschikbare voorraad voor artikel ' +
        artikelCode +
        '. Beschikbaar: ' +
        beschikbaar
      );
    }
  });
}

function saveReturnLines(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  const retourId = safeText(payload.retourId);
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!retourId) throw new Error('RetourID ontbreekt.');
  if (!lines.length) throw new Error('Geen retourlijnen ontvangen.');

  const retour = getReturnById(retourId);
  if (!retour) throw new Error('Retour niet gevonden.');

  assertReturnEditAccess(retour, user);

  if ([RETURN_STATUS.SUBMITTED, RETURN_STATUS.APPROVED].includes(retour.status)) {
    throw new Error('Deze retour kan niet meer aangepast worden.');
  }

  const cleaned = lines
    .map(line => ({
      artikelCode: safeText(line.artikelCode),
      artikelOmschrijving: safeText(line.artikelOmschrijving),
      eenheid: safeText(line.eenheid),
      aantal: safeNumber(line.aantal, 0),
      reden: normalizeReturnReason_(line.reden),
      opmerking: safeText(line.opmerking)
    }))
    .filter(line => line.artikelCode && line.artikelOmschrijving && line.aantal > 0);

  validateReturnLinesForContext(retour, cleaned, user);

  const lineObjects = cleaned.map(line => buildReturnLineObject(retourId, line));
  replaceReturnLinesForReturn(retourId, lineObjects);

  const returnType = determineReturnMaterialType(cleaned);

  updateReturnHeader(retourId, {
    TypeMateriaal: returnType,
    Leverancier: 'Fluvius'
  });

  writeAudit(
    'Retourlijnen opgeslagen',
    user.rol,
    user.naam || user.email || user.techniekerCode,
    'Retour',
    retourId,
    {
      lijnen: lineObjects.length,
      returnMode: getReturnMode(retour),
      typeMateriaal: returnType
    }
  );

  return {
    success: true,
    lines: lineObjects.length,
    message: 'Retourlijnen opgeslagen.'
  };
}

function submitReturn(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  const retourId = safeText(payload.retourId);
  const actor = safeText(payload.actor || user.naam || user.email || user.techniekerCode);

  if (!retourId) throw new Error('RetourID ontbreekt.');

  const retour = getReturnById(retourId);
  if (!retour) throw new Error('Retour niet gevonden.');

  assertReturnEditAccess(retour, user);

  const lineRows = getReturnLinesByReturnId(retourId);
  if (!lineRows.length) {
    throw new Error('Deze retour bevat nog geen actieve lijnen.');
  }

  updateReturnHeader(retourId, {
    Status: RETURN_STATUS.SUBMITTED,
    IngediendDoor: actor,
    IngediendOp: nowStamp()
  });

  pushManagerNotification(
    'RetourGoedTeKeuren',
    'Retour ingediend',
    `Retour ${retourId} is ingediend en wacht op goedkeuring.`,
    'Retour',
    retourId
  );

  if (user.rol === ROLE.TECHNICIAN) {
    pushWarehouseNotification(
      'TechniekerRetourIngediend',
      'Techniekerretour ingediend',
      `${retour.techniekerNaam || retour.techniekerCode} diende retour ${retourId} in.`,
      'Retour',
      retourId
    );
  }

  writeAudit(
    'Retour ingediend',
    user.rol,
    actor,
    'Retour',
    retourId,
    {
      status: RETURN_STATUS.SUBMITTED,
      returnMode: getReturnMode(retour),
      lijnen: lineRows.length
    }
  );

  return {
    success: true,
    message: 'Retour ingediend voor goedkeuring.'
  };
}

function buildApprovalMovementsForBusReturn(retour, lines, actor, note) {
  const busLocation = getBusLocationCode(retour.techniekerCode);
  const documentDatum = safeText(retour.documentDatumIso || retour.documentDatum);

  return (lines || [])
    .filter(line => safeNumber(line.aantal, 0) > 0)
    .map(line => {
      const aantal = safeNumber(line.aantal, 0);

      return {
        datumDocument: documentDatum,
        typeMutatie: 'RetourIn',
        bronId: retour.retourId,
        typeMateriaal: determineMaterialTypeFromArticle(line.artikelCode),
        artikelCode: line.artikelCode,
        artikelOmschrijving: line.artikelOmschrijving,
        eenheid: line.eenheid,
        aantalIn: aantal,
        aantalUit: aantal,
        nettoAantal: 0,
        locatieVan: busLocation,
        locatieNaar: LOCATION.CENTRAL,
        reden: line.reden,
        opmerking: safeText(line.opmerking || note),
        goedgekeurdDoor: actor,
        goedgekeurdOp: nowStamp()
      };
    });
}

function buildApprovalMovementsForCentralReturn(retour, lines, actor, note) {
  const documentDatum = safeText(retour.documentDatumIso || retour.documentDatum);

  return (lines || [])
    .filter(line => safeNumber(line.aantal, 0) > 0)
    .map(line => {
      const aantal = safeNumber(line.aantal, 0);

      return {
        datumDocument: documentDatum,
        typeMutatie: 'RetourFluvius',
        bronId: retour.retourId,
        typeMateriaal: determineMaterialTypeFromArticle(line.artikelCode),
        artikelCode: line.artikelCode,
        artikelOmschrijving: line.artikelOmschrijving,
        eenheid: line.eenheid,
        aantalIn: 0,
        aantalUit: aantal,
        nettoAantal: -aantal,
        locatieVan: LOCATION.CENTRAL,
        locatieNaar: 'Fluvius',
        reden: line.reden,
        opmerking: safeText(line.opmerking || note),
        goedgekeurdDoor: actor,
        goedgekeurdOp: nowStamp()
      };
    });
}

function validateReturnStockBeforeApproval(retour, lines) {
  if (isBusReturnDocument(retour)) {
    const busStockMap = getBusStockMapForTechnician(retour.techniekerCode);

    (lines || []).forEach(line => {
      const code = safeText(line.artikelCode);
      const aantal = safeNumber(line.aantal, 0);
      const currentBusQty = busStockMap[code] ? safeNumber(busStockMap[code].voorraadBus, 0) : 0;

      if (aantal > currentBusQty) {
        throw new Error(`Retour ${code} overschrijdt busvoorraad. Huidig in bus: ${currentBusQty}`);
      }
    });

    return;
  }

  const centralMap = buildCentralWarehouseMap();

  (lines || []).forEach(line => {
    const code = safeText(line.artikelCode);
    const aantal = safeNumber(line.aantal, 0);
    const currentCentral = centralMap[code] ? safeNumber(centralMap[code].voorraadCentraal, 0) : 0;

    if (aantal > currentCentral) {
      throw new Error(`Retour ${code} overschrijdt centrale voorraad. Huidig centraal: ${currentCentral}`);
    }
  });
}

function approveReturn(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertManagerAccess(sessionId);

  const retourId = safeText(payload.retourId);
  const actor = safeText(payload.actor || user.naam || 'Manager');
  const note = safeText(payload.note);

  if (!retourId) throw new Error('RetourID ontbreekt.');

  const retour = getReturnById(retourId);
  if (!retour) throw new Error('Retour niet gevonden.');

  const lineRows = getReturnLinesByReturnId(retourId);
  if (!lineRows.length) {
    throw new Error('Deze retour bevat geen actieve lijnen.');
  }

  validateReturnStockBeforeApproval(retour, lineRows);

  updateReturnHeader(retourId, {
    Status: RETURN_STATUS.APPROVED,
    GoedgekeurdDoor: actor,
    GoedgekeurdOp: nowStamp(),
    ManagerOpmerking: note
  });

  const movementPayloads = isBusReturnDocument(retour)
    ? buildApprovalMovementsForBusReturn(retour, lineRows, actor, note)
    : buildApprovalMovementsForCentralReturn(retour, lineRows, actor, note);

  replaceSourceMovements(
    isBusReturnDocument(retour) ? 'RetourIn' : 'RetourFluvius',
    retourId,
    movementPayloads
  );

  rebuildCentralWarehouseOverview();

  markManagerNotificationsBySource('Retour', retourId);

  pushWarehouseNotification(
    'RetourGoedgekeurd',
    'Retour goedgekeurd',
    isBusReturnDocument(retour)
      ? `Retour ${retourId} is goedgekeurd en terug geboekt naar centraal magazijn.`
      : `Retour ${retourId} is goedgekeurd en afgeboekt van centraal magazijn naar Fluvius.`,
    'Retour',
    retourId
  );

  if (isBusReturnDocument(retour)) {
    pushTechnicianNotification(
      retour.techniekerCode,
      retour.techniekerNaam,
      'RetourGoedgekeurd',
      'Je retour is goedgekeurd',
      `Retour ${retourId} werd goedgekeurd door de manager.`,
      'Retour',
      retourId
    );
  }

  writeAudit(
    'Retour goedgekeurd',
    user.rol,
    actor,
    'Retour',
    retourId,
    {
      note: note,
      returnMode: getReturnMode(retour),
      techniekerCode: retour.techniekerCode,
      lijnen: movementPayloads.length
    }
  );

  return {
    success: true,
    message: isBusReturnDocument(retour)
      ? 'Retour goedgekeurd en terug geboekt naar centraal magazijn.'
      : 'Retour goedgekeurd en afgeboekt naar Fluvius.'
  };
}