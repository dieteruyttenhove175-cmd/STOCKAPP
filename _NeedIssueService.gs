/* =========================================================
   42_NeedIssueService.gs — behoefte-uitgiftes centraal → bus
   ========================================================= */

function makeNeedIssueId() {
  return makeStampedId('U');
}

function getAllNeedIssues() {
  return readObjectsSafe(TABS.NEED_ISSUES)
    .map(mapNeedIssue)
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.uitgifteId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.uitgifteId)}`
      )
    );
}

function getAllNeedIssueLines() {
  return readObjectsSafe(TABS.NEED_ISSUE_LINES)
    .map(mapNeedIssueLine)
    .filter(line => line.actief);
}

function getNeedIssueById(uitgifteId) {
  const id = safeText(uitgifteId);
  if (!id) return null;
  return getAllNeedIssues().find(item => item.uitgifteId === id) || null;
}

function getNeedIssueLinesByIssueId(uitgifteId) {
  const id = safeText(uitgifteId);
  if (!id) return [];
  return getAllNeedIssueLines().filter(line => line.uitgifteId === id);
}

function buildNeedIssueLineObject(uitgifteId, line) {
  const artikelCode = safeText(line.artikelCode);
  const article = getArticleMaster(artikelCode);

  return {
    UitgifteID: safeText(uitgifteId),
    ArtikelCode: artikelCode,
    ArtikelOmschrijving: safeText(line.artikelOmschrijving) || safeText(article && article.artikelOmschrijving),
    Eenheid: safeText(line.eenheid) || safeText(article && article.eenheid),
    Aantal: safeNumber(line.aantal, 0),
    Actief: 'Ja'
  };
}

function replaceNeedIssueLinesForIssue(uitgifteId, lineObjects) {
  const id = safeText(uitgifteId);

  const sheet = getSheetOrThrow(TABS.NEED_ISSUE_LINES);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.NEED_ISSUE_LINES);
  const col = getColMap(headers);
  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => safeText(row[col['UitgifteID']]) !== id);
  const newRows = (lineObjects || []).map(obj => buildRowFromHeaders(headers, obj));

  writeFullTable(TABS.NEED_ISSUE_LINES, headers, keptRows.concat(newRows));

  return {
    success: true,
    lines: newRows.length
  };
}

function updateNeedIssueHeader(uitgifteId, updates) {
  const id = safeText(uitgifteId);

  const sheet = getSheetOrThrow(TABS.NEED_ISSUES);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab BehoefteUitgiftes is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['UitgifteID']]) !== id) continue;

    Object.keys(updates || {}).forEach(field => {
      if (col[field] !== undefined) {
        values[i][col[field]] = updates[field];
      }
    });

    updated = true;
    break;
  }

  if (!updated) {
    throw new Error('Behoefte-uitgifte niet gevonden.');
  }

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return { success: true };
}

function buildNeedIssuesWithLines(issues, lines) {
  const linesByIssue = {};

  (lines || []).forEach(line => {
    const id = safeText(line.uitgifteId);
    if (!id) return;

    if (!linesByIssue[id]) linesByIssue[id] = [];
    linesByIssue[id].push(line);
  });

  return (issues || []).map(issue => ({
    ...issue,
    lines: linesByIssue[issue.uitgifteId] || [],
    lineCount: (linesByIssue[issue.uitgifteId] || []).length
  }));
}

function getNeedIssueData(sessionId) {
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten voor behoefte-uitgiftes.');
  }

  const issues = getAllNeedIssues();
  const lines = getAllNeedIssueLines();

  return {
    technicians: getActiveTechnicians().map(t => ({ code: t.code, naam: t.naam })),
    issues: buildNeedIssuesWithLines(issues, lines),
    centralWarehouse: getCentralWarehouseOverview(),
    busStock: buildBusStockRows()
  };
}

function createNeedIssue(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om een behoefte-uitgifte aan te maken.');
  }

  const techniekerCode = safeText(payload.techniekerCode);
  const documentDatum = safeText(payload.documentDatum);
  const reden = safeText(payload.reden);
  const opmerking = safeText(payload.opmerking);

  if (!techniekerCode) throw new Error('Technieker is verplicht.');
  if (!documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!reden) throw new Error('Reden is verplicht.');

  const techniekerNaam = getTechnicianNameByCode(techniekerCode);
  const uitgifteId = makeNeedIssueId();

  appendObjects(TABS.NEED_ISSUES, [{
    UitgifteID: uitgifteId,
    TechniekerCode: techniekerCode,
    TechniekerNaam: techniekerNaam,
    DocumentDatum: documentDatum,
    Status: NEED_ISSUE_STATUS.OPEN,
    Reden: reden,
    Opmerking: opmerking,
    GeboektDoor: '',
    GeboektOp: '',
    TypeMateriaal: MATERIAL_TYPE.NEED,
    LocatieVan: LOCATION.CENTRAL,
    LocatieNaar: getBusLocationCode(techniekerCode),
    InitiatiefRol: user.rol
  }]);

  writeAudit(
    'Behoefte-uitgifte aangemaakt',
    user.rol,
    user.naam || user.email || 'Magazijn',
    'BehoefteUitgifte',
    uitgifteId,
    {
      techniekerCode: techniekerCode,
      techniekerNaam: techniekerNaam,
      documentDatum: documentDatum,
      reden: reden
    }
  );

  return {
    success: true,
    uitgifteId,
    message: 'Behoefte-uitgifte aangemaakt.'
  };
}

function validateNeedIssueLines(cleanedLines) {
  if (!(cleanedLines || []).length) {
    throw new Error('Geen geldige uitgiftelijnen om te bewaren.');
  }

  cleanedLines.forEach(line => {
    const artikelCode = safeText(line.artikelCode);
    const aantal = safeNumber(line.aantal, 0);

    if (!artikelCode) {
      throw new Error('Artikelcode ontbreekt op een uitgiftelijn.');
    }

    if (!isBehoefteArticle(artikelCode)) {
      throw new Error('Behoefte-uitgifte is enkel toegestaan voor behoeftemateriaal. Artikel: ' + artikelCode);
    }

    if (aantal <= 0) {
      throw new Error('Aantal moet groter zijn dan 0 voor artikel ' + artikelCode);
    }
  });
}

function saveNeedIssueLines(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om uitgiftelijnen te bewaren.');
  }

  const uitgifteId = safeText(payload.uitgifteId);
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!uitgifteId) throw new Error('UitgifteID ontbreekt.');
  if (!lines.length) throw new Error('Geen uitgiftelijnen ontvangen.');

  const issue = getNeedIssueById(uitgifteId);
  if (!issue) throw new Error('Behoefte-uitgifte niet gevonden.');

  if (issue.status === NEED_ISSUE_STATUS.BOOKED) {
    throw new Error('Deze behoefte-uitgifte kan niet meer aangepast worden.');
  }

  const cleaned = lines
    .map(line => ({
      artikelCode: safeText(line.artikelCode),
      artikelOmschrijving: safeText(line.artikelOmschrijving),
      eenheid: safeText(line.eenheid),
      aantal: safeNumber(line.aantal, 0)
    }))
    .filter(line => line.artikelCode && line.artikelOmschrijving && line.aantal > 0);

  validateNeedIssueLines(cleaned);

  const lineObjects = cleaned.map(line => buildNeedIssueLineObject(uitgifteId, line));
  replaceNeedIssueLinesForIssue(uitgifteId, lineObjects);

  const issueType = determineNeedIssueMaterialType(cleaned);

  updateNeedIssueHeader(uitgifteId, {
    TypeMateriaal: issueType || MATERIAL_TYPE.NEED
  });

  writeAudit(
    'Behoefte-uitgiftelijnen opgeslagen',
    user.rol,
    user.naam || user.email || 'Magazijn',
    'BehoefteUitgifte',
    uitgifteId,
    {
      lijnen: lineObjects.length,
      typeMateriaal: issueType || MATERIAL_TYPE.NEED
    }
  );

  return {
    success: true,
    lines: lineObjects.length,
    message: 'Uitgiftelijnen opgeslagen.'
  };
}

function validateNeedIssueStockBeforeBooking(issue, issueLines) {
  const centralMap = buildCentralWarehouseMap();

  (issueLines || []).forEach(line => {
    const code = safeText(line.artikelCode);
    const aantal = safeNumber(line.aantal, 0);
    const currentCentral = centralMap[code] ? safeNumber(centralMap[code].voorraadCentraal, 0) : 0;

    if (aantal > currentCentral) {
      throw new Error(`Onvoldoende voorraad in centraal magazijn voor artikel ${code}. Beschikbaar: ${currentCentral}`);
    }
  });
}

function buildNeedIssueBookingMovements(issue, issueLines, actor) {
  const documentDatum = safeText(issue.documentDatumIso || issue.documentDatum);
  const busLocation = getBusLocationCode(issue.techniekerCode);

  return (issueLines || [])
    .filter(line => safeNumber(line.aantal, 0) > 0)
    .map(line => {
      const aantal = safeNumber(line.aantal, 0);

      return {
        datumDocument: documentDatum,
        typeMutatie: 'BehoefteUitgifte',
        bronId: issue.uitgifteId,
        typeMateriaal: MATERIAL_TYPE.NEED,
        artikelCode: line.artikelCode,
        artikelOmschrijving: line.artikelOmschrijving,
        eenheid: line.eenheid,
        aantalIn: aantal,
        aantalUit: aantal,
        nettoAantal: 0,
        locatieVan: LOCATION.CENTRAL,
        locatieNaar: busLocation,
        reden: safeText(issue.reden),
        opmerking: safeText(issue.opmerking),
        goedgekeurdDoor: actor,
        goedgekeurdOp: nowStamp()
      };
    });
}

function bookNeedIssue(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om een behoefte-uitgifte te boeken.');
  }

  const uitgifteId = safeText(payload.uitgifteId);
  const actor = safeText(payload.actor || user.naam || 'Magazijn');

  if (!uitgifteId) throw new Error('UitgifteID ontbreekt.');

  const issue = getNeedIssueById(uitgifteId);
  if (!issue) throw new Error('Behoefte-uitgifte niet gevonden.');

  const issueLines = getNeedIssueLinesByIssueId(uitgifteId);
  if (!issueLines.length) {
    throw new Error('Deze behoefte-uitgifte bevat nog geen actieve lijnen.');
  }

  if (issue.status === NEED_ISSUE_STATUS.BOOKED) {
    throw new Error('Deze behoefte-uitgifte is al geboekt.');
  }

  validateNeedIssueStockBeforeBooking(issue, issueLines);

  updateNeedIssueHeader(uitgifteId, {
    Status: NEED_ISSUE_STATUS.BOOKED,
    GeboektDoor: actor,
    GeboektOp: nowStamp()
  });

  const movementPayloads = buildNeedIssueBookingMovements(issue, issueLines, actor);

  replaceSourceMovements('BehoefteUitgifte', uitgifteId, movementPayloads);
  rebuildCentralWarehouseOverview();

  pushManagerNotification(
    'BehoefteUitgifteGeboekt',
    'Behoefte-uitgifte geboekt',
    `Behoefte-uitgifte ${uitgifteId} naar ${issue.techniekerNaam || issue.techniekerCode} is geboekt.`,
    'BehoefteUitgifte',
    uitgifteId
  );

  pushTechnicianNotification(
    issue.techniekerCode,
    issue.techniekerNaam,
    'BehoefteUitgifteGeboekt',
    'Behoeftestock aangevuld',
    `Je bus werd aangevuld via behoefte-uitgifte ${uitgifteId}.`,
    'BehoefteUitgifte',
    uitgifteId
  );

  writeAudit(
    'Behoefte-uitgifte geboekt',
    user.rol,
    actor,
    'BehoefteUitgifte',
    uitgifteId,
    {
      techniekerCode: issue.techniekerCode,
      techniekerNaam: issue.techniekerNaam,
      lijnen: movementPayloads.length
    }
  );

  return {
    success: true,
    message: 'Behoefte-uitgifte geboekt.'
  };
}