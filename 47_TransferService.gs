/* =========================================================
   47_TransferService.gs — transfers
   Centraal / Mobiel / Bus transfers
   ========================================================= */

/* ---------------------------------------------------------
   ID
   --------------------------------------------------------- */

function makeTransferId() {
  return makeStampedId('TRF');
}

/* ---------------------------------------------------------
   Mobile location helpers
   --------------------------------------------------------- */

function getMobileWarehouseLocationCode(mobileWarehouseCode) {
  const code = safeText(mobileWarehouseCode);
  return code ? `${LOCATION.MOBILE}:${code}` : LOCATION.MOBILE;
}

function parseMobileWarehouseLocation(location) {
  const text = safeText(location);
  if (!text) return '';
  if (text === LOCATION.MOBILE) return '';
  if (text.indexOf(`${LOCATION.MOBILE}:`) === 0) {
    return text.substring((`${LOCATION.MOBILE}:`).length);
  }
  return '';
}

/* ---------------------------------------------------------
   Notification helper voor mobiel magazijn
   --------------------------------------------------------- */

function pushMobileWarehouseNotification(type, title, message, bronType, bronId) {
  if (typeof pushNotification !== 'function') {
    return { success: true, skipped: true };
  }

  return pushNotification({
    rol: NOTIFICATION_ROLE.MOBILE_WAREHOUSE,
    ontvangerCode: 'MOBIEL_MAGAZIJN',
    ontvangerNaam: 'Mobiel magazijn',
    type,
    titel: title,
    bericht: message,
    bronType,
    bronId
  });
}

/* ---------------------------------------------------------
   Mappers
   --------------------------------------------------------- */

function mapTransfer(row) {
  return {
    transferId: safeText(row.TransferID),
    flowType: safeText(row.FlowType),
    vanLocatie: safeText(row.VanLocatie),
    naarLocatie: safeText(row.NaarLocatie),

    bronTechniekerCode: safeText(row.BronTechniekerCode),
    bronTechniekerNaam: safeText(row.BronTechniekerNaam),

    doelTechniekerCode: safeText(row.DoelTechniekerCode),
    doelTechniekerNaam: safeText(row.DoelTechniekerNaam),

    mobileWarehouseCode: safeText(row.MobileWarehouseCode),

    documentDatumIso: toIsoDate(row.DocumentDatum),
    documentDatum: toDisplayDate(row.DocumentDatum),

    status: safeText(row.Status),
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),

    aangemaaktDoor: safeText(row.AangemaaktDoor),
    aangemaaktOp: toDisplayDateTime(row.AangemaaktOp),

    klaargezetDoor: safeText(row.KlaargezetDoor),
    klaargezetOp: toDisplayDateTime(row.KlaargezetOp),

    meegegevenDoor: safeText(row.MeegegevenDoor),
    meegegevenOp: toDisplayDateTime(row.MeegegevenOp),

    ontvangenDoor: safeText(row.OntvangenDoor),
    ontvangenOp: toDisplayDateTime(row.OntvangenOp),

    geboektDoor: safeText(row.GeboektDoor),
    geboektOp: toDisplayDateTime(row.GeboektOp),

    goedgekeurdDoor: safeText(row.GoedgekeurdDoor),
    goedgekeurdOp: toDisplayDateTime(row.GoedgekeurdOp),
    managerOpmerking: safeText(row.ManagerOpmerking)
  };
}

function mapTransferLine(row) {
  return {
    transferId: safeText(row.TransferID),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    aantal: safeNumber(row.Aantal, 0),
    actief: safeText(row.Actief) !== 'Nee',
    typeMateriaal: safeText(row.TypeMateriaal)
  };
}

/* ---------------------------------------------------------
   Readers
   --------------------------------------------------------- */

function getAllTransfers() {
  return readObjectsSafe(TABS.TRANSFERS)
    .map(mapTransfer)
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.transferId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.transferId)}`
      )
    );
}

function getAllTransferLines() {
  return readObjectsSafe(TABS.TRANSFER_LINES)
    .map(mapTransferLine)
    .filter(line => line.actief);
}

function getTransferById(transferId) {
  const id = safeText(transferId);
  if (!id) return null;
  return getAllTransfers().find(item => item.transferId === id) || null;
}

function getTransferLinesById(transferId) {
  const id = safeText(transferId);
  if (!id) return [];
  return getAllTransferLines().filter(line => line.transferId === id);
}

function buildTransfersWithLines(transfers, lines) {
  const linesByTransfer = {};

  (lines || []).forEach(line => {
    const id = safeText(line.transferId);
    if (!id) return;
    if (!linesByTransfer[id]) linesByTransfer[id] = [];
    linesByTransfer[id].push(line);
  });

  return (transfers || []).map(transfer => ({
    ...transfer,
    lines: linesByTransfer[transfer.transferId] || [],
    lineCount: (linesByTransfer[transfer.transferId] || []).length
  }));
}

function getTransferData(sessionId) {
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om transferdata te bekijken.');
  }

  return {
    transfers: buildTransfersWithLines(getAllTransfers(), getAllTransferLines()),
    centralWarehouse: getCentralWarehouseOverview(),
    busStock: buildBusStockRows(),
    technicians: getActiveTechnicians().map(t => ({ code: t.code, naam: t.naam }))
  };
}

/* ---------------------------------------------------------
   Stock helpers generiek per locatie
   --------------------------------------------------------- */

function buildLocationStockRows(locationCode) {
  const targetLocation = safeText(locationCode);
  if (!targetLocation) return [];

  const articles = readObjectsSafe(TABS.SUPPLIER_ARTICLES).map(mapSupplierArticle);
  const articleMap = {};
  articles.forEach(item => {
    articleMap[item.artikelCode] = item;
  });

  const moves = readObjectsSafe(TABS.WAREHOUSE_MOVEMENTS).map(mapWarehouseMovement);
  const grouped = {};

  moves.forEach(move => {
    const code = safeText(move.artikelCode);
    if (!code) return;

    const qtyIn = safeNumber(move.aantalIn, 0);
    const qtyOut = safeNumber(move.aantalUit, 0);
    let delta = 0;

    if (safeText(move.locatieNaar) === targetLocation) {
      delta += qtyIn || Math.abs(safeNumber(move.nettoAantal, 0));
    }

    if (safeText(move.locatieVan) === targetLocation) {
      delta -= qtyOut || Math.abs(safeNumber(move.nettoAantal, 0));
    }

    if (!delta) return;

    if (!grouped[code]) {
      grouped[code] = {
        locatie: targetLocation,
        artikelCode: code,
        artikelOmschrijving: safeText(move.artikelOmschrijving) || safeText(articleMap[code] && articleMap[code].artikelOmschrijving),
        eenheid: safeText(move.eenheid) || safeText(articleMap[code] && articleMap[code].eenheid),
        voorraad: 0,
        laatsteMutatie: safeText(move.datumBoeking),
        laatsteMutatieRaw: safeText(move.datumBoekingRaw)
      };
    }

    grouped[code].voorraad += delta;

    if (safeText(move.datumBoekingRaw) > safeText(grouped[code].laatsteMutatieRaw)) {
      grouped[code].laatsteMutatie = safeText(move.datumBoeking);
      grouped[code].laatsteMutatieRaw = safeText(move.datumBoekingRaw);
    }
  });

  return Object.keys(grouped)
    .map(key => grouped[key])
    .filter(item => safeNumber(item.voorraad, 0) !== 0)
    .sort((a, b) => safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)));
}

function buildLocationStockMap(locationCode) {
  const map = {};
  buildLocationStockRows(locationCode).forEach(item => {
    map[item.artikelCode] = item;
  });
  return map;
}

/* ---------------------------------------------------------
   Artikel helpers
   --------------------------------------------------------- */

function buildTransferArticleMasterMap() {
  const map = {};
  readObjectsSafe(TABS.SUPPLIER_ARTICLES)
    .map(mapSupplierArticle)
    .filter(item => item.actief !== false)
    .forEach(item => {
      map[item.artikelCode] = item;
    });
  return map;
}

/* ---------------------------------------------------------
   Flow resolvers
   --------------------------------------------------------- */

function resolveTransferLocations(flowType, payload) {
  const flow = safeText(flowType);

  const bronTechniekerCode = safeText(payload.bronTechniekerCode);
  const doelTechniekerCode = safeText(payload.doelTechniekerCode);
  const mobileWarehouseCode = safeText(payload.mobileWarehouseCode);

  switch (flow) {
    case TRANSFER_FLOW.CENTRAL_TO_BUS:
      if (!doelTechniekerCode) throw new Error('DoelTechniekerCode is verplicht.');
      return {
        vanLocatie: LOCATION.CENTRAL,
        naarLocatie: getBusLocationCode(doelTechniekerCode)
      };

    case TRANSFER_FLOW.MOBILE_TO_BUS:
      if (!doelTechniekerCode) throw new Error('DoelTechniekerCode is verplicht.');
      return {
        vanLocatie: getMobileWarehouseLocationCode(mobileWarehouseCode),
        naarLocatie: getBusLocationCode(doelTechniekerCode)
      };

    case TRANSFER_FLOW.BUS_TO_MOBILE:
      if (!bronTechniekerCode) throw new Error('BronTechniekerCode is verplicht.');
      return {
        vanLocatie: getBusLocationCode(bronTechniekerCode),
        naarLocatie: getMobileWarehouseLocationCode(mobileWarehouseCode)
      };

    case TRANSFER_FLOW.BUS_TO_CENTRAL:
      if (!bronTechniekerCode) throw new Error('BronTechniekerCode is verplicht.');
      return {
        vanLocatie: getBusLocationCode(bronTechniekerCode),
        naarLocatie: LOCATION.CENTRAL
      };

    case TRANSFER_FLOW.MOBILE_TO_CENTRAL:
      return {
        vanLocatie: getMobileWarehouseLocationCode(mobileWarehouseCode),
        naarLocatie: LOCATION.CENTRAL
      };

    default:
      throw new Error(`FlowType niet ondersteund in deze versie: ${flow}`);
  }
}

function buildTransferHeaderObject(transferId, payload) {
  const flowType = safeText(payload.flowType);
  const locations = resolveTransferLocations(flowType, payload);

  const bronTechniekerCode = safeText(payload.bronTechniekerCode);
  const doelTechniekerCode = safeText(payload.doelTechniekerCode);

  return {
    TransferID: transferId,
    FlowType: flowType,
    VanLocatie: locations.vanLocatie,
    NaarLocatie: locations.naarLocatie,

    BronTechniekerCode: bronTechniekerCode,
    BronTechniekerNaam: bronTechniekerCode ? getTechnicianNameByCode(bronTechniekerCode) : '',

    DoelTechniekerCode: doelTechniekerCode,
    DoelTechniekerNaam: doelTechniekerCode ? getTechnicianNameByCode(doelTechniekerCode) : '',

    MobileWarehouseCode: safeText(payload.mobileWarehouseCode),

    DocumentDatum: safeText(payload.documentDatum),
    Status: TRANSFER_STATUS.OPEN,
    Reden: safeText(payload.reden),
    Opmerking: safeText(payload.opmerking),

    AangemaaktDoor: safeText(payload.actor),
    AangemaaktOp: nowStamp(),

    KlaargezetDoor: '',
    KlaargezetOp: '',
    MeegegevenDoor: '',
    MeegegevenOp: '',
    OntvangenDoor: '',
    OntvangenOp: '',
    GeboektDoor: '',
    GeboektOp: '',
    GoedgekeurdDoor: '',
    GoedgekeurdOp: '',
    ManagerOpmerking: ''
  };
}

/* ---------------------------------------------------------
   Transfer line helpers
   --------------------------------------------------------- */

function buildTransferLineObject(transferId, line) {
  const article = buildTransferArticleMasterMap()[safeText(line.artikelCode)] || null;

  return {
    TransferID: safeText(transferId),
    ArtikelCode: safeText(line.artikelCode),
    ArtikelOmschrijving: safeText(line.artikelOmschrijving) || safeText(article && article.artikelOmschrijving),
    Eenheid: safeText(line.eenheid) || safeText(article && article.eenheid),
    Aantal: safeNumber(line.aantal, 0),
    Actief: 'Ja',
    TypeMateriaal: safeText(line.typeMateriaal) || determineMaterialTypeFromArticle(line.artikelCode)
  };
}

function replaceTransferLinesForTransfer(transferId, lineObjects) {
  const id = safeText(transferId);

  const sheet = getSheetOrThrow(TABS.TRANSFER_LINES);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.TRANSFER_LINES);
  const col = getColMap(headers);
  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => safeText(row[col['TransferID']]) !== id);
  const newRows = (lineObjects || []).map(obj => buildRowFromHeaders(headers, obj));

  writeFullTable(TABS.TRANSFER_LINES, headers, keptRows.concat(newRows));

  return {
    success: true,
    lines: newRows.length
  };
}

function validateTransferLinePayloads(lines) {
  if (!(lines || []).length) throw new Error('Geen transferlijnen ontvangen.');

  const articleMap = buildTransferArticleMasterMap();

  (lines || []).forEach(line => {
    const code = safeText(line.artikelCode);
    const qty = safeNumber(line.aantal, 0);

    if (!code) throw new Error('ArtikelCode ontbreekt op een transferlijn.');
    if (!articleMap[code]) throw new Error(`ArtikelCode niet gevonden: ${code}`);
    if (qty <= 0) throw new Error(`Aantal moet groter zijn dan 0 voor artikel ${code}`);
  });

  return { success: true };
}

/* ---------------------------------------------------------
   Update header helper
   --------------------------------------------------------- */

function updateTransferHeader(transferId, updates) {
  const id = safeText(transferId);

  const sheet = getSheetOrThrow(TABS.TRANSFERS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab Transfers is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['TransferID']]) !== id) continue;

    Object.keys(updates || {}).forEach(field => {
      if (col[field] !== undefined) {
        values[i][col[field]] = updates[field];
      }
    });

    updated = true;
    break;
  }

  if (!updated) throw new Error('Transfer niet gevonden.');

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return { success: true };
}

/* ---------------------------------------------------------
   Permissions
   --------------------------------------------------------- */

function canUserManageTransfer(user) {
  return roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER]);
}

function canUserReceiveTransfer(transfer, user) {
  if (!transfer || !user) return false;
  if (user.rol === ROLE.ADMIN || user.rol === ROLE.MANAGER) return true;

  const target = safeText(transfer.naarLocatie);

  if (target === LOCATION.CENTRAL) {
    return user.rol === ROLE.WAREHOUSE;
  }

  if (target === LOCATION.MOBILE || target.indexOf(`${LOCATION.MOBILE}:`) === 0) {
    return user.rol === ROLE.MOBILE_WAREHOUSE;
  }

  const targetBusTech = parseBusLocation(target);
  if (targetBusTech) {
    return user.rol === ROLE.TECHNICIAN && normalizeRef(user.techniekerCode) === normalizeRef(targetBusTech);
  }

  return false;
}

/* ---------------------------------------------------------
   Create / save
   --------------------------------------------------------- */

function createTransfer(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!canUserManageTransfer(user)) {
    throw new Error('Geen rechten om een transfer aan te maken.');
  }

  const flowType = safeText(payload.flowType);
  const documentDatum = safeText(payload.documentDatum);
  const reden = safeText(payload.reden);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');

  if (!flowType) throw new Error('FlowType is verplicht.');
  if (!documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!reden) throw new Error('Reden is verplicht.');

  const transferId = makeTransferId();
  const headerObject = buildTransferHeaderObject(transferId, {
    ...payload,
    actor: actor
  });

  appendObjects(TABS.TRANSFERS, [headerObject]);

  writeAudit(
    'Transfer aangemaakt',
    user.rol,
    actor,
    'Transfer',
    transferId,
    {
      flowType: flowType,
      vanLocatie: headerObject.VanLocatie,
      naarLocatie: headerObject.NaarLocatie,
      reden: reden
    }
  );

  return {
    success: true,
    transferId,
    message: 'Transfer aangemaakt.'
  };
}

function saveTransferLines(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!canUserManageTransfer(user)) {
    throw new Error('Geen rechten om transferlijnen op te slaan.');
  }

  const transferId = safeText(payload.transferId);
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!transferId) throw new Error('TransferID ontbreekt.');
  if (!lines.length) throw new Error('Geen transferlijnen ontvangen.');

  const transfer = getTransferById(transferId);
  if (!transfer) throw new Error('Transfer niet gevonden.');

  if ([TRANSFER_STATUS.DISPATCHED, TRANSFER_STATUS.RECEIVED, TRANSFER_STATUS.BOOKED, TRANSFER_STATUS.APPROVED, TRANSFER_STATUS.CANCELLED].includes(transfer.status)) {
    throw new Error('Deze transfer kan niet meer aangepast worden.');
  }

  validateTransferLinePayloads(lines);

  const cleanedObjects = lines
    .map(line => buildTransferLineObject(transferId, line))
    .filter(line => line.ArtikelCode && safeNumber(line.Aantal, 0) > 0);

  if (!cleanedObjects.length) {
    throw new Error('Geen geldige transferlijnen om te bewaren.');
  }

  replaceTransferLinesForTransfer(transferId, cleanedObjects);

  writeAudit(
    'Transferlijnen opgeslagen',
    user.rol,
    user.naam || user.email,
    'Transfer',
    transferId,
    {
      lijnen: cleanedObjects.length
    }
  );

  return {
    success: true,
    lines: cleanedObjects.length,
    message: 'Transferlijnen opgeslagen.'
  };
}

/* ---------------------------------------------------------
   Statusflow
   --------------------------------------------------------- */

function setTransferReady(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!canUserManageTransfer(user)) {
    throw new Error('Geen rechten om transfer klaar te zetten.');
  }

  const transferId = safeText(payload.transferId);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');

  if (!transferId) throw new Error('TransferID ontbreekt.');

  const transfer = getTransferById(transferId);
  if (!transfer) throw new Error('Transfer niet gevonden.');
  if (![TRANSFER_STATUS.OPEN, TRANSFER_STATUS.READY].includes(transfer.status)) {
    throw new Error('Transfer kan niet op Klaargezet gezet worden.');
  }

  const lines = getTransferLinesById(transferId);
  if (!lines.length) throw new Error('Deze transfer bevat nog geen actieve lijnen.');

  updateTransferHeader(transferId, {
    Status: TRANSFER_STATUS.READY,
    KlaargezetDoor: actor,
    KlaargezetOp: nowStamp()
  });

  return {
    success: true,
    message: 'Transfer op Klaargezet gezet.'
  };
}

function dispatchTransfer(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!canUserManageTransfer(user)) {
    throw new Error('Geen rechten om transfer mee te geven.');
  }

  const transferId = safeText(payload.transferId);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');

  if (!transferId) throw new Error('TransferID ontbreekt.');

  const transfer = getTransferById(transferId);
  if (!transfer) throw new Error('Transfer niet gevonden.');
  if (![TRANSFER_STATUS.READY, TRANSFER_STATUS.DISPATCHED].includes(transfer.status)) {
    throw new Error('Transfer kan enkel meegegeven worden vanuit Klaargezet.');
  }

  updateTransferHeader(transferId, {
    Status: TRANSFER_STATUS.DISPATCHED,
    MeegegevenDoor: actor,
    MeegegevenOp: nowStamp()
  });

  if (transfer.doelTechniekerCode && safeText(transfer.naarLocatie).indexOf('Bus:') === 0) {
    pushTechnicianNotification(
      transfer.doelTechniekerCode,
      transfer.doelTechniekerNaam,
      'TransferMeegegeven',
      'Transfer onderweg',
      `Transfer ${transfer.transferId} is meegegeven naar jouw bus.`,
      'Transfer',
      transfer.transferId
    );
  }

  if (safeText(transfer.naarLocatie) === LOCATION.CENTRAL) {
    pushWarehouseNotification(
      'TransferOnderweg',
      'Transfer onderweg naar centraal magazijn',
      `Transfer ${transfer.transferId} is meegegeven naar centraal magazijn.`,
      'Transfer',
      transfer.transferId
    );
  }

  if (safeText(transfer.naarLocatie) === LOCATION.MOBILE || safeText(transfer.naarLocatie).indexOf(`${LOCATION.MOBILE}:`) === 0) {
    pushMobileWarehouseNotification(
      'TransferOnderweg',
      'Transfer onderweg naar mobiel magazijn',
      `Transfer ${transfer.transferId} is meegegeven naar mobiel magazijn.`,
      'Transfer',
      transfer.transferId
    );
  }

  writeAudit(
    'Transfer meegegeven',
    user.rol,
    actor,
    'Transfer',
    transferId,
    {
      flowType: transfer.flowType,
      vanLocatie: transfer.vanLocatie,
      naarLocatie: transfer.naarLocatie
    }
  );

  return {
    success: true,
    message: 'Transfer op Meegegeven gezet.'
  };
}

function receiveTransfer(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  const transferId = safeText(payload.transferId);
  const actor = safeText(payload.actor || user.naam || user.email || 'Ontvanger');

  if (!transferId) throw new Error('TransferID ontbreekt.');

  const transfer = getTransferById(transferId);
  if (!transfer) throw new Error('Transfer niet gevonden.');

  if (!canUserReceiveTransfer(transfer, user)) {
    throw new Error('Geen rechten om deze transfer te ontvangen.');
  }

  if (![TRANSFER_STATUS.DISPATCHED, TRANSFER_STATUS.RECEIVED].includes(transfer.status)) {
    throw new Error('Transfer kan enkel ontvangen worden vanuit Meegegeven.');
  }

  updateTransferHeader(transferId, {
    Status: TRANSFER_STATUS.RECEIVED,
    OntvangenDoor: actor,
    OntvangenOp: nowStamp()
  });

  pushManagerNotification(
    'TransferOntvangen',
    'Transfer ontvangen',
    `Transfer ${transfer.transferId} is ontvangen en klaar om geboekt te worden.`,
    'Transfer',
    transfer.transferId
  );

  writeAudit(
    'Transfer ontvangen',
    user.rol,
    actor,
    'Transfer',
    transferId,
    {
      flowType: transfer.flowType,
      naarLocatie: transfer.naarLocatie
    }
  );

  return {
    success: true,
    message: 'Transfer ontvangen.'
  };
}

/* ---------------------------------------------------------
   Booking
   --------------------------------------------------------- */

function validateTransferSourceStock(transfer, lines) {
  const sourceLocation = safeText(transfer.vanLocatie);
  const sourceStockMap = buildLocationStockMap(sourceLocation);

  (lines || []).forEach(line => {
    const code = safeText(line.artikelCode);
    const qty = safeNumber(line.aantal, 0);
    const current = sourceStockMap[code] ? safeNumber(sourceStockMap[code].voorraad, 0) : 0;

    if (qty > current) {
      throw new Error(`Onvoldoende voorraad op bronlocatie voor artikel ${code}. Beschikbaar: ${current}`);
    }
  });

  return { success: true };
}

function getMovementTypeForTransferFlow(flowType) {
  switch (safeText(flowType)) {
    case TRANSFER_FLOW.CENTRAL_TO_BUS:
      return MOVEMENT_TYPE.TRANSFER_CENTRAL_TO_BUS;
    case TRANSFER_FLOW.MOBILE_TO_BUS:
      return MOVEMENT_TYPE.TRANSFER_MOBILE_TO_BUS;
    case TRANSFER_FLOW.BUS_TO_MOBILE:
      return MOVEMENT_TYPE.TRANSFER_BUS_TO_MOBILE;
    case TRANSFER_FLOW.BUS_TO_CENTRAL:
      return MOVEMENT_TYPE.TRANSFER_BUS_TO_CENTRAL;
    case TRANSFER_FLOW.MOBILE_TO_CENTRAL:
      return MOVEMENT_TYPE.TRANSFER_MOBILE_TO_CENTRAL;
    default:
      throw new Error(`Geen movement type mapping voor flow ${flowType}`);
  }
}

function buildTransferMovementPayloads(transfer, lines, actor) {
  const typeMutatie = getMovementTypeForTransferFlow(transfer.flowType);
  const documentDatum = safeText(transfer.documentDatumIso || transfer.documentDatum);

  return (lines || [])
    .filter(line => safeNumber(line.aantal, 0) > 0)
    .map(line => ({
      datumDocument: documentDatum,
      typeMutatie: typeMutatie,
      typeMateriaal: safeText(line.typeMateriaal) || determineMaterialTypeFromArticle(line.artikelCode),
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      eenheid: line.eenheid,
      aantalIn: safeNumber(line.aantal, 0),
      aantalUit: safeNumber(line.aantal, 0),
      nettoAantal: 0,
      locatieVan: safeText(transfer.vanLocatie),
      locatieNaar: safeText(transfer.naarLocatie),
      reden: safeText(transfer.reden),
      opmerking: safeText(transfer.opmerking),
      goedgekeurdDoor: actor,
      goedgekeurdOp: nowStamp()
    }));
}

function replaceTransferBookingMovements(transferId, movementPayloads) {
  const id = safeText(transferId);

  const sheet = getSheetOrThrow(TABS.WAREHOUSE_MOVEMENTS);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.WAREHOUSE_MOVEMENTS);
  const col = getColMap(headers);
  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => safeText(row[col['BronID']]) !== id);
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

function bookTransfer(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!canUserManageTransfer(user)) {
    throw new Error('Geen rechten om transfer te boeken.');
  }

  const transferId = safeText(payload.transferId);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');

  if (!transferId) throw new Error('TransferID ontbreekt.');

  const transfer = getTransferById(transferId);
  if (!transfer) throw new Error('Transfer niet gevonden.');

  if (![TRANSFER_STATUS.RECEIVED, TRANSFER_STATUS.BOOKED].includes(transfer.status)) {
    throw new Error('Transfer kan enkel geboekt worden vanuit Ontvangen.');
  }

  const lines = getTransferLinesById(transferId);
  if (!lines.length) throw new Error('Deze transfer bevat nog geen actieve lijnen.');

  validateTransferSourceStock(transfer, lines);

  const movementPayloads = buildTransferMovementPayloads(transfer, lines, actor);
  replaceTransferBookingMovements(transferId, movementPayloads);

  updateTransferHeader(transferId, {
    Status: TRANSFER_STATUS.BOOKED,
    GeboektDoor: actor,
    GeboektOp: nowStamp()
  });

  rebuildCentralWarehouseOverview();

  pushManagerNotification(
    'TransferGeboekt',
    'Transfer geboekt',
    `Transfer ${transfer.transferId} is geboekt.`,
    'Transfer',
    transfer.transferId
  );

  writeAudit(
    'Transfer geboekt',
    user.rol,
    actor,
    'Transfer',
    transferId,
    {
      flowType: transfer.flowType,
      vanLocatie: transfer.vanLocatie,
      naarLocatie: transfer.naarLocatie,
      lijnen: movementPayloads.length
    }
  );

  return {
    success: true,
    message: 'Transfer geboekt.'
  };
}

function approveTransfer(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertManagerAccess(sessionId);

  const transferId = safeText(payload.transferId);
  const actor = safeText(payload.actor || user.naam || 'Manager');
  const note = safeText(payload.note);

  if (!transferId) throw new Error('TransferID ontbreekt.');

  const transfer = getTransferById(transferId);
  if (!transfer) throw new Error('Transfer niet gevonden.');

  if (![TRANSFER_STATUS.BOOKED, TRANSFER_STATUS.APPROVED].includes(transfer.status)) {
    throw new Error('Transfer kan enkel goedgekeurd worden na boeking.');
  }

  updateTransferHeader(transferId, {
    Status: TRANSFER_STATUS.APPROVED,
    GoedgekeurdDoor: actor,
    GoedgekeurdOp: nowStamp(),
    ManagerOpmerking: note
  });

  writeAudit(
    'Transfer goedgekeurd',
    user.rol,
    actor,
    'Transfer',
    transferId,
    {
      note: note
    }
  );

  return {
    success: true,
    message: 'Transfer goedgekeurd.'
  };
}