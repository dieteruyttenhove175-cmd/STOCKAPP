/* =========================================================
   53_MobileWarehouseRequestService.gs
   aanvragen technieker -> mobiel magazijn
   ========================================================= */

const MW_REQUEST_STATUS = {
  OPEN: 'Open',
  SUBMITTED: 'Ingediend',
  APPROVED: 'Goedgekeurd',
  BOOKED: 'Geboekt',
  REJECTED: 'Afgekeurd',
  CANCELLED: 'Geannuleerd'
};

/* =========================================================
   53.1 — lokale helpers / fallback mappers
   ========================================================= */

function makeMobileRequestId() {
  return 'MWR-' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
}

function mapMobileWarehouseSafe(row) {
  if (typeof mapMobileWarehouse === 'function') return mapMobileWarehouse(row);
  return {
    code: safeText(row.Code),
    naam: safeText(row.Naam),
    type: safeText(row.Type),
    active: row.Actief === undefined ? true : isTrue(row.Actief),
    managerNaam: safeText(row.ManagerNaam),
    opmerking: safeText(row.Opmerking)
  };
}

function mapMobileRequestSafe(row) {
  if (typeof mapMobileWarehouseRequest === 'function') return mapMobileWarehouseRequest(row);
  return {
    aanvraagId: safeText(row.AanvraagID),
    mobielMagazijnCode: safeText(row.MobielMagazijnCode),
    techniekerCode: safeText(row.TechniekerCode),
    techniekerNaam: safeText(row.TechniekerNaam),
    documentDatumIso: toIsoDate(row.DocumentDatum),
    documentDatum: toDisplayDate(row.DocumentDatum),
    status: safeText(row.Status),
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),
    aangemaaktDoor: safeText(row.AangemaaktDoor),
    aangemaaktOp: toDisplayDateTime(row.AangemaaktOp),
    ingediendDoor: safeText(row.IngediendDoor),
    ingediendOp: toDisplayDateTime(row.IngediendOp),
    goedgekeurdDoor: safeText(row.GoedgekeurdDoor),
    goedgekeurdOp: toDisplayDateTime(row.GoedgekeurdOp),
    geboektDoor: safeText(row.GeboektDoor),
    geboektOp: toDisplayDateTime(row.GeboektOp),
    afgekeurdDoor: safeText(row.AfgekeurdDoor),
    afgekeurdOp: toDisplayDateTime(row.AfgekeurdOp),
    afkeurReden: safeText(row.AfkeurReden),
    transferId: safeText(row.TransferID)
  };
}

function mapMobileRequestLineSafe(row) {
  if (typeof mapMobileWarehouseRequestLine === 'function') return mapMobileWarehouseRequestLine(row);
  return {
    aanvraagId: safeText(row.AanvraagID),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    aantal: safeNumber(row.Aantal, 0),
    actief: safeText(row.Actief) !== 'Nee'
  };
}

function getAllMobileWarehouses() {
  return readObjectsSafe(TABS.MOBILE_WAREHOUSES)
    .map(mapMobileWarehouseSafe)
    .filter(x => x.active)
    .sort((a, b) => safeText(a.naam).localeCompare(safeText(b.naam)));
}

function getMobileWarehouseByCode(code) {
  return getAllMobileWarehouses().find(x => normalizeRef(x.code) === normalizeRef(code)) || null;
}

function assertMobileWarehouseRequestTabsExist() {
  getSheetOrThrow(TABS.MOBILE_REQUESTS);
  getSheetOrThrow(TABS.MOBILE_REQUEST_LINES);
}

function assertTechnicianOwnRequestAccess(request, sessionId) {
  const user = requireLoggedInUser(sessionId);

  if (user.rol === ROLE.ADMIN) return user;

  if (
    user.rol === ROLE.TECHNICIAN &&
    normalizeRef(user.techniekerCode) === normalizeRef(request.techniekerCode)
  ) {
    return user;
  }

  throw new Error('Geen rechten voor deze aanvraag.');
}

function resolveMobileWarehouseRequestAccess(request, sessionId) {
  const user = assertRoleAllowed([ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER], sessionId);

  if (user.rol === ROLE.MOBILE_WAREHOUSE) {
    const ownCode = safeText(user.mobileWarehouseCode);
    if (!ownCode) throw new Error('Gebruiker is niet gekoppeld aan een mobiel magazijn.');
    if (normalizeRef(ownCode) !== normalizeRef(request.mobielMagazijnCode)) {
      throw new Error('Geen rechten voor deze aanvraag van ander mobiel magazijn.');
    }
  }

  return user;
}

function getMobileRequestLinesMap() {
  const map = {};
  readObjectsSafe(TABS.MOBILE_REQUEST_LINES)
    .map(mapMobileRequestLineSafe)
    .filter(x => x.actief)
    .forEach(line => {
      if (!map[line.aanvraagId]) map[line.aanvraagId] = [];
      map[line.aanvraagId].push(line);
    });
  return map;
}

function buildMobileWarehouseRequestRows(mobielMagazijnCode) {
  assertMobileWarehouseRequestTabsExist();

  const linesMap = getMobileRequestLinesMap();

  let rows = readObjectsSafe(TABS.MOBILE_REQUESTS)
    .map(mapMobileRequestSafe)
    .map(request => {
      const lines = linesMap[request.aanvraagId] || [];
      return {
        ...request,
        lines,
        lineCount: lines.length
      };
    });

  if (mobielMagazijnCode) {
    rows = rows.filter(x => normalizeRef(x.mobielMagazijnCode) === normalizeRef(mobielMagazijnCode));
  }

  return rows.sort((a, b) =>
    `${safeText(b.documentDatumIso)} ${safeText(b.aanvraagId)}`.localeCompare(
      `${safeText(a.documentDatumIso)} ${safeText(a.aanvraagId)}`
    )
  );
}

function buildMobileWarehouseRequestQueue(mobielMagazijnCode) {
  return buildMobileWarehouseRequestRows(mobielMagazijnCode)
    .filter(x => [MW_REQUEST_STATUS.SUBMITTED, MW_REQUEST_STATUS.APPROVED].includes(x.status));
}

function getMobileWarehouseRequestsForTechnician(techRef, sessionId) {
  const access = assertTechnicianAccessToRef(techRef, sessionId);
  return buildMobileWarehouseRequestRows('')
    .filter(x => normalizeRef(x.techniekerCode) === normalizeRef(access.technician.code));
}

/* =========================================================
   53.2 — create / save / submit door technieker
   ========================================================= */

function createMobileWarehouseRequest(payload) {
  payload = payload || {};

  const sessionId = getPayloadSessionId(payload);
  const techRef = safeText(payload.techRef || payload.techniekerCode);
  const access = assertTechnicianAccessToRef(techRef, sessionId);

  const mobielMagazijnCode = safeText(payload.mobielMagazijnCode);
  const documentDatum = safeText(payload.documentDatum);
  const reden = safeText(payload.reden);
  const opmerking = safeText(payload.opmerking);

  if (!mobielMagazijnCode) throw new Error('Mobiel magazijn is verplicht.');
  if (!documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!reden) throw new Error('Reden is verplicht.');

  const warehouse = getMobileWarehouseByCode(mobielMagazijnCode);
  if (!warehouse) throw new Error('Mobiel magazijn niet gevonden of niet actief.');

  const aanvraagId = makeMobileRequestId();

  appendObjects(TABS.MOBILE_REQUESTS, [{
    AanvraagID: aanvraagId,
    MobielMagazijnCode: warehouse.code,
    TechniekerCode: access.technician.code,
    TechniekerNaam: access.technician.naam,
    DocumentDatum: documentDatum,
    Status: MW_REQUEST_STATUS.OPEN,
    Reden: reden,
    Opmerking: opmerking,
    AangemaaktDoor: access.user.naam || access.technician.naam || access.technician.code,
    AangemaaktOp: nowStamp(),
    IngediendDoor: '',
    IngediendOp: '',
    GoedgekeurdDoor: '',
    GoedgekeurdOp: '',
    GeboektDoor: '',
    GeboektOp: '',
    AfgekeurdDoor: '',
    AfgekeurdOp: '',
    AfkeurReden: '',
    TransferID: ''
  }]);

  writeAudit('Mobiel magazijn aanvraag aangemaakt', access.user.rol, access.user.naam || access.user.email, 'MobielMagazijnAanvraag', aanvraagId, {
    techniekerCode: access.technician.code,
    mobielMagazijnCode: warehouse.code,
    reden: reden
  });

  return {
    success: true,
    aanvraagId,
    message: 'Aanvraag naar mobiel magazijn aangemaakt.'
  };
}

function saveMobileWarehouseRequestLines(payload) {
  payload = payload || {};

  const sessionId = getPayloadSessionId(payload);
  const aanvraagId = safeText(payload.aanvraagId);
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!aanvraagId) throw new Error('AanvraagID ontbreekt.');
  if (!lines.length) throw new Error('Geen aanvraaglijnen ontvangen.');

  const request = buildMobileWarehouseRequestRows('').find(x => x.aanvraagId === aanvraagId);
  if (!request) throw new Error('Aanvraag niet gevonden.');

  assertTechnicianOwnRequestAccess(request, sessionId);

  if (![MW_REQUEST_STATUS.OPEN].includes(request.status)) {
    throw new Error('Deze aanvraag kan niet meer aangepast worden.');
  }

  const cleaned = lines
    .map(line => ({
      artikelCode: safeText(line.artikelCode),
      artikelOmschrijving: safeText(line.artikelOmschrijving),
      eenheid: safeText(line.eenheid),
      aantal: safeNumber(line.aantal, 0),
      actief: 'Ja'
    }))
    .filter(line => line.artikelCode && line.artikelOmschrijving && line.aantal > 0);

  if (!cleaned.length) throw new Error('Geen geldige aanvraaglijnen om te bewaren.');

  const sheet = getSheetOrThrow(TABS.MOBILE_REQUEST_LINES);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => String(h || '').trim()) : [];
  const col = getColMap(headers);
  const existing = values.length > 1 ? values.slice(1) : [];

  const kept = existing.filter(row => safeText(row[col['AanvraagID']]) !== aanvraagId);
  const newRows = cleaned.map(line => buildRowFromHeaders(headers, {
    AanvraagID: aanvraagId,
    ArtikelCode: line.artikelCode,
    ArtikelOmschrijving: line.artikelOmschrijving,
    Eenheid: line.eenheid,
    Aantal: line.aantal,
    Actief: line.actief
  }));

  writeFullTable(TABS.MOBILE_REQUEST_LINES, headers, kept.concat(newRows));

  writeAudit('Mobiel magazijn aanvraaglijnen opgeslagen', request.status || 'Technieker', request.techniekerNaam || request.techniekerCode, 'MobielMagazijnAanvraag', aanvraagId, {
    lijnen: newRows.length
  });

  return { success: true, lines: newRows.length, message: 'Aanvraaglijnen opgeslagen.' };
}

function submitMobileWarehouseRequest(payload) {
  payload = payload || {};

  const sessionId = getPayloadSessionId(payload);
  const aanvraagId = safeText(payload.aanvraagId);
  const actor = safeText(payload.actor);

  if (!aanvraagId) throw new Error('AanvraagID ontbreekt.');

  const request = buildMobileWarehouseRequestRows('').find(x => x.aanvraagId === aanvraagId);
  if (!request) throw new Error('Aanvraag niet gevonden.');

  const user = assertTechnicianOwnRequestAccess(request, sessionId);

  if (request.status !== MW_REQUEST_STATUS.OPEN) {
    throw new Error('Deze aanvraag kan niet meer ingediend worden.');
  }

  if (!(request.lines || []).length) {
    throw new Error('Deze aanvraag bevat nog geen actieve lijnen.');
  }

  const sheet = getSheetOrThrow(TABS.MOBILE_REQUESTS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['AanvraagID']]) !== aanvraagId) continue;

    values[i][col['Status']] = MW_REQUEST_STATUS.SUBMITTED;
    values[i][col['IngediendDoor']] = actor || user.naam || user.techniekerCode || user.email;
    values[i][col['IngediendOp']] = nowStamp();
    updated = true;
    break;
  }

  if (!updated) throw new Error('Aanvraag niet gevonden.');

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  pushNotification({
    rol: NOTIFICATION_ROLE.MOBILE_WAREHOUSE,
    ontvangerCode: request.mobielMagazijnCode,
    ontvangerNaam: request.mobielMagazijnCode,
    type: 'MobielMagazijnAanvraag',
    titel: 'Nieuwe aanvraag naar mobiel magazijn',
    bericht: `${request.techniekerNaam || request.techniekerCode} diende een aanvraag in naar mobiel magazijn ${request.mobielMagazijnCode}.`,
    bronType: 'MobielMagazijnAanvraag',
    bronId: aanvraagId
  });

  writeAudit('Mobiel magazijn aanvraag ingediend', user.rol, user.naam || user.email, 'MobielMagazijnAanvraag', aanvraagId, {
    techniekerCode: request.techniekerCode,
    mobielMagazijnCode: request.mobielMagazijnCode
  });

  return { success: true, message: 'Aanvraag ingediend naar mobiel magazijn.' };
}

/* =========================================================
   53.3 — goedkeuren / afkeuren door mobiel magazijn
   ========================================================= */

function approveMobileWarehouseRequest(payload) {
  payload = payload || {};

  const sessionId = getPayloadSessionId(payload);
  const aanvraagId = safeText(payload.aanvraagId);
  const actor = safeText(payload.actor);

  if (!aanvraagId) throw new Error('AanvraagID ontbreekt.');

  const request = buildMobileWarehouseRequestRows('').find(x => x.aanvraagId === aanvraagId);
  if (!request) throw new Error('Aanvraag niet gevonden.');

  const user = resolveMobileWarehouseRequestAccess(request, sessionId);

  if (request.status !== MW_REQUEST_STATUS.SUBMITTED) {
    throw new Error('Enkel ingediende aanvragen kunnen goedgekeurd worden.');
  }

  const sheet = getSheetOrThrow(TABS.MOBILE_REQUESTS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['AanvraagID']]) !== aanvraagId) continue;

    values[i][col['Status']] = MW_REQUEST_STATUS.APPROVED;
    values[i][col['GoedgekeurdDoor']] = actor || user.naam || user.email;
    values[i][col['GoedgekeurdOp']] = nowStamp();
    updated = true;
    break;
  }

  if (!updated) throw new Error('Aanvraag niet gevonden.');

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  markNotificationsReadBySource({
    rol: NOTIFICATION_ROLE.MOBILE_WAREHOUSE,
    ontvangerCode: request.mobielMagazijnCode,
    bronType: 'MobielMagazijnAanvraag',
    bronId: aanvraagId
  });

  pushTechnicianNotification(
    request.techniekerCode,
    request.techniekerNaam,
    'MobielMagazijnAanvraagGoedgekeurd',
    'Aanvraag mobiel magazijn goedgekeurd',
    `Je aanvraag ${aanvraagId} naar mobiel magazijn ${request.mobielMagazijnCode} werd goedgekeurd.`,
    'MobielMagazijnAanvraag',
    aanvraagId
  );

  writeAudit('Mobiel magazijn aanvraag goedgekeurd', user.rol, user.naam || user.email, 'MobielMagazijnAanvraag', aanvraagId, {
    techniekerCode: request.techniekerCode,
    mobielMagazijnCode: request.mobielMagazijnCode
  });

  return { success: true, message: 'Aanvraag goedgekeurd.' };
}

function rejectMobileWarehouseRequest(payload) {
  payload = payload || {};

  const sessionId = getPayloadSessionId(payload);
  const aanvraagId = safeText(payload.aanvraagId);
  const actor = safeText(payload.actor);
  const reason = safeText(payload.reason || payload.note);

  if (!aanvraagId) throw new Error('AanvraagID ontbreekt.');
  if (!reason) throw new Error('Afkeurreden is verplicht.');

  const request = buildMobileWarehouseRequestRows('').find(x => x.aanvraagId === aanvraagId);
  if (!request) throw new Error('Aanvraag niet gevonden.');

  const user = resolveMobileWarehouseRequestAccess(request, sessionId);

  if (![MW_REQUEST_STATUS.SUBMITTED, MW_REQUEST_STATUS.APPROVED].includes(request.status)) {
    throw new Error('Deze aanvraag kan niet meer afgekeurd worden.');
  }

  const sheet = getSheetOrThrow(TABS.MOBILE_REQUESTS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['AanvraagID']]) !== aanvraagId) continue;

    values[i][col['Status']] = MW_REQUEST_STATUS.REJECTED;
    values[i][col['AfgekeurdDoor']] = actor || user.naam || user.email;
    values[i][col['AfgekeurdOp']] = nowStamp();
    values[i][col['AfkeurReden']] = reason;
    updated = true;
    break;
  }

  if (!updated) throw new Error('Aanvraag niet gevonden.');

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  markNotificationsReadBySource({
    rol: NOTIFICATION_ROLE.MOBILE_WAREHOUSE,
    ontvangerCode: request.mobielMagazijnCode,
    bronType: 'MobielMagazijnAanvraag',
    bronId: aanvraagId
  });

  pushTechnicianNotification(
    request.techniekerCode,
    request.techniekerNaam,
    'MobielMagazijnAanvraagAfgekeurd',
    'Aanvraag mobiel magazijn afgekeurd',
    `Je aanvraag ${aanvraagId} werd afgekeurd. Reden: ${reason}`,
    'MobielMagazijnAanvraag',
    aanvraagId
  );

  writeAudit('Mobiel magazijn aanvraag afgekeurd', user.rol, user.naam || user.email, 'MobielMagazijnAanvraag', aanvraagId, {
    afkeurReden: reason
  });

  return { success: true, message: 'Aanvraag afgekeurd.' };
}

/* =========================================================
   53.4 — aanvraag omzetten naar transfer mobiel -> bus
   ========================================================= */

function bookMobileWarehouseRequestToTransfer(payload) {
  payload = payload || {};

  const sessionId = getPayloadSessionId(payload);
  const aanvraagId = safeText(payload.aanvraagId);
  const actor = safeText(payload.actor);

  if (!aanvraagId) throw new Error('AanvraagID ontbreekt.');
  if (typeof createDirectMobileToBusTransfer !== 'function') {
    throw new Error('Transferfunctie createDirectMobileToBusTransfer ontbreekt.');
  }

  const request = buildMobileWarehouseRequestRows('').find(x => x.aanvraagId === aanvraagId);
  if (!request) throw new Error('Aanvraag niet gevonden.');

  const user = resolveMobileWarehouseRequestAccess(request, sessionId);

  if (request.status !== MW_REQUEST_STATUS.APPROVED) {
    throw new Error('Enkel goedgekeurde aanvragen kunnen geboekt worden naar transfer.');
  }

  if (!(request.lines || []).length) {
    throw new Error('Deze aanvraag bevat geen actieve lijnen.');
  }

  const transferResult = createDirectMobileToBusTransfer({
    sessionId: sessionId,
    mobileWarehouseCode: request.mobielMagazijnCode,
    doelTechniekerCode: request.techniekerCode,
    documentDatum: request.documentDatumIso,
    reden: request.reden,
    opmerking: request.opmerking,
    actor: actor || user.naam || user.email,
    lines: request.lines.map(line => ({
      artikelCode: line.artikelCode,
      aantal: line.aantal
    }))
  });

  const transferId = safeText((transferResult && transferResult.transferId) || '');

  const sheet = getSheetOrThrow(TABS.MOBILE_REQUESTS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['AanvraagID']]) !== aanvraagId) continue;

    values[i][col['Status']] = MW_REQUEST_STATUS.BOOKED;
    values[i][col['GeboektDoor']] = actor || user.naam || user.email;
    values[i][col['GeboektOp']] = nowStamp();
    if (col['TransferID'] !== undefined) values[i][col['TransferID']] = transferId;
    updated = true;
    break;
  }

  if (!updated) throw new Error('Aanvraag niet gevonden.');

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  markNotificationsReadBySource({
    rol: NOTIFICATION_ROLE.MOBILE_WAREHOUSE,
    ontvangerCode: request.mobielMagazijnCode,
    bronType: 'MobielMagazijnAanvraag',
    bronId: aanvraagId
  });

  writeAudit('Mobiel magazijn aanvraag geboekt naar transfer', user.rol, user.naam || user.email, 'MobielMagazijnAanvraag', aanvraagId, {
    transferId: transferId,
    techniekerCode: request.techniekerCode,
    mobielMagazijnCode: request.mobielMagazijnCode
  });

  return {
    success: true,
    transferId,
    message: 'Aanvraag geboekt naar transfer mobiel → bus.'
  };
}