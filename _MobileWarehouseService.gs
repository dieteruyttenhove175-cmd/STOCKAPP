/* =========================================================
   52_MobileWarehouseService.gs — mobiel magazijn service
   ========================================================= */

/* ---------------------------------------------------------
   Helpers locatie
   --------------------------------------------------------- */

function getMobileWarehouseLocationCode(code) {
  return 'Mobiel:' + safeText(code);
}

function parseMobileWarehouseLocation(location) {
  const text = safeText(location);
  if (!/^Mobiel:/i.test(text)) return '';
  return safeText(text.split(':').slice(1).join(':'));
}

function isMobileWarehouseLocation(location) {
  return /^Mobiel:/i.test(safeText(location));
}

/* ---------------------------------------------------------
   Mapper
   --------------------------------------------------------- */

function mapMobileWarehouse(row) {
  return {
    code: safeText(row.Code),
    naam: safeText(row.Naam),
    type: safeText(row.Type),
    active: row.Actief === undefined ? true : isTrue(row.Actief),
    managerNaam: safeText(row.ManagerNaam),
    opmerking: safeText(row.Opmerking)
  };
}

/* ---------------------------------------------------------
   Readers
   --------------------------------------------------------- */

function getAllMobileWarehouses() {
  return readObjectsSafe(TABS.MOBILE_WAREHOUSES)
    .map(mapMobileWarehouse)
    .filter(x => x.active)
    .sort((a, b) => safeText(a.naam).localeCompare(safeText(b.naam)));
}

function getMobileWarehouseByCode(code) {
  const value = safeText(code);
  if (!value) return null;
  return getAllMobileWarehouses().find(x => safeText(x.code) === value) || null;
}

function getDefaultMobileWarehouseCode() {
  const rows = getAllMobileWarehouses();
  if (!rows.length) return '';
  return safeText(rows[0].code);
}

function getEffectiveMobileWarehouseCodeForUser(user, requestedCode) {
  const asked = safeText(requestedCode);
  const own = safeText((user && user.mobileWarehouseCode) || '');

  if (user && user.rol === ROLE.ADMIN) {
    return asked || own || getDefaultMobileWarehouseCode();
  }

  if (user && user.rol === ROLE.MANAGER) {
    return asked || own || getDefaultMobileWarehouseCode();
  }

  if (user && user.rol === ROLE.MOBILE_WAREHOUSE) {
    if (own && asked && own !== asked) {
      throw new Error('Je kan enkel je eigen mobiel magazijn openen.');
    }
    return own || asked || getDefaultMobileWarehouseCode();
  }

  throw new Error('Geen rechten voor mobiel magazijn.');
}

function assertMobileWarehouseAccess(sessionId, requestedCode) {
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten voor mobiel magazijn.');
  }

  const mobileWarehouseCode = getEffectiveMobileWarehouseCodeForUser(user, requestedCode);
  if (!mobileWarehouseCode) throw new Error('Geen mobiel magazijn gekoppeld of beschikbaar.');

  const warehouse = getMobileWarehouseByCode(mobileWarehouseCode);
  if (!warehouse) throw new Error('Mobiel magazijn niet gevonden.');

  return {
    user,
    mobileWarehouseCode,
    warehouse
  };
}

/* ---------------------------------------------------------
   Notificaties mobiel magazijn
   --------------------------------------------------------- */

function pushMobileWarehouseNotification(type, title, message, bronType, bronId, mobileWarehouseCode) {
  return pushNotification({
    rol: NOTIFICATION_ROLE.MOBILE_WAREHOUSE,
    ontvangerCode: safeText(mobileWarehouseCode) || 'MOBIEL',
    ontvangerNaam: 'Mobiel magazijn',
    type,
    titel: title,
    bericht: message,
    bronType,
    bronId
  });
}

function getNotificationsForMobileWarehouse(mobileWarehouseCode) {
  const code = safeText(mobileWarehouseCode);

  return readObjectsSafe(TABS.NOTIFICATIONS)
    .map(mapNotification)
    .filter(item =>
      item.rol === NOTIFICATION_ROLE.MOBILE_WAREHOUSE &&
      (!code || safeText(item.ontvangerCode) === code || safeText(item.ontvangerCode) === 'MOBIEL')
    )
    .sort((a, b) => safeText(b.aangemaaktOpRaw).localeCompare(safeText(a.aangemaaktOpRaw)));
}

/* ---------------------------------------------------------
   Stock mobiel magazijn opbouwen
   --------------------------------------------------------- */

function buildAllMobileWarehouseStockRows() {
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

    const from = safeText(move.locatieVan);
    const to = safeText(move.locatieNaar);

    const qtyIn = safeNumber(move.aantalIn, 0);
    const qtyOut = safeNumber(move.aantalUit, 0);
    const net = safeNumber(move.nettoAantal, 0);

    if (isMobileWarehouseLocation(to)) {
      const mwCode = parseMobileWarehouseLocation(to);
      const key = `${mwCode}|${code}`;

      if (!grouped[key]) {
        grouped[key] = {
          mobileWarehouseCode: mwCode,
          mobileWarehouseNaam: (getMobileWarehouseByCode(mwCode) || {}).naam || mwCode,
          artikelCode: code,
          artikelOmschrijving: safeText(move.artikelOmschrijving) || safeText(articleMap[code] && articleMap[code].artikelOmschrijving),
          eenheid: safeText(move.eenheid) || safeText(articleMap[code] && articleMap[code].eenheid),
          typeMateriaal: safeText(move.typeMateriaal) || determineMaterialTypeFromArticle(code),
          voorraadMobiel: 0,
          laatsteMutatie: safeText(move.datumBoeking),
          laatsteMutatieRaw: safeText(move.datumBoekingRaw)
        };
      }

      grouped[key].voorraadMobiel += qtyIn || Math.abs(net);

      if (safeText(move.datumBoekingRaw) > safeText(grouped[key].laatsteMutatieRaw)) {
        grouped[key].laatsteMutatie = safeText(move.datumBoeking);
        grouped[key].laatsteMutatieRaw = safeText(move.datumBoekingRaw);
      }
    }

    if (isMobileWarehouseLocation(from)) {
      const mwCode = parseMobileWarehouseLocation(from);
      const key = `${mwCode}|${code}`;

      if (!grouped[key]) {
        grouped[key] = {
          mobileWarehouseCode: mwCode,
          mobileWarehouseNaam: (getMobileWarehouseByCode(mwCode) || {}).naam || mwCode,
          artikelCode: code,
          artikelOmschrijving: safeText(move.artikelOmschrijving) || safeText(articleMap[code] && articleMap[code].artikelOmschrijving),
          eenheid: safeText(move.eenheid) || safeText(articleMap[code] && articleMap[code].eenheid),
          typeMateriaal: safeText(move.typeMateriaal) || determineMaterialTypeFromArticle(code),
          voorraadMobiel: 0,
          laatsteMutatie: safeText(move.datumBoeking),
          laatsteMutatieRaw: safeText(move.datumBoekingRaw)
        };
      }

      grouped[key].voorraadMobiel -= qtyOut || Math.abs(net);

      if (safeText(move.datumBoekingRaw) > safeText(grouped[key].laatsteMutatieRaw)) {
        grouped[key].laatsteMutatie = safeText(move.datumBoeking);
        grouped[key].laatsteMutatieRaw = safeText(move.datumBoekingRaw);
      }
    }
  });

  return Object.keys(grouped)
    .map(key => grouped[key])
    .filter(item => safeNumber(item.voorraadMobiel, 0) !== 0)
    .sort((a, b) =>
      safeText(a.mobileWarehouseNaam).localeCompare(safeText(b.mobileWarehouseNaam)) ||
      safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
    );
}

function buildMobileWarehouseStockRows(mobileWarehouseCode) {
  const code = safeText(mobileWarehouseCode);
  return buildAllMobileWarehouseStockRows().filter(item => safeText(item.mobileWarehouseCode) === code);
}

function buildMobileWarehouseStockMap(mobileWarehouseCode) {
  const map = {};
  buildMobileWarehouseStockRows(mobileWarehouseCode).forEach(item => {
    map[item.artikelCode] = item;
  });
  return map;
}

function buildMobileWarehouseSummary(mobileWarehouseCode) {
  const rows = buildMobileWarehouseStockRows(mobileWarehouseCode);

  return {
    mobileWarehouseCode: safeText(mobileWarehouseCode),
    artikels: rows.length,
    totaalAantal: rows.reduce((sum, row) => sum + safeNumber(row.voorraadMobiel, 0), 0),
    grabbelAantal: rows
      .filter(row => safeText(row.typeMateriaal) === 'Grabbel')
      .reduce((sum, row) => sum + safeNumber(row.voorraadMobiel, 0), 0),
    behoefteAantal: rows
      .filter(row => safeText(row.typeMateriaal) === 'Behoefte')
      .reduce((sum, row) => sum + safeNumber(row.voorraadMobiel, 0), 0)
  };
}

/* ---------------------------------------------------------
   Aanvragen voor mobiel magazijn
   --------------------------------------------------------- */

function buildMobileWarehouseRequestRows(mobileWarehouseCode) {
  const code = safeText(mobileWarehouseCode);

  if (typeof getAllMobileRequests !== 'function' || typeof getAllMobileRequestLines !== 'function') {
    return [];
  }

  return buildMobileRequestsWithLines(getAllMobileRequests(), getAllMobileRequestLines())
    .filter(item => !code || safeText(item.mobileWarehouseCode) === code)
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.aanvraagId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.aanvraagId)}`
      )
    );
}

function buildMobileWarehouseRequestQueue(mobileWarehouseCode) {
  return buildMobileWarehouseRequestRows(mobileWarehouseCode)
    .filter(item => [
      MOBILE_REQUEST_STATUS.SUBMITTED,
      MOBILE_REQUEST_STATUS.APPROVED,
      MOBILE_REQUEST_STATUS.BOOKED
    ].includes(safeText(item.status)));
}

/* ---------------------------------------------------------
   Vaste filters / snelle zoeklogica
   --------------------------------------------------------- */

function filterMobileWarehouseStockRows(rows, filters) {
  const source = rows || [];
  const f = filters || {};

  const materialType = safeText(f.materialType);
  const articleCodePrefix = safeText(f.articleCodePrefix).toLowerCase();
  const inStockOnly = !!f.inStockOnly;
  const minQty = f.minQty === '' || f.minQty === null || f.minQty === undefined ? null : Number(f.minQty);
  const sortBy = safeText(f.sortBy || 'artikel');
  const sortDir = safeText(f.sortDir || 'asc').toLowerCase();

  let result = source.slice();

  if (materialType) {
    result = result.filter(row => safeText(row.typeMateriaal) === materialType);
  }

  if (articleCodePrefix) {
    result = result.filter(row => safeText(row.artikelCode).toLowerCase().startsWith(articleCodePrefix));
  }

  if (inStockOnly) {
    result = result.filter(row => safeNumber(row.voorraadMobiel, 0) > 0);
  }

  if (minQty !== null && !isNaN(minQty)) {
    result = result.filter(row => safeNumber(row.voorraadMobiel, 0) >= minQty);
  }

  const factor = sortDir === 'desc' ? -1 : 1;

  result.sort((a, b) => {
    if (sortBy === 'qty') {
      return factor * (safeNumber(a.voorraadMobiel, 0) - safeNumber(b.voorraadMobiel, 0));
    }
    if (sortBy === 'type') {
      return factor * safeText(a.typeMateriaal).localeCompare(safeText(b.typeMateriaal));
    }
    return factor * (
      safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
      safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
    );
  });

  return result;
}

function getMobileWarehouseFilteredStock(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const access = assertMobileWarehouseAccess(sessionId, safeText(payload.mobileWarehouseCode));

  const rows = buildMobileWarehouseStockRows(access.mobileWarehouseCode);
  return filterMobileWarehouseStockRows(rows, payload.filters || {});
}

/* ---------------------------------------------------------
   Dashboarddata
   --------------------------------------------------------- */

function getMobileWarehouseDashboardData(payload) {
  payload = payload || {};

  const sessionId = getPayloadSessionId(payload);
  const access = assertMobileWarehouseAccess(sessionId, safeText(payload.mobileWarehouseCode));

  const warehouseCode = access.mobileWarehouseCode;
  const stockRows = buildMobileWarehouseStockRows(warehouseCode);
  const filteredStockRows = filterMobileWarehouseStockRows(stockRows, payload.filters || {});
  const requestQueue = buildMobileWarehouseRequestQueue(warehouseCode);
  const allRequests = buildMobileWarehouseRequestRows(warehouseCode);

  return {
    warehouse: access.warehouse,
    selectedMobileWarehouseCode: warehouseCode,
    availableWarehouses: getAllMobileWarehouses(),
    summary: buildMobileWarehouseSummary(warehouseCode),
    stockRows: filteredStockRows,
    centralWarehouse: getCentralWarehouseOverview(),
    technicians: getActiveTechnicians().map(t => ({ code: t.code, naam: t.naam })),
    requestQueue: requestQueue,
    requestHistory: allRequests,
    notifications: getNotificationsForMobileWarehouse(warehouseCode),
    recurringBusCountSummary: typeof buildBusCountTriggerSummary === 'function'
      ? buildBusCountTriggerSummary()
      : null,
    generatedAt: Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm')
  };
}

/* ---------------------------------------------------------
   Directe transfers
   Vereist bestaand transferblok met createTransfer/saveTransferLines
   --------------------------------------------------------- */

function assertTransferServiceAvailable() {
  if (typeof createTransfer !== 'function') {
    throw new Error('Transferservice ontbreekt. Werk eerst het transferblok in.');
  }
  if (typeof saveTransferLines !== 'function') {
    throw new Error('Transferlijnservice ontbreekt. Werk eerst het transferblok in.');
  }
}

function createDirectMobileToBusTransfer(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');
  assertTransferServiceAvailable();

  const sessionId = getPayloadSessionId(payload);
  const access = assertMobileWarehouseAccess(sessionId, safeText(payload.mobileWarehouseCode));

  const doelTechniekerCode = safeText(payload.doelTechniekerCode);
  const documentDatum = safeText(payload.documentDatum);
  const reden = safeText(payload.reden);
  const opmerking = safeText(payload.opmerking);
  const actor = safeText(payload.actor || access.user.naam || access.user.email || 'MobielMagazijn');
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!doelTechniekerCode) throw new Error('DoelTechniekerCode is verplicht.');
  if (!documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!reden) throw new Error('Reden is verplicht.');
  if (!lines.length) throw new Error('Geen transferlijnen ontvangen.');

  const stockMap = buildMobileWarehouseStockMap(access.mobileWarehouseCode);

  lines.forEach(line => {
    const code = safeText(line.artikelCode);
    const qty = safeNumber(line.aantal, 0);
    const current = stockMap[code] ? safeNumber(stockMap[code].voorraadMobiel, 0) : 0;

    if (!code) throw new Error('ArtikelCode ontbreekt.');
    if (qty <= 0) throw new Error(`Aantal moet groter zijn dan 0 voor artikel ${code}.`);
    if (qty > current) {
      throw new Error(`Onvoldoende voorraad in mobiel magazijn voor artikel ${code}. Beschikbaar: ${current}`);
    }
  });

  const result = createTransfer({
    sessionId: sessionId,
    flowType: TRANSFER_FLOW.MOBILE_TO_BUS,
    mobileWarehouseCode: access.mobileWarehouseCode,
    doelTechniekerCode: doelTechniekerCode,
    documentDatum: documentDatum,
    reden: reden,
    opmerking: opmerking,
    actor: actor
  });

  saveTransferLines({
    sessionId: sessionId,
    transferId: result.transferId,
    lines: lines
  });

  pushTechnicianNotification(
    doelTechniekerCode,
    getTechnicianNameByCode(doelTechniekerCode),
    'MobielTransferAangemaakt',
    'Transfer vanuit mobiel magazijn',
    `Er is een transfer vanuit mobiel magazijn aangemaakt naar jouw bus (${result.transferId}).`,
    'Transfer',
    result.transferId
  );

  writeAudit(
    'Mobiel -> bus transfer aangemaakt',
    access.user.rol,
    actor,
    'Transfer',
    result.transferId,
    {
      mobileWarehouseCode: access.mobileWarehouseCode,
      doelTechniekerCode: doelTechniekerCode,
      lijnen: lines.length,
      reden: reden
    }
  );

  return {
    success: true,
    transferId: result.transferId,
    message: 'Transfer mobiel -> bus aangemaakt.'
  };
}

function createDirectCentralToMobileTransfer(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');
  assertTransferServiceAvailable();

  const sessionId = getPayloadSessionId(payload);
  const access = assertMobileWarehouseAccess(sessionId, safeText(payload.mobileWarehouseCode));

  const documentDatum = safeText(payload.documentDatum);
  const reden = safeText(payload.reden);
  const opmerking = safeText(payload.opmerking);
  const actor = safeText(payload.actor || access.user.naam || access.user.email || 'MobielMagazijn');
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!reden) throw new Error('Reden is verplicht.');
  if (!lines.length) throw new Error('Geen transferlijnen ontvangen.');

  const centralMap = buildCentralWarehouseMap();

  lines.forEach(line => {
    const code = safeText(line.artikelCode);
    const qty = safeNumber(line.aantal, 0);
    const current = centralMap[code] ? safeNumber(centralMap[code].voorraadCentraal, 0) : 0;

    if (!code) throw new Error('ArtikelCode ontbreekt.');
    if (qty <= 0) throw new Error(`Aantal moet groter zijn dan 0 voor artikel ${code}.`);
    if (qty > current) {
      throw new Error(`Onvoldoende voorraad in centraal magazijn voor artikel ${code}. Beschikbaar: ${current}`);
    }
  });

  const result = createTransfer({
    sessionId: sessionId,
    flowType: TRANSFER_FLOW.CENTRAL_TO_MOBILE,
    mobileWarehouseCode: access.mobileWarehouseCode,
    documentDatum: documentDatum,
    reden: reden,
    opmerking: opmerking,
    actor: actor
  });

  saveTransferLines({
    sessionId: sessionId,
    transferId: result.transferId,
    lines: lines
  });

  writeAudit(
    'Centraal -> mobiel transfer aangemaakt',
    access.user.rol,
    actor,
    'Transfer',
    result.transferId,
    {
      mobileWarehouseCode: access.mobileWarehouseCode,
      lijnen: lines.length,
      reden: reden
    }
  );

  return {
    success: true,
    transferId: result.transferId,
    message: 'Transfer centraal -> mobiel aangemaakt.'
  };
}

function createDirectMobileToCentralTransfer(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');
  assertTransferServiceAvailable();

  const sessionId = getPayloadSessionId(payload);
  const access = assertMobileWarehouseAccess(sessionId, safeText(payload.mobileWarehouseCode));

  const documentDatum = safeText(payload.documentDatum);
  const reden = safeText(payload.reden);
  const opmerking = safeText(payload.opmerking);
  const actor = safeText(payload.actor || access.user.naam || access.user.email || 'MobielMagazijn');
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!reden) throw new Error('Reden is verplicht.');
  if (!lines.length) throw new Error('Geen transferlijnen ontvangen.');

  const stockMap = buildMobileWarehouseStockMap(access.mobileWarehouseCode);

  lines.forEach(line => {
    const code = safeText(line.artikelCode);
    const qty = safeNumber(line.aantal, 0);
    const current = stockMap[code] ? safeNumber(stockMap[code].voorraadMobiel, 0) : 0;

    if (!code) throw new Error('ArtikelCode ontbreekt.');
    if (qty <= 0) throw new Error(`Aantal moet groter zijn dan 0 voor artikel ${code}.`);
    if (qty > current) {
      throw new Error(`Onvoldoende voorraad in mobiel magazijn voor artikel ${code}. Beschikbaar: ${current}`);
    }
  });

  const result = createTransfer({
    sessionId: sessionId,
    flowType: TRANSFER_FLOW.MOBILE_TO_CENTRAL,
    mobileWarehouseCode: access.mobileWarehouseCode,
    documentDatum: documentDatum,
    reden: reden,
    opmerking: opmerking,
    actor: actor
  });

  saveTransferLines({
    sessionId: sessionId,
    transferId: result.transferId,
    lines: lines
  });

  writeAudit(
    'Mobiel -> centraal transfer aangemaakt',
    access.user.rol,
    actor,
    'Transfer',
    result.transferId,
    {
      mobileWarehouseCode: access.mobileWarehouseCode,
      lijnen: lines.length,
      reden: reden
    }
  );

  return {
    success: true,
    transferId: result.transferId,
    message: 'Transfer mobiel -> centraal aangemaakt.'
  };
}