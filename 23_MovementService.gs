/* =========================================================
   23_MovementService.gs — centrale mutatielogica
   ========================================================= */

function makeMovementId() {
  return makeStampedId('M');
}

function getAllMovements() {
  return readObjectsSafe(TABS.WAREHOUSE_MOVEMENTS)
    .map(mapWarehouseMovement)
    .sort((a, b) => String(a.datumBoekingRaw || '').localeCompare(String(b.datumBoekingRaw || '')));
}

function buildMovementObject(payload) {
  payload = payload || {};

  const artikelCode = safeText(payload.artikelCode);
  const articleMaster = getArticleMaster(artikelCode);

  const aantalIn = safeNumber(payload.aantalIn, 0);
  const aantalUit = safeNumber(payload.aantalUit, 0);
  const nettoAantal = payload.nettoAantal === undefined || payload.nettoAantal === null || payload.nettoAantal === ''
    ? (aantalIn - aantalUit)
    : safeNumber(payload.nettoAantal, 0);

  return {
    MutatieID: safeText(payload.mutatieId) || makeMovementId(),
    DatumBoeking: safeText(payload.datumBoeking) || nowStamp(),
    DatumDocument: safeText(payload.datumDocument),
    TypeMutatie: safeText(payload.typeMutatie),
    BronID: safeText(payload.bronId),
    TypeMateriaal: safeText(payload.typeMateriaal) || determineMaterialTypeFromArticle(artikelCode),
    ArtikelCode: artikelCode,
    ArtikelOmschrijving: safeText(payload.artikelOmschrijving) || safeText(articleMaster && articleMaster.artikelOmschrijving),
    Eenheid: safeText(payload.eenheid) || safeText(articleMaster && articleMaster.eenheid),
    AantalIn: aantalIn,
    AantalUit: aantalUit,
    NettoAantal: nettoAantal,
    LocatieVan: safeText(payload.locatieVan),
    LocatieNaar: safeText(payload.locatieNaar),
    Reden: safeText(payload.reden),
    Opmerking: safeText(payload.opmerking),
    GoedgekeurdDoor: safeText(payload.goedgekeurdDoor),
    GoedgekeurdOp: safeText(payload.goedgekeurdOp) || nowStamp()
  };
}

function appendMovement(payload) {
  const movement = buildMovementObject(payload);
  appendObjects(TABS.WAREHOUSE_MOVEMENTS, [movement]);
  return movement;
}

function appendMovementBatch(payloadList) {
  const rows = (payloadList || []).map(buildMovementObject);
  if (rows.length) {
    appendObjects(TABS.WAREHOUSE_MOVEMENTS, rows);
  }
  return rows;
}

function getMovementsBySource(bronId, typeMutatie) {
  const sourceId = safeText(bronId);
  const movementType = safeText(typeMutatie);

  return getAllMovements().filter(move => {
    if (safeText(move.bronId) !== sourceId) return false;
    if (movementType && safeText(move.typeMutatie) !== movementType) return false;
    return true;
  });
}

function replaceSourceMovements(typeMutatieOrList, bronId, movementPayloads) {
  const typeList = Array.isArray(typeMutatieOrList)
    ? typeMutatieOrList.map(safeText).filter(Boolean)
    : [safeText(typeMutatieOrList)].filter(Boolean);

  const sourceId = safeText(bronId);

  const sheet = getSheetOrThrow(TABS.WAREHOUSE_MOVEMENTS);
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values[0].map(h => safeText(h)) : getHeaders(TABS.WAREHOUSE_MOVEMENTS);
  const col = getColMap(headers);

  const existingRows = values.length > 1 ? values.slice(1) : [];

  const keptRows = existingRows.filter(row => {
    const rowType = safeText(row[col['TypeMutatie']]);
    const rowSource = safeText(row[col['BronID']]);

    if (rowSource !== sourceId) return true;
    if (!typeList.length) return false;
    return !typeList.includes(rowType);
  });

  const newRows = (movementPayloads || [])
    .map(buildMovementObject)
    .map(obj => buildRowFromHeaders(headers, obj));

  writeFullTable(TABS.WAREHOUSE_MOVEMENTS, headers, keptRows.concat(newRows));

  return {
    success: true,
    replaced: newRows.length
  };
}

function getMovementInboundQty(move) {
  const qtyIn = safeNumber(move.aantalIn, 0);
  if (qtyIn > 0) return qtyIn;

  const net = safeNumber(move.nettoAantal, 0);
  return net > 0 ? net : 0;
}

function getMovementOutboundQty(move) {
  const qtyOut = safeNumber(move.aantalUit, 0);
  if (qtyOut > 0) return qtyOut;

  const net = safeNumber(move.nettoAantal, 0);
  return net < 0 ? Math.abs(net) : 0;
}

function isBusLocation(location) {
  return /^Bus:/i.test(safeText(location));
}

function getLocationDeltaFromMovement(move, locationCode) {
  const code = safeText(locationCode);
  if (!code) return 0;

  let delta = 0;

  if (safeText(move.locatieNaar) === code) {
    delta += getMovementInboundQty(move);
  }

  if (safeText(move.locatieVan) === code) {
    delta -= getMovementOutboundQty(move);
  }

  return delta;
}