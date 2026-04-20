/* =========================================================
   24_StockSnapshotService.gs — stock snapshots / overzichten
   ========================================================= */

function buildLocationStockRows(locationCode, extraFieldsBuilder) {
  const articleMap = buildArticleMasterMap();
  const moves = getAllMovements();
  const grouped = {};

  moves.forEach(move => {
    const code = safeText(move.artikelCode);
    if (!code) return;

    const delta = getLocationDeltaFromMovement(move, locationCode);
    if (!delta) return;

    if (!grouped[code]) {
      const article = articleMap[code] || {};
      grouped[code] = {
        artikelCode: code,
        artikelOmschrijving: safeText(move.artikelOmschrijving) || safeText(article.artikelOmschrijving),
        eenheid: safeText(move.eenheid) || safeText(article.eenheid),
        voorraad: 0,
        min: article.min === undefined ? '' : article.min,
        max: article.max === undefined ? '' : article.max,
        laatsteMutatie: move.datumBoeking || '',
        laatsteMutatieRaw: move.datumBoekingRaw || ''
      };

      if (typeof extraFieldsBuilder === 'function') {
        const extra = extraFieldsBuilder(code, article, move) || {};
        Object.keys(extra).forEach(key => {
          grouped[code][key] = extra[key];
        });
      }
    }

    grouped[code].voorraad += delta;

    if (String(move.datumBoekingRaw || '') > String(grouped[code].laatsteMutatieRaw || '')) {
      grouped[code].laatsteMutatie = move.datumBoeking || '';
      grouped[code].laatsteMutatieRaw = move.datumBoekingRaw || '';
    }
  });

  return Object.keys(grouped)
    .map(key => grouped[key])
    .filter(item => Number(item.voorraad || 0) !== 0)
    .sort((a, b) =>
      String(a.artikelOmschrijving || '').localeCompare(String(b.artikelOmschrijving || '')) ||
      String(a.artikelCode || '').localeCompare(String(b.artikelCode || ''))
    );
}

function buildCentralWarehouseRows() {
  return buildLocationStockRows(LOCATION.CENTRAL).map(item => ({
    artikelCode: item.artikelCode,
    artikelOmschrijving: item.artikelOmschrijving,
    eenheid: item.eenheid,
    voorraadCentraal: Number(item.voorraad || 0),
    min: item.min,
    max: item.max,
    laatsteMutatie: item.laatsteMutatie || '',
    laatsteMutatieRaw: item.laatsteMutatieRaw || ''
  }));
}

function buildCentralWarehouseMap() {
  const map = {};
  buildCentralWarehouseRows().forEach(item => {
    map[item.artikelCode] = item;
  });
  return map;
}

function rebuildCentralWarehouseOverview() {
  const rows = buildCentralWarehouseRows().map(item => [
    item.artikelCode,
    item.artikelOmschrijving,
    item.eenheid,
    item.voorraadCentraal,
    item.min,
    item.max,
    item.laatsteMutatie
  ]);

  writeFullTable(
    TABS.CENTRAL_WAREHOUSE,
    ['ArtikelCode', 'ArtikelOmschrijving', 'Eenheid', 'VoorraadCentraal', 'Min', 'Max', 'LaatsteMutatie'],
    rows
  );

  return {
    success: true,
    lines: rows.length,
    message: 'Centraal magazijn vernieuwd.'
  };
}

function getCentralWarehouseOverview() {
  return buildCentralWarehouseRows();
}

function buildMobileWarehouseRows() {
  return buildLocationStockRows(LOCATION.MOBILE).map(item => ({
    artikelCode: item.artikelCode,
    artikelOmschrijving: item.artikelOmschrijving,
    eenheid: item.eenheid,
    voorraadMobiel: Number(item.voorraad || 0),
    min: item.min,
    max: item.max,
    laatsteMutatie: item.laatsteMutatie || '',
    laatsteMutatieRaw: item.laatsteMutatieRaw || ''
  }));
}

function buildMobileWarehouseMap() {
  const map = {};
  buildMobileWarehouseRows().forEach(item => {
    map[item.artikelCode] = item;
  });
  return map;
}

function buildBusStockRows() {
  const articleMap = buildArticleMasterMap();
  const moves = getAllMovements();
  const grouped = {};

  moves.forEach(move => {
    const code = safeText(move.artikelCode);
    if (!code) return;

    const fromLocation = safeText(move.locatieVan);
    const toLocation = safeText(move.locatieNaar);

    if (isBusLocation(toLocation)) {
      const busCode = parseBusLocation(toLocation);
      const key = `${busCode}||${code}`;

      if (!grouped[key]) {
        const article = articleMap[code] || {};
        grouped[key] = {
          techniekerCode: busCode,
          techniekerNaam: getTechnicianNameByCode(busCode),
          artikelCode: code,
          artikelOmschrijving: safeText(move.artikelOmschrijving) || safeText(article.artikelOmschrijving),
          eenheid: safeText(move.eenheid) || safeText(article.eenheid),
          voorraadBus: 0,
          laatsteMutatie: move.datumBoeking || '',
          laatsteMutatieRaw: move.datumBoekingRaw || ''
        };
      }

      grouped[key].voorraadBus += getMovementInboundQty(move);

      if (String(move.datumBoekingRaw || '') > String(grouped[key].laatsteMutatieRaw || '')) {
        grouped[key].laatsteMutatie = move.datumBoeking || '';
        grouped[key].laatsteMutatieRaw = move.datumBoekingRaw || '';
      }
    }

    if (isBusLocation(fromLocation)) {
      const busCode = parseBusLocation(fromLocation);
      const key = `${busCode}||${code}`;

      if (!grouped[key]) {
        const article = articleMap[code] || {};
        grouped[key] = {
          techniekerCode: busCode,
          techniekerNaam: getTechnicianNameByCode(busCode),
          artikelCode: code,
          artikelOmschrijving: safeText(move.artikelOmschrijving) || safeText(article.artikelOmschrijving),
          eenheid: safeText(move.eenheid) || safeText(article.eenheid),
          voorraadBus: 0,
          laatsteMutatie: move.datumBoeking || '',
          laatsteMutatieRaw: move.datumBoekingRaw || ''
        };
      }

      grouped[key].voorraadBus -= getMovementOutboundQty(move);

      if (String(move.datumBoekingRaw || '') > String(grouped[key].laatsteMutatieRaw || '')) {
        grouped[key].laatsteMutatie = move.datumBoeking || '';
        grouped[key].laatsteMutatieRaw = move.datumBoekingRaw || '';
      }
    }
  });

  return Object.keys(grouped)
    .map(key => grouped[key])
    .filter(item => Number(item.voorraadBus || 0) !== 0)
    .sort((a, b) =>
      String(a.techniekerNaam || '').localeCompare(String(b.techniekerNaam || '')) ||
      String(a.artikelOmschrijving || '').localeCompare(String(b.artikelOmschrijving || '')) ||
      String(a.artikelCode || '').localeCompare(String(b.artikelCode || ''))
    );
}

function buildBusStockMap() {
  const map = {};
  buildBusStockRows().forEach(item => {
    const busCode = safeText(item.techniekerCode);
    if (!map[busCode]) map[busCode] = {};
    map[busCode][item.artikelCode] = item;
  });
  return map;
}

function getBusStockMapForTechnician(techniekerCode) {
  const code = safeText(techniekerCode);
  if (!code) return {};

  return buildBusStockMap()[code] || {};
}

function buildCombinedStockRows() {
  const centralMap = buildCentralWarehouseMap();
  const busRows = buildBusStockRows();
  const groupedBus = {};
  const articleMaster = buildArticleMasterMap();

  busRows.forEach(row => {
    const code = safeText(row.artikelCode);
    if (!code) return;
    groupedBus[code] = (groupedBus[code] || 0) + Number(row.voorraadBus || 0);
  });

  const codeSet = {};
  Object.keys(centralMap).forEach(code => { codeSet[code] = true; });
  Object.keys(groupedBus).forEach(code => { codeSet[code] = true; });
  Object.keys(articleMaster).forEach(code => { codeSet[code] = true; });

  return Object.keys(codeSet)
    .map(code => {
      const central = centralMap[code];
      const article = articleMaster[code] || {};

      return {
        artikelCode: code,
        artikelOmschrijving: safeText((central && central.artikelOmschrijving) || article.artikelOmschrijving),
        eenheid: safeText((central && central.eenheid) || article.eenheid),
        voorraadCentraal: Number((central && central.voorraadCentraal) || 0),
        voorraadBus: Number(groupedBus[code] || 0),
        voorraadTotaal: Number((central && central.voorraadCentraal) || 0) + Number(groupedBus[code] || 0)
      };
    })
    .filter(item =>
      item.voorraadCentraal !== 0 ||
      item.voorraadBus !== 0 ||
      item.voorraadTotaal !== 0
    )
    .sort((a, b) =>
      String(a.artikelOmschrijving || '').localeCompare(String(b.artikelOmschrijving || '')) ||
      String(a.artikelCode || '').localeCompare(String(b.artikelCode || ''))
    );
}

function buildStockScopeSummary() {
  const centralRows = buildCentralWarehouseRows();
  const mobileRows = buildMobileWarehouseRows();
  const busRows = buildBusStockRows();
  const combinedRows = buildCombinedStockRows();

  return {
    centraalArtikels: centralRows.length,
    mobielArtikels: mobileRows.length,
    busArtikels: busRows.length,
    totaalArtikels: combinedRows.length,

    centraalAantal: centralRows.reduce((sum, row) => sum + Number(row.voorraadCentraal || 0), 0),
    mobielAantal: mobileRows.reduce((sum, row) => sum + Number(row.voorraadMobiel || 0), 0),
    busAantal: busRows.reduce((sum, row) => sum + Number(row.voorraadBus || 0), 0),
    totaalAantal: combinedRows.reduce((sum, row) => sum + Number(row.voorraadTotaal || 0), 0)
  };
}