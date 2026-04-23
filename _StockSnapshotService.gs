/* =========================================================
   24_StockSnapshotService.gs
   Refactor: stock snapshot core service
   Doel:
   - stock afleiden uit mutaties
   - centraal magazijn snapshot
   - busstock snapshot
   - mobiel magazijn snapshot helpers
   - gecombineerde stockbeelden voor schermen
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getCentralStockOverviewTab_() {
  return TABS.CENTRAL_STOCK_OVERVIEW || TABS.CENTRAL_STOCK || 'CentraalMagazijnOverzicht';
}

function getWarehouseMovementsTab_() {
  return TABS.WAREHOUSE_MOVEMENTS || TABS.MOVEMENTS || 'MagazijnMutaties';
}

/* ---------------------------------------------------------
   Movement mapping fallback
   --------------------------------------------------------- */

function mapSnapshotMovement_(row) {
  if (typeof mapWarehouseMovement === 'function') {
    return mapWarehouseMovement(row);
  }

  return {
    movementId: safeText(row.MovementID || row.MovementId || row.ID),
    movementType: safeText(row.MovementType || row.Type),
    bronType: safeText(row.BronType),
    bronId: safeText(row.BronID || row.BronId),
    datumBoeking: safeText(row.DatumBoeking || row.DocumentDatum || row.Datum),
    datumBoekingRaw: safeText(row.DatumBoekingRaw || row.DatumBoeking || row.DocumentDatum || row.Datum),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode || row.ArtikelNr))),
    eenheid: safeText(row.Eenheid || row.Unit),
    aantalIn: safeNumber(row.AantalIn, 0),
    aantalUit: safeNumber(row.AantalUit, 0),
    nettoAantal: safeNumber(
      row.NettoAantal,
      safeNumber(row.AantalIn, 0) - safeNumber(row.AantalUit, 0)
    ),
    locatieVan: safeText(row.LocatieVan),
    locatieNaar: safeText(row.LocatieNaar),
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),
    actor: safeText(row.Actor),
  };
}

function getAllSnapshotMovements_() {
  return readObjectsSafe(getWarehouseMovementsTab_())
    .map(mapSnapshotMovement_)
    .filter(function (item) {
      return !!safeText(item.artikelCode);
    })
    .sort(function (a, b) {
      return (
        safeText(a.datumBoekingRaw).localeCompare(safeText(b.datumBoekingRaw)) ||
        safeText(a.movementId).localeCompare(safeText(b.movementId))
      );
    });
}

/* ---------------------------------------------------------
   Generic location helpers
   --------------------------------------------------------- */

function isCentralLocation_(location) {
  return safeText(location) === safeText(LOCATION.CENTRAL);
}

function isBusLocation_(location) {
  if (typeof isBusLocation === 'function') {
    return isBusLocation(location);
  }
  return /^Bus:/i.test(safeText(location));
}

function parseBusLocation_(location) {
  if (typeof parseBusLocation === 'function') {
    return parseBusLocation(location);
  }
  var text = safeText(location);
  if (!/^Bus:/i.test(text)) return '';
  return safeText(text.split(':').slice(1).join(':'));
}

function isMobileWarehouseLocation_(location) {
  if (typeof isMobileWarehouseLocation === 'function') {
    return isMobileWarehouseLocation(location);
  }
  return /^Mobiel:/i.test(safeText(location));
}

function parseMobileWarehouseLocation_(location) {
  if (typeof parseMobileWarehouseLocation === 'function') {
    return parseMobileWarehouseLocation(location);
  }
  var text = safeText(location);
  if (!/^Mobiel:/i.test(text)) return '';
  return safeText(text.split(':').slice(1).join(':'));
}

/* ---------------------------------------------------------
   Generic aggregation
   --------------------------------------------------------- */

function buildLocationStockRows(movements, options) {
  var cfg = Object.assign(
    {
      includePredicate: function () { return true; },
      locationExtractor: function () { return ''; },
      qtyExtractor: function () { return 0; },
      rowDecorator: function () { return {}; },
    },
    options || {}
  );

  var grouped = {};

  (movements || []).forEach(function (move) {
    if (!cfg.includePredicate(move)) {
      return;
    }

    var locationKey = safeText(cfg.locationExtractor(move));
    var artikelCode = safeText(move.artikelCode);
    if (!locationKey || !artikelCode) {
      return;
    }

    var key = locationKey + '|' + artikelCode;
    if (!grouped[key]) {
      grouped[key] = Object.assign({
        locatie: locationKey,
        artikelCode: artikelCode,
        artikelOmschrijving: safeText(move.artikelOmschrijving),
        typeMateriaal: safeText(move.typeMateriaal || determineMaterialTypeFromArticle(artikelCode)),
        eenheid: safeText(move.eenheid),
        voorraad: 0,
        laatsteMutatie: safeText(move.datumBoeking),
        laatsteMutatieRaw: safeText(move.datumBoekingRaw),
      }, cfg.rowDecorator(move, locationKey) || {});
    }

    grouped[key].voorraad += safeNumber(cfg.qtyExtractor(move), 0);

    if (safeText(move.datumBoekingRaw) > safeText(grouped[key].laatsteMutatieRaw)) {
      grouped[key].laatsteMutatie = safeText(move.datumBoeking);
      grouped[key].laatsteMutatieRaw = safeText(move.datumBoekingRaw);
    }

    if (!grouped[key].artikelOmschrijving) {
      grouped[key].artikelOmschrijving = safeText(move.artikelOmschrijving);
    }
    if (!grouped[key].typeMateriaal) {
      grouped[key].typeMateriaal = safeText(move.typeMateriaal || determineMaterialTypeFromArticle(artikelCode));
    }
    if (!grouped[key].eenheid) {
      grouped[key].eenheid = safeText(move.eenheid);
    }
  });

  return Object.keys(grouped)
    .map(function (key) { return grouped[key]; })
    .filter(function (item) {
      return safeNumber(item.voorraad, 0) !== 0;
    })
    .sort(function (a, b) {
      return (
        safeText(a.locatie).localeCompare(safeText(b.locatie)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

/* ---------------------------------------------------------
   Central warehouse snapshot
   --------------------------------------------------------- */

function buildCentralWarehouseStockRows() {
  var movements = getAllSnapshotMovements_();

  return buildLocationStockRows(movements, {
    includePredicate: function (move) {
      return isCentralLocation_(move.locatieNaar) || isCentralLocation_(move.locatieVan);
    },
    locationExtractor: function () {
      return safeText(LOCATION.CENTRAL);
    },
    qtyExtractor: function (move) {
      var qty = 0;
      if (isCentralLocation_(move.locatieNaar)) {
        qty += safeNumber(move.aantalIn || move.nettoAantal, 0);
      }
      if (isCentralLocation_(move.locatieVan)) {
        qty -= safeNumber(move.aantalUit || Math.abs(move.nettoAantal), 0);
      }
      return qty;
    },
    rowDecorator: function () {
      return {
        locatieLabel: 'Centraal magazijn',
        voorraadCentraal: 0,
      };
    },
  }).map(function (row) {
    row.voorraadCentraal = safeNumber(row.voorraad, 0);
    return row;
  });
}

function buildCentralStockMap() {
  var map = {};
  buildCentralWarehouseStockRows().forEach(function (row) {
    map[row.artikelCode] = row;
  });
  return map;
}

function buildCentralWarehouseSummary() {
  var rows = buildCentralWarehouseStockRows();

  return {
    artikels: rows.length,
    totaalAantal: rows.reduce(function (sum, row) {
      return sum + safeNumber(row.voorraadCentraal, 0);
    }, 0),
    grabbelAantal: rows
      .filter(function (row) { return safeText(row.typeMateriaal) === 'Grabbel'; })
      .reduce(function (sum, row) {
        return sum + safeNumber(row.voorraadCentraal, 0);
      }, 0),
    behoefteAantal: rows
      .filter(function (row) { return safeText(row.typeMateriaal) === 'Behoefte'; })
      .reduce(function (sum, row) {
        return sum + safeNumber(row.voorraadCentraal, 0);
      }, 0),
    artikelenMetVoorraad: rows.filter(function (row) {
      return safeNumber(row.voorraadCentraal, 0) > 0;
    }).length,
  };
}

function getCentralWarehouseOverview() {
  return {
    summary: buildCentralWarehouseSummary(),
    items: buildCentralWarehouseStockRows(),
  };
}

function rebuildCentralWarehouseOverview() {
  var rows = buildCentralWarehouseStockRows();
  var headers = [
    'Locatie',
    'LocatieLabel',
    'ArtikelCode',
    'ArtikelOmschrijving',
    'TypeMateriaal',
    'Eenheid',
    'Voorraad',
    'VoorraadCentraal',
    'LaatsteMutatie',
    'LaatsteMutatieRaw',
  ];

  writeFullTable(
    getCentralStockOverviewTab_(),
    headers,
    rows.map(function (row) {
      return buildRowFromHeaders(headers, {
        Locatie: row.locatie,
        LocatieLabel: row.locatieLabel,
        ArtikelCode: row.artikelCode,
        ArtikelOmschrijving: row.artikelOmschrijving,
        TypeMateriaal: row.typeMateriaal,
        Eenheid: row.eenheid,
        Voorraad: row.voorraad,
        VoorraadCentraal: row.voorraadCentraal,
        LaatsteMutatie: row.laatsteMutatie,
        LaatsteMutatieRaw: row.laatsteMutatieRaw,
      });
    })
  );

  return {
    updated: true,
    rowCount: rows.length,
  };
}

/* ---------------------------------------------------------
   Bus stock snapshot
   --------------------------------------------------------- */

function buildAllBusStockRows() {
  var movements = getAllSnapshotMovements_();

  return buildLocationStockRows(movements, {
    includePredicate: function (move) {
      return isBusLocation_(move.locatieNaar) || isBusLocation_(move.locatieVan);
    },
    locationExtractor: function (move) {
      if (isBusLocation_(move.locatieNaar)) {
        return safeText(move.locatieNaar);
      }
      return safeText(move.locatieVan);
    },
    qtyExtractor: function (move) {
      var qty = 0;
      if (isBusLocation_(move.locatieNaar)) {
        qty += safeNumber(move.aantalIn || move.nettoAantal, 0);
      }
      if (isBusLocation_(move.locatieVan)) {
        qty -= safeNumber(move.aantalUit || Math.abs(move.nettoAantal), 0);
      }
      return qty;
    },
    rowDecorator: function (move, locationKey) {
      return {
        busLocatie: locationKey,
        techniekerCode: parseBusLocation_(locationKey),
        voorraadBus: 0,
      };
    },
  }).map(function (row) {
    row.voorraadBus = safeNumber(row.voorraad, 0);
    return row;
  });
}

function buildBusStockRowsForTechnician(technicianCode) {
  var code = safeText(technicianCode);
  return buildAllBusStockRows().filter(function (row) {
    return safeText(row.techniekerCode) === code;
  });
}

function buildBusStockMapForTechnician(technicianCode) {
  var map = {};
  buildBusStockRowsForTechnician(technicianCode).forEach(function (row) {
    map[row.artikelCode] = row;
  });
  return map;
}

function buildBusStockSummaryForTechnician(technicianCode) {
  var rows = buildBusStockRowsForTechnician(technicianCode);

  return {
    techniekerCode: safeText(technicianCode),
    artikels: rows.length,
    totaalAantal: rows.reduce(function (sum, row) {
      return sum + safeNumber(row.voorraadBus, 0);
    }, 0),
  };
}

/* ---------------------------------------------------------
   Mobile warehouse stock snapshot
   --------------------------------------------------------- */

function buildAllMobileWarehouseSnapshotRows() {
  var movements = getAllSnapshotMovements_();

  return buildLocationStockRows(movements, {
    includePredicate: function (move) {
      return isMobileWarehouseLocation_(move.locatieNaar) || isMobileWarehouseLocation_(move.locatieVan);
    },
    locationExtractor: function (move) {
      if (isMobileWarehouseLocation_(move.locatieNaar)) {
        return safeText(move.locatieNaar);
      }
      return safeText(move.locatieVan);
    },
    qtyExtractor: function (move) {
      var qty = 0;
      if (isMobileWarehouseLocation_(move.locatieNaar)) {
        qty += safeNumber(move.aantalIn || move.nettoAantal, 0);
      }
      if (isMobileWarehouseLocation_(move.locatieVan)) {
        qty -= safeNumber(move.aantalUit || Math.abs(move.nettoAantal), 0);
      }
      return qty;
    },
    rowDecorator: function (move, locationKey) {
      return {
        mobileWarehouseCode: parseMobileWarehouseLocation_(locationKey),
        voorraadMobiel: 0,
      };
    },
  }).map(function (row) {
    row.voorraadMobiel = safeNumber(row.voorraad, 0);
    return row;
  });
}

function buildMobileWarehouseSnapshotRows(mobileWarehouseCode) {
  var code = safeText(mobileWarehouseCode);
  return buildAllMobileWarehouseSnapshotRows().filter(function (row) {
    return safeText(row.mobileWarehouseCode) === code;
  });
}

function buildMobileWarehouseSnapshotMap(mobileWarehouseCode) {
  var map = {};
  buildMobileWarehouseSnapshotRows(mobileWarehouseCode).forEach(function (row) {
    map[row.artikelCode] = row;
  });
  return map;
}

/* ---------------------------------------------------------
   Combined stock views
   --------------------------------------------------------- */

function buildCombinedStockRowsForTechnician(technicianCode) {
  var centralMap = buildCentralStockMap();
  var busMap = buildBusStockMapForTechnician(technicianCode);
  var keys = {};

  Object.keys(centralMap).forEach(function (key) { keys[key] = true; });
  Object.keys(busMap).forEach(function (key) { keys[key] = true; });

  return Object.keys(keys)
    .map(function (artikelCode) {
      var central = centralMap[artikelCode] || {};
      var bus = busMap[artikelCode] || {};

      return {
        artikelCode: artikelCode,
        artikelOmschrijving: safeText(central.artikelOmschrijving || bus.artikelOmschrijving),
        typeMateriaal: safeText(central.typeMateriaal || bus.typeMateriaal),
        eenheid: safeText(central.eenheid || bus.eenheid),
        voorraadCentraal: safeNumber(central.voorraadCentraal, 0),
        voorraadBus: safeNumber(bus.voorraadBus, 0),
        voorraadTotaal:
          safeNumber(central.voorraadCentraal, 0) +
          safeNumber(bus.voorraadBus, 0),
      };
    })
    .sort(function (a, b) {
      return (
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function buildCombinedStockSummaryForTechnician(technicianCode) {
  var rows = buildCombinedStockRowsForTechnician(technicianCode);

  return {
    techniekerCode: safeText(technicianCode),
    artikels: rows.length,
    totaalCentraal: rows.reduce(function (sum, row) {
      return sum + safeNumber(row.voorraadCentraal, 0);
    }, 0),
    totaalBus: rows.reduce(function (sum, row) {
      return sum + safeNumber(row.voorraadBus, 0);
    }, 0),
    totaalSamen: rows.reduce(function (sum, row) {
      return sum + safeNumber(row.voorraadTotaal, 0);
    }, 0),
  };
}

/* ---------------------------------------------------------
   Queries for screens
   --------------------------------------------------------- */

function filterCentralStockRows_(rows, filters) {
  var f = filters || {};
  var query = safeText(f.query).toLowerCase();
  var typeMateriaal = safeText(f.typeMateriaal);
  var inStockOnly = isTrue(f.inStockOnly);

  return (rows || []).filter(function (row) {
    if (query) {
      var hit = [
        row.artikelCode,
        row.artikelOmschrijving,
        row.typeMateriaal,
      ].some(function (value) {
        return safeText(value).toLowerCase().indexOf(query) >= 0;
      });

      if (!hit) return false;
    }

    if (typeMateriaal && safeText(row.typeMateriaal) !== typeMateriaal) {
      return false;
    }

    if (inStockOnly && safeNumber(row.voorraadCentraal, 0) <= 0) {
      return false;
    }

    return true;
  });
}

function getCentralStockData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!canViewCentralStock(actor)) {
    throw new Error('Geen rechten om centrale stock te bekijken.');
  }

  var rows = filterCentralStockRows_(buildCentralWarehouseStockRows(), payload.filters || {});
  var summary = buildCentralWarehouseSummary();

  return {
    items: rows,
    stock: rows,
    summary: summary,
  };
}
