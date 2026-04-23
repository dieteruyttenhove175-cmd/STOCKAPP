/* =========================================================
   52_MobileWarehouseService.gs
   Refactor for mobile warehouse backend
   - keeps mobile warehouse access + dashboard in one place
   - aligns payload names with the new mobile UI shell
   ========================================================= */

/* ---------------------------------------------------------
   Helpers locatie
   --------------------------------------------------------- */

function getMobileWarehouseLocationCode(code) {
  return 'Mobiel:' + safeText(code);
}

function parseMobileWarehouseLocation(location) {
  var text = safeText(location);
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
    opmerking: safeText(row.Opmerking),
  };
}

/* ---------------------------------------------------------
   Readers
   --------------------------------------------------------- */

function getAllMobileWarehouses() {
  return readObjectsSafe(TABS.MOBILE_WAREHOUSES)
    .map(mapMobileWarehouse)
    .filter(function (x) { return x.active; })
    .sort(function (a, b) {
      return safeText(a.naam).localeCompare(safeText(b.naam));
    });
}

function getMobileWarehouseByCode(code) {
  var value = safeText(code);
  if (!value) return null;
  return getAllMobileWarehouses().find(function (x) {
    return safeText(x.code) === value;
  }) || null;
}

function getDefaultMobileWarehouseCode() {
  var rows = getAllMobileWarehouses();
  if (!rows.length) return '';
  return safeText(rows[0].code);
}

function getEffectiveMobileWarehouseCodeForUser(user, requestedCode) {
  var asked = safeText(requestedCode);
  var own = safeText((user && user.mobileWarehouseCode) || '');

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
  var user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten voor mobiel magazijn.');
  }

  var mobileWarehouseCode = getEffectiveMobileWarehouseCodeForUser(user, requestedCode);
  if (!mobileWarehouseCode) {
    throw new Error('Geen mobiel magazijn gekoppeld of beschikbaar.');
  }

  var warehouse = getMobileWarehouseByCode(mobileWarehouseCode);
  if (!warehouse) {
    throw new Error('Mobiel magazijn niet gevonden.');
  }

  return {
    user: user,
    mobileWarehouseCode: mobileWarehouseCode,
    warehouse: warehouse,
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
    type: type,
    titel: title,
    bericht: message,
    bronType: bronType,
    bronId: bronId,
  });
}

function getNotificationsForMobileWarehouse(mobileWarehouseCode) {
  var code = safeText(mobileWarehouseCode);

  return readObjectsSafe(TABS.NOTIFICATIONS)
    .map(mapNotification)
    .filter(function (item) {
      return item.rol === NOTIFICATION_ROLE.MOBILE_WAREHOUSE &&
        (!code ||
          safeText(item.ontvangerCode) === code ||
          safeText(item.ontvangerCode) === 'MOBIEL');
    })
    .sort(function (a, b) {
      return safeText(b.aangemaaktOpRaw).localeCompare(safeText(a.aangemaaktOpRaw));
    });
}

function markAllMobileWarehouseNotificationsRead(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var access = assertMobileWarehouseAccess(sessionId, safeText(payload.mobileWarehouseCode));
  var mobileWarehouseCode = access.mobileWarehouseCode;

  var sheet = getSheetOrThrow(TABS.NOTIFICATIONS);
  var values = getAllValues(TABS.NOTIFICATIONS);
  if (!values.length) return { updatedCount: 0 };

  var headers = values[0];
  var dataRows = values.slice(1);

  var updatedCount = 0;
  var updatedRows = dataRows.map(function (row) {
    var obj = rowToObject(headers, row);
    var mapped = mapNotification(obj);

    var matchesRole = mapped.rol === NOTIFICATION_ROLE.MOBILE_WAREHOUSE;
    var matchesWarehouse =
      !mobileWarehouseCode ||
      safeText(mapped.ontvangerCode) === mobileWarehouseCode ||
      safeText(mapped.ontvangerCode) === 'MOBIEL';

    if (!matchesRole || !matchesWarehouse) {
      return row;
    }

    var currentStatus = safeText(obj.Status || obj.status);
    if (currentStatus && currentStatus !== NOTIFICATION_STATUS.OPEN) {
      return row;
    }

    obj.Status = NOTIFICATION_STATUS.READ;
    obj.GelezenOp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
    updatedCount += 1;
    return buildRowFromHeaders(headers, obj);
  });

  writeFullTable(TABS.NOTIFICATIONS, headers, updatedRows);

  if (updatedCount) {
    writeAudit({
      actie: 'MARK_MOBILE_NOTIFICATIONS_READ',
      actor: access.user,
      documentType: 'Notificaties',
      documentId: mobileWarehouseCode || 'MOBIEL',
      details: {
        updatedCount: updatedCount,
        mobileWarehouseCode: mobileWarehouseCode,
      },
    });
  }

  return {
    updatedCount: updatedCount,
    mobileWarehouseCode: mobileWarehouseCode,
  };
}

/* ---------------------------------------------------------
   Stock mobiel magazijn opbouwen
   --------------------------------------------------------- */

function buildAllMobileWarehouseStockRows() {
  var articles = readObjectsSafe(TABS.SUPPLIER_ARTICLES).map(mapSupplierArticle);
  var articleMap = {};
  articles.forEach(function (item) {
    articleMap[item.artikelCode] = item;
  });

  var moves = readObjectsSafe(TABS.WAREHOUSE_MOVEMENTS).map(mapWarehouseMovement);
  var grouped = {};

  moves.forEach(function (move) {
    var code = safeText(move.artikelCode);
    if (!code) return;

    var from = safeText(move.locatieVan);
    var to = safeText(move.locatieNaar);
    var qtyIn = safeNumber(move.aantalIn, 0);
    var qtyOut = safeNumber(move.aantalUit, 0);
    var net = safeNumber(move.nettoAantal, 0);

    function ensureGroup(mwCode) {
      var key = mwCode + '|' + code;
      if (!grouped[key]) {
        grouped[key] = {
          mobileWarehouseCode: mwCode,
          mobileWarehouseNaam: (getMobileWarehouseByCode(mwCode) || {}).naam || mwCode,
          artikelCode: code,
          artikelOmschrijving:
            safeText(move.artikelOmschrijving) ||
            safeText(articleMap[code] && articleMap[code].artikelOmschrijving),
          eenheid:
            safeText(move.eenheid) ||
            safeText(articleMap[code] && articleMap[code].eenheid),
          typeMateriaal:
            safeText(move.typeMateriaal) ||
            determineMaterialTypeFromArticle(code),
          voorraadMobiel: 0,
          laatsteMutatie: safeText(move.datumBoeking),
          laatsteMutatieRaw: safeText(move.datumBoekingRaw),
        };
      }
      return grouped[key];
    }

    if (isMobileWarehouseLocation(to)) {
      var mwTo = parseMobileWarehouseLocation(to);
      var rowTo = ensureGroup(mwTo);
      rowTo.voorraadMobiel += qtyIn || Math.abs(net);

      if (safeText(move.datumBoekingRaw) > safeText(rowTo.laatsteMutatieRaw)) {
        rowTo.laatsteMutatie = safeText(move.datumBoeking);
        rowTo.laatsteMutatieRaw = safeText(move.datumBoekingRaw);
      }
    }

    if (isMobileWarehouseLocation(from)) {
      var mwFrom = parseMobileWarehouseLocation(from);
      var rowFrom = ensureGroup(mwFrom);
      rowFrom.voorraadMobiel -= qtyOut || Math.abs(net);

      if (safeText(move.datumBoekingRaw) > safeText(rowFrom.laatsteMutatieRaw)) {
        rowFrom.laatsteMutatie = safeText(move.datumBoeking);
        rowFrom.laatsteMutatieRaw = safeText(move.datumBoekingRaw);
      }
    }
  });

  return Object.keys(grouped)
    .map(function (key) { return grouped[key]; })
    .filter(function (item) { return safeNumber(item.voorraadMobiel, 0) !== 0; })
    .sort(function (a, b) {
      return (
        safeText(a.mobileWarehouseNaam).localeCompare(safeText(b.mobileWarehouseNaam)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
      );
    });
}

function buildMobileWarehouseStockRows(mobileWarehouseCode) {
  var code = safeText(mobileWarehouseCode);
  return buildAllMobileWarehouseStockRows().filter(function (item) {
    return safeText(item.mobileWarehouseCode) === code;
  });
}

function buildMobileWarehouseStockMap(mobileWarehouseCode) {
  var map = {};
  buildMobileWarehouseStockRows(mobileWarehouseCode).forEach(function (item) {
    map[item.artikelCode] = item;
  });
  return map;
}

function buildMobileWarehouseSummary(mobileWarehouseCode) {
  var rows = buildMobileWarehouseStockRows(mobileWarehouseCode);

  return {
    mobileWarehouseCode: safeText(mobileWarehouseCode),
    artikels: rows.length,
    totaalAantal: rows.reduce(function (sum, row) {
      return sum + safeNumber(row.voorraadMobiel, 0);
    }, 0),
    grabbelAantal: rows
      .filter(function (row) { return safeText(row.typeMateriaal) === 'Grabbel'; })
      .reduce(function (sum, row) {
        return sum + safeNumber(row.voorraadMobiel, 0);
      }, 0),
    behoefteAantal: rows
      .filter(function (row) { return safeText(row.typeMateriaal) === 'Behoefte'; })
      .reduce(function (sum, row) {
        return sum + safeNumber(row.voorraadMobiel, 0);
      }, 0),
  };
}

/* ---------------------------------------------------------
   Aanvragen voor mobiel magazijn
   --------------------------------------------------------- */

function buildMobileWarehouseRequestRows(mobileWarehouseCode) {
  var code = safeText(mobileWarehouseCode);

  if (
    typeof getAllMobileRequests !== 'function' ||
    typeof getAllMobileRequestLines !== 'function' ||
    typeof buildMobileRequestsWithLines !== 'function'
  ) {
    return [];
  }

  return buildMobileRequestsWithLines(getAllMobileRequests(), getAllMobileRequestLines())
    .filter(function (item) {
      return !code || safeText(item.mobileWarehouseCode) === code;
    })
    .sort(function (a, b) {
      return (
        (safeText(b.documentDatumIso) + ' ' + safeText(b.aanvraagId)).localeCompare(
          safeText(a.documentDatumIso) + ' ' + safeText(a.aanvraagId)
        )
      );
    });
}

function buildMobileWarehouseRequestQueue(mobileWarehouseCode) {
  return buildMobileWarehouseRequestRows(mobileWarehouseCode).filter(function (item) {
    return [
      MOBILE_REQUEST_STATUS.SUBMITTED,
      MOBILE_REQUEST_STATUS.APPROVED,
      MOBILE_REQUEST_STATUS.BOOKED,
    ].indexOf(safeText(item.status)) >= 0;
  });
}

/* ---------------------------------------------------------
   Transfer history
   --------------------------------------------------------- */

function buildMobileWarehouseTransferHistory(mobileWarehouseCode) {
  var code = safeText(mobileWarehouseCode);
  if (!code) return [];

  if (typeof getAllTransfers === 'function') {
    return getAllTransfers()
      .filter(function (item) {
        return (
          safeText(item.vanLocatie) === getMobileWarehouseLocationCode(code) ||
          safeText(item.naarLocatie) === getMobileWarehouseLocationCode(code)
        );
      })
      .sort(function (a, b) {
        return safeText(b.documentDatumIso || b.documentDatum).localeCompare(
          safeText(a.documentDatumIso || a.documentDatum)
        );
      });
  }

  if (!TABS.TRANSFERS) return [];

  return readObjectsSafe(TABS.TRANSFERS)
    .map(function (row) {
      return {
        transferId: safeText(row.TransferID || row.TransferId || row.ID),
        documentDatum: safeText(row.DocumentDatum || row.Datum),
        documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum || row.Datum),
        vanLocatie: safeText(row.VanLocatie),
        naarLocatie: safeText(row.NaarLocatie),
        status: safeText(row.Status),
      };
    })
    .filter(function (item) {
      return (
        safeText(item.vanLocatie) === getMobileWarehouseLocationCode(code) ||
        safeText(item.naarLocatie) === getMobileWarehouseLocationCode(code)
      );
    })
    .sort(function (a, b) {
      return safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso));
    });
}

/* ---------------------------------------------------------
   Vaste filters / snelle zoeklogica
   --------------------------------------------------------- */

function filterMobileWarehouseStockRows(rows, filters) {
  var source = rows || [];
  var f = filters || {};
  var materialType = safeText(f.materialType);
  var articleCodePrefix = safeText(f.articleCodePrefix).toLowerCase();
  var inStockOnly = !!f.inStockOnly;
  var minQty =
    f.minQty === '' || f.minQty === null || f.minQty === undefined
      ? null
      : Number(f.minQty);
  var sortBy = safeText(f.sortBy || 'artikel');
  var sortDir = safeText(f.sortDir || 'asc').toLowerCase();

  var result = source.slice();

  if (materialType) {
    result = result.filter(function (row) {
      return safeText(row.typeMateriaal) === materialType;
    });
  }

  if (articleCodePrefix) {
    result = result.filter(function (row) {
      return safeText(row.artikelCode).toLowerCase().startsWith(articleCodePrefix);
    });
  }

  if (inStockOnly) {
    result = result.filter(function (row) {
      return safeNumber(row.voorraadMobiel, 0) > 0;
    });
  }

  if (minQty !== null && !isNaN(minQty)) {
    result = result.filter(function (row) {
      return safeNumber(row.voorraadMobiel, 0) >= minQty;
    });
  }

  var factor = sortDir === 'desc' ? -1 : 1;
  result.sort(function (a, b) {
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

  var sessionId = getPayloadSessionId(payload);
  var access = assertMobileWarehouseAccess(sessionId, safeText(payload.mobileWarehouseCode));
  var rows = buildMobileWarehouseStockRows(access.mobileWarehouseCode);

  return filterMobileWarehouseStockRows(rows, payload.filters || {});
}

/* ---------------------------------------------------------
   Dashboarddata
   --------------------------------------------------------- */

function getMobileWarehouseDashboardData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var access = assertMobileWarehouseAccess(sessionId, safeText(payload.mobileWarehouseCode));
  var warehouseCode = access.mobileWarehouseCode;

  var stockRows = buildMobileWarehouseStockRows(warehouseCode);
  var filteredStockRows = filterMobileWarehouseStockRows(stockRows, payload.filters || {});
  var requestQueue = buildMobileWarehouseRequestQueue(warehouseCode);
  var allRequests = buildMobileWarehouseRequestRows(warehouseCode);

  return {
    warehouse: access.warehouse,
    selectedMobileWarehouseCode: warehouseCode,
    availableWarehouses: getAllMobileWarehouses(),
    summary: buildMobileWarehouseSummary(warehouseCode),
    stockRows: filteredStockRows,
    centralWarehouse:
      typeof getCentralWarehouseOverview === 'function'
        ? getCentralWarehouseOverview()
        : null,
    technicians:
      typeof getActiveTechnicians === 'function'
        ? getActiveTechnicians().map(function (t) {
            return { code: t.code, naam: t.naam };
          })
        : [],
    requestQueue: requestQueue,
    requestHistory: allRequests,
    transferHistory: buildMobileWarehouseTransferHistory(warehouseCode),
    notifications: getNotificationsForMobileWarehouse(warehouseCode),
    recurringBusCountSummary:
      typeof buildBusCountTriggerSummary === 'function'
        ? buildBusCountTriggerSummary()
        : null,
    generatedAt: Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm'),
  };
}

/* ---------------------------------------------------------
   Directe transfers
   Vereist bestaand transferblok met createTransfer / saveTransferLines
   --------------------------------------------------------- */

function assertTransferServiceAvailable() {
  if (typeof createTransfer !== 'function') {
    throw new Error('Transferservice ontbreekt. Werk eerst het transferblok in.');
  }
  if (typeof saveTransferLines !== 'function') {
    throw new Error('Transferlijnservice ontbreekt. Werk eerst het transferblok in.');
  }
}

function normalizeMobileToBusTransferPayload(payload) {
  var normalized = Object.assign({}, payload || {});
  normalized.doelTechniekerCode = safeText(
    payload.doelTechniekerCode || payload.technicianCode || payload.techCode
  );
  normalized.documentDatum = safeText(
    payload.documentDatum || payload.documentDate
  );
  normalized.reden = safeText(payload.reden || payload.reason);
  normalized.opmerking = safeText(payload.opmerking || payload.remark);
  normalized.lines = Array.isArray(payload.lines) ? payload.lines : [];
  return normalized;
}

function createDirectMobileToBusTransfer(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  assertTransferServiceAvailable();

  var normalized = normalizeMobileToBusTransferPayload(payload);
  var sessionId = getPayloadSessionId(payload);
  var access = assertMobileWarehouseAccess(sessionId, safeText(payload.mobileWarehouseCode));

  var doelTechniekerCode = safeText(normalized.doelTechniekerCode);
  var documentDatum = safeText(normalized.documentDatum);
  var reden = safeText(normalized.reden);
  var opmerking = safeText(normalized.opmerking);
  var actor = safeText(payload.actor || access.user.naam || access.user.email || 'MobielMagazijn');
  var lines = normalized.lines;

  if (!doelTechniekerCode) throw new Error('DoelTechniekerCode is verplicht.');
  if (!documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!reden) throw new Error('Reden is verplicht.');
  if (!lines.length) throw new Error('Geen transferlijnen ontvangen.');

  var stockMap = buildMobileWarehouseStockMap(access.mobileWarehouseCode);

  lines.forEach(function (line) {
    var code = safeText(line.artikelCode);
    var qty = safeNumber(line.aantal, 0);
    var current = stockMap[code] ? safeNumber(stockMap[code].voorraadMobiel, 0) : 0;

    if (!code) throw new Error('ArtikelCode ontbreekt.');
    if (qty <= 0) throw new Error('Aantal moet groter zijn dan 0 voor artikel ' + code + '.');
    if (qty > current) {
      throw new Error(
        'Onvoldoende voorraad in mobiel magazijn voor artikel ' + code +
        '. Beschikbaar: ' + current + ', gevraagd: ' + qty + '.'
      );
    }
  });

  var transferPayload = {
    sessionId: sessionId,
    flowType: 'MOBILE_TO_BUS',
    vanLocatie: getMobileWarehouseLocationCode(access.mobileWarehouseCode),
    naarLocatie:
      typeof getBusLocationCode === 'function'
        ? getBusLocationCode(doelTechniekerCode)
        : ('Bus:' + doelTechniekerCode),
    doelTechniekerCode: doelTechniekerCode,
    documentDatum: documentDatum,
    reden: reden,
    opmerking: opmerking,
    actor: actor,
  };

  var transfer = createTransfer(transferPayload);

  saveTransferLines({
    sessionId: sessionId,
    transferId: transfer.transferId || transfer.TransferID || transfer.id,
    lines: lines.map(function (line) {
      return {
        artikelCode: safeText(line.artikelCode),
        artikelOmschrijving: safeText(line.artikelOmschrijving),
        eenheid: safeText(line.eenheid) || 'Stuk',
        aantal: safeNumber(line.aantal, 0),
      };
    }),
  });

  if (typeof submitTransfer === 'function' && isTrue(payload.autoSubmit)) {
    submitTransfer({
      sessionId: sessionId,
      transferId: transfer.transferId || transfer.TransferID || transfer.id,
    });
  }

  if (typeof pushTechnicianNotification === 'function') {
    pushTechnicianNotification(
      'Transfer',
      'Aanvulling vanuit mobiel magazijn',
      'Er werd een transfer klaargezet vanuit mobiel magazijn.',
      'Transfer',
      transfer.transferId || transfer.TransferID || transfer.id,
      doelTechniekerCode
    );
  }

  writeAudit({
    actie: 'CREATE_DIRECT_MOBILE_TO_BUS_TRANSFER',
    actor: access.user,
    documentType: 'Transfer',
    documentId: transfer.transferId || transfer.TransferID || transfer.id,
    details: {
      mobileWarehouseCode: access.mobileWarehouseCode,
      doelTechniekerCode: doelTechniekerCode,
      lineCount: lines.length,
    },
  });

  return transfer;
}
