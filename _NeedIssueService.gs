/* =========================================================
   42_NeedIssueService.gs
   Refactor: need issue core service
   Doel:
   - centrale documentlaag voor behoefte-uitgiftes
   - lijnen bewaren
   - indienen
   - boeken naar doeltechnieker of doelmagazijn
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getNeedIssueHeaderTab_() {
  return TABS.NEED_ISSUES || 'BehoefteUitgiftes';
}

function getNeedIssueLineTab_() {
  return TABS.NEED_ISSUE_LINES || 'BehoefteUitgifteLijnen';
}

function getNeedIssueStatusOpen_() {
  if (typeof NEED_ISSUE_STATUS !== 'undefined' && NEED_ISSUE_STATUS && NEED_ISSUE_STATUS.OPEN) {
    return NEED_ISSUE_STATUS.OPEN;
  }
  return 'Open';
}

function getNeedIssueStatusSubmitted_() {
  if (typeof NEED_ISSUE_STATUS !== 'undefined' && NEED_ISSUE_STATUS && NEED_ISSUE_STATUS.SUBMITTED) {
    return NEED_ISSUE_STATUS.SUBMITTED;
  }
  return 'Ingediend';
}

function getNeedIssueStatusBooked_() {
  if (typeof NEED_ISSUE_STATUS !== 'undefined' && NEED_ISSUE_STATUS && NEED_ISSUE_STATUS.BOOKED) {
    return NEED_ISSUE_STATUS.BOOKED;
  }
  return 'Geboekt';
}

function getNeedIssueStatusClosed_() {
  if (typeof NEED_ISSUE_STATUS !== 'undefined' && NEED_ISSUE_STATUS && NEED_ISSUE_STATUS.CLOSED) {
    return NEED_ISSUE_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function getNeedIssueTargetTypeTechnician_() {
  return 'TECHNICIAN';
}

function getNeedIssueTargetTypeMobileWarehouse_() {
  return 'MOBILE_WAREHOUSE';
}

function isNeedIssueEditable_(status) {
  var value = safeText(status);
  return !value || value === getNeedIssueStatusOpen_();
}

function makeNeedIssueId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'NIS-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function makeNeedIssueLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'NIL-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapNeedIssueHeader(row) {
  return {
    needIssueId: safeText(row.NeedIssueID || row.NeedIssueId || row.BehoefteUitgifteID || row.ID),
    targetType: safeText(row.TargetType || getNeedIssueTargetTypeTechnician_()),
    targetCode: safeText(row.TargetCode),
    targetName: safeText(row.TargetName),
    vanLocatie: safeText(row.VanLocatie || LOCATION.CENTRAL),
    naarLocatie: safeText(row.NaarLocatie),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    status: safeText(row.Status),
    actor: safeText(row.Actor),
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),
    typeMateriaal: safeText(row.TypeMateriaal),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    ingediendOp: safeText(row.IngediendOp),
    geboektOp: safeText(row.GeboektOp),
  };
}

function mapNeedIssueLine(row) {
  return {
    needIssueLineId: safeText(row.NeedIssueLineID || row.NeedIssueLineId || row.BehoefteUitgifteLijnID || row.ID),
    needIssueId: safeText(row.NeedIssueID || row.NeedIssueId || row.BehoefteUitgifteID),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode || row.ArtikelNr))),
    eenheid: safeText(row.Eenheid || row.Unit),
    aantal: safeNumber(row.Aantal || row.Quantity, 0),
    opmerking: safeText(row.Opmerking),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllNeedIssueHeaders() {
  return readObjectsSafe(getNeedIssueHeaderTab_())
    .map(mapNeedIssueHeader)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.needIssueId).localeCompare(safeText(a.needIssueId))
      );
    });
}

function getAllNeedIssueLines() {
  return readObjectsSafe(getNeedIssueLineTab_())
    .map(mapNeedIssueLine)
    .sort(function (a, b) {
      return (
        safeText(a.needIssueId).localeCompare(safeText(b.needIssueId)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
      );
    });
}

function getNeedIssueHeaderById(needIssueId) {
  var id = safeText(needIssueId);
  if (!id) return null;

  return getAllNeedIssueHeaders().find(function (item) {
    return safeText(item.needIssueId) === id;
  }) || null;
}

function getNeedIssueLinesById(needIssueId) {
  var id = safeText(needIssueId);
  return getAllNeedIssueLines().filter(function (item) {
    return safeText(item.needIssueId) === id;
  });
}

function buildNeedIssuesWithLines(headers, lines) {
  var lineMap = {};

  (lines || []).forEach(function (line) {
    var id = safeText(line.needIssueId);
    if (!lineMap[id]) lineMap[id] = [];
    lineMap[id].push(line);
  });

  return (headers || []).map(function (header) {
    var issueLines = lineMap[safeText(header.needIssueId)] || [];

    return Object.assign({}, header, {
      lines: issueLines,
      lineCount: issueLines.length,
      totaalAantal: issueLines.reduce(function (sum, line) {
        return sum + safeNumber(line.aantal, 0);
      }, 0),
    });
  });
}

function getNeedIssuesWithLines() {
  return buildNeedIssuesWithLines(getAllNeedIssueHeaders(), getAllNeedIssueLines());
}

function getNeedIssueWithLines(needIssueId) {
  var header = getNeedIssueHeaderById(needIssueId);
  if (!header) return null;
  return buildNeedIssuesWithLines([header], getNeedIssueLinesById(needIssueId))[0] || null;
}

/* ---------------------------------------------------------
   Normalization / helpers
   --------------------------------------------------------- */

function buildNeedIssueTargetLocation_(targetType, targetCode) {
  var type = safeText(targetType);
  var code = safeText(targetCode);

  if (type === getNeedIssueTargetTypeTechnician_()) {
    if (typeof getBusLocationCode === 'function') {
      return getBusLocationCode(code);
    }
    return 'Bus:' + code;
  }

  if (type === getNeedIssueTargetTypeMobileWarehouse_()) {
    if (typeof getMobileWarehouseLocationCode === 'function') {
      return getMobileWarehouseLocationCode(code);
    }
    return 'Mobiel:' + code;
  }

  throw new Error('Onbekend doellocatietype.');
}

function resolveNeedIssueTargetName_(targetType, targetCode) {
  var type = safeText(targetType);
  var code = safeText(targetCode);

  if (type === getNeedIssueTargetTypeTechnician_()) {
    if (typeof resolveTechnicianByRef === 'function') {
      var tech = resolveTechnicianByRef(code);
      return tech ? safeText(tech.naam) : code;
    }
    return code;
  }

  if (type === getNeedIssueTargetTypeMobileWarehouse_()) {
    if (typeof getMobileWarehouseByCode === 'function') {
      var mw = getMobileWarehouseByCode(code);
      return mw ? safeText(mw.naam) : code;
    }
    return code;
  }

  return code;
}

function normalizeNeedIssueHeaderPayload_(payload) {
  payload = payload || {};

  var targetType = safeText(payload.targetType || payload.doelType || getNeedIssueTargetTypeTechnician_());
  var targetCode = safeText(payload.targetCode || payload.doelCode || payload.techniekerCode || payload.mobileWarehouseCode);

  return {
    sessionId: getPayloadSessionId(payload),
    targetType: targetType,
    targetCode: targetCode,
    targetName: safeText(payload.targetName || resolveNeedIssueTargetName_(targetType, targetCode)),
    vanLocatie: safeText(payload.vanLocatie || LOCATION.CENTRAL),
    naarLocatie: safeText(payload.naarLocatie || buildNeedIssueTargetLocation_(targetType, targetCode)),
    documentDatum: safeText(payload.documentDatum || payload.documentDate),
    reden: safeText(payload.reden || payload.reason),
    opmerking: safeText(payload.opmerking || payload.remark),
    actor: safeText(payload.actor),
    typeMateriaal: safeText(payload.typeMateriaal || payload.materialType),
  };
}

function normalizeNeedIssueLines_(lines) {
  return (Array.isArray(lines) ? lines : []).map(function (line) {
    return {
      artikelCode: safeText(line.artikelCode || line.artikelNr),
      artikelOmschrijving: safeText(line.artikelOmschrijving || line.artikel),
      typeMateriaal: safeText(line.typeMateriaal || determineMaterialTypeFromArticle(safeText(line.artikelCode || line.artikelNr))),
      eenheid: safeText(line.eenheid || line.unit || 'Stuk'),
      aantal: safeNumber(line.aantal || line.quantity, 0),
      opmerking: safeText(line.opmerking),
    };
  });
}

function deriveNeedIssueMaterialType_(lines) {
  var types = {};
  (lines || []).forEach(function (line) {
    var t = safeText(line.typeMateriaal);
    if (t) {
      types[t] = true;
    }
  });

  var keys = Object.keys(types);
  if (!keys.length) return '';
  if (keys.length === 1) return keys[0];
  return 'Gemengd';
}

/* ---------------------------------------------------------
   Validation
   --------------------------------------------------------- */

function validateNeedIssueHeader_(payload) {
  if (!payload.sessionId) throw new Error('Sessie ontbreekt.');
  if (!payload.targetType) throw new Error('TargetType is verplicht.');
  if (!payload.targetCode) throw new Error('TargetCode is verplicht.');
  if (!payload.documentDatum) throw new Error('Documentdatum is verplicht.');
  if (!payload.reden) throw new Error('Reden is verplicht.');
}

function validateNeedIssueLines_(lines) {
  if (!lines.length) {
    throw new Error('Geen behoefte-uitgiftelijnen ontvangen.');
  }

  lines.forEach(function (line, index) {
    var rowNr = index + 1;

    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.aantal, 0) <= 0) {
      throw new Error('Aantal moet groter zijn dan 0 op lijn ' + rowNr + '.');
    }
  });

  return true;
}

function validateNeedIssueSourceStock_(header, lines) {
  if (typeof buildCentralStockMap !== 'function') {
    throw new Error('Central stock service ontbreekt. Werk eerst het stockblok in.');
  }

  var stockMap = buildCentralStockMap();

  lines.forEach(function (line) {
    var code = safeText(line.artikelCode);
    var requested = safeNumber(line.aantal, 0);
    var available = stockMap[code] ? safeNumber(stockMap[code].voorraadCentraal, 0) : 0;

    if (requested > available) {
      throw new Error(
        'Onvoldoende centrale voorraad voor artikel ' + code +
        '. Beschikbaar: ' + available + ', gevraagd: ' + requested + '.'
      );
    }
  });
}

/* ---------------------------------------------------------
   Access policy
   --------------------------------------------------------- */

function assertNeedIssueWriteAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER],
    'Geen rechten voor behoefte-uitgiftes.'
  );
  return user;
}

function assertNeedIssueReadAccess_(sessionId) {
  return assertNeedIssueWriteAccess_(sessionId);
}

/* ---------------------------------------------------------
   Create / save
   --------------------------------------------------------- */

function createNeedIssue(payload) {
  var normalized = normalizeNeedIssueHeaderPayload_(payload);
  validateNeedIssueHeader_(normalized);

  var actor = assertNeedIssueWriteAccess_(normalized.sessionId);
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var needIssueId = makeNeedIssueId_();

  var obj = {
    NeedIssueID: needIssueId,
    TargetType: normalized.targetType,
    TargetCode: normalized.targetCode,
    TargetName: normalized.targetName,
    VanLocatie: normalized.vanLocatie,
    NaarLocatie: normalized.naarLocatie,
    DocumentDatum: normalized.documentDatum,
    DocumentDatumIso: normalized.documentDatum,
    Status: getNeedIssueStatusOpen_(),
    Actor: normalized.actor || safeText(actor.naam || actor.email),
    Reden: normalized.reden,
    Opmerking: normalized.opmerking,
    TypeMateriaal: normalized.typeMateriaal,
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    IngediendOp: '',
    GeboektOp: '',
  };

  appendObjects(getNeedIssueHeaderTab_(), [obj]);

  writeAudit({
    actie: 'CREATE_NEED_ISSUE',
    actor: actor,
    documentType: 'BehoefteUitgifte',
    documentId: needIssueId,
    details: {
      targetType: obj.TargetType,
      targetCode: obj.TargetCode,
      documentDatum: obj.DocumentDatum,
    },
  });

  return mapNeedIssueHeader(obj);
}

function saveNeedIssueLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertNeedIssueWriteAccess_(sessionId);
  var needIssueId = safeText(payload.needIssueId);
  if (!needIssueId) throw new Error('NeedIssueId ontbreekt.');

  var header = getNeedIssueHeaderById(needIssueId);
  if (!header) throw new Error('Behoefte-uitgifte niet gevonden.');
  if (!isNeedIssueEditable_(header.status)) {
    throw new Error('Behoefte-uitgifte is niet meer bewerkbaar.');
  }

  var lines = normalizeNeedIssueLines_(payload.lines);
  validateNeedIssueLines_(lines);

  var table = getAllValues(getNeedIssueLineTab_());
  var headerRow = table.length ? table[0] : null;
  var currentRows = readObjectsSafe(getNeedIssueLineTab_());

  var kept = currentRows.filter(function (row) {
    return safeText(row.NeedIssueID || row.NeedIssueId || row.BehoefteUitgifteID) !== needIssueId;
  });

  var newRows = lines.map(function (line) {
    return {
      NeedIssueLineID: makeNeedIssueLineId_(),
      NeedIssueID: needIssueId,
      ArtikelCode: line.artikelCode,
      ArtikelOmschrijving: line.artikelOmschrijving,
      TypeMateriaal: line.typeMateriaal,
      Eenheid: line.eenheid,
      Aantal: line.aantal,
      Opmerking: line.opmerking,
    };
  });

  var finalObjects = kept.concat(newRows);

  if (!headerRow && finalObjects.length) {
    appendObjects(getNeedIssueLineTab_(), newRows);
  } else if (headerRow) {
    writeFullTable(
      getNeedIssueLineTab_(),
      headerRow,
      finalObjects.map(function (obj) {
        return buildRowFromHeaders(headerRow, obj);
      })
    );
  } else {
    appendObjects(getNeedIssueLineTab_(), newRows);
  }

  updateNeedIssueDerivedFields_(needIssueId, {
    typeMateriaal: deriveNeedIssueMaterialType_(lines),
  });

  writeAudit({
    actie: 'SAVE_NEED_ISSUE_LINES',
    actor: actor,
    documentType: 'BehoefteUitgifte',
    documentId: needIssueId,
    details: {
      lineCount: lines.length,
    },
  });

  return getNeedIssueWithLines(needIssueId);
}

function updateNeedIssueDerivedFields_(needIssueId, values) {
  var table = getAllValues(getNeedIssueHeaderTab_());
  if (!table.length) throw new Error('Behoefte-uitgiftetab is leeg of ongeldig.');

  var headerRow = table[0];
  var rows = readObjectsSafe(getNeedIssueHeaderTab_()).map(function (row) {
    var current = mapNeedIssueHeader(row);
    if (safeText(current.needIssueId) !== safeText(needIssueId)) {
      return row;
    }

    row.TypeMateriaal = safeText(values.typeMateriaal || row.TypeMateriaal);
    return row;
  });

  writeFullTable(
    getNeedIssueHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );
}

/* ---------------------------------------------------------
   Submit / book
   --------------------------------------------------------- */

function submitNeedIssue(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertNeedIssueWriteAccess_(sessionId);
  var needIssueId = safeText(payload.needIssueId);
  if (!needIssueId) throw new Error('NeedIssueId ontbreekt.');

  var issue = getNeedIssueWithLines(needIssueId);
  if (!issue) throw new Error('Behoefte-uitgifte niet gevonden.');
  if (!isNeedIssueEditable_(issue.status)) {
    throw new Error('Behoefte-uitgifte kan niet meer ingediend worden.');
  }

  validateNeedIssueLines_(issue.lines || []);
  validateNeedIssueSourceStock_(issue, issue.lines || []);

  var table = getAllValues(getNeedIssueHeaderTab_());
  if (!table.length) throw new Error('Behoefte-uitgiftetab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getNeedIssueHeaderTab_()).map(function (row) {
    var current = mapNeedIssueHeader(row);
    if (safeText(current.needIssueId) !== needIssueId) {
      return row;
    }

    row.Status = getNeedIssueStatusSubmitted_();
    row.IngediendOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getNeedIssueHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  writeAudit({
    actie: 'SUBMIT_NEED_ISSUE',
    actor: actor,
    documentType: 'BehoefteUitgifte',
    documentId: needIssueId,
    details: {
      lineCount: (issue.lines || []).length,
      targetType: issue.targetType,
      targetCode: issue.targetCode,
    },
  });

  return getNeedIssueWithLines(needIssueId);
}

function getNeedIssueMovementTypePair_() {
  if (typeof MOVEMENT_TYPE === 'undefined') {
    return { outType: 'NeedIssueOut', inType: 'NeedIssueIn' };
  }

  return {
    outType: MOVEMENT_TYPE.NEED_ISSUE_OUT || 'NeedIssueOut',
    inType: MOVEMENT_TYPE.NEED_ISSUE_IN || 'NeedIssueIn',
  };
}

function buildNeedIssueMovements_(issue) {
  var header = issue || {};
  var lines = header.lines || [];
  var pair = getNeedIssueMovementTypePair_();

  var outMovements = lines.map(function (line) {
    return buildMovementObject({
      movementType: pair.outType,
      bronType: 'BehoefteUitgifte',
      bronId: header.needIssueId,
      datumBoeking: header.documentDatum,
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      typeMateriaal: line.typeMateriaal,
      eenheid: line.eenheid,
      aantalUit: line.aantal,
      aantalIn: 0,
      nettoAantal: -safeNumber(line.aantal, 0),
      locatieVan: header.vanLocatie,
      locatieNaar: header.naarLocatie,
      reden: header.reden,
      opmerking: line.opmerking || header.opmerking,
      actor: header.actor,
    });
  });

  var inMovements = lines.map(function (line) {
    return buildMovementObject({
      movementType: pair.inType,
      bronType: 'BehoefteUitgifte',
      bronId: header.needIssueId,
      datumBoeking: header.documentDatum,
      artikelCode: line.artikelCode,
      artikelOmschrijving: line.artikelOmschrijving,
      typeMateriaal: line.typeMateriaal,
      eenheid: line.eenheid,
      aantalUit: 0,
      aantalIn: line.aantal,
      nettoAantal: safeNumber(line.aantal, 0),
      locatieVan: header.vanLocatie,
      locatieNaar: header.naarLocatie,
      reden: header.reden,
      opmerking: line.opmerking || header.opmerking,
      actor: header.actor,
    });
  });

  return outMovements.concat(inMovements);
}

function bookNeedIssue(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertNeedIssueWriteAccess_(sessionId);
  var needIssueId = safeText(payload.needIssueId);
  if (!needIssueId) throw new Error('NeedIssueId ontbreekt.');

  if (typeof replaceSourceMovements !== 'function') {
    throw new Error('Movement service ontbreekt. Werk eerst het movementblok in.');
  }

  var issue = getNeedIssueWithLines(needIssueId);
  if (!issue) throw new Error('Behoefte-uitgifte niet gevonden.');

  var allowedStatuses = [getNeedIssueStatusSubmitted_(), getNeedIssueStatusOpen_()];
  if (allowedStatuses.indexOf(safeText(issue.status)) < 0) {
    throw new Error('Behoefte-uitgifte kan niet geboekt worden vanuit status "' + safeText(issue.status) + '".');
  }

  validateNeedIssueLines_(issue.lines || []);
  validateNeedIssueSourceStock_(issue, issue.lines || []);

  var movements = buildNeedIssueMovements_(issue);
  replaceSourceMovements('BehoefteUitgifte', issue.needIssueId, movements);

  var table = getAllValues(getNeedIssueHeaderTab_());
  if (!table.length) throw new Error('Behoefte-uitgiftetab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getNeedIssueHeaderTab_()).map(function (row) {
    var current = mapNeedIssueHeader(row);
    if (safeText(current.needIssueId) !== needIssueId) {
      return row;
    }

    row.Status = getNeedIssueStatusBooked_();
    row.GeboektOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getNeedIssueHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  if (typeof rebuildCentralWarehouseOverview === 'function') {
    rebuildCentralWarehouseOverview();
  }

  writeAudit({
    actie: 'BOOK_NEED_ISSUE',
    actor: actor,
    documentType: 'BehoefteUitgifte',
    documentId: needIssueId,
    details: {
      movementCount: movements.length,
      targetType: issue.targetType,
      targetCode: issue.targetCode,
      totalIssued: (issue.lines || []).reduce(function (sum, line) {
        return sum + safeNumber(line.aantal, 0);
      }, 0),
    },
  });

  return getNeedIssueWithLines(needIssueId);
}

/* ---------------------------------------------------------
   Queries for screens
   --------------------------------------------------------- */

function filterNeedIssuesForUser_(rows, user) {
  if (roleAllowed(user, [ROLE.MANAGER, ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE])) {
    return rows;
  }

  return [];
}

function getNeedIssuesData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertNeedIssueReadAccess_(sessionId);
  var rows = filterNeedIssuesForUser_(getNeedIssuesWithLines(), actor);

  return {
    items: rows,
    needIssues: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getNeedIssueStatusOpen_(); }).length,
      ingediend: rows.filter(function (x) { return safeText(x.status) === getNeedIssueStatusSubmitted_(); }).length,
      geboekt: rows.filter(function (x) { return safeText(x.status) === getNeedIssueStatusBooked_(); }).length,
      gesloten: rows.filter(function (x) { return safeText(x.status) === getNeedIssueStatusClosed_(); }).length,
      naarTechnieker: rows.filter(function (x) { return safeText(x.targetType) === getNeedIssueTargetTypeTechnician_(); }).length,
      naarMobiel: rows.filter(function (x) { return safeText(x.targetType) === getNeedIssueTargetTypeMobileWarehouse_(); }).length,
    }
  };
}
