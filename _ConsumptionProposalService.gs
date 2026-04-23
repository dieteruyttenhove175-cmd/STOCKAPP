/* =========================================================
   46_ConsumptionProposalService.gs
   Refactor: consumption proposal service
   Doel:
   - aanvulvoorstellen opbouwen uit historisch verbruik
   - voorstellen voor busjes en mobiel magazijn
   - lijnen bewaren / goedkeuren
   - voorstel materialiseren naar behoefte-uitgifte of transferdraft
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getConsumptionProposalHeaderTab_() {
  return TABS.CONSUMPTION_PROPOSALS || 'VerbruiksVoorstellen';
}

function getConsumptionProposalLineTab_() {
  return TABS.CONSUMPTION_PROPOSAL_LINES || 'VerbruiksVoorstelLijnen';
}

function getConsumptionProposalStatusOpen_() {
  if (typeof CONSUMPTION_PROPOSAL_STATUS !== 'undefined' && CONSUMPTION_PROPOSAL_STATUS && CONSUMPTION_PROPOSAL_STATUS.OPEN) {
    return CONSUMPTION_PROPOSAL_STATUS.OPEN;
  }
  return 'Open';
}

function getConsumptionProposalStatusApproved_() {
  if (typeof CONSUMPTION_PROPOSAL_STATUS !== 'undefined' && CONSUMPTION_PROPOSAL_STATUS && CONSUMPTION_PROPOSAL_STATUS.APPROVED) {
    return CONSUMPTION_PROPOSAL_STATUS.APPROVED;
  }
  return 'Goedgekeurd';
}

function getConsumptionProposalStatusMaterialized_() {
  if (typeof CONSUMPTION_PROPOSAL_STATUS !== 'undefined' && CONSUMPTION_PROPOSAL_STATUS && CONSUMPTION_PROPOSAL_STATUS.MATERIALIZED) {
    return CONSUMPTION_PROPOSAL_STATUS.MATERIALIZED;
  }
  return 'Gematerialiseerd';
}

function getConsumptionProposalStatusClosed_() {
  if (typeof CONSUMPTION_PROPOSAL_STATUS !== 'undefined' && CONSUMPTION_PROPOSAL_STATUS && CONSUMPTION_PROPOSAL_STATUS.CLOSED) {
    return CONSUMPTION_PROPOSAL_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function getConsumptionProposalTargetTechnician_() {
  return 'TECHNICIAN';
}

function getConsumptionProposalTargetMobileWarehouse_() {
  return 'MOBILE_WAREHOUSE';
}

function getConsumptionProposalDefaultWindowDays_() {
  var cfg = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG) ? APP_CONFIG : {};
  return safeNumber(cfg.CONSUMPTION_PROPOSAL_WINDOW_DAYS, 30);
}

function getConsumptionProposalDefaultCoverageDays_() {
  var cfg = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG) ? APP_CONFIG : {};
  return safeNumber(cfg.CONSUMPTION_PROPOSAL_COVERAGE_DAYS, 14);
}

function makeConsumptionProposalId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'CPS-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function makeConsumptionProposalLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'CPLN-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function isConsumptionProposalEditable_(status) {
  var value = safeText(status);
  return !value || value === getConsumptionProposalStatusOpen_();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapConsumptionProposalHeader(row) {
  return {
    proposalId: safeText(row.ProposalID || row.ConsumptionProposalID || row.ConsumptionProposalId || row.ID),
    targetType: safeText(row.TargetType || getConsumptionProposalTargetTechnician_()),
    targetCode: safeText(row.TargetCode),
    targetName: safeText(row.TargetName),
    sourceLocation: safeText(row.SourceLocation || row.VanLocatie || LOCATION.CENTRAL),
    targetLocation: safeText(row.TargetLocation || row.NaarLocatie),
    linkedDocumentType: safeText(row.LinkedDocumentType),
    linkedDocumentId: safeText(row.LinkedDocumentID || row.LinkedDocumentId),
    windowDays: safeNumber(row.WindowDays, getConsumptionProposalDefaultWindowDays_()),
    coverageDays: safeNumber(row.CoverageDays, getConsumptionProposalDefaultCoverageDays_()),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    status: safeText(row.Status || getConsumptionProposalStatusOpen_()),
    actor: safeText(row.Actor),
    reason: safeText(row.Reason || row.Reden),
    opmerking: safeText(row.Opmerking),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    goedgekeurdOp: safeText(row.GoedgekeurdOp),
    geslotenOp: safeText(row.GeslotenOp),
  };
}

function mapConsumptionProposalLine(row) {
  return {
    proposalLineId: safeText(row.ProposalLineID || row.ConsumptionProposalLineID || row.ConsumptionProposalLineId || row.ID),
    proposalId: safeText(row.ProposalID || row.ConsumptionProposalID || row.ConsumptionProposalId),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode || row.ArtikelNr))),
    eenheid: safeText(row.Eenheid || row.Unit),
    recentConsumptionQty: safeNumber(row.RecentConsumptionQty || row.RecentQty, 0),
    averageDailyQty: safeNumber(row.AverageDailyQty || row.DailyQty, 0),
    currentStockQty: safeNumber(row.CurrentStockQty || row.CurrentQty, 0),
    targetStockQty: safeNumber(row.TargetStockQty || row.TargetQty, 0),
    suggestedQty: safeNumber(row.SuggestedQty || row.SuggestedAantal || 0),
    opmerking: safeText(row.Opmerking),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllConsumptionProposals() {
  return readObjectsSafe(getConsumptionProposalHeaderTab_())
    .map(mapConsumptionProposalHeader)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.aangemaaktOpRaw).localeCompare(safeText(a.aangemaaktOpRaw)) ||
        safeText(b.proposalId).localeCompare(safeText(a.proposalId))
      );
    });
}

function getAllConsumptionProposalLines() {
  return readObjectsSafe(getConsumptionProposalLineTab_())
    .map(mapConsumptionProposalLine)
    .sort(function (a, b) {
      return (
        safeText(a.proposalId).localeCompare(safeText(b.proposalId)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function getConsumptionProposalById(proposalId) {
  var id = safeText(proposalId);
  if (!id) return null;

  return getAllConsumptionProposals().find(function (item) {
    return safeText(item.proposalId) === id;
  }) || null;
}

function getConsumptionProposalLinesById(proposalId) {
  var id = safeText(proposalId);
  return getAllConsumptionProposalLines().filter(function (item) {
    return safeText(item.proposalId) === id;
  });
}

function buildConsumptionProposalsWithLines(headers, lines) {
  var lineMap = {};

  (lines || []).forEach(function (line) {
    var id = safeText(line.proposalId);
    if (!lineMap[id]) lineMap[id] = [];
    lineMap[id].push(line);
  });

  return (headers || []).map(function (header) {
    var proposalLines = lineMap[safeText(header.proposalId)] || [];
    return Object.assign({}, header, {
      lines: proposalLines,
      lineCount: proposalLines.length,
      totaalSuggestedQty: proposalLines.reduce(function (sum, line) {
        return sum + safeNumber(line.suggestedQty, 0);
      }, 0),
    });
  });
}

function getConsumptionProposalWithLines(proposalId) {
  var header = getConsumptionProposalById(proposalId);
  if (!header) return null;
  return buildConsumptionProposalsWithLines(
    [header],
    getConsumptionProposalLinesById(proposalId)
  )[0] || null;
}

/* ---------------------------------------------------------
   Access policy
   --------------------------------------------------------- */

function assertConsumptionProposalAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.MOBILE_WAREHOUSE],
    'Geen rechten voor verbruiksvoorstellen.'
  );
  return user;
}

function assertConsumptionProposalApproveAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.MANAGER, ROLE.WAREHOUSE],
    'Geen rechten om verbruiksvoorstellen goed te keuren.'
  );
  return user;
}

/* ---------------------------------------------------------
   History / stock helpers
   --------------------------------------------------------- */

function getConsumptionProposalCutoffRaw_(windowDays) {
  var dt = new Date();
  dt.setDate(dt.getDate() - safeNumber(windowDays, getConsumptionProposalDefaultWindowDays_()));
  return Utilities.formatDate(dt, TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function getBookedConsumptionLinesForProposal_() {
  if (typeof getAllConsumptionBookings !== 'function' || typeof getAllConsumptionBookingLines !== 'function') {
    return [];
  }

  var headers = getAllConsumptionBookings().filter(function (header) {
    return safeText(header.status) === getConsumptionBookingStatusBooked_();
  });

  if (!headers.length) return [];

  var allowed = {};
  headers.forEach(function (header) {
    allowed[safeText(header.bookingId)] = header;
  });

  return getAllConsumptionBookingLines()
    .filter(function (line) {
      return !!allowed[safeText(line.bookingId)];
    })
    .map(function (line) {
      var header = allowed[safeText(line.bookingId)] || {};
      return Object.assign({}, line, {
        techniekerCode: safeText(header.techniekerCode),
        techniekerNaam: safeText(header.techniekerNaam),
        documentDatumIso: safeText(header.documentDatumIso || header.documentDatum),
      });
    });
}

function buildRecentConsumptionByTechnician_(techniekerCode, windowDays) {
  var code = safeText(techniekerCode);
  var cutoffRaw = getConsumptionProposalCutoffRaw_(windowDays);
  var grouped = {};

  getBookedConsumptionLinesForProposal_()
    .filter(function (line) {
      return safeText(line.techniekerCode) === code &&
        safeText(line.documentDatumIso) >= cutoffRaw.slice(0, 10);
    })
    .forEach(function (line) {
      var artikelCode = safeText(line.artikelCode);
      if (!artikelCode) return;

      if (!grouped[artikelCode]) {
        grouped[artikelCode] = {
          artikelCode: artikelCode,
          artikelOmschrijving: safeText(line.artikelOmschrijving),
          typeMateriaal: safeText(line.typeMateriaal),
          eenheid: safeText(line.eenheid),
          recentConsumptionQty: 0,
        };
      }

      grouped[artikelCode].recentConsumptionQty += safeNumber(line.aantal, 0);
    });

  return Object.keys(grouped)
    .map(function (key) { return grouped[key]; })
    .sort(function (a, b) {
      return (
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function buildRecentConsumptionByMobileWarehouse_(mobileWarehouseCode, windowDays) {
  var mwCode = safeText(mobileWarehouseCode);
  if (!mwCode) return [];

  var technicianMap = {};
  if (typeof getAllActiveTechniciansForRights_ === 'function') {
    getAllActiveTechniciansForRights_().forEach(function (tech) {
      if (safeText(tech.mobileWarehouseCode) === mwCode) {
        technicianMap[safeText(tech.code || tech.ref)] = tech;
      }
    });
  }

  var grouped = {};
  Object.keys(technicianMap).forEach(function (techCode) {
    buildRecentConsumptionByTechnician_(techCode, windowDays).forEach(function (item) {
      var artikelCode = safeText(item.artikelCode);
      if (!artikelCode) return;

      if (!grouped[artikelCode]) {
        grouped[artikelCode] = {
          artikelCode: artikelCode,
          artikelOmschrijving: safeText(item.artikelOmschrijving),
          typeMateriaal: safeText(item.typeMateriaal),
          eenheid: safeText(item.eenheid),
          recentConsumptionQty: 0,
        };
      }

      grouped[artikelCode].recentConsumptionQty += safeNumber(item.recentConsumptionQty, 0);
    });
  });

  return Object.keys(grouped)
    .map(function (key) { return grouped[key]; })
    .sort(function (a, b) {
      return (
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function getCurrentProposalStockQty_(targetType, targetCode, artikelCode) {
  var type = safeText(targetType);
  var code = safeText(targetCode);
  var article = safeText(artikelCode);

  if (!article) return 0;

  if (type === getConsumptionProposalTargetTechnician_()) {
    if (typeof buildBusStockMapForTechnician === 'function') {
      var busMap = buildBusStockMapForTechnician(code);
      return busMap[article] ? safeNumber(busMap[article].voorraadBus, 0) : 0;
    }
    return 0;
  }

  if (type === getConsumptionProposalTargetMobileWarehouse_()) {
    if (typeof buildMobileWarehouseSnapshotMap === 'function') {
      var mwMap = buildMobileWarehouseSnapshotMap(code);
      return mwMap[article] ? safeNumber(mwMap[article].voorraadMobiel, 0) : 0;
    }
    if (typeof buildMobileWarehouseStockMap === 'function') {
      var oldMap = buildMobileWarehouseStockMap(code);
      return oldMap[article] ? safeNumber(oldMap[article].voorraadMobiel, 0) : 0;
    }
    return 0;
  }

  return 0;
}

function resolveConsumptionProposalTargetName_(targetType, targetCode) {
  var type = safeText(targetType);
  var code = safeText(targetCode);

  if (type === getConsumptionProposalTargetTechnician_()) {
    if (typeof resolveTechnicianByRef === 'function') {
      var tech = resolveTechnicianByRef(code);
      return tech ? safeText(tech.naam) : code;
    }
    return code;
  }

  if (type === getConsumptionProposalTargetMobileWarehouse_()) {
    if (typeof getMobileWarehouseByCode === 'function') {
      var mw = getMobileWarehouseByCode(code);
      return mw ? safeText(mw.naam) : code;
    }
    return code;
  }

  return code;
}

function buildConsumptionProposalDraftLines_(targetType, targetCode, windowDays, coverageDays) {
  var basisRows = safeText(targetType) === getConsumptionProposalTargetMobileWarehouse_()
    ? buildRecentConsumptionByMobileWarehouse_(targetCode, windowDays)
    : buildRecentConsumptionByTechnician_(targetCode, windowDays);

  return basisRows
    .map(function (item) {
      var recentQty = safeNumber(item.recentConsumptionQty, 0);
      var averageDaily = safeNumber(windowDays, 0) > 0
        ? recentQty / safeNumber(windowDays, 1)
        : 0;
      var targetQty = Math.ceil(averageDaily * safeNumber(coverageDays, 0));
      var currentQty = getCurrentProposalStockQty_(targetType, targetCode, item.artikelCode);
      var suggestedQty = Math.max(targetQty - currentQty, 0);

      return {
        artikelCode: safeText(item.artikelCode),
        artikelOmschrijving: safeText(item.artikelOmschrijving),
        typeMateriaal: safeText(item.typeMateriaal),
        eenheid: safeText(item.eenheid),
        recentConsumptionQty: recentQty,
        averageDailyQty: averageDaily,
        currentStockQty: currentQty,
        targetStockQty: targetQty,
        suggestedQty: suggestedQty,
        opmerking: '',
      };
    })
    .filter(function (item) {
      return safeNumber(item.suggestedQty, 0) > 0;
    })
    .sort(function (a, b) {
      return (
        safeNumber(b.suggestedQty, 0) - safeNumber(a.suggestedQty, 0) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
      );
    });
}

/* ---------------------------------------------------------
   Header / lines normalization
   --------------------------------------------------------- */

function normalizeConsumptionProposalHeaderPayload_(payload) {
  payload = payload || {};

  var targetType = safeText(payload.targetType || payload.doelType || getConsumptionProposalTargetTechnician_());
  var targetCode = safeText(payload.targetCode || payload.doelCode || payload.techniekerCode || payload.mobileWarehouseCode);

  return {
    sessionId: getPayloadSessionId(payload),
    targetType: targetType,
    targetCode: targetCode,
    targetName: safeText(payload.targetName || resolveConsumptionProposalTargetName_(targetType, targetCode)),
    sourceLocation: safeText(payload.sourceLocation || payload.vanLocatie || LOCATION.CENTRAL),
    targetLocation: safeText(
      payload.targetLocation ||
      payload.naarLocatie ||
      (
        targetType === getConsumptionProposalTargetMobileWarehouse_()
          ? (typeof getMobileWarehouseLocationCode === 'function' ? getMobileWarehouseLocationCode(targetCode) : ('Mobiel:' + targetCode))
          : (typeof getBusLocationCode === 'function' ? getBusLocationCode(targetCode) : ('Bus:' + targetCode))
      )
    ),
    windowDays: safeNumber(payload.windowDays, getConsumptionProposalDefaultWindowDays_()),
    coverageDays: safeNumber(payload.coverageDays, getConsumptionProposalDefaultCoverageDays_()),
    documentDatum: safeText(payload.documentDatum || payload.documentDate),
    reason: safeText(payload.reason || payload.reden || 'Verbruiksvoorstel'),
    opmerking: safeText(payload.opmerking || payload.remark),
    actor: safeText(payload.actor),
  };
}

function normalizeConsumptionProposalLines_(lines) {
  return (Array.isArray(lines) ? lines : []).map(function (line) {
    return {
      artikelCode: safeText(line.artikelCode || line.artikelNr),
      artikelOmschrijving: safeText(line.artikelOmschrijving || line.artikel),
      typeMateriaal: safeText(line.typeMateriaal || determineMaterialTypeFromArticle(safeText(line.artikelCode || line.artikelNr))),
      eenheid: safeText(line.eenheid || line.unit || 'Stuk'),
      recentConsumptionQty: safeNumber(line.recentConsumptionQty || line.recentQty, 0),
      averageDailyQty: safeNumber(line.averageDailyQty || line.dailyQty, 0),
      currentStockQty: safeNumber(line.currentStockQty || line.currentQty, 0),
      targetStockQty: safeNumber(line.targetStockQty || line.targetQty, 0),
      suggestedQty: safeNumber(line.suggestedQty || line.suggestedAantal, 0),
      opmerking: safeText(line.opmerking),
    };
  });
}

function validateConsumptionProposalHeader_(payload) {
  if (!payload.sessionId) throw new Error('Sessie ontbreekt.');
  if (!payload.targetType) throw new Error('TargetType is verplicht.');
  if (!payload.targetCode) throw new Error('TargetCode is verplicht.');
  if (!payload.documentDatum) throw new Error('Documentdatum is verplicht.');
  return true;
}

function validateConsumptionProposalLines_(lines) {
  if (!lines.length) {
    throw new Error('Geen voorstelregels ontvangen.');
  }

  lines.forEach(function (line, index) {
    var rowNr = index + 1;
    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.suggestedQty, 0) <= 0) {
      throw new Error('Voorgesteld aantal moet groter zijn dan 0 op lijn ' + rowNr + '.');
    }
  });

  return true;
}

/* ---------------------------------------------------------
   Create / save
   --------------------------------------------------------- */

function createConsumptionProposal(payload) {
  var normalized = normalizeConsumptionProposalHeaderPayload_(payload);
  var actor = assertConsumptionProposalAccess_(normalized.sessionId);

  validateConsumptionProposalHeader_(normalized);

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var proposalId = makeConsumptionProposalId_();

  var obj = {
    ProposalID: proposalId,
    TargetType: normalized.targetType,
    TargetCode: normalized.targetCode,
    TargetName: normalized.targetName,
    SourceLocation: normalized.sourceLocation,
    TargetLocation: normalized.targetLocation,
    LinkedDocumentType: '',
    LinkedDocumentID: '',
    WindowDays: normalized.windowDays,
    CoverageDays: normalized.coverageDays,
    DocumentDatum: normalized.documentDatum,
    DocumentDatumIso: normalized.documentDatum,
    Status: getConsumptionProposalStatusOpen_(),
    Actor: normalized.actor || safeText(actor.naam || actor.email),
    Reason: normalized.reason,
    Opmerking: normalized.opmerking,
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    GoedgekeurdOp: '',
    GeslotenOp: '',
  };

  appendObjects(getConsumptionProposalHeaderTab_(), [obj]);

  writeAudit({
    actie: 'CREATE_CONSUMPTION_PROPOSAL',
    actor: actor,
    documentType: 'VerbruiksVoorstel',
    documentId: proposalId,
    details: {
      targetType: normalized.targetType,
      targetCode: normalized.targetCode,
      coverageDays: normalized.coverageDays,
      windowDays: normalized.windowDays,
    },
  });

  return mapConsumptionProposalHeader(obj);
}

function createConsumptionProposalDraftFromHistory(payload) {
  payload = payload || {};

  var normalized = normalizeConsumptionProposalHeaderPayload_(payload);
  var actor = assertConsumptionProposalAccess_(normalized.sessionId);

  var header = createConsumptionProposal({
    sessionId: normalized.sessionId,
    targetType: normalized.targetType,
    targetCode: normalized.targetCode,
    targetName: normalized.targetName,
    sourceLocation: normalized.sourceLocation,
    targetLocation: normalized.targetLocation,
    windowDays: normalized.windowDays,
    coverageDays: normalized.coverageDays,
    documentDatum: normalized.documentDatum,
    reason: normalized.reason,
    opmerking: normalized.opmerking,
    actor: normalized.actor || safeText(actor.naam || actor.email),
  });

  var lines = buildConsumptionProposalDraftLines_(
    normalized.targetType,
    normalized.targetCode,
    normalized.windowDays,
    normalized.coverageDays
  );

  return saveConsumptionProposalLines({
    sessionId: normalized.sessionId,
    proposalId: header.proposalId,
    lines: lines,
  });
}

function saveConsumptionProposalLines(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionProposalAccess_(sessionId);
  var proposalId = safeText(payload.proposalId);
  if (!proposalId) throw new Error('ProposalId ontbreekt.');

  var header = getConsumptionProposalById(proposalId);
  if (!header) throw new Error('Verbruiksvoorstel niet gevonden.');
  if (!isConsumptionProposalEditable_(header.status)) {
    throw new Error('Verbruiksvoorstel is niet meer bewerkbaar.');
  }

  var lines = normalizeConsumptionProposalLines_(payload.lines);
  validateConsumptionProposalLines_(lines);

  var table = getAllValues(getConsumptionProposalLineTab_());
  var headerRow = table.length ? table[0] : null;
  var currentRows = readObjectsSafe(getConsumptionProposalLineTab_());

  var kept = currentRows.filter(function (row) {
    return safeText(row.ProposalID || row.ConsumptionProposalID || row.ConsumptionProposalId) !== proposalId;
  });

  var newRows = lines.map(function (line) {
    return {
      ProposalLineID: makeConsumptionProposalLineId_(),
      ProposalID: proposalId,
      ArtikelCode: line.artikelCode,
      ArtikelOmschrijving: line.artikelOmschrijving,
      TypeMateriaal: line.typeMateriaal,
      Eenheid: line.eenheid,
      RecentConsumptionQty: line.recentConsumptionQty,
      AverageDailyQty: line.averageDailyQty,
      CurrentStockQty: line.currentStockQty,
      TargetStockQty: line.targetStockQty,
      SuggestedQty: line.suggestedQty,
      Opmerking: line.opmerking,
    };
  });

  var finalObjects = kept.concat(newRows);

  if (!headerRow && finalObjects.length) {
    appendObjects(getConsumptionProposalLineTab_(), newRows);
  } else if (headerRow) {
    writeFullTable(
      getConsumptionProposalLineTab_(),
      headerRow,
      finalObjects.map(function (obj) {
        return buildRowFromHeaders(headerRow, obj);
      })
    );
  } else if (newRows.length) {
    appendObjects(getConsumptionProposalLineTab_(), newRows);
  }

  writeAudit({
    actie: 'SAVE_CONSUMPTION_PROPOSAL_LINES',
    actor: actor,
    documentType: 'VerbruiksVoorstel',
    documentId: proposalId,
    details: {
      lineCount: lines.length,
    },
  });

  return getConsumptionProposalWithLines(proposalId);
}

/* ---------------------------------------------------------
   Approve / materialize
   --------------------------------------------------------- */

function approveConsumptionProposal(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionProposalApproveAccess_(sessionId);
  var proposalId = safeText(payload.proposalId);
  if (!proposalId) throw new Error('ProposalId ontbreekt.');

  var proposal = getConsumptionProposalWithLines(proposalId);
  if (!proposal) throw new Error('Verbruiksvoorstel niet gevonden.');
  if (safeText(proposal.status) !== getConsumptionProposalStatusOpen_()) {
    throw new Error('Verbruiksvoorstel staat niet in open status.');
  }

  validateConsumptionProposalLines_(proposal.lines || []);

  var table = getAllValues(getConsumptionProposalHeaderTab_());
  if (!table.length) throw new Error('Voorsteltab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getConsumptionProposalHeaderTab_()).map(function (row) {
    var current = mapConsumptionProposalHeader(row);
    if (safeText(current.proposalId) !== proposalId) {
      return row;
    }

    row.Status = getConsumptionProposalStatusApproved_();
    row.GoedgekeurdOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getConsumptionProposalHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  writeAudit({
    actie: 'APPROVE_CONSUMPTION_PROPOSAL',
    actor: actor,
    documentType: 'VerbruiksVoorstel',
    documentId: proposalId,
    details: {
      targetType: proposal.targetType,
      targetCode: proposal.targetCode,
      lineCount: (proposal.lines || []).length,
    },
  });

  return getConsumptionProposalWithLines(proposalId);
}

function materializeConsumptionProposal(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionProposalAccess_(sessionId);
  var proposalId = safeText(payload.proposalId);
  if (!proposalId) throw new Error('ProposalId ontbreekt.');

  var proposal = getConsumptionProposalWithLines(proposalId);
  if (!proposal) throw new Error('Verbruiksvoorstel niet gevonden.');

  if ([getConsumptionProposalStatusApproved_(), getConsumptionProposalStatusOpen_()].indexOf(safeText(proposal.status)) < 0) {
    throw new Error('Verbruiksvoorstel kan niet gematerialiseerd worden vanuit status "' + safeText(proposal.status) + '".');
  }

  validateConsumptionProposalLines_(proposal.lines || []);

  var linkedDocumentType = '';
  var linkedDocumentId = '';

  if (safeText(proposal.targetType) === getConsumptionProposalTargetTechnician_()) {
    if (typeof createNeedIssue !== 'function' || typeof saveNeedIssueLines !== 'function') {
      throw new Error('NeedIssue service ontbreekt. Werk eerst het needissueblok in.');
    }

    var needIssue = createNeedIssue({
      sessionId: sessionId,
      targetType: getNeedIssueTargetTypeTechnician_(),
      targetCode: proposal.targetCode,
      documentDatum: proposal.documentDatum,
      reason: proposal.reason || 'Aanvulling op basis van verbruik',
      remark: proposal.opmerking,
      actor: safeText(actor.naam || actor.email),
    });

    saveNeedIssueLines({
      sessionId: sessionId,
      needIssueId: needIssue.needIssueId,
      lines: (proposal.lines || []).map(function (line) {
        return {
          artikelCode: line.artikelCode,
          artikelOmschrijving: line.artikelOmschrijving,
          typeMateriaal: line.typeMateriaal,
          eenheid: line.eenheid,
          aantal: line.suggestedQty,
          opmerking: line.opmerking,
        };
      }),
    });

    linkedDocumentType = 'BehoefteUitgifte';
    linkedDocumentId = safeText(needIssue.needIssueId);
  } else if (safeText(proposal.targetType) === getConsumptionProposalTargetMobileWarehouse_()) {
    if (typeof createTransfer !== 'function' || typeof saveTransferLines !== 'function') {
      throw new Error('Transfer service ontbreekt. Werk eerst het transferblok in.');
    }

    var transfer = createTransfer({
      sessionId: sessionId,
      flowType: TRANSFER_FLOW.CENTRAL_TO_MOBILE,
      vanLocatie: proposal.sourceLocation || LOCATION.CENTRAL,
      naarLocatie: proposal.targetLocation,
      documentDatum: proposal.documentDatum,
      reason: proposal.reason || 'Aanvulling op basis van verbruik',
      remark: proposal.opmerking,
      actor: safeText(actor.naam || actor.email),
      mobileWarehouseCode: proposal.targetCode,
    });

    saveTransferLines({
      sessionId: sessionId,
      transferId: transfer.transferId,
      lines: (proposal.lines || []).map(function (line) {
        return {
          artikelCode: line.artikelCode,
          artikelOmschrijving: line.artikelOmschrijving,
          typeMateriaal: line.typeMateriaal,
          eenheid: line.eenheid,
          aantal: line.suggestedQty,
          opmerking: line.opmerking,
        };
      }),
    });

    linkedDocumentType = 'Transfer';
    linkedDocumentId = safeText(transfer.transferId);
  } else {
    throw new Error('Onbekend targetType voor verbruiksvoorstel.');
  }

  var table = getAllValues(getConsumptionProposalHeaderTab_());
  if (!table.length) throw new Error('Voorsteltab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getConsumptionProposalHeaderTab_()).map(function (row) {
    var current = mapConsumptionProposalHeader(row);
    if (safeText(current.proposalId) !== proposalId) {
      return row;
    }

    row.Status = getConsumptionProposalStatusMaterialized_();
    row.LinkedDocumentType = linkedDocumentType;
    row.LinkedDocumentID = linkedDocumentId;
    row.GeslotenOp = toDisplayDateTime(nowRaw);
    return row;
  });

  writeFullTable(
    getConsumptionProposalHeaderTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  writeAudit({
    actie: 'MATERIALIZE_CONSUMPTION_PROPOSAL',
    actor: actor,
    documentType: 'VerbruiksVoorstel',
    documentId: proposalId,
    details: {
      linkedDocumentType: linkedDocumentType,
      linkedDocumentId: linkedDocumentId,
      targetType: proposal.targetType,
      targetCode: proposal.targetCode,
    },
  });

  return getConsumptionProposalWithLines(proposalId);
}

/* ---------------------------------------------------------
   Queries
   --------------------------------------------------------- */

function getConsumptionProposalData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertConsumptionProposalAccess_(sessionId);
  var rows = buildConsumptionProposalsWithLines(
    getAllConsumptionProposals(),
    getAllConsumptionProposalLines()
  );

  return {
    items: rows,
    proposals: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getConsumptionProposalStatusOpen_(); }).length,
      goedgekeurd: rows.filter(function (x) { return safeText(x.status) === getConsumptionProposalStatusApproved_(); }).length,
      gematerialiseerd: rows.filter(function (x) { return safeText(x.status) === getConsumptionProposalStatusMaterialized_(); }).length,
      gesloten: rows.filter(function (x) { return safeText(x.status) === getConsumptionProposalStatusClosed_(); }).length,
      totaalSuggestedQty: rows.reduce(function (sum, item) {
        return sum + safeNumber(item.totaalSuggestedQty, 0);
      }, 0),
      actorRol: safeText(actor.rol),
    }
  };
}
