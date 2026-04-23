/* =========================================================
   51_RecurringBusCountService.gs
   Refactor: recurring bus count service
   Doel:
   - recurrente bustellingen sturen vanuit risicoscores
   - open suggesties / open requests / historiek
   - draft-aanmaak van gerichte bustellingen per technieker
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getRecurringBusCountTab_() {
  return TABS.RECURRING_BUS_COUNTS || 'RecurrenteBusTellingen';
}

function getRecurringBusCountStatusOpen_() {
  if (typeof RECURRING_BUS_COUNT_STATUS !== 'undefined' && RECURRING_BUS_COUNT_STATUS && RECURRING_BUS_COUNT_STATUS.OPEN) {
    return RECURRING_BUS_COUNT_STATUS.OPEN;
  }
  return 'Open';
}

function getRecurringBusCountStatusCreated_() {
  if (typeof RECURRING_BUS_COUNT_STATUS !== 'undefined' && RECURRING_BUS_COUNT_STATUS && RECURRING_BUS_COUNT_STATUS.CREATED) {
    return RECURRING_BUS_COUNT_STATUS.CREATED;
  }
  return 'Aangemaakt';
}

function getRecurringBusCountStatusClosed_() {
  if (typeof RECURRING_BUS_COUNT_STATUS !== 'undefined' && RECURRING_BUS_COUNT_STATUS && RECURRING_BUS_COUNT_STATUS.CLOSED) {
    return RECURRING_BUS_COUNT_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function makeRecurringBusCountId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'RBC-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapRecurringBusCount(row) {
  return {
    recurringBusCountId: safeText(row.RecurringBusCountID || row.RecurringBusCountId || row.ID),
    techniekerCode: safeText(row.TechniekerCode || row.TechnicianCode || row.TechCode),
    techniekerNaam: safeText(row.TechniekerNaam || row.TechnicianName),
    score: safeNumber(row.Score, 0),
    sourceType: safeText(row.SourceType || 'RISK_SCORE'),
    linkedBusCountId: safeText(row.LinkedBusCountID || row.LinkedBusCountId),
    status: safeText(row.Status || getRecurringBusCountStatusOpen_()),
    reason: safeText(row.Reason || row.Reden),
    articleCodesCsv: safeText(row.ArticleCodesCsv || row.ArtikelCodesCsv),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    actor: safeText(row.Actor),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    geslotenOp: safeText(row.GeslotenOp),
    opmerking: safeText(row.Opmerking),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllRecurringBusCountRequests() {
  return readObjectsSafe(getRecurringBusCountTab_())
    .map(mapRecurringBusCount)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(b.aangemaaktOpRaw).localeCompare(safeText(a.aangemaaktOpRaw)) ||
        safeText(b.recurringBusCountId).localeCompare(safeText(a.recurringBusCountId))
      );
    });
}

function getRecurringBusCountById(recurringBusCountId) {
  var id = safeText(recurringBusCountId);
  if (!id) return null;

  return getAllRecurringBusCountRequests().find(function (item) {
    return safeText(item.recurringBusCountId) === id;
  }) || null;
}

function getOpenRecurringBusCountRequests() {
  return getAllRecurringBusCountRequests().filter(function (item) {
    return [
      getRecurringBusCountStatusOpen_(),
      getRecurringBusCountStatusCreated_()
    ].indexOf(safeText(item.status)) >= 0;
  });
}

function getRecurringBusCountRequestsByTechnician(techniekerCode) {
  var code = safeText(techniekerCode);
  return getAllRecurringBusCountRequests().filter(function (item) {
    return safeText(item.techniekerCode) === code;
  });
}

/* ---------------------------------------------------------
   Helpers
   --------------------------------------------------------- */

function parseRecurringArticleCodesCsv_(csv) {
  return safeText(csv)
    .split(',')
    .map(function (item) { return safeText(item); })
    .filter(Boolean);
}

function buildRecurringArticleCodesCsv_(codes) {
  var seen = {};
  return (Array.isArray(codes) ? codes : [])
    .map(function (item) { return safeText(item); })
    .filter(function (item) {
      if (!item) return false;
      if (seen[item]) return false;
      seen[item] = true;
      return true;
    })
    .join(',');
}

function resolveTechnicianNameForRecurring_(techniekerCode) {
  var code = safeText(techniekerCode);
  if (!code) return '';

  if (typeof resolveTechnicianByRef === 'function') {
    var tech = resolveTechnicianByRef(code);
    return tech ? safeText(tech.naam) : code;
  }

  return code;
}

function hasOpenRecurringRequestForTechnician_(techniekerCode) {
  var code = safeText(techniekerCode);

  return getOpenRecurringBusCountRequests().some(function (item) {
    return safeText(item.techniekerCode) === code;
  });
}

function hasOpenBusCountForTechnician_(techniekerCode) {
  var code = safeText(techniekerCode);

  if (typeof getBusCountsWithLines !== 'function') {
    return false;
  }

  return getBusCountsWithLines().some(function (item) {
    return safeText(item.techniekerCode) === code &&
      [getBusCountStatusOpen_(), getBusCountStatusSubmitted_()].indexOf(safeText(item.status)) >= 0;
  });
}

function getRecurringBusCountSuggestedArticleCodes_(techniekerCode) {
  var code = safeText(techniekerCode);
  if (!code) return [];

  if (typeof getStockRiskSignals !== 'function') {
    return [];
  }

  var seen = {};
  return getStockRiskSignals()
    .filter(function (signal) {
      return safeText(signal.techniekerCode) === code &&
        safeText(signal.signalType) === 'BUSCOUNT_DELTA';
    })
    .sort(function (a, b) {
      return safeText(b.rawDate).localeCompare(safeText(a.rawDate));
    })
    .map(function (signal) {
      return safeText(signal.artikelCode);
    })
    .filter(function (artikelCode) {
      if (!artikelCode) return false;
      if (seen[artikelCode]) return false;
      seen[artikelCode] = true;
      return true;
    })
    .slice(0, 25);
}

/* ---------------------------------------------------------
   Suggestion builder from risk scores
   --------------------------------------------------------- */

function buildRecurringBusCountCandidates() {
  var suggestions =
    typeof buildRecurringBusCountSuggestions === 'function'
      ? buildRecurringBusCountSuggestions()
      : [];

  return suggestions
    .map(function (item) {
      var techniekerCode = safeText(item.techniekerCode);
      return {
        techniekerCode: techniekerCode,
        techniekerNaam: safeText(item.techniekerNaam || resolveTechnicianNameForRecurring_(techniekerCode)),
        score: safeNumber(item.score, 0),
        reason: safeText(item.reason),
        articleCodes: getRecurringBusCountSuggestedArticleCodes_(techniekerCode),
        hasOpenRecurringRequest: hasOpenRecurringRequestForTechnician_(techniekerCode),
        hasOpenBusCount: hasOpenBusCountForTechnician_(techniekerCode),
      };
    })
    .sort(function (a, b) {
      return (
        safeNumber(b.score, 0) - safeNumber(a.score, 0) ||
        safeText(a.techniekerNaam).localeCompare(safeText(b.techniekerNaam))
      );
    });
}

function getPendingRecurringBusCountCandidates() {
  return buildRecurringBusCountCandidates().filter(function (item) {
    return !item.hasOpenRecurringRequest && !item.hasOpenBusCount;
  });
}

/* ---------------------------------------------------------
   Access policy
   --------------------------------------------------------- */

function assertRecurringBusCountReadAccess_(sessionId) {
  var user = requireLoggedInUser(sessionId);
  assertRoleAllowed(
    user,
    [ROLE.WAREHOUSE, ROLE.MANAGER],
    'Geen rechten voor recurrente bustellingen.'
  );
  return user;
}

function assertRecurringBusCountWriteAccess_(sessionId) {
  return assertRecurringBusCountReadAccess_(sessionId);
}

/* ---------------------------------------------------------
   Create / close recurring requests
   --------------------------------------------------------- */

function createRecurringBusCountRequest(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertRecurringBusCountWriteAccess_(sessionId);

  var techniekerCode = safeText(payload.techniekerCode || payload.technicianCode || payload.techCode);
  if (!techniekerCode) {
    throw new Error('TechniekerCode ontbreekt.');
  }

  if (hasOpenRecurringRequestForTechnician_(techniekerCode)) {
    throw new Error('Er bestaat al een open recurrente bustelling voor deze technieker.');
  }

  var score = safeNumber(payload.score, 0);
  var reason = safeText(payload.reason || payload.reden || 'Recurrente bustelling');
  var articleCodes = Array.isArray(payload.articleCodes) ? payload.articleCodes : getRecurringBusCountSuggestedArticleCodes_(techniekerCode);
  var documentDatum = safeText(payload.documentDatum || payload.documentDate || Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd'));
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var obj = {
    RecurringBusCountID: makeRecurringBusCountId_(),
    TechniekerCode: techniekerCode,
    TechniekerNaam: safeText(payload.techniekerNaam || resolveTechnicianNameForRecurring_(techniekerCode)),
    Score: score,
    SourceType: safeText(payload.sourceType || 'RISK_SCORE'),
    LinkedBusCountID: '',
    Status: getRecurringBusCountStatusOpen_(),
    Reason: reason,
    ArticleCodesCsv: buildRecurringArticleCodesCsv_(articleCodes),
    DocumentDatum: documentDatum,
    DocumentDatumIso: documentDatum,
    Actor: safeText(payload.actor || actor.naam || actor.email),
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    GeslotenOp: '',
    Opmerking: safeText(payload.opmerking || payload.remark),
  };

  appendObjects(getRecurringBusCountTab_(), [obj]);

  writeAudit({
    actie: 'CREATE_RECURRING_BUS_COUNT_REQUEST',
    actor: actor,
    documentType: 'RecurrenteBusTelling',
    documentId: obj.RecurringBusCountID,
    details: {
      techniekerCode: techniekerCode,
      score: score,
      articleCodeCount: parseRecurringArticleCodesCsv_(obj.ArticleCodesCsv).length,
    },
  });

  if (typeof pushTechnicianNotification === 'function') {
    pushTechnicianNotification(
      'BusTelling',
      'Recurrente bustelling gevraagd',
      'Er werd een recurrente bustelling voor jouw bus gevraagd.',
      'RecurrenteBusTelling',
      obj.RecurringBusCountID,
      techniekerCode,
      obj.TechniekerNaam
    );
  }

  return mapRecurringBusCount(obj);
}

function closeRecurringBusCountRequest(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertRecurringBusCountWriteAccess_(sessionId);
  var recurringBusCountId = safeText(payload.recurringBusCountId);
  if (!recurringBusCountId) throw new Error('RecurringBusCountId ontbreekt.');

  var request = getRecurringBusCountById(recurringBusCountId);
  if (!request) throw new Error('Recurrente bustelling niet gevonden.');

  var table = getAllValues(getRecurringBusCountTab_());
  if (!table.length) throw new Error('Recurrente bustellingtab is leeg of ongeldig.');

  var headerRow = table[0];
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var rows = readObjectsSafe(getRecurringBusCountTab_()).map(function (row) {
    var current = mapRecurringBusCount(row);
    if (safeText(current.recurringBusCountId) !== recurringBusCountId) {
      return row;
    }

    row.Status = getRecurringBusCountStatusClosed_();
    row.GeslotenOp = toDisplayDateTime(nowRaw);
    if (payload.linkedBusCountId) {
      row.LinkedBusCountID = safeText(payload.linkedBusCountId);
    }
    return row;
  });

  writeFullTable(
    getRecurringBusCountTab_(),
    headerRow,
    rows.map(function (row) {
      return buildRowFromHeaders(headerRow, row);
    })
  );

  writeAudit({
    actie: 'CLOSE_RECURRING_BUS_COUNT_REQUEST',
    actor: actor,
    documentType: 'RecurrenteBusTelling',
    documentId: recurringBusCountId,
    details: {
      linkedBusCountId: safeText(payload.linkedBusCountId),
    },
  });

  return getRecurringBusCountById(recurringBusCountId);
}

/* ---------------------------------------------------------
   Create actual bus count from recurring request
   --------------------------------------------------------- */

function createBusCountFromRecurringRequest(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertRecurringBusCountWriteAccess_(sessionId);
  var recurringBusCountId = safeText(payload.recurringBusCountId);
  if (!recurringBusCountId) throw new Error('RecurringBusCountId ontbreekt.');

  var request = getRecurringBusCountById(recurringBusCountId);
  if (!request) throw new Error('Recurrente bustelling niet gevonden.');

  if (safeText(request.status) !== getRecurringBusCountStatusOpen_()) {
    throw new Error('Recurrente bustelling staat niet in open status.');
  }

  if (typeof createBusCountDraftFromSnapshot !== 'function') {
    throw new Error('BusCount service ontbreekt. Werk eerst het buscountblok in.');
  }

  var articleCodes = parseRecurringArticleCodesCsv_(request.articleCodesCsv);
  var busCount = createBusCountDraftFromSnapshot({
    sessionId: sessionId,
    techniekerCode: request.techniekerCode,
    techniekerNaam: request.techniekerNaam,
    scopeType: articleCodes.length ? getBusCountScopeTargeted_() : getBusCountScopeFull_(),
    requestedBy: safeText(actor.naam || actor.email),
    reason: safeText(request.reason || 'Recurrente bustelling'),
    documentDatum: safeText(payload.documentDatum || payload.documentDate || request.documentDatum),
    opmerking: safeText(payload.opmerking || payload.remark || request.opmerking),
    actor: safeText(actor.naam || actor.email),
    articleCodes: articleCodes,
  });

  var requestTable = getAllValues(getRecurringBusCountTab_());
  if (!requestTable.length) throw new Error('Recurrente bustellingtab is leeg of ongeldig.');

  var requestHeaderRow = requestTable[0];
  var updatedRequestRows = readObjectsSafe(getRecurringBusCountTab_()).map(function (row) {
    var current = mapRecurringBusCount(row);
    if (safeText(current.recurringBusCountId) !== recurringBusCountId) {
      return row;
    }

    row.Status = getRecurringBusCountStatusCreated_();
    row.LinkedBusCountID = safeText(busCount && busCount.busCountId);
    return row;
  });

  writeFullTable(
    getRecurringBusCountTab_(),
    requestHeaderRow,
    updatedRequestRows.map(function (row) {
      return buildRowFromHeaders(requestHeaderRow, row);
    })
  );

  writeAudit({
    actie: 'CREATE_BUS_COUNT_FROM_RECURRING',
    actor: actor,
    documentType: 'RecurrenteBusTelling',
    documentId: recurringBusCountId,
    details: {
      busCountId: safeText(busCount && busCount.busCountId),
      techniekerCode: request.techniekerCode,
    },
  });

  return {
    recurringRequest: getRecurringBusCountById(recurringBusCountId),
    busCount: busCount,
  };
}

/* ---------------------------------------------------------
   Auto materialization from current scoring
   --------------------------------------------------------- */

function createPendingRecurringBusCountRequests(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertRecurringBusCountWriteAccess_(sessionId);
  var candidates = getPendingRecurringBusCountCandidates();
  var created = [];

  candidates.forEach(function (candidate) {
    var request = createRecurringBusCountRequest({
      sessionId: sessionId,
      techniekerCode: candidate.techniekerCode,
      techniekerNaam: candidate.techniekerNaam,
      score: candidate.score,
      sourceType: 'RISK_SCORE',
      reason: candidate.reason,
      articleCodes: candidate.articleCodes,
      documentDatum: Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd'),
      actor: safeText(actor.naam || actor.email),
    });

    created.push(request);
  });

  if (created.length) {
    writeAudit({
      actie: 'CREATE_PENDING_RECURRING_BUS_COUNT_REQUESTS',
      actor: actor,
      documentType: 'RecurrenteBusTelling',
      documentId: 'BATCH',
      details: {
        createdCount: created.length,
      },
    });
  }

  return {
    createdCount: created.length,
    items: created,
  };
}

/* ---------------------------------------------------------
   Queries for screens
   --------------------------------------------------------- */

function buildRecurringBusCountSummary() {
  var allRows = getAllRecurringBusCountRequests();
  var pendingCandidates = getPendingRecurringBusCountCandidates();

  return {
    totaal: allRows.length,
    open: allRows.filter(function (x) { return safeText(x.status) === getRecurringBusCountStatusOpen_(); }).length,
    aangemaakt: allRows.filter(function (x) { return safeText(x.status) === getRecurringBusCountStatusCreated_(); }).length,
    gesloten: allRows.filter(function (x) { return safeText(x.status) === getRecurringBusCountStatusClosed_(); }).length,
    pendingCandidates: pendingCandidates.length,
  };
}

function getRecurringBusCountDashboardData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertRecurringBusCountReadAccess_(sessionId);

  var requests = getAllRecurringBusCountRequests();
  var candidates = buildRecurringBusCountCandidates();

  return {
    requests: requests,
    candidates: candidates,
    pendingCandidates: candidates.filter(function (item) {
      return !item.hasOpenRecurringRequest && !item.hasOpenBusCount;
    }),
    summary: Object.assign(
      {},
      buildRecurringBusCountSummary(),
      { actorRol: safeText(actor.rol) }
    ),
  };
}
