/* =========================================================
   49_StockScoringService.gs
   Refactor: stock scoring core service
   Doel:
   - centrale risicoscores opbouwen uit afwijkingen
   - artikelscores voor centraal magazijn en algemene stock
   - techniekerscores voor recurrente bustellingen
   - gerichte alerts en suggesties voor tellingen
   ========================================================= */

/* ---------------------------------------------------------
   Config / fallbacks
   --------------------------------------------------------- */

function getStockScoringConfig_() {
  var cfg = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG) ? APP_CONFIG : {};

  return {
    grabbelDeltaWeight: safeNumber(cfg.GRABBEL_DELTA_WEIGHT, 2),
    receiptDeltaWeight: safeNumber(cfg.RECEIPT_DELTA_WEIGHT, 3),
    busCountDeltaWeight: safeNumber(cfg.BUS_COUNT_DELTA_WEIGHT, 4),
    centralCountDeltaWeight: safeNumber(cfg.CENTRAL_COUNT_DELTA_WEIGHT, 5),
    correctionMovementWeight: safeNumber(cfg.CORRECTION_MOVEMENT_WEIGHT, 2),

    recentWindowDays: safeNumber(cfg.STOCK_SCORE_RECENT_WINDOW_DAYS, 90),

    articleAlertThreshold: safeNumber(cfg.STOCK_SCORE_ARTICLE_ALERT_THRESHOLD, 8),
    centralCountSuggestionThreshold: safeNumber(cfg.STOCK_SCORE_CENTRAL_COUNT_THRESHOLD, 12),
    technicianAlertThreshold: safeNumber(cfg.STOCK_SCORE_TECHNICIAN_ALERT_THRESHOLD, 8),
    recurringBusCountThreshold: safeNumber(cfg.STOCK_SCORE_RECURRING_BUSCOUNT_THRESHOLD, 12),
  };
}

function getStockScoringTodayRaw_() {
  return Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function isRecentForStockScoring_(rawDate, maxDays) {
  var raw = safeText(rawDate);
  if (!raw) return false;

  var dt = new Date(raw);
  if (isNaN(dt.getTime())) return false;

  var now = new Date();
  var diffMs = now.getTime() - dt.getTime();
  var maxMs = safeNumber(maxDays, 90) * 24 * 60 * 60 * 1000;

  return diffMs >= 0 && diffMs <= maxMs;
}

function createEmptyArticleRisk_(artikelCode, artikelOmschrijving, typeMateriaal, eenheid) {
  return {
    artikelCode: safeText(artikelCode),
    artikelOmschrijving: safeText(artikelOmschrijving),
    typeMateriaal: safeText(typeMateriaal),
    eenheid: safeText(eenheid),
    score: 0,

    grabbelDeltaCount: 0,
    receiptDeltaCount: 0,
    busCountDeltaCount: 0,
    centralCountDeltaCount: 0,
    correctionMovementCount: 0,

    lastSeenRaw: '',
    lastSeenType: '',
  };
}

function createEmptyTechnicianRisk_(techniekerCode, techniekerNaam) {
  return {
    techniekerCode: safeText(techniekerCode),
    techniekerNaam: safeText(techniekerNaam),
    score: 0,

    busCountDeltaCount: 0,
    grabbelIssueCount: 0,
    correctionMovementCount: 0,

    lastSeenRaw: '',
    lastSeenType: '',
  };
}

function touchRiskTimestamp_(target, rawDate, signalType) {
  var raw = safeText(rawDate);
  if (!raw) return;

  if (!safeText(target.lastSeenRaw) || raw > safeText(target.lastSeenRaw)) {
    target.lastSeenRaw = raw;
    target.lastSeenType = safeText(signalType);
  }
}

/* ---------------------------------------------------------
   Source readers with graceful fallback
   --------------------------------------------------------- */

function getStockScoringGrabbelGroups_() {
  if (typeof buildGrabbelOrderGroups === 'function' && typeof getAllGrabbelOrders === 'function') {
    return buildGrabbelOrderGroups(getAllGrabbelOrders());
  }
  return [];
}

function getStockScoringReceipts_() {
  if (typeof getReceiptsWithLines === 'function') {
    return getReceiptsWithLines();
  }
  return [];
}

function getStockScoringBusCounts_() {
  if (typeof getBusCountsWithLines === 'function') {
    return getBusCountsWithLines();
  }
  return [];
}

function getStockScoringCentralCounts_() {
  if (typeof getCentralCountsWithLines === 'function') {
    return getCentralCountsWithLines();
  }
  return [];
}

function getStockScoringMovements_() {
  if (typeof getAllMovements === 'function') {
    return getAllMovements();
  }
  return [];
}

/* ---------------------------------------------------------
   Signal builders
   --------------------------------------------------------- */

function buildGrabbelDeltaSignals_() {
  var groups = getStockScoringGrabbelGroups_();

  var signals = [];
  groups.forEach(function (group) {
    var raw = '';
    if (group.lines && group.lines.length) {
      raw = safeText(group.lines[0].aangemaaktOpRaw || group.lines[0].documentDatumIso || '');
    }

    (group.lines || []).forEach(function (line) {
      var delta = Math.abs(
        safeNumber(line.voorzienAantal, 0) - safeNumber(line.gevraagdAantal, 0)
      );
      if (!delta) return;

      signals.push({
        signalType: 'GRABBEL_DELTA',
        artikelCode: safeText(line.artikelCode),
        artikelOmschrijving: safeText(line.artikelOmschrijving),
        typeMateriaal: safeText(line.typeMateriaal),
        eenheid: safeText(line.eenheid),
        techniekerCode: safeText(group.techniekerCode),
        techniekerNaam: safeText(group.techniekerNaam),
        rawDate: raw,
        impactUnits: delta,
      });
    });
  });

  return signals;
}

function buildReceiptDeltaSignals_() {
  var receipts = getStockScoringReceipts_();

  var signals = [];
  receipts.forEach(function (receipt) {
    var raw = safeText(
      receipt.aangemaaktOpRaw ||
      receipt.documentDatumIso ||
      receipt.ontvangstDatumIso ||
      ''
    );

    (receipt.lines || []).forEach(function (line) {
      var delta = Math.abs(safeNumber(line.deltaAantal, 0));
      if (!delta) return;

      signals.push({
        signalType: 'RECEIPT_DELTA',
        artikelCode: safeText(line.artikelCode),
        artikelOmschrijving: safeText(line.artikelOmschrijving),
        typeMateriaal: safeText(line.typeMateriaal),
        eenheid: safeText(line.eenheid),
        techniekerCode: '',
        techniekerNaam: '',
        rawDate: raw,
        impactUnits: delta,
      });
    });
  });

  return signals;
}

function buildBusCountDeltaSignals_() {
  var counts = getStockScoringBusCounts_();

  var signals = [];
  counts.forEach(function (count) {
    var raw = safeText(count.aangemaaktOpRaw || count.documentDatumIso || '');

    (count.lines || []).forEach(function (line) {
      var delta = Math.abs(safeNumber(line.deltaAantal, 0));
      if (!delta) return;

      signals.push({
        signalType: 'BUSCOUNT_DELTA',
        artikelCode: safeText(line.artikelCode),
        artikelOmschrijving: safeText(line.artikelOmschrijving),
        typeMateriaal: safeText(line.typeMateriaal),
        eenheid: safeText(line.eenheid),
        techniekerCode: safeText(count.techniekerCode),
        techniekerNaam: safeText(count.techniekerNaam),
        rawDate: raw,
        impactUnits: delta,
      });
    });
  });

  return signals;
}

function buildCentralCountDeltaSignals_() {
  var counts = getStockScoringCentralCounts_();

  var signals = [];
  counts.forEach(function (count) {
    var raw = safeText(count.aangemaaktOpRaw || count.documentDatumIso || '');

    (count.lines || []).forEach(function (line) {
      var delta = Math.abs(safeNumber(line.deltaAantal, 0));
      if (!delta) return;

      signals.push({
        signalType: 'CENTRALCOUNT_DELTA',
        artikelCode: safeText(line.artikelCode),
        artikelOmschrijving: safeText(line.artikelOmschrijving),
        typeMateriaal: safeText(line.typeMateriaal),
        eenheid: safeText(line.eenheid),
        techniekerCode: '',
        techniekerNaam: '',
        rawDate: raw,
        impactUnits: delta,
      });
    });
  });

  return signals;
}

function buildCorrectionMovementSignals_() {
  var movements = getStockScoringMovements_();

  var correctionTypes = {};
  correctionTypes['BusCorrectieIn'] = true;
  correctionTypes['BusCorrectieUit'] = true;
  correctionTypes['CentralCountIn'] = true;
  correctionTypes['CentralCountOut'] = true;
  correctionTypes['CentralCountIn'.toLowerCase()] = true;
  correctionTypes['CentralCountOut'.toLowerCase()] = true;

  if (typeof getBusCountMovementIn_ === 'function') {
    correctionTypes[safeText(getBusCountMovementIn_())] = true;
  }
  if (typeof getBusCountMovementOut_ === 'function') {
    correctionTypes[safeText(getBusCountMovementOut_())] = true;
  }
  if (typeof getCentralCountMovementIn_ === 'function') {
    correctionTypes[safeText(getCentralCountMovementIn_())] = true;
  }
  if (typeof getCentralCountMovementOut_ === 'function') {
    correctionTypes[safeText(getCentralCountMovementOut_())] = true;
  }

  var signals = [];
  movements.forEach(function (move) {
    var type = safeText(move.movementType);
    if (!correctionTypes[type] && !correctionTypes[type.toLowerCase()]) {
      return;
    }

    var impact = Math.abs(safeNumber(move.nettoAantal, 0)) ||
      safeNumber(move.aantalIn, 0) ||
      safeNumber(move.aantalUit, 0);

    if (!impact) return;

    var techniekerCode = '';
    if (typeof parseBusLocation === 'function' && safeText(move.locatieVan)) {
      try {
        techniekerCode = safeText(parseBusLocation(move.locatieVan));
      } catch (e) {}
    }
    if (!techniekerCode && typeof parseBusLocation === 'function' && safeText(move.locatieNaar)) {
      try {
        techniekerCode = safeText(parseBusLocation(move.locatieNaar));
      } catch (e) {}
    }

    signals.push({
      signalType: 'CORRECTION_MOVEMENT',
      artikelCode: safeText(move.artikelCode),
      artikelOmschrijving: safeText(move.artikelOmschrijving),
      typeMateriaal: safeText(move.typeMateriaal),
      eenheid: safeText(move.eenheid),
      techniekerCode: techniekerCode,
      techniekerNaam: '',
      rawDate: safeText(move.datumBoekingRaw),
      impactUnits: impact,
    });
  });

  return signals;
}

/* ---------------------------------------------------------
   Scoring logic
   --------------------------------------------------------- */

function getStockRiskSignals() {
  return []
    .concat(buildGrabbelDeltaSignals_())
    .concat(buildReceiptDeltaSignals_())
    .concat(buildBusCountDeltaSignals_())
    .concat(buildCentralCountDeltaSignals_())
    .concat(buildCorrectionMovementSignals_());
}

function buildArticleRiskScores() {
  var cfg = getStockScoringConfig_();
  var scores = {};

  getStockRiskSignals().forEach(function (signal) {
    if (!isRecentForStockScoring_(signal.rawDate, cfg.recentWindowDays)) {
      return;
    }

    var artikelCode = safeText(signal.artikelCode);
    if (!artikelCode) return;

    if (!scores[artikelCode]) {
      scores[artikelCode] = createEmptyArticleRisk_(
        artikelCode,
        signal.artikelOmschrijving,
        signal.typeMateriaal,
        signal.eenheid
      );
    }

    var item = scores[artikelCode];
    var impact = Math.max(1, safeNumber(signal.impactUnits, 1));

    if (signal.signalType === 'GRABBEL_DELTA') {
      item.score += cfg.grabbelDeltaWeight * impact;
      item.grabbelDeltaCount += 1;
    } else if (signal.signalType === 'RECEIPT_DELTA') {
      item.score += cfg.receiptDeltaWeight * impact;
      item.receiptDeltaCount += 1;
    } else if (signal.signalType === 'BUSCOUNT_DELTA') {
      item.score += cfg.busCountDeltaWeight * impact;
      item.busCountDeltaCount += 1;
    } else if (signal.signalType === 'CENTRALCOUNT_DELTA') {
      item.score += cfg.centralCountDeltaWeight * impact;
      item.centralCountDeltaCount += 1;
    } else if (signal.signalType === 'CORRECTION_MOVEMENT') {
      item.score += cfg.correctionMovementWeight * impact;
      item.correctionMovementCount += 1;
    }

    touchRiskTimestamp_(item, signal.rawDate, signal.signalType);

    if (!item.artikelOmschrijving) item.artikelOmschrijving = safeText(signal.artikelOmschrijving);
    if (!item.typeMateriaal) item.typeMateriaal = safeText(signal.typeMateriaal);
    if (!item.eenheid) item.eenheid = safeText(signal.eenheid);
  });

  return Object.keys(scores)
    .map(function (key) { return scores[key]; })
    .sort(function (a, b) {
      return (
        safeNumber(b.score, 0) - safeNumber(a.score, 0) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function buildTechnicianRiskScores() {
  var cfg = getStockScoringConfig_();
  var scores = {};

  getStockRiskSignals().forEach(function (signal) {
    if (!isRecentForStockScoring_(signal.rawDate, cfg.recentWindowDays)) {
      return;
    }

    var techniekerCode = safeText(signal.techniekerCode);
    if (!techniekerCode) return;

    if (!scores[techniekerCode]) {
      scores[techniekerCode] = createEmptyTechnicianRisk_(
        techniekerCode,
        signal.techniekerNaam
      );
    }

    var item = scores[techniekerCode];
    var impact = Math.max(1, safeNumber(signal.impactUnits, 1));

    if (signal.signalType === 'BUSCOUNT_DELTA') {
      item.score += cfg.busCountDeltaWeight * impact;
      item.busCountDeltaCount += 1;
    } else if (signal.signalType === 'GRABBEL_DELTA') {
      item.score += cfg.grabbelDeltaWeight * impact;
      item.grabbelIssueCount += 1;
    } else if (signal.signalType === 'CORRECTION_MOVEMENT') {
      item.score += cfg.correctionMovementWeight * impact;
      item.correctionMovementCount += 1;
    }

    touchRiskTimestamp_(item, signal.rawDate, signal.signalType);

    if (!item.techniekerNaam) {
      item.techniekerNaam = safeText(signal.techniekerNaam);
    }
  });

  return Object.keys(scores)
    .map(function (key) { return scores[key]; })
    .sort(function (a, b) {
      return (
        safeNumber(b.score, 0) - safeNumber(a.score, 0) ||
        safeText(a.techniekerNaam).localeCompare(safeText(b.techniekerNaam)) ||
        safeText(a.techniekerCode).localeCompare(safeText(b.techniekerCode))
      );
    });
}

/* ---------------------------------------------------------
   Alerts / suggestions
   --------------------------------------------------------- */

function buildArticleRiskAlerts() {
  var cfg = getStockScoringConfig_();

  return buildArticleRiskScores()
    .filter(function (item) {
      return safeNumber(item.score, 0) >= cfg.articleAlertThreshold;
    })
    .map(function (item) {
      return Object.assign({}, item, {
        alertType: 'ARTICLE_RISK',
        suggestedAction:
          safeNumber(item.score, 0) >= cfg.centralCountSuggestionThreshold
            ? 'Gerichte centrale telling'
            : 'Artikel opvolgen',
      });
    });
}

function buildCentralCountSuggestions() {
  var cfg = getStockScoringConfig_();

  return buildArticleRiskScores()
    .filter(function (item) {
      return safeNumber(item.score, 0) >= cfg.centralCountSuggestionThreshold;
    })
    .map(function (item) {
      return {
        suggestionType: 'CENTRAL_COUNT',
        artikelCode: item.artikelCode,
        artikelOmschrijving: item.artikelOmschrijving,
        typeMateriaal: item.typeMateriaal,
        eenheid: item.eenheid,
        score: item.score,
        reason:
          'Artikel heeft verhoogd risicoprofiel door delta’s, tellingen of correctiemutaties.',
      };
    });
}

function buildRecurringBusCountSuggestions() {
  var cfg = getStockScoringConfig_();

  return buildTechnicianRiskScores()
    .filter(function (item) {
      return safeNumber(item.score, 0) >= cfg.recurringBusCountThreshold;
    })
    .map(function (item) {
      return {
        suggestionType: 'RECURRING_BUS_COUNT',
        techniekerCode: item.techniekerCode,
        techniekerNaam: item.techniekerNaam,
        score: item.score,
        reason:
          'Technieker vertoont herhaalde afwijkingen of correcties en vraagt recurrente bustelling.',
      };
    });
}

function buildBusCountTriggerSummary() {
  var technicianSuggestions = buildRecurringBusCountSuggestions();
  var articleSuggestions = buildArticleRiskAlerts();

  return {
    technicianSuggestions: technicianSuggestions,
    articleSuggestions: articleSuggestions,
    technicianSuggestionCount: technicianSuggestions.length,
    articleSuggestionCount: articleSuggestions.length,
  };
}

/* ---------------------------------------------------------
   Dashboard query
   --------------------------------------------------------- */

function getStockRiskDashboardData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  assertRoleAllowed(
    actor,
    [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.MOBILE_WAREHOUSE],
    'Geen rechten om stockrisico te bekijken.'
  );

  var articleScores = buildArticleRiskScores();
  var technicianScores = buildTechnicianRiskScores();
  var articleAlerts = buildArticleRiskAlerts();
  var centralSuggestions = buildCentralCountSuggestions();
  var recurringBusSuggestions = buildRecurringBusCountSuggestions();

  return {
    articleScores: articleScores,
    technicianScores: technicianScores,
    articleAlerts: articleAlerts,
    centralCountSuggestions: centralSuggestions,
    recurringBusCountSuggestions: recurringBusSuggestions,
    summary: {
      articleRiskCount: articleAlerts.length,
      centralCountSuggestionCount: centralSuggestions.length,
      recurringBusCountSuggestionCount: recurringBusSuggestions.length,
      topArticleScore: articleScores.length ? safeNumber(articleScores[0].score, 0) : 0,
      topTechnicianScore: technicianScores.length ? safeNumber(technicianScores[0].score, 0) : 0,
    },
  };
}
