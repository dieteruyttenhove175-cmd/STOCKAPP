/* =========================================================
   49_StockScoringService.gs — stock scoring + tellingtriggers
   ========================================================= */

/* ---------------------------------------------------------
   Rules
   --------------------------------------------------------- */

function getDefaultStockScoreRules() {
  return {
    grabbel_delta_weight: 3,
    technician_diff_weight: 4,
    receipt_delta_weight: 2,
    buscount_delta_weight: 5,
    correction_weight: 2,
    qty_impact_factor: 0.25,

    alert_threshold: 8,
    central_count_threshold: 12,
    technician_recurrent_threshold: 10,
    minimum_signal_count: 2,

    annual_central_count_days: 365,
    periodic_central_count_days: 90,
    max_refs: 8,
    periodic_bus_count_days: 60,
    high_risk_bus_count_days: 14,
    targeted_bus_count_max_articles: 5
  };
}

function getStockScoreRules() {
  const defaults = getDefaultStockScoreRules();
  const rows = readObjectsSafe(TABS.STOCK_SCORE_RULES);

  if (!rows.length) return defaults;

  const result = { ...defaults };

  rows.forEach(row => {
    const actief = row.Actief === undefined ? true : isTrue(row.Actief);
    if (!actief) return;

    const key = safeText(row.Sleutel || row.Key || row.RuleKey);
    const rawValue = row.Waarde !== undefined ? row.Waarde : (row.Value !== undefined ? row.Value : row.WaardeTekst);

    if (!key) return;
    if (defaults[key] === undefined) return;

    const defaultType = typeof defaults[key];

    if (defaultType === 'number') {
      const parsed = Number(rawValue);
      if (!isNaN(parsed)) result[key] = parsed;
      return;
    }

    result[key] = rawValue;
  });

  return result;
}

function getStockScoreRuleValue(key, defaultValue) {
  const rules = getStockScoreRules();
  return rules[key] !== undefined ? rules[key] : defaultValue;
}

/* ---------------------------------------------------------
   Helpers
   --------------------------------------------------------- */

function isBusLocation(location) {
  return /^Bus:/i.test(safeText(location));
}

function parseDateTimeForDiff(value) {
  const text = safeText(value);
  if (!text) return null;

  const parsed = new Date(text);
  if (!isNaN(parsed)) return parsed;

  return null;
}

function daysBetweenDates(dateA, dateB) {
  if (!dateA || !dateB) return null;
  const ms = Math.abs(dateB.getTime() - dateA.getTime());
  return Math.floor(ms / (24 * 60 * 60 * 1000));
}

function uniquePush(targetArray, value, maxSize) {
  const v = safeText(value);
  if (!v) return;
  if (targetArray.includes(v)) return;
  if (maxSize && targetArray.length >= maxSize) return;
  targetArray.push(v);
}

function ensureArticleScoreBucket(map, articleCode, artikelOmschrijving) {
  const code = safeText(articleCode);
  if (!code) return null;

  if (!map[code]) {
    map[code] = {
      artikelCode: code,
      artikelOmschrijving: safeText(artikelOmschrijving),

      signalCount: 0,
      qtyImpact: 0,
      riskScore: 0,

      grabbelDeltaCount: 0,
      technicianDiffCount: 0,
      receiptDeltaCount: 0,
      busCountDeltaCount: 0,
      correctionCount: 0,

      refs: [],
      affectedTechnicians: [],
      currentCentralStock: 0,
      currentBusStock: 0,
      currentTotalStock: 0
    };
  }

  return map[code];
}

function ensureTechnicianScoreBucket(map, techniekerCode, techniekerNaam) {
  const code = safeText(techniekerCode);
  if (!code) return null;

  if (!map[code]) {
    map[code] = {
      techniekerCode: code,
      techniekerNaam: safeText(techniekerNaam) || getTechnicianNameByCode(code),

      signalCount: 0,
      qtyImpact: 0,
      riskScore: 0,

      grabbelDeltaCount: 0,
      technicianDiffCount: 0,
      busCountDeltaCount: 0,

      articleCodes: [],
      refs: []
    };
  }

  return map[code];
}

function applyArticleSignal(bucket, options) {
  const rules = getStockScoreRules();

  const signalWeight = Number(options.signalWeight || 0);
  const qtyImpact = Number(options.qtyImpact || 0);

  bucket.signalCount += 1;
  bucket.qtyImpact += qtyImpact;
  bucket.riskScore += signalWeight + (qtyImpact * Number(rules.qty_impact_factor || 0));

  if (options.grabbelDelta) bucket.grabbelDeltaCount += 1;
  if (options.technicianDiff) bucket.technicianDiffCount += 1;
  if (options.receiptDelta) bucket.receiptDeltaCount += 1;
  if (options.busCountDelta) bucket.busCountDeltaCount += 1;
  if (options.correction) bucket.correctionCount += 1;

  uniquePush(bucket.refs, options.ref, Number(rules.max_refs || 8));
  uniquePush(bucket.affectedTechnicians, options.techniekerCode, 50);
}

function applyTechnicianSignal(bucket, options) {
  const rules = getStockScoreRules();

  const signalWeight = Number(options.signalWeight || 0);
  const qtyImpact = Number(options.qtyImpact || 0);

  bucket.signalCount += 1;
  bucket.qtyImpact += qtyImpact;
  bucket.riskScore += signalWeight + (qtyImpact * Number(rules.qty_impact_factor || 0));

  if (options.grabbelDelta) bucket.grabbelDeltaCount += 1;
  if (options.technicianDiff) bucket.technicianDiffCount += 1;
  if (options.busCountDelta) bucket.busCountDeltaCount += 1;

  uniquePush(bucket.articleCodes, options.artikelCode, 100);
  uniquePush(bucket.refs, options.ref, Number(rules.max_refs || 8));
}

/* ---------------------------------------------------------
   Busstock opbouwen
   Deze functie wordt in veel andere blokken gebruikt
   --------------------------------------------------------- */

function buildBusStockRows() {
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

    const locationFrom = safeText(move.locatieVan);
    const locationTo = safeText(move.locatieNaar);

    const qtyIn = safeNumber(move.aantalIn, 0);
    const qtyOut = safeNumber(move.aantalUit, 0);

    if (isBusLocation(locationTo)) {
      const techCode = parseBusLocation(locationTo);
      const key = `${techCode}|${code}`;

      if (!grouped[key]) {
        grouped[key] = {
          techniekerCode: techCode,
          techniekerNaam: getTechnicianNameByCode(techCode),
          artikelCode: code,
          artikelOmschrijving: safeText(move.artikelOmschrijving) || safeText(articleMap[code] && articleMap[code].artikelOmschrijving),
          eenheid: safeText(move.eenheid) || safeText(articleMap[code] && articleMap[code].eenheid),
          voorraadBus: 0,
          laatsteMutatie: safeText(move.datumBoeking),
          laatsteMutatieRaw: safeText(move.datumBoekingRaw)
        };
      }

      grouped[key].voorraadBus += qtyIn || Math.abs(safeNumber(move.nettoAantal, 0));

      if (safeText(move.datumBoekingRaw) > safeText(grouped[key].laatsteMutatieRaw)) {
        grouped[key].laatsteMutatie = safeText(move.datumBoeking);
        grouped[key].laatsteMutatieRaw = safeText(move.datumBoekingRaw);
      }
    }

    if (isBusLocation(locationFrom)) {
      const techCode = parseBusLocation(locationFrom);
      const key = `${techCode}|${code}`;

      if (!grouped[key]) {
        grouped[key] = {
          techniekerCode: techCode,
          techniekerNaam: getTechnicianNameByCode(techCode),
          artikelCode: code,
          artikelOmschrijving: safeText(move.artikelOmschrijving) || safeText(articleMap[code] && articleMap[code].artikelOmschrijving),
          eenheid: safeText(move.eenheid) || safeText(articleMap[code] && articleMap[code].eenheid),
          voorraadBus: 0,
          laatsteMutatie: safeText(move.datumBoeking),
          laatsteMutatieRaw: safeText(move.datumBoekingRaw)
        };
      }

      grouped[key].voorraadBus -= qtyOut || Math.abs(safeNumber(move.nettoAantal, 0));

      if (safeText(move.datumBoekingRaw) > safeText(grouped[key].laatsteMutatieRaw)) {
        grouped[key].laatsteMutatie = safeText(move.datumBoeking);
        grouped[key].laatsteMutatieRaw = safeText(move.datumBoekingRaw);
      }
    }
  });

  return Object.keys(grouped)
    .map(key => grouped[key])
    .filter(item => safeNumber(item.voorraadBus, 0) !== 0)
    .sort((a, b) =>
      safeText(a.techniekerNaam).localeCompare(safeText(b.techniekerNaam)) ||
      safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
    );
}

/* ---------------------------------------------------------
   Gecombineerde stock + samenvatting
   --------------------------------------------------------- */

function buildCombinedStockRows() {
  const centralRows = buildCentralWarehouseRows();
  const busRows = buildBusStockRows();

  const map = {};

  centralRows.forEach(item => {
    const code = safeText(item.artikelCode);
    if (!code) return;

    if (!map[code]) {
      map[code] = {
        artikelCode: code,
        artikelOmschrijving: safeText(item.artikelOmschrijving),
        eenheid: safeText(item.eenheid),
        voorraadCentraal: 0,
        voorraadBus: 0,
        voorraadTotaal: 0
      };
    }

    map[code].voorraadCentraal += safeNumber(item.voorraadCentraal, 0);
  });

  busRows.forEach(item => {
    const code = safeText(item.artikelCode);
    if (!code) return;

    if (!map[code]) {
      map[code] = {
        artikelCode: code,
        artikelOmschrijving: safeText(item.artikelOmschrijving),
        eenheid: safeText(item.eenheid),
        voorraadCentraal: 0,
        voorraadBus: 0,
        voorraadTotaal: 0
      };
    }

    map[code].voorraadBus += safeNumber(item.voorraadBus, 0);
  });

  return Object.keys(map)
    .map(key => ({
      ...map[key],
      voorraadTotaal: safeNumber(map[key].voorraadCentraal, 0) + safeNumber(map[key].voorraadBus, 0)
    }))
    .sort((a, b) => safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)));
}

function buildStockScopeSummary() {
  const centralRows = buildCentralWarehouseRows();
  const busRows = buildBusStockRows();
  const totalRows = buildCombinedStockRows();

  return {
    centraalArtikels: centralRows.length,
    busArtikels: busRows.length,
    totaalArtikels: totalRows.length,
    centraalAantal: centralRows.reduce((sum, row) => sum + safeNumber(row.voorraadCentraal, 0), 0),
    busAantal: busRows.reduce((sum, row) => sum + safeNumber(row.voorraadBus, 0), 0),
    totaalAantal: totalRows.reduce((sum, row) => sum + safeNumber(row.voorraadTotaal, 0), 0)
  };
}

/* ---------------------------------------------------------
   Article risk scoring
   --------------------------------------------------------- */

function buildArticleRiskScores() {
  const rules = getStockScoreRules();
  const articleScores = {};

  /* Grabbelstock deltas */
  const orders = readObjectsSafe(TABS.ORDERS).map(mapWarehouseOrder);
  orders.forEach(order => {
    const hasWarehouseDelta =
      safeNumber(order.deltaDozen, 0) > 0 || safeNumber(order.deltaStuks, 0) > 0;

    const hasTechnicianDiff =
      safeText(order.techniekerVerschilReden) !== '';

    if (!hasWarehouseDelta && !hasTechnicianDiff) return;

    const bucket = ensureArticleScoreBucket(
      articleScores,
      order.artikelCode,
      order.artikelOmschrijving
    );
    if (!bucket) return;

    if (hasWarehouseDelta) {
      applyArticleSignal(bucket, {
        signalWeight: Number(rules.grabbel_delta_weight || 0),
        qtyImpact: Math.max(
          safeNumber(order.deltaDozen, 0),
          safeNumber(order.deltaStuks, 0)
        ),
        grabbelDelta: true,
        ref: order.beleveringId || order.bestellingId,
        techniekerCode: order.techniekerCode
      });
    }

    if (hasTechnicianDiff) {
      applyArticleSignal(bucket, {
        signalWeight: Number(rules.technician_diff_weight || 0),
        qtyImpact: Math.abs(
          safeNumber(order.aantalDozenVoorzien, 0) - safeNumber(order.techniekerOntvangenDozen, 0)
        ),
        technicianDiff: true,
        ref: order.beleveringId || order.bestellingId,
        techniekerCode: order.techniekerCode
      });
    }
  });

  /* Receipt deltas */
  const receiptLines = readObjectsSafe(TABS.RECEIPT_LINES)
    .map(mapReceiptLine)
    .filter(line => line.actief);

  receiptLines.forEach(line => {
    if (safeNumber(line.deltaAantal, 0) === 0) return;

    const bucket = ensureArticleScoreBucket(
      articleScores,
      line.artikelCode,
      line.artikelOmschrijving
    );
    if (!bucket) return;

    applyArticleSignal(bucket, {
      signalWeight: Number(rules.receipt_delta_weight || 0),
      qtyImpact: Math.abs(safeNumber(line.deltaAantal, 0)),
      receiptDelta: true,
      ref: line.ontvangstId
    });
  });

  /* Buscount deltas */
  const busCountLines = readObjectsSafe(TABS.BUS_COUNT_LINES)
    .map(mapBusCountLine)
    .filter(line => line.actief);

  const counts = readObjectsSafe(TABS.BUS_COUNTS).map(mapBusCount);
  const countMap = {};
  counts.forEach(count => {
    countMap[count.tellingId] = count;
  });

  busCountLines.forEach(line => {
    if (safeNumber(line.deltaAantal, 0) === 0) return;

    const parent = countMap[line.tellingId] || null;

    const bucket = ensureArticleScoreBucket(
      articleScores,
      line.artikelCode,
      line.artikelOmschrijving
    );
    if (!bucket) return;

    applyArticleSignal(bucket, {
      signalWeight: Number(rules.buscount_delta_weight || 0),
      qtyImpact: Math.abs(safeNumber(line.deltaAantal, 0)),
      busCountDelta: true,
      ref: line.tellingId,
      techniekerCode: parent ? parent.techniekerCode : ''
    });
  });

  /* Correctiemutaties */
  const moves = readObjectsSafe(TABS.WAREHOUSE_MOVEMENTS).map(mapWarehouseMovement);
  moves.forEach(move => {
    const typeMutatie = safeText(move.typeMutatie);
    const isCorrection =
      typeMutatie === MOVEMENT_TYPE.BUS_COUNT_IN ||
      typeMutatie === MOVEMENT_TYPE.BUS_COUNT_OUT ||
      typeMutatie === MOVEMENT_TYPE.RECEIPT_CORRECTION_IN ||
      typeMutatie === MOVEMENT_TYPE.RECEIPT_CORRECTION_OUT ||
      typeMutatie === 'BusCorrectieIn' ||
      typeMutatie === 'BusCorrectieUit' ||
      typeMutatie === 'CorrectieIn' ||
      typeMutatie === 'CorrectieUit';

    if (!isCorrection) return;

    const bucket = ensureArticleScoreBucket(
      articleScores,
      move.artikelCode,
      move.artikelOmschrijving
    );
    if (!bucket) return;

    applyArticleSignal(bucket, {
      signalWeight: Number(rules.correction_weight || 0),
      qtyImpact: Math.abs(safeNumber(move.nettoAantal, 0)) || Math.abs(safeNumber(move.aantalIn, 0)) || Math.abs(safeNumber(move.aantalUit, 0)),
      correction: true,
      ref: move.bronId
    });
  });

  /* Stock snapshot toevoegen */
  const centralMap = buildCentralWarehouseMap();
  const combinedRows = buildCombinedStockRows();
  const combinedMap = {};
  combinedRows.forEach(item => {
    combinedMap[item.artikelCode] = item;
  });

  return Object.keys(articleScores)
    .map(key => {
      const entry = articleScores[key];
      const central = centralMap[key] || null;
      const combined = combinedMap[key] || null;

      return {
        ...entry,
        currentCentralStock: central ? safeNumber(central.voorraadCentraal, 0) : 0,
        currentBusStock: combined ? safeNumber(combined.voorraadBus, 0) : 0,
        currentTotalStock: combined ? safeNumber(combined.voorraadTotaal, 0) : 0,
        shouldTriggerTargetedCount:
          entry.riskScore >= Number(rules.alert_threshold || 0) &&
          entry.signalCount >= Number(rules.minimum_signal_count || 0),
        shouldTriggerCentralCount:
          entry.riskScore >= Number(rules.central_count_threshold || 0)
      };
    })
    .sort((a, b) =>
      Number(b.riskScore || 0) - Number(a.riskScore || 0) ||
      Number(b.signalCount || 0) - Number(a.signalCount || 0) ||
      safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
    );
}

/* ---------------------------------------------------------
   Technician risk scoring
   --------------------------------------------------------- */

function buildTechnicianRiskScores() {
  const rules = getStockScoreRules();
  const technicianScores = {};

  const orders = readObjectsSafe(TABS.ORDERS).map(mapWarehouseOrder);
  orders.forEach(order => {
    const hasWarehouseDelta =
      safeNumber(order.deltaDozen, 0) > 0 || safeNumber(order.deltaStuks, 0) > 0;

    const hasTechnicianDiff =
      safeText(order.techniekerVerschilReden) !== '';

    if (!hasWarehouseDelta && !hasTechnicianDiff) return;
    if (!safeText(order.techniekerCode)) return;

    const bucket = ensureTechnicianScoreBucket(
      technicianScores,
      order.techniekerCode,
      order.techniekerNaam
    );
    if (!bucket) return;

    if (hasWarehouseDelta) {
      applyTechnicianSignal(bucket, {
        signalWeight: Number(rules.grabbel_delta_weight || 0),
        qtyImpact: Math.max(
          safeNumber(order.deltaDozen, 0),
          safeNumber(order.deltaStuks, 0)
        ),
        grabbelDelta: true,
        artikelCode: order.artikelCode,
        ref: order.beleveringId || order.bestellingId
      });
    }

    if (hasTechnicianDiff) {
      applyTechnicianSignal(bucket, {
        signalWeight: Number(rules.technician_diff_weight || 0),
        qtyImpact: Math.abs(
          safeNumber(order.aantalDozenVoorzien, 0) - safeNumber(order.techniekerOntvangenDozen, 0)
        ),
        technicianDiff: true,
        artikelCode: order.artikelCode,
        ref: order.beleveringId || order.bestellingId
      });
    }
  });

  const busCountLines = readObjectsSafe(TABS.BUS_COUNT_LINES)
    .map(mapBusCountLine)
    .filter(line => line.actief);

  const counts = readObjectsSafe(TABS.BUS_COUNTS).map(mapBusCount);
  const countMap = {};
  counts.forEach(count => {
    countMap[count.tellingId] = count;
  });

  busCountLines.forEach(line => {
    if (safeNumber(line.deltaAantal, 0) === 0) return;

    const parent = countMap[line.tellingId] || null;
    if (!parent || !safeText(parent.techniekerCode)) return;

    const bucket = ensureTechnicianScoreBucket(
      technicianScores,
      parent.techniekerCode,
      parent.techniekerNaam
    );
    if (!bucket) return;

    applyTechnicianSignal(bucket, {
      signalWeight: Number(rules.buscount_delta_weight || 0),
      qtyImpact: Math.abs(safeNumber(line.deltaAantal, 0)),
      busCountDelta: true,
      artikelCode: line.artikelCode,
      ref: line.tellingId
    });
  });

  return Object.keys(technicianScores)
    .map(key => ({
      ...technicianScores[key],
      shouldTriggerRecurringBusCount:
        Number(technicianScores[key].riskScore || 0) >= Number(rules.technician_recurrent_threshold || 0)
    }))
    .sort((a, b) =>
      Number(b.riskScore || 0) - Number(a.riskScore || 0) ||
      Number(b.signalCount || 0) - Number(a.signalCount || 0) ||
      safeText(a.techniekerNaam).localeCompare(safeText(b.techniekerNaam))
    );
}

/* ---------------------------------------------------------
   Alerts die bestaande dashboards al gebruiken
   --------------------------------------------------------- */

function buildDeltaCountAlerts() {
  return buildArticleRiskScores()
    .filter(item => item.shouldTriggerTargetedCount)
    .map(item => ({
      artikelCode: item.artikelCode,
      artikelOmschrijving: item.artikelOmschrijving,
      signalCount: item.signalCount,
      qtyImpact: item.qtyImpact,
      riskScore: Number(item.riskScore.toFixed(2)),
      grabbelDeltaCount: item.grabbelDeltaCount,
      receiptDeltaCount: item.receiptDeltaCount,
      technicianDiffCount: item.technicianDiffCount,
      busCountDeltaCount: item.busCountDeltaCount,
      correctionCount: item.correctionCount,
      currentCentralStock: item.currentCentralStock,
      currentBusStock: item.currentBusStock,
      currentTotalStock: item.currentTotalStock,
      refs: item.refs,
      affectedTechnicians: item.affectedTechnicians
    }));
}

function buildRecurringBusCountAlerts() {
  return buildTechnicianRiskScores()
    .filter(item => item.shouldTriggerRecurringBusCount)
    .map(item => ({
      techniekerCode: item.techniekerCode,
      techniekerNaam: item.techniekerNaam,
      signalCount: item.signalCount,
      qtyImpact: item.qtyImpact,
      riskScore: Number(item.riskScore.toFixed(2)),
      grabbelDeltaCount: item.grabbelDeltaCount,
      technicianDiffCount: item.technicianDiffCount,
      busCountDeltaCount: item.busCountDeltaCount,
      articleCodes: item.articleCodes,
      refs: item.refs
    }));
}

/* ---------------------------------------------------------
   Centrale telling triggers
   --------------------------------------------------------- */

function getLastAuditMomentForActionText(searchText) {
  const needle = safeText(searchText).toLowerCase();
  if (!needle) return null;

  const rows = readObjectsSafe(TABS.AUDIT_LOG);
  const matching = rows
    .filter(row => safeText(row.Actie).toLowerCase().includes(needle))
    .sort((a, b) => safeText(b.Tijdstip).localeCompare(safeText(a.Tijdstip)));

  if (!matching.length) return null;

  return parseDateTimeForDiff(matching[0].Tijdstip);
}

function buildCentralCountTriggerSummary() {
  const rules = getStockScoreRules();
  const highRiskArticles = buildArticleRiskScores().filter(item => item.shouldTriggerCentralCount);

  const now = new Date();
  const lastYearly = getLastAuditMomentForActionText('centrale stocktelling');
  const lastPeriodic = getLastAuditMomentForActionText('periodieke centrale telling');

  const daysSinceLastYearly = lastYearly ? daysBetweenDates(lastYearly, now) : null;
  const daysSinceLastPeriodic = lastPeriodic ? daysBetweenDates(lastPeriodic, now) : null;

  return {
    yearlyDue: !lastYearly || daysSinceLastYearly >= Number(rules.annual_central_count_days || 365),
    periodicDue:
      highRiskArticles.length > 0 &&
      (
        !lastPeriodic ||
        daysSinceLastPeriodic >= Number(rules.periodic_central_count_days || 90)
      ),
    daysSinceLastYearly: daysSinceLastYearly,
    daysSinceLastPeriodic: daysSinceLastPeriodic,
    highRiskArticleCount: highRiskArticles.length,
    highRiskArticles: highRiskArticles.slice(0, 20).map(item => ({
      artikelCode: item.artikelCode,
      artikelOmschrijving: item.artikelOmschrijving,
      riskScore: Number(item.riskScore.toFixed(2)),
      signalCount: item.signalCount
    }))
  };
}

/* ---------------------------------------------------------
   Snapshot voor latere dashboards / services
   --------------------------------------------------------- */

function getStockScoringSnapshot(sessionId) {
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER, ROLE.ANALYSIS])) {
    throw new Error('Geen rechten om stock scoring te bekijken.');
  }

  return {
    articleScores: buildArticleRiskScores(),
    articleAlerts: buildDeltaCountAlerts(),
    technicianAlerts: buildRecurringBusCountAlerts(),
    centralCountSummary: buildCentralCountTriggerSummary(),
    stockScopeSummary: buildStockScopeSummary(),
    combinedStock: buildCombinedStockRows()
  };
}