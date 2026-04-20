/* =========================================================
   51_RecurringBusCountService.gs — recurrente bustellingen
   ========================================================= */

/* ---------------------------------------------------------
   Helpers
   --------------------------------------------------------- */

function parseScopeArtikelCodes(scopeArtikelCodes) {
  return safeText(scopeArtikelCodes)
    .split(',')
    .map(x => safeText(x))
    .filter(Boolean);
}

function isBusCountOpenLikeStatus(status) {
  const value = safeText(status);
  return value === BUS_COUNT_STATUS.OPEN || value === BUS_COUNT_STATUS.SUBMITTED;
}

function getOpenLikeBusCounts() {
  return readObjectsSafe(TABS.BUS_COUNTS)
    .map(mapBusCount)
    .filter(count => isBusCountOpenLikeStatus(count.status));
}

function getApprovedBusCounts() {
  return readObjectsSafe(TABS.BUS_COUNTS)
    .map(mapBusCount)
    .filter(count => safeText(count.status) === BUS_COUNT_STATUS.APPROVED);
}

function getLastApprovedBusCountForTechnician(techniekerCode) {
  const code = normalizeRef(techniekerCode);

  const rows = getApprovedBusCounts()
    .filter(count => normalizeRef(count.techniekerCode) === code)
    .sort((a, b) =>
      `${safeText(b.documentDatumIso)} ${safeText(b.tellingId)}`.localeCompare(
        `${safeText(a.documentDatumIso)} ${safeText(a.tellingId)}`
      )
    );

  return rows.length ? rows[0] : null;
}

function hasOpenFullBusCountForTechnician(techniekerCode) {
  const code = normalizeRef(techniekerCode);

  return getOpenLikeBusCounts().some(count =>
    normalizeRef(count.techniekerCode) === code &&
    safeText(count.scopeType || 'Volledig') === 'Volledig'
  );
}

function hasOpenTargetedBusCountForTechnicianArticle(techniekerCode, artikelCode) {
  const techCode = normalizeRef(techniekerCode);
  const articleCode = safeText(artikelCode);

  return getOpenLikeBusCounts().some(count => {
    if (normalizeRef(count.techniekerCode) !== techCode) return false;

    const scopeType = safeText(count.scopeType || 'Volledig');
    if (scopeType === 'Volledig') return true;

    const codes = parseScopeArtikelCodes(count.scopeArtikelCodes);
    return codes.includes(articleCode);
  });
}

function hasOpenAnyBusCountForTechnician(techniekerCode) {
  const code = normalizeRef(techniekerCode);

  return getOpenLikeBusCounts().some(count =>
    normalizeRef(count.techniekerCode) === code
  );
}

function buildBusCountSuggestionKey(type, techniekerCode, articleCode) {
  return [
    safeText(type),
    normalizeRef(techniekerCode),
    safeText(articleCode)
  ].join('|');
}

function daysSinceIsoDate(isoDate) {
  const iso = safeText(isoDate);
  if (!iso) return null;

  const start = new Date(iso + 'T00:00:00');
  if (isNaN(start)) return null;

  const now = new Date();
  return daysBetweenDates(start, now);
}

/* ---------------------------------------------------------
   Suggesties opbouwen
   --------------------------------------------------------- */

function buildPeriodicBusCountSuggestions() {
  const rules = getStockScoreRules();
  const technicians = getActiveTechnicians();
  const suggestions = [];

  technicians.forEach(technician => {
    const hasOpen = hasOpenFullBusCountForTechnician(technician.code);
    if (hasOpen) return;

    const lastApproved = getLastApprovedBusCountForTechnician(technician.code);
    const daysSinceLast = lastApproved ? daysSinceIsoDate(lastApproved.documentDatumIso) : null;

    const due =
      !lastApproved ||
      daysSinceLast === null ||
      daysSinceLast >= Number(rules.periodic_bus_count_days || 60);

    if (!due) return;

    suggestions.push({
      suggestionType: 'PERIODIC_FULL',
      techniekerCode: technician.code,
      techniekerNaam: technician.naam,
      scopeType: 'Volledig',
      requestedArticles: [],
      riskScore: 0,
      signalCount: 0,
      reason: !lastApproved
        ? 'Eerste periodieke bustelling'
        : `Periodieke bustelling (${daysSinceLast} dagen sinds laatste goedgekeurde telling)`,
      sourceRefs: lastApproved ? [lastApproved.tellingId] : []
    });
  });

  return suggestions;
}

function buildRecurringTechnicianBusCountSuggestions() {
  const rules = getStockScoreRules();
  const technicianAlerts = typeof buildRecurringBusCountAlerts === 'function'
    ? buildRecurringBusCountAlerts()
    : [];

  return technicianAlerts
    .filter(alert => !hasOpenFullBusCountForTechnician(alert.techniekerCode))
    .map(alert => ({
      suggestionType: 'RECURRING_FULL',
      techniekerCode: alert.techniekerCode,
      techniekerNaam: alert.techniekerNaam,
      scopeType: 'Volledig',
      requestedArticles: [],
      riskScore: Number(alert.riskScore || 0),
      signalCount: Number(alert.signalCount || 0),
      reason: `Recurrente bustelling wegens verhoogd risicoprofiel (${alert.signalCount} signalen, score ${alert.riskScore})`,
      sourceRefs: alert.refs || []
    }));
}

function buildTargetedArticleBusCountSuggestions() {
  const rules = getStockScoreRules();
  const articleAlerts = typeof buildDeltaCountAlerts === 'function'
    ? buildDeltaCountAlerts()
    : [];

  const suggestions = [];
  const seen = {};

  articleAlerts.forEach(alert => {
    const affectedTechnicians = Array.isArray(alert.affectedTechnicians)
      ? alert.affectedTechnicians
      : [];

    affectedTechnicians.forEach(techniekerCode => {
      if (!safeText(techniekerCode)) return;

      if (hasOpenTargetedBusCountForTechnicianArticle(techniekerCode, alert.artikelCode)) {
        return;
      }

      const key = buildBusCountSuggestionKey('TARGETED_ARTICLE', techniekerCode, alert.artikelCode);
      if (seen[key]) return;
      seen[key] = true;

      suggestions.push({
        suggestionType: 'TARGETED_ARTICLE',
        techniekerCode: techniekerCode,
        techniekerNaam: getTechnicianNameByCode(techniekerCode),
        scopeType: 'Gericht',
        requestedArticles: [alert.artikelCode],
        riskScore: Number(alert.riskScore || 0),
        signalCount: Number(alert.signalCount || 0),
        reason: `Gerichte telling voor artikel ${alert.artikelCode} wegens verhoogde risicoscore (${alert.signalCount} signalen, score ${alert.riskScore})`,
        sourceRefs: alert.refs || [],
        artikelCode: alert.artikelCode,
        artikelOmschrijving: alert.artikelOmschrijving
      });
    });
  });

  return suggestions;
}

function buildGroupedTargetedBusCountSuggestions() {
  const maxArticles = Number(getStockScoreRuleValue('targeted_bus_count_max_articles', 5));
  const targeted = buildTargetedArticleBusCountSuggestions();

  const grouped = {};

  targeted.forEach(item => {
    const techCode = safeText(item.techniekerCode);
    if (!techCode) return;

    if (!grouped[techCode]) {
      grouped[techCode] = {
        suggestionType: 'TARGETED_MULTI',
        techniekerCode: item.techniekerCode,
        techniekerNaam: item.techniekerNaam,
        scopeType: 'Gericht',
        requestedArticles: [],
        riskScore: 0,
        signalCount: 0,
        reason: 'Gerichte bustelling op risicovolle artikels',
        sourceRefs: []
      };
    }

    if (grouped[techCode].requestedArticles.length < maxArticles) {
      uniquePush(grouped[techCode].requestedArticles, item.artikelCode, maxArticles);
    }

    grouped[techCode].riskScore += Number(item.riskScore || 0);
    grouped[techCode].signalCount += Number(item.signalCount || 0);

    (item.sourceRefs || []).forEach(ref => uniquePush(grouped[techCode].sourceRefs, ref, 25));
  });

  return Object.keys(grouped)
    .map(key => grouped[key])
    .filter(item => item.requestedArticles.length > 0)
    .sort((a, b) =>
      Number(b.riskScore || 0) - Number(a.riskScore || 0) ||
      safeText(a.techniekerNaam).localeCompare(safeText(b.techniekerNaam))
    );
}

function buildBusCountSuggestions() {
  const suggestions = []
    .concat(buildPeriodicBusCountSuggestions())
    .concat(buildRecurringTechnicianBusCountSuggestions());

  const groupedTargeted = buildGroupedTargetedBusCountSuggestions();

  const map = {};
  suggestions.forEach(item => {
    const key = buildBusCountSuggestionKey(item.suggestionType, item.techniekerCode, '');
    if (!map[key]) map[key] = item;
  });

  groupedTargeted.forEach(item => {
    const fullExists = Object.values(map).some(existing =>
      normalizeRef(existing.techniekerCode) === normalizeRef(item.techniekerCode) &&
      existing.scopeType === 'Volledig'
    );

    if (!fullExists) {
      const key = buildBusCountSuggestionKey(item.suggestionType, item.techniekerCode, '');
      if (!map[key]) map[key] = item;
    }
  });

  return Object.keys(map)
    .map(key => map[key])
    .sort((a, b) =>
      safeText(a.techniekerNaam).localeCompare(safeText(b.techniekerNaam)) ||
      Number(b.riskScore || 0) - Number(a.riskScore || 0)
    );
}

/* ---------------------------------------------------------
   Suggestiedata ophalen
   --------------------------------------------------------- */

function getRecurringBusCountData(sessionId) {
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om bustellingssuggesties te bekijken.');
  }

  return {
    suggestions: buildBusCountSuggestions(),
    recurringTechnicianAlerts: typeof buildRecurringBusCountAlerts === 'function'
      ? buildRecurringBusCountAlerts()
      : [],
    articleAlerts: typeof buildDeltaCountAlerts === 'function'
      ? buildDeltaCountAlerts()
      : [],
    openBusCounts: getOpenLikeBusCounts()
  };
}

/* ---------------------------------------------------------
   Aanmaken vanuit suggestie
   --------------------------------------------------------- */

function createBusCountFromSuggestion(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om bustelling vanuit suggestie aan te maken.');
  }

  const suggestionType = safeText(payload.suggestionType);
  const techniekerCode = safeText(payload.techniekerCode);
  const documentDatum = safeText(payload.documentDatum);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');
  const requestedArticles = Array.isArray(payload.requestedArticles) ? payload.requestedArticles : [];
  const reason = safeText(payload.reason);

  if (!techniekerCode) throw new Error('TechniekerCode ontbreekt.');
  if (!documentDatum) throw new Error('Documentdatum ontbreekt.');

  if (requestedArticles.length) {
    const uniqueCodes = [...new Set(requestedArticles.map(x => safeText(x)).filter(Boolean))];

    return createBusCountRequest({
      sessionId: sessionId,
      techniekerCode: techniekerCode,
      documentDatum: documentDatum,
      actor: actor,
      reden: reason || `Bustelling (${suggestionType})`,
      requestedArticles: uniqueCodes
    });
  }

  return createBusCountRequest({
    sessionId: sessionId,
    techniekerCode: techniekerCode,
    documentDatum: documentDatum,
    actor: actor,
    reden: reason || `Bustelling (${suggestionType})`,
    requestedArticles: []
  });
}

/* ---------------------------------------------------------
   Bulk create alle missende suggesties
   --------------------------------------------------------- */

function createAllMissingBusCountSuggestions(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = requireLoggedInUser(sessionId);

  if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER])) {
    throw new Error('Geen rechten om bustellingssuggesties aan te maken.');
  }

  const documentDatum = safeText(payload.documentDatum);
  const actor = safeText(payload.actor || user.naam || user.email || 'Magazijn');

  if (!documentDatum) throw new Error('Documentdatum ontbreekt.');

  const suggestions = buildBusCountSuggestions();
  if (!suggestions.length) {
    return {
      success: true,
      created: 0,
      skipped: 0,
      createdIds: [],
      message: 'Geen nieuwe bustellingssuggesties om aan te maken.'
    };
  }

  const createdIds = [];
  let skipped = 0;

  suggestions.forEach(suggestion => {
    try {
      const result = createBusCountFromSuggestion({
        sessionId: sessionId,
        suggestionType: suggestion.suggestionType,
        techniekerCode: suggestion.techniekerCode,
        documentDatum: documentDatum,
        actor: actor,
        requestedArticles: suggestion.requestedArticles || [],
        reason: suggestion.reason
      });

      if (result && result.tellingId) {
        createdIds.push(result.tellingId);
      } else {
        skipped++;
      }
    } catch (e) {
      skipped++;
    }
  });

  writeAudit(
    'Bustellingssuggesties bulk aangemaakt',
    user.rol,
    actor,
    'BusstocktellingSuggesties',
    createdIds.join(', '),
    {
      created: createdIds.length,
      skipped: skipped
    }
  );

  return {
    success: true,
    created: createdIds.length,
    skipped: skipped,
    createdIds: createdIds,
    message: `Bustellingssuggesties verwerkt. Aangemaakt: ${createdIds.length}, overgeslagen: ${skipped}.`
  };
}

/* ---------------------------------------------------------
   Jaarlijkse / periodieke samenvatting
   --------------------------------------------------------- */

function buildBusCountTriggerSummary() {
  const rules = getStockScoreRules();
  const technicians = getActiveTechnicians();

  let periodicDueCount = 0;

  technicians.forEach(technician => {
    const lastApproved = getLastApprovedBusCountForTechnician(technician.code);
    const daysSinceLast = lastApproved ? daysSinceIsoDate(lastApproved.documentDatumIso) : null;

    const due =
      !lastApproved ||
      daysSinceLast === null ||
      daysSinceLast >= Number(rules.periodic_bus_count_days || 60);

    if (due) periodicDueCount++;
  });

  const recurringAlerts = buildRecurringBusCountAlerts();
  const targetedAlerts = buildGroupedTargetedBusCountSuggestions();

  return {
    periodicDueCount: periodicDueCount,
    recurringRiskCount: recurringAlerts.length,
    targetedRiskCount: targetedAlerts.length,
    totalSuggestedCounts: buildBusCountSuggestions().length
  };
}