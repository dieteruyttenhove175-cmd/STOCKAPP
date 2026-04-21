/* =========================================================
   45A_ConsumptionProposalService.gs — voorstellen behoefte
   Snapshots + delivery moments -> automatische behoeftevoorstellen
   ========================================================= */

function buildAutoNeedDeliveryMoments_(daysAhead) {
  const todayIso = toIsoDate(nowDate());
  const maxIso = addDaysToIsoDate(todayIso, safeNumber(daysAhead, 7));

  const moments = [];

  getActiveTechnicians().forEach(technician => {
    const allDeliveries = getDeliveriesForTechnician(technician)
      .map(enrichDeliveryWithTimingState)
      .filter(item => safeText(item.datumIso))
      .sort((a, b) =>
        `${safeText(a.datumIso)} ${safeText(a.tijdslot)}`.localeCompare(
          `${safeText(b.datumIso)} ${safeText(b.tijdslot)}`
        )
      );

    allDeliveries.forEach((delivery, idx) => {
      if (safeText(delivery.datumIso) < todayIso) return;
      if (safeText(delivery.datumIso) > maxIso) return;

      const previous = idx > 0 ? allDeliveries[idx - 1] : null;

      moments.push({
        key: [
          safeText(technician.code),
          safeText(delivery.beleveringId || ''),
          safeText(delivery.datumIso),
          safeText(delivery.tijdslot)
        ].join('|'),
        techniekerCode: safeText(technician.code),
        techniekerNaam: safeText(technician.naam),
        deliveryId: safeText(delivery.beleveringId || ''),
        deliveryDateIso: safeText(delivery.datumIso),
        deliveryTime: safeText(delivery.tijdslot),
        deliverySortKey: safeText(delivery.sortKey || `${delivery.datumIso} ${delivery.tijdslot}`),
        deliveryLabel: `${toDisplayDate(delivery.datumIso)} ${safeText(delivery.tijdslot)}`,
        currentCutoffIso: addDaysToIsoDate(delivery.datumIso, -1),
        previousDeliveryDateIso: previous ? safeText(previous.datumIso) : '',
        previousCutoffIso: previous ? addDaysToIsoDate(previous.datumIso, -1) : ''
      });
    });
  });

  return moments.sort((a, b) => safeText(a.deliverySortKey).localeCompare(safeText(b.deliverySortKey)));
}

function getSnapshotArticleCodesForTechnician_(techniekerCode, cutoffIso) {
  const cutoff = safeText(cutoffIso);
  if (!cutoff) return [];

  const codes = {};
  getAllConsumptionSnapshots()
    .filter(row =>
      normalizeRef(row.techniekerCode) === normalizeRef(techniekerCode) &&
      safeText(row.snapshotDatum) <= cutoff
    )
    .forEach(row => {
      const code = safeText(row.artikelCode);
      if (code) codes[code] = true;
    });

  return Object.keys(codes).sort();
}

function getTransferredNeedQtyFromMobileToBus_(techniekerCode, artikelCode, fromIsoExclusive, toIsoInclusive) {
  const techCode = safeText(techniekerCode);
  const articleCode = safeText(artikelCode);
  if (!techCode || !articleCode) return 0;

  const toLoc = getBusLocationCode(techCode);

  return getAllTransfers()
    .filter(transfer =>
      normalizeRef(transfer.doelTechniekerCode) === normalizeRef(techCode) &&
      safeText(transfer.naarLocatie) === toLoc &&
      (
        safeText(transfer.vanLocatie) === LOCATION.MOBILE ||
        safeText(transfer.vanLocatie).indexOf(`${LOCATION.MOBILE}:`) === 0
      ) &&
      ['Geboekt', 'Goedgekeurd', 'Ontvangen'].includes(safeText(transfer.status))
    )
    .filter(transfer => {
      const dateIso = safeText(transfer.documentDatumIso);
      if (!dateIso) return false;
      if (fromIsoExclusive && dateIso <= fromIsoExclusive) return false;
      if (toIsoInclusive && dateIso > toIsoInclusive) return false;
      return true;
    })
    .reduce((sum, transfer) => {
      const qty = getTransferLinesById(transfer.transferId)
        .filter(line =>
          safeText(line.artikelCode) === articleCode &&
          (
            safeText(line.typeMateriaal) === MATERIAL_TYPE.NEED ||
            determineMaterialTypeFromArticle(line.artikelCode) === MATERIAL_TYPE.NEED
          )
        )
        .reduce((lineSum, line) => lineSum + safeNumber(line.aantal, 0), 0);

      return sum + qty;
    }, 0);
}

function getOpenNeedShortfallForPeriod_(techniekerCode, artikelCode, fromIsoExclusive, toIsoInclusive) {
  // TODO:
  // Hier later open rest van vorige levering verwerken.
  // Voor nu 0 zodat de eerste automatische flow stabiel blijft.
  return 0;
}

function buildAutoNeedProposalRows_(daysAhead) {
  const moments = buildAutoNeedDeliveryMoments_(daysAhead);
  const rows = [];

  moments.forEach(moment => {
    const articleCodes = getSnapshotArticleCodesForTechnician_(
      moment.techniekerCode,
      moment.currentCutoffIso
    );

    articleCodes.forEach(artikelCode => {
      const currentSnapshot = getLatestConsumptionSnapshotOnOrBefore_(
        moment.techniekerCode,
        artikelCode,
        moment.currentCutoffIso
      );

      const previousSnapshot = moment.previousCutoffIso
        ? getLatestConsumptionSnapshotOnOrBefore_(
            moment.techniekerCode,
            artikelCode,
            moment.previousCutoffIso
          )
        : null;

      const currentValue = currentSnapshot ? safeNumber(currentSnapshot.cumulatiefVerbruik, 0) : 0;
      const previousValue = previousSnapshot ? safeNumber(previousSnapshot.cumulatiefVerbruik, 0) : 0;
      const verbruikDelta = Math.max(0, currentValue - previousValue);

      const transfersQty = getTransferredNeedQtyFromMobileToBus_(
        moment.techniekerCode,
        artikelCode,
        moment.previousCutoffIso,
        moment.currentCutoffIso
      );

      const shortfallQty = getOpenNeedShortfallForPeriod_(
        moment.techniekerCode,
        artikelCode,
        moment.previousCutoffIso,
        moment.currentCutoffIso
      );

      const voorstelAantal = Math.max(0, verbruikDelta - transfersQty + shortfallQty);
      if (voorstelAantal <= 0) return;

      const article = getArticleMaster(artikelCode) || {};

      rows.push({
        proposalKey: [
          moment.key,
          safeText(artikelCode)
        ].join('|'),
        techniekerCode: moment.techniekerCode,
        techniekerNaam: moment.techniekerNaam,
        deliveryId: moment.deliveryId,
        deliveryDateIso: moment.deliveryDateIso,
        deliveryLabel: moment.deliveryLabel,
        currentCutoffIso: moment.currentCutoffIso,
        previousCutoffIso: moment.previousCutoffIso,
        artikelCode: artikelCode,
        artikelOmschrijving: safeText(
          (currentSnapshot && currentSnapshot.artikelOmschrijving) ||
          (previousSnapshot && previousSnapshot.artikelOmschrijving) ||
          article.artikelOmschrijving
        ),
        eenheid: safeText(
          (currentSnapshot && currentSnapshot.eenheid) ||
          (previousSnapshot && previousSnapshot.eenheid) ||
          article.eenheid
        ),
        previousSnapshot: previousValue,
        currentSnapshot: currentValue,
        verbruikDelta: verbruikDelta,
        transfersQty: transfersQty,
        shortfallQty: shortfallQty,
        voorstelAantal: voorstelAantal
      });
    });
  });

  return rows.sort((a, b) =>
    `${safeText(a.deliveryDateIso)} ${safeText(a.techniekerNaam)} ${safeText(a.artikelCode)}`.localeCompare(
      `${safeText(b.deliveryDateIso)} ${safeText(b.techniekerNaam)} ${safeText(b.artikelCode)}`
    )
  );
}

function getAutoNeedProposals(payload) {
  const sessionId = getPayloadSessionId(payload);
  assertWarehouseOrMobileAccess(sessionId);

  const daysAhead = safeNumber((payload && payload.daysAhead) || 7, 7);

  return {
    proposals: buildAutoNeedProposalRows_(daysAhead)
  };
}

function findOpenAutoNeedIssueForMoment_(techniekerCode, deliveryDateIso) {
  return getAllNeedIssues().find(issue =>
    normalizeRef(issue.techniekerCode) === normalizeRef(techniekerCode) &&
    safeText(issue.documentDatumIso) === safeText(deliveryDateIso) &&
    safeText(issue.status) === NEED_ISSUE_STATUS.OPEN &&
    safeText(issue.reden) === 'Auto aanvulling verbruik'
  ) || null;
}

function upsertProposalLineInNeedIssue_(uitgifteId, proposal) {
  const currentIssue = getNeedIssueById(uitgifteId);
  if (!currentIssue) {
    throw new Error('Behoefte-uitgifte niet gevonden voor voorstel.');
  }

  const existingLines = getNeedIssueLinesByIssueId(uitgifteId).map(line => ({
    artikelCode: safeText(line.artikelCode),
    artikelOmschrijving: safeText(line.artikelOmschrijving),
    eenheid: safeText(line.eenheid),
    aantal: safeNumber(line.aantal, 0)
  }));

  let found = false;

  const mergedLines = existingLines.map(line => {
    if (safeText(line.artikelCode) !== safeText(proposal.artikelCode)) {
      return line;
    }

    found = true;

    return {
      artikelCode: proposal.artikelCode,
      artikelOmschrijving: proposal.artikelOmschrijving,
      eenheid: proposal.eenheid,
      aantal: safeNumber(proposal.voorstelAantal, 0)
    };
  });

  if (!found) {
    mergedLines.push({
      artikelCode: proposal.artikelCode,
      artikelOmschrijving: proposal.artikelOmschrijving,
      eenheid: proposal.eenheid,
      aantal: safeNumber(proposal.voorstelAantal, 0)
    });
  }

  return mergedLines;
}

function createNeedIssueFromProposal(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertWarehouseOrMobileAccess(sessionId);

  const proposal = payload.proposal || {};
  const techniekerCode = safeText(proposal.techniekerCode);
  const deliveryDateIso = safeText(proposal.deliveryDateIso);
  const voorstelAantal = safeNumber(proposal.voorstelAantal, 0);

  if (!techniekerCode) throw new Error('Technieker ontbreekt op voorstel.');
  if (!deliveryDateIso) throw new Error('Beleveringsdatum ontbreekt op voorstel.');
  if (!safeText(proposal.artikelCode)) throw new Error('Artikelcode ontbreekt op voorstel.');
  if (voorstelAantal <= 0) throw new Error('Voorstelaantal moet groter zijn dan 0.');

  let issue = findOpenAutoNeedIssueForMoment_(techniekerCode, deliveryDateIso);

  if (!issue) {
    const createResult = createNeedIssue({
      sessionId: sessionId,
      techniekerCode: techniekerCode,
      documentDatum: deliveryDateIso,
      reden: 'Auto aanvulling verbruik',
      opmerking: `Automatisch voorstel voor beleveringsmoment ${safeText(proposal.deliveryLabel || '')}`
    });

    issue = getNeedIssueById(createResult.uitgifteId);
  }

  const mergedLines = upsertProposalLineInNeedIssue_(issue.uitgifteId, proposal);

  saveNeedIssueLines({
    sessionId: sessionId,
    uitgifteId: issue.uitgifteId,
    lines: mergedLines
  });

  writeAudit(
    'Automatisch behoeftevoorstel omgezet',
    user.rol,
    user.naam || user.email || 'Gebruiker',
    'BehoefteUitgifte',
    issue.uitgifteId,
    {
      techniekerCode: techniekerCode,
      artikelCode: safeText(proposal.artikelCode),
      voorstelAantal: voorstelAantal,
      deliveryDateIso: deliveryDateIso
    }
  );

  return {
    success: true,
    uitgifteId: issue.uitgifteId,
    message: 'Automatisch behoeftevoorstel toegevoegd aan behoefte-uitgifte.'
  };
}