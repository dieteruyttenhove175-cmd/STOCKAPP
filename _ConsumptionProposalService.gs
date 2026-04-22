/* =========================================================
   45A_ConsumptionProposalService.gs — voorstellen behoefte
   Snapshots + delivery moments -> automatische behoeftevoorstellen
   Werkt met:
   1) echte snapshot-tab(s), als die later toegevoegd worden
   2) fallback op VerbruikImportRaw als snapshotbron
   ========================================================= */

/* ---------------------------------------------------------
   Snapshot source helpers
   --------------------------------------------------------- */

function getOptionalSheetByName_(sheetName) {
  const name = safeText(sheetName);
  if (!name) return null;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    return ss.getSheetByName(name);
  } catch (e) {
    return null;
  }
}

function readObjectsFromOptionalSheet_(sheetName) {
  const sheet = getOptionalSheetByName_(sheetName);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (!values || !values.length) return [];

  const headers = values[0].map(h => safeText(h));
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

function getConsumptionSnapshotSheetNames_() {
  const names = [];

  // toekomstige snapshot-tabben, als je die later toevoegt
  if (typeof TABS !== 'undefined') {
    if (safeText(TABS.CONSUMPTION_SNAPSHOTS)) names.push(TABS.CONSUMPTION_SNAPSHOTS);
    if (safeText(TABS.CONSUMPTION_SNAPSHOT_RUNS)) names.push(TABS.CONSUMPTION_SNAPSHOT_RUNS);
  }

  // harde fallback namen als er later manueel tabben bijkomen
  names.push('VerbruikSnapshots');

  return [...new Set(names.filter(Boolean))];
}

function mapSnapshotRowFromSnapshotTab_(row) {
  return {
    snapshotId: safeText(row.SnapshotID || row.snapshotId || ''),
    runId: safeText(row.RunID || row.runId || ''),
    snapshotDatum: toIsoDate(row.SnapshotDatum || row.DocumentDatum || row.snapshotDatum || ''),
    techniekerCode: safeText(row.TechniekerCode || row.techniekerCode || ''),
    techniekerNaam: safeText(row.TechniekerNaam || row.techniekerNaam || ''),
    artikelCode: safeText(row.ArtikelCode || row.artikelCode || ''),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.artikelOmschrijving || ''),
    eenheid: safeText(row.Eenheid || row.eenheid || ''),
    cumulatiefVerbruik: safeNumber(
      row.CumulatiefVerbruik || row.cumulatiefVerbruik || row.Aantal || row.aantal || 0,
      0
    ),
    bronType: safeText(row.BronType || row.bronType || 'Snapshot'),
    bronBestand: safeText(row.BronBestand || row.bronBestand || ''),
    aangemaaktOp: safeText(row.AangemaaktOp || row.aangemaaktOp || ''),
    opmerking: safeText(row.Opmerking || row.opmerking || '')
  };
}

function getSnapshotsFromExplicitSnapshotTabs_() {
  const rows = [];

  getConsumptionSnapshotSheetNames_().forEach(sheetName => {
    readObjectsFromOptionalSheet_(sheetName).forEach(row => {
      const mapped = mapSnapshotRowFromSnapshotTab_(row);

      if (!safeText(mapped.snapshotDatum)) return;
      if (!safeText(mapped.techniekerCode)) return;
      if (!safeText(mapped.artikelCode)) return;

      rows.push(mapped);
    });
  });

  return rows;
}

function mapSnapshotFromRawImportRow_(row) {
  return {
    snapshotId: '',
    runId: safeText(row.RunID || ''),
    snapshotDatum: toIsoDate(row.DocumentDatum || row.documentDatum || ''),
    techniekerCode: safeText(row.TechniekerCode || row.techniekerCode || ''),
    techniekerNaam: safeText(row.TechniekerNaam || row.techniekerNaam || ''),
    artikelCode: safeText(row.ArtikelCode || row.artikelCode || ''),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.artikelOmschrijving || ''),
    eenheid: safeText(row.Eenheid || row.eenheid || ''),
    cumulatiefVerbruik: safeNumber(row.Aantal || row.aantal || 0, 0),
    bronType: 'ImportRaw',
    bronBestand: '',
    aangemaaktOp: safeText(row.VerwerktOp || row.ImportStart || ''),
    opmerking: ''
  };
}

function aggregateRawImportRowsToSnapshots_(rows) {
  const grouped = {};

  (rows || []).forEach(row => {
    const key = [
      safeText(row.snapshotDatum),
      normalizeRef(row.techniekerCode),
      safeText(row.artikelCode)
    ].join('|');

    if (!grouped[key]) {
      grouped[key] = {
        snapshotId: '',
        runId: safeText(row.runId),
        snapshotDatum: safeText(row.snapshotDatum),
        techniekerCode: safeText(row.techniekerCode),
        techniekerNaam: safeText(row.techniekerNaam),
        artikelCode: safeText(row.artikelCode),
        artikelOmschrijving: safeText(row.artikelOmschrijving),
        eenheid: safeText(row.eenheid),
        cumulatiefVerbruik: safeNumber(row.cumulatiefVerbruik, 0),
        bronType: 'ImportRaw',
        bronBestand: '',
        aangemaaktOp: safeText(row.aangemaaktOp),
        opmerking: ''
      };
      return;
    }

    // Neem hoogste waarde als snapshot voor die dag/artikel/technieker
    grouped[key].cumulatiefVerbruik = Math.max(
      safeNumber(grouped[key].cumulatiefVerbruik, 0),
      safeNumber(row.cumulatiefVerbruik, 0)
    );

    if (!grouped[key].techniekerNaam && row.techniekerNaam) {
      grouped[key].techniekerNaam = safeText(row.techniekerNaam);
    }

    if (!grouped[key].artikelOmschrijving && row.artikelOmschrijving) {
      grouped[key].artikelOmschrijving = safeText(row.artikelOmschrijving);
    }

    if (!grouped[key].eenheid && row.eenheid) {
      grouped[key].eenheid = safeText(row.eenheid);
    }

    if (safeText(row.aangemaaktOp) > safeText(grouped[key].aangemaaktOp)) {
      grouped[key].aangemaaktOp = safeText(row.aangemaaktOp);
    }
  });

  return Object.keys(grouped).map(key => grouped[key]);
}

function getSnapshotsFromImportRawFallback_() {
  if (typeof readObjectsSafe !== 'function') return [];
  if (typeof TABS === 'undefined' || !safeText(TABS.CONSUMPTION_IMPORT_RAW)) return [];

  const rawRows = readObjectsSafe(TABS.CONSUMPTION_IMPORT_RAW)
    .map(mapSnapshotFromRawImportRow_)
    .filter(row =>
      safeText(row.snapshotDatum) &&
      safeText(row.techniekerCode) &&
      safeText(row.artikelCode) &&
      safeNumber(row.cumulatiefVerbruik, 0) >= 0
    );

  return aggregateRawImportRowsToSnapshots_(rawRows);
}

function getAllConsumptionSnapshots() {
  const snapshotRows = getSnapshotsFromExplicitSnapshotTabs_();
  if (snapshotRows.length) {
    return snapshotRows.sort((a, b) =>
      `${safeText(a.snapshotDatum)} ${safeText(a.techniekerCode)} ${safeText(a.artikelCode)} ${safeText(a.aangemaaktOp)}`.localeCompare(
        `${safeText(b.snapshotDatum)} ${safeText(b.techniekerCode)} ${safeText(b.artikelCode)} ${safeText(b.aangemaaktOp)}`
      )
    );
  }

  const fallbackRows = getSnapshotsFromImportRawFallback_();
  return fallbackRows.sort((a, b) =>
    `${safeText(a.snapshotDatum)} ${safeText(a.techniekerCode)} ${safeText(a.artikelCode)} ${safeText(a.aangemaaktOp)}`.localeCompare(
      `${safeText(b.snapshotDatum)} ${safeText(b.techniekerCode)} ${safeText(b.artikelCode)} ${safeText(b.aangemaaktOp)}`
    )
  );
}

function getSnapshotsForTechnicianArticle_(techniekerCode, artikelCode) {
  return getAllConsumptionSnapshots()
    .filter(row =>
      normalizeRef(row.techniekerCode) === normalizeRef(techniekerCode) &&
      safeText(row.artikelCode) === safeText(artikelCode)
    )
    .sort((a, b) =>
      `${safeText(a.snapshotDatum)} ${safeText(a.aangemaaktOp)}`.localeCompare(
        `${safeText(b.snapshotDatum)} ${safeText(b.aangemaaktOp)}`
      )
    );
}

function getLatestConsumptionSnapshotOnOrBefore_(techniekerCode, artikelCode, cutoffIsoDate) {
  const cutoff = safeText(cutoffIsoDate);
  if (!cutoff) return null;

  const rows = getSnapshotsForTechnicianArticle_(techniekerCode, artikelCode)
    .filter(row => safeText(row.snapshotDatum) <= cutoff);

  return rows.length ? rows[rows.length - 1] : null;
}

/* ---------------------------------------------------------
   Delivery selection
   --------------------------------------------------------- */

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

  return moments.sort((a, b) =>
    safeText(a.deliverySortKey).localeCompare(safeText(b.deliverySortKey))
  );
}

/* ---------------------------------------------------------
   Proposal calculations
   --------------------------------------------------------- */

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
  // later uitbreidbaar
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
        proposalKey: [moment.key, safeText(artikelCode)].join('|'),
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

/* ---------------------------------------------------------
   Need issue creation from proposal
   --------------------------------------------------------- */

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