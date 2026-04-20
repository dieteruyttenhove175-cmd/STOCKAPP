/* =========================================================
   45_ConsumptionProcessingService.gs — verwerking verbruiksimport
   Raw import -> Verbruiken / VerbruikLijnen (nog niet geboekt)
   ========================================================= */

/* ---------------------------------------------------------
   Raw mapping
   --------------------------------------------------------- */

function mapConsumptionImportRaw(row) {
  return {
    runId: safeText(row.RunID),
    bronRijNr: safeNumber(row.BronRijNr, 0),
    documentDatum: toIsoDate(row.DocumentDatum),
    techniekerCode: safeText(row.TechniekerCode),
    techniekerNaam: safeText(row.TechniekerNaam),
    werfRef: safeText(row.WerfRef),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    aantal: safeNumber(row.Aantal, 0),
    eenheid: safeText(row.Eenheid),
    bronHash: safeText(row.BronHash),
    validatieStatus: safeText(row.ValidatieStatus),
    validatieFout: safeText(row.ValidatieFout),
    boekStatus: safeText(row.BoekStatus),
    boekFout: safeText(row.BoekFout),
    verbruikId: safeText(row.VerbruikID),
    verwerktOp: safeText(row.VerwerktOp),
    rawJson: safeText(row.RawJson)
  };
}

function readConsumptionImportRawRowsWithMeta() {
  const sheet = getSheetOrThrow(TABS.CONSUMPTION_IMPORT_RAW);
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const headers = values[0].map(h => safeText(h));

  return values.slice(1).map((row, idx) => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });

    return {
      ...mapConsumptionImportRaw(obj),
      _sheetRowIndex: idx + 2
    };
  });
}

function assertConsumptionImportRawProcessingColumns() {
  const headers = getHeaders(TABS.CONSUMPTION_IMPORT_RAW);

  ['VerbruikID', 'VerwerktOp'].forEach(requiredHeader => {
    if (!headers.includes(requiredHeader)) {
      throw new Error(`Tab ${TABS.CONSUMPTION_IMPORT_RAW} mist verplichte kolom ${requiredHeader}.`);
    }
  });
}

/* ---------------------------------------------------------
   Helpers — bronselectie
   --------------------------------------------------------- */

function getConsumptionImportRunRows(runId) {
  const id = safeText(runId);
  return readConsumptionImportRawRowsWithMeta()
    .filter(row => !id || row.runId === id);
}

function getPendingConsumptionImportRows(runId) {
  return getConsumptionImportRunRows(runId).filter(row => {
    if (row.verbruikId) return false;
    if (row.validatieStatus === IMPORT_VALIDATION_STATUS.DUPLICATE) return false;
    if (row.validatieStatus === IMPORT_VALIDATION_STATUS.ERROR) return false;
    if (row.validatieStatus === IMPORT_VALIDATION_STATUS.SKIPPED) return false;
    return true;
  });
}

function getLatestConsumptionImportRunId() {
  const runs = readObjectsSafe(TABS.CONSUMPTION_IMPORT_RUNS)
    .map(row => ({
      runId: safeText(row.RunID),
      importStart: safeText(row.ImportStart)
    }))
    .filter(x => x.runId)
    .sort((a, b) => safeText(b.importStart).localeCompare(safeText(a.importStart)));

  return runs.length ? runs[0].runId : '';
}

/* ---------------------------------------------------------
   Helpers — masterdata
   --------------------------------------------------------- */

function buildConsumptionTechnicianMap() {
  const map = {};
  getActiveTechnicians().forEach(item => {
    map[normalizeRef(item.code)] = item;
  });
  return map;
}

function buildConsumptionArticleMap() {
  const map = {};
  readObjectsSafe(TABS.SUPPLIER_ARTICLES)
    .map(mapSupplierArticle)
    .filter(item => item.actief !== false)
    .forEach(item => {
      map[safeText(item.artikelCode)] = item;
    });
  return map;
}

function isArticleAllowedForConsumptionScope(article) {
  if (!article) return false;

  const allowedScopesRaw = safeText(getConsumptionImportConfigValue('ToegelatenCategorieen', ''));
  if (!allowedScopesRaw) return true;

  const allowedScopes = allowedScopesRaw
    .split(',')
    .map(x => safeText(x).toUpperCase())
    .filter(Boolean);

  if (!allowedScopes.length) return true;

  const articleScope = safeText(
    article.CategorieScope ||
    article.ProjectScope ||
    article.Categorie ||
    article.CAT ||
    article.Cat ||
    article.Scope
  ).toUpperCase();

  if (!articleScope) return true;
  return allowedScopes.includes(articleScope);
}

/* ---------------------------------------------------------
   Business-validatie
   --------------------------------------------------------- */

function validateConsumptionRawBusinessRow(rawRow, technicianMap, articleMap) {
  const errors = [];

  const tech = technicianMap[normalizeRef(rawRow.techniekerCode)] || null;
  const article = articleMap[safeText(rawRow.artikelCode)] || null;

  if (!rawRow.documentDatum) {
    errors.push('Documentdatum ontbreekt of is ongeldig');
  }

  if (!rawRow.techniekerCode) {
    errors.push('TechniekerCode ontbreekt');
  } else if (!tech) {
    errors.push('TechniekerCode niet gevonden of niet actief');
  }

  if (!rawRow.artikelCode) {
    errors.push('ArtikelCode ontbreekt');
  } else if (!article) {
    errors.push('ArtikelCode niet gevonden');
  }

  if (safeNumber(rawRow.aantal, 0) <= 0) {
    errors.push('Aantal moet groter zijn dan 0');
  }

  if (article && !isArticleAllowedForConsumptionScope(article)) {
    errors.push('Artikel valt buiten toegelaten scope');
  }

  return {
    isValid: !errors.length,
    errors,
    technician: tech,
    article: article
  };
}

/* ---------------------------------------------------------
   Raw update
   --------------------------------------------------------- */

function updateConsumptionRawProcessingResults(rawResults) {
  if (!(rawResults || []).length) return { success: true, updated: 0 };

  const sheet = getSheetOrThrow(TABS.CONSUMPTION_IMPORT_RAW);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab VerbruikImportRaw is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = 0;

  rawResults.forEach(result => {
    const rowIndex = safeNumber(result._sheetRowIndex, 0);
    if (!rowIndex || rowIndex < 2 || rowIndex > values.length) return;

    const arrIndex = rowIndex - 1;

    if (col['ValidatieStatus'] !== undefined) values[arrIndex][col['ValidatieStatus']] = safeText(result.validatieStatus);
    if (col['ValidatieFout'] !== undefined) values[arrIndex][col['ValidatieFout']] = safeText(result.validatieFout);
    if (col['VerbruikID'] !== undefined) values[arrIndex][col['VerbruikID']] = safeText(result.verbruikId);
    if (col['VerwerktOp'] !== undefined) values[arrIndex][col['VerwerktOp']] = safeText(result.verwerktOp);

    updated++;
  });

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return { success: true, updated };
}

/* ---------------------------------------------------------
   Groepering
   --------------------------------------------------------- */

function buildConsumptionProcessingGroupKey(rawRow) {
  return [
    safeText(rawRow.documentDatum),
    safeText(rawRow.techniekerCode),
    safeText(rawRow.werfRef)
  ].join('|');
}

function makeImportedConsumptionId(runId, groupIndex) {
  return `VIMP-${safeText(runId)}-${String(groupIndex).padStart(3, '0')}`;
}

function groupValidatedConsumptionRows(validatedRows) {
  const groups = {};

  (validatedRows || []).forEach(row => {
    const key = buildConsumptionProcessingGroupKey(row);

    if (!groups[key]) {
      groups[key] = {
        key,
        documentDatum: row.documentDatum,
        techniekerCode: row.techniekerCode,
        techniekerNaam: row._matchedTechnician ? row._matchedTechnician.naam : row.techniekerNaam,
        werfRef: row.werfRef,
        rows: []
      };
    }

    groups[key].rows.push(row);
  });

  return Object.keys(groups)
    .map(key => groups[key])
    .sort((a, b) =>
      `${safeText(a.documentDatum)} ${safeText(a.techniekerCode)} ${safeText(a.werfRef)}`.localeCompare(
        `${safeText(b.documentDatum)} ${safeText(b.techniekerCode)} ${safeText(b.werfRef)}`
      )
    );
}

function aggregateGroupLines(group) {
  const map = {};

  (group.rows || []).forEach(row => {
    const code = safeText(row.artikelCode);
    if (!code) return;

    if (!map[code]) {
      map[code] = {
        artikelCode: code,
        artikelOmschrijving: safeText(row.artikelOmschrijving),
        eenheid: safeText(row.eenheid),
        aantal: 0,
        bronHashes: [],
        bronRijNrs: []
      };
    }

    map[code].aantal += safeNumber(row.aantal, 0);
    map[code].bronHashes.push(safeText(row.bronHash));
    map[code].bronRijNrs.push(safeNumber(row.bronRijNr, 0));
  });

  return Object.keys(map)
    .map(key => map[key])
    .sort((a, b) => safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)));
}

/* ---------------------------------------------------------
   Verbruiken / lijnen opbouwen
   --------------------------------------------------------- */

function buildConsumptionDocumentObjectFromGroup(runId, verbruikId, group) {
  return {
    VerbruikID: verbruikId,
    TechniekerCode: safeText(group.techniekerCode),
    TechniekerNaam: safeText(group.techniekerNaam),
    DocumentDatum: safeText(group.documentDatum),
    WerfRef: safeText(group.werfRef),
    Status: CONSUMPTION_STATUS.OPEN,
    Reden: 'Import',
    Opmerking: `Automatisch opgebouwd uit import run ${safeText(runId)}`,
    GeboektDoor: '',
    GeboektOp: '',
    BronType: 'Import',
    BronRunID: safeText(runId)
  };
}

function buildConsumptionLineObjectsFromGroup(runId, verbruikId, group) {
  const aggregatedLines = aggregateGroupLines(group);

  return aggregatedLines.map(line => ({
    VerbruikID: verbruikId,
    ArtikelCode: safeText(line.artikelCode),
    ArtikelOmschrijving: safeText(line.artikelOmschrijving),
    Eenheid: safeText(line.eenheid),
    Aantal: safeNumber(line.aantal, 0),
    Actief: 'Ja',
    BronRunID: safeText(runId),
    BronHash: line.bronHashes.join(', '),
    BronRijNr: line.bronRijNrs.join(', ')
  }));
}

function replaceImportedConsumptionsForRun(runId, documentObjects, lineObjects) {
  const runPrefix = `VIMP-${safeText(runId)}-`;

  /* Verbruiken */
  const consumptionSheet = getSheetOrThrow(TABS.CONSUMPTIONS);
  const consumptionValues = consumptionSheet.getDataRange().getValues();
  const consumptionHeaders = consumptionValues.length
    ? consumptionValues[0].map(h => safeText(h))
    : getHeaders(TABS.CONSUMPTIONS);
  const cc = getColMap(consumptionHeaders);
  const existingConsumptionRows = consumptionValues.length > 1 ? consumptionValues.slice(1) : [];

  const keptConsumptionRows = existingConsumptionRows.filter(row => {
    const verbruikId = safeText(row[cc['VerbruikID']]);
    return !verbruikId.startsWith(runPrefix);
  });

  const newConsumptionRows = (documentObjects || []).map(obj => buildRowFromHeaders(consumptionHeaders, obj));
  writeFullTable(TABS.CONSUMPTIONS, consumptionHeaders, keptConsumptionRows.concat(newConsumptionRows));

  /* VerbruikLijnen */
  const lineSheet = getSheetOrThrow(TABS.CONSUMPTION_LINES);
  const lineValues = lineSheet.getDataRange().getValues();
  const lineHeaders = lineValues.length
    ? lineValues[0].map(h => safeText(h))
    : getHeaders(TABS.CONSUMPTION_LINES);
  const lc = getColMap(lineHeaders);
  const existingLineRows = lineValues.length > 1 ? lineValues.slice(1) : [];

  const keptLineRows = existingLineRows.filter(row => {
    const verbruikId = safeText(row[lc['VerbruikID']]);
    return !verbruikId.startsWith(runPrefix);
  });

  const newLineRows = (lineObjects || []).map(obj => buildRowFromHeaders(lineHeaders, obj));
  writeFullTable(TABS.CONSUMPTION_LINES, lineHeaders, keptLineRows.concat(newLineRows));

  return {
    success: true,
    documents: newConsumptionRows.length,
    lines: newLineRows.length
  };
}

/* ---------------------------------------------------------
   Main processing
   --------------------------------------------------------- */

function processConsumptionImportRun(runId) {
  const id = safeText(runId);
  if (!id) throw new Error('RunID ontbreekt.');

  assertConsumptionImportRawProcessingColumns();

  const rawRows = getPendingConsumptionImportRows(id);
  if (!rawRows.length) {
    return {
      success: true,
      runId: id,
      processedRows: 0,
      validRows: 0,
      errorRows: 0,
      documentsCreated: 0,
      linesCreated: 0,
      message: 'Geen nieuwe raw lijnen om te verwerken.'
    };
  }

  const technicianMap = buildConsumptionTechnicianMap();
  const articleMap = buildConsumptionArticleMap();

  const rawUpdateResults = [];
  const validatedRows = [];

  rawRows.forEach(rawRow => {
    const validation = validateConsumptionRawBusinessRow(rawRow, technicianMap, articleMap);

    if (!validation.isValid) {
      rawUpdateResults.push({
        _sheetRowIndex: rawRow._sheetRowIndex,
        validatieStatus: IMPORT_VALIDATION_STATUS.ERROR,
        validatieFout: validation.errors.join(' | '),
        verbruikId: '',
        verwerktOp: nowStamp()
      });
      return;
    }

    const enriched = {
      ...rawRow,
      _matchedTechnician: validation.technician,
      _matchedArticle: validation.article
    };

    validatedRows.push(enriched);
  });

  const groups = groupValidatedConsumptionRows(validatedRows);

  const documentObjects = [];
  const lineObjects = [];

  groups.forEach((group, idx) => {
    const verbruikId = makeImportedConsumptionId(id, idx + 1);

    documentObjects.push(buildConsumptionDocumentObjectFromGroup(id, verbruikId, group));
    lineObjects.push(...buildConsumptionLineObjectsFromGroup(id, verbruikId, group));

    (group.rows || []).forEach(rawRow => {
      rawUpdateResults.push({
        _sheetRowIndex: rawRow._sheetRowIndex,
        validatieStatus: IMPORT_VALIDATION_STATUS.VALIDATED,
        validatieFout: '',
        verbruikId: verbruikId,
        verwerktOp: nowStamp()
      });
    });
  });

  replaceImportedConsumptionsForRun(id, documentObjects, lineObjects);
  updateConsumptionRawProcessingResults(rawUpdateResults);

  const validCount = rawUpdateResults.filter(x => safeText(x.validatieStatus) === IMPORT_VALIDATION_STATUS.VALIDATED).length;
  const errorCount = rawUpdateResults.filter(x => safeText(x.validatieStatus) === IMPORT_VALIDATION_STATUS.ERROR).length;

  writeConsumptionImportLog(id, 'INFO', 'Verwerking afgerond', {
    processedRows: rawRows.length,
    validRows: validCount,
    errorRows: errorCount,
    documentsCreated: documentObjects.length,
    linesCreated: lineObjects.length
  });

  writeAudit(
    'Verbruiksimport verwerkt',
    'Systeem',
    'Automatisch',
    'VerbruikImportRun',
    id,
    {
      processedRows: rawRows.length,
      validRows: validCount,
      errorRows: errorCount,
      documentsCreated: documentObjects.length,
      linesCreated: lineObjects.length
    }
  );

  return {
    success: true,
    runId: id,
    processedRows: rawRows.length,
    validRows: validCount,
    errorRows: errorCount,
    documentsCreated: documentObjects.length,
    linesCreated: lineObjects.length,
    message: 'Verbruiksimport verwerkt.'
  };
}

function processLatestConsumptionImportRun() {
  const latestRunId = getLatestConsumptionImportRunId();
  if (!latestRunId) {
    return {
      success: true,
      processedRows: 0,
      message: 'Geen import run gevonden.'
    };
  }

  return processConsumptionImportRun(latestRunId);
}

function testProcessLatestConsumptionImportRun() {
  return processLatestConsumptionImportRun();
}