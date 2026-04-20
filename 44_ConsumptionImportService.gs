/* =========================================================
   44_ConsumptionImportService.gs — verbruiksimport
   Google Drive bestand -> VerbruikImportRuns / VerbruikImportRaw
   ========================================================= */

function makeConsumptionImportRunId() {
  return makeStampedId('VRUN');
}

function makeConsumptionImportLogId() {
  return makeStampedId('VLOG');
}

/* ---------------------------------------------------------
   Config
   --------------------------------------------------------- */

function getConsumptionImportConfigMap() {
  const rows = readObjectsSafe(TABS.CONSUMPTION_IMPORT_CONFIG);
  const map = {};

  rows.forEach(row => {
    const key = safeText(row.Sleutel || row.Key);
    if (!key) return;
    map[key] = safeText(row.Waarde || row.Value);
  });

  return map;
}

function getConsumptionImportConfigValue(key, fallbackValue) {
  const map = getConsumptionImportConfigMap();
  return map[key] !== undefined && map[key] !== '' ? map[key] : fallbackValue;
}

/* ---------------------------------------------------------
   Logging
   --------------------------------------------------------- */

function writeConsumptionImportLog(runId, level, message, details) {
  appendObjects(TABS.CONSUMPTION_IMPORT_LOG, [{
    LogID: makeConsumptionImportLogId(),
    RunID: safeText(runId),
    Tijdstip: nowStamp(),
    Niveau: safeText(level),
    Bericht: safeText(message),
    Details: typeof details === 'object' ? JSON.stringify(details) : safeText(details)
  }]);
}

/* ---------------------------------------------------------
   Run helpers
   --------------------------------------------------------- */

function startConsumptionImportRun(sourceFile) {
  const runId = makeConsumptionImportRunId();

  appendObjects(TABS.CONSUMPTION_IMPORT_RUNS, [{
    RunID: runId,
    BronBestand: sourceFile ? safeText(sourceFile.getName()) : '',
    BronVersie: sourceFile ? safeText(sourceFile.getLastUpdated()) : '',
    ImportStart: nowStamp(),
    ImportEinde: '',
    Status: IMPORT_RUN_STATUS.STARTED,
    AantalRuweLijnen: 0,
    AantalGeldigeLijnen: 0,
    AantalFoutLijnen: 0,
    AantalDubbeleLijnen: 0,
    Opmerking: ''
  }]);

  return runId;
}

function finishConsumptionImportRun(runId, updates) {
  const id = safeText(runId);
  const sheet = getSheetOrThrow(TABS.CONSUMPTION_IMPORT_RUNS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab VerbruikImportRuns is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['RunID']]) !== id) continue;

    Object.keys(updates || {}).forEach(field => {
      if (col[field] !== undefined) {
        values[i][col[field]] = updates[field];
      }
    });

    if (col['ImportEinde'] !== undefined) {
      values[i][col['ImportEinde']] = nowStamp();
    }

    updated = true;
    break;
  }

  if (!updated) throw new Error('ImportRun niet gevonden.');

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return { success: true };
}

/* ---------------------------------------------------------
   Drive bronbestand
   --------------------------------------------------------- */

function getConsumptionSourceDriveFile() {
  const sourceFileId = safeText(getConsumptionImportConfigValue('ImportBronBestandId', ''));
  const sourceFileName = safeText(getConsumptionImportConfigValue('ImportBestandsNaam', ''));

  if (sourceFileId) {
    return DriveApp.getFileById(sourceFileId);
  }

  if (sourceFileName) {
    const files = DriveApp.getFilesByName(sourceFileName);
    if (files.hasNext()) return files.next();
  }

  throw new Error('Geen bronbestand gevonden. Vul ImportBronBestandId of ImportBestandsNaam in.');
}

function isGoogleSpreadsheetMimeType(mimeType) {
  return safeText(mimeType) === 'application/vnd.google-apps.spreadsheet';
}

function isExcelMimeType(mimeType) {
  const mime = safeText(mimeType);
  return mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      || mime === 'application/vnd.ms-excel';
}

function openConsumptionSourceSpreadsheetFromDrive(file) {
  if (!file) throw new Error('Bronbestand ontbreekt.');

  const mimeType = safeText(file.getMimeType());

  if (isGoogleSpreadsheetMimeType(mimeType)) {
    return {
      spreadsheet: SpreadsheetApp.openById(file.getId()),
      tempFileId: ''
    };
  }

  if (isExcelMimeType(mimeType)) {
    if (typeof Drive === 'undefined' || !Drive.Files || !Drive.Files.copy) {
      throw new Error('Advanced Drive Service is niet actief. Activeer Drive API advanced service voor XLSX-conversie.');
    }

    const tempName = `tmp_verbruik_import_${Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd_HHmmss')}`;

    const copied = Drive.Files.copy(
      {
        title: tempName,
        mimeType: 'application/vnd.google-apps.spreadsheet'
      },
      file.getId()
    );

    const tempFileId = safeText(copied.id || copied.fileId);
    if (!tempFileId) {
      throw new Error('Tijdelijke conversie naar Google Sheet mislukt.');
    }

    return {
      spreadsheet: SpreadsheetApp.openById(tempFileId),
      tempFileId: tempFileId
    };
  }

  throw new Error(`Niet ondersteund bronbestandstype: ${mimeType}`);
}

function cleanupTemporaryImportedSpreadsheet(tempFileId) {
  const id = safeText(tempFileId);
  if (!id) return;

  try {
    DriveApp.getFileById(id).setTrashed(true);
  } catch (e) {
    // niet blokkeren
  }
}

/* ---------------------------------------------------------
   Brondata lezen
   --------------------------------------------------------- */

function getConsumptionSourceSheet(spreadsheet) {
  const configuredName = safeText(getConsumptionImportConfigValue('ImportSheetNaam', ''));

  if (configuredName) {
    const namedSheet = spreadsheet.getSheetByName(configuredName);
    if (!namedSheet) {
      throw new Error(`Bronblad niet gevonden: ${configuredName}`);
    }
    return namedSheet;
  }

  const firstSheet = spreadsheet.getSheets()[0];
  if (!firstSheet) {
    throw new Error('Geen werkblad gevonden in bronbestand.');
  }

  return firstSheet;
}

function getSourceHeaderAliases() {
  return {
    documentDatum: safeText(getConsumptionImportConfigValue('KolomDocumentDatum', 'DocumentDatum')),
    techniekerCode: safeText(getConsumptionImportConfigValue('KolomTechniekerCode', 'TechniekerCode')),
    techniekerNaam: safeText(getConsumptionImportConfigValue('KolomTechniekerNaam', 'TechniekerNaam')),
    werfRef: safeText(getConsumptionImportConfigValue('KolomWerfRef', 'WerfRef')),
    artikelCode: safeText(getConsumptionImportConfigValue('KolomArtikelCode', 'ArtikelCode')),
    artikelOmschrijving: safeText(getConsumptionImportConfigValue('KolomArtikelOmschrijving', 'ArtikelOmschrijving')),
    aantal: safeText(getConsumptionImportConfigValue('KolomAantal', 'Aantal')),
    eenheid: safeText(getConsumptionImportConfigValue('KolomEenheid', 'Eenheid'))
  };
}

function buildHeaderIndexMap(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[safeText(header)] = index;
  });
  return map;
}

function getCellByHeader(row, headerIndexMap, headerName) {
  const idx = headerIndexMap[safeText(headerName)];
  if (idx === undefined) return '';
  return row[idx];
}

function normalizeConsumptionImportRow(rawRow, headerIndexMap, rowNumber) {
  const aliases = getSourceHeaderAliases();

  const obj = {
    bronRijNr: rowNumber,
    documentDatum: toIsoDate(getCellByHeader(rawRow, headerIndexMap, aliases.documentDatum)),
    techniekerCode: safeText(getCellByHeader(rawRow, headerIndexMap, aliases.techniekerCode)),
    techniekerNaam: safeText(getCellByHeader(rawRow, headerIndexMap, aliases.techniekerNaam)),
    werfRef: safeText(getCellByHeader(rawRow, headerIndexMap, aliases.werfRef)),
    artikelCode: safeText(getCellByHeader(rawRow, headerIndexMap, aliases.artikelCode)),
    artikelOmschrijving: safeText(getCellByHeader(rawRow, headerIndexMap, aliases.artikelOmschrijving)),
    aantal: safeNumber(getCellByHeader(rawRow, headerIndexMap, aliases.aantal), 0),
    eenheid: safeText(getCellByHeader(rawRow, headerIndexMap, aliases.eenheid))
  };

  obj.bronHash = buildConsumptionImportHash(obj);
  return obj;
}

function readConsumptionSourceRowsFromSpreadsheet(spreadsheet) {
  const sheet = getConsumptionSourceSheet(spreadsheet);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(h => safeText(h));
  const headerIndexMap = buildHeaderIndexMap(headers);

  return values.slice(1).map((row, idx) => normalizeConsumptionImportRow(row, headerIndexMap, idx + 2));
}

/* ---------------------------------------------------------
   Hash / duplicate
   --------------------------------------------------------- */

function bytesToHexString(bytes) {
  return (bytes || []).map(b => {
    const v = b < 0 ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function buildConsumptionImportHash(row) {
  const raw = [
    safeText(row.documentDatum),
    safeText(row.techniekerCode),
    safeText(row.werfRef),
    safeText(row.artikelCode),
    String(safeNumber(row.aantal, 0))
  ].join('|');

  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return bytesToHexString(digest);
}

function getExistingConsumptionImportHashes() {
  const rows = readObjectsSafe(TABS.CONSUMPTION_IMPORT_RAW);
  const map = {};

  rows.forEach(row => {
    const hash = safeText(row.BronHash);
    if (!hash) return;
    map[hash] = true;
  });

  return map;
}

/* ---------------------------------------------------------
   Raw object build
   --------------------------------------------------------- */

function determineRawImportValidationStatus(row, existingHashMap, seenThisRunMap) {
  if (!safeText(row.documentDatum) || !safeText(row.techniekerCode) || !safeText(row.artikelCode)) {
    return {
      status: IMPORT_VALIDATION_STATUS.ERROR,
      error: 'Verplichte bronvelden ontbreken'
    };
  }

  if (safeNumber(row.aantal, 0) <= 0) {
    return {
      status: IMPORT_VALIDATION_STATUS.ERROR,
      error: 'Aantal moet groter zijn dan 0'
    };
  }

  if (existingHashMap[row.bronHash] || seenThisRunMap[row.bronHash]) {
    return {
      status: IMPORT_VALIDATION_STATUS.DUPLICATE,
      error: 'BronHash bestaat al'
    };
  }

  return {
    status: IMPORT_VALIDATION_STATUS.NEW,
    error: ''
  };
}

function buildConsumptionImportRawObject(runId, normalizedRow, validationResult) {
  return {
    RunID: safeText(runId),
    BronRijNr: safeNumber(normalizedRow.bronRijNr, 0),
    DocumentDatum: safeText(normalizedRow.documentDatum),
    TechniekerCode: safeText(normalizedRow.techniekerCode),
    TechniekerNaam: safeText(normalizedRow.techniekerNaam),
    WerfRef: safeText(normalizedRow.werfRef),
    ArtikelCode: safeText(normalizedRow.artikelCode),
    ArtikelOmschrijving: safeText(normalizedRow.artikelOmschrijving),
    Aantal: safeNumber(normalizedRow.aantal, 0),
    Eenheid: safeText(normalizedRow.eenheid),
    BronHash: safeText(normalizedRow.bronHash),
    ValidatieStatus: safeText(validationResult.status),
    ValidatieFout: safeText(validationResult.error),
    BoekStatus: IMPORT_BOOK_STATUS.NOT_BOOKED,
    BoekFout: '',
    RawJson: JSON.stringify(normalizedRow)
  };
}

function appendConsumptionImportRawObjects(rawObjects) {
  if (!(rawObjects || []).length) {
    return { success: true, lines: 0 };
  }

  appendObjects(TABS.CONSUMPTION_IMPORT_RAW, rawObjects);
  return { success: true, lines: rawObjects.length };
}

/* ---------------------------------------------------------
   Main import
   --------------------------------------------------------- */

function importConsumptionFromDrive() {
  const sourceFile = getConsumptionSourceDriveFile();
  const runId = startConsumptionImportRun(sourceFile);

  let tempFileId = '';
  let rawObjects = [];

  try {
    writeConsumptionImportLog(runId, 'INFO', 'Import gestart', {
      bronBestand: sourceFile.getName(),
      bronMimeType: sourceFile.getMimeType()
    });

    const source = openConsumptionSourceSpreadsheetFromDrive(sourceFile);
    tempFileId = safeText(source.tempFileId);

    const normalizedRows = readConsumptionSourceRowsFromSpreadsheet(source.spreadsheet);
    const existingHashMap = getExistingConsumptionImportHashes();
    const seenThisRunMap = {};

    rawObjects = normalizedRows.map(row => {
      const validationResult = determineRawImportValidationStatus(row, existingHashMap, seenThisRunMap);

      if (validationResult.status === IMPORT_VALIDATION_STATUS.NEW) {
        seenThisRunMap[row.bronHash] = true;
      }

      return buildConsumptionImportRawObject(runId, row, validationResult);
    });

    appendConsumptionImportRawObjects(rawObjects);

    const duplicateCount = rawObjects.filter(x => safeText(x.ValidatieStatus) === IMPORT_VALIDATION_STATUS.DUPLICATE).length;
    const errorCount = rawObjects.filter(x => safeText(x.ValidatieStatus) === IMPORT_VALIDATION_STATUS.ERROR).length;
    const newCount = rawObjects.filter(x => safeText(x.ValidatieStatus) === IMPORT_VALIDATION_STATUS.NEW).length;

    finishConsumptionImportRun(runId, {
      Status: IMPORT_RUN_STATUS.IMPORTED,
      AantalRuweLijnen: rawObjects.length,
      AantalGeldigeLijnen: newCount,
      AantalFoutLijnen: errorCount,
      AantalDubbeleLijnen: duplicateCount,
      Opmerking: ''
    });

    writeConsumptionImportLog(runId, 'INFO', 'Import afgerond', {
      totaal: rawObjects.length,
      nieuw: newCount,
      fout: errorCount,
      dubbel: duplicateCount
    });

    writeAudit(
      'Verbruiksimport uitgevoerd',
      'Systeem',
      'Automatisch',
      'VerbruikImportRun',
      runId,
      {
        bronBestand: sourceFile.getName(),
        totaal: rawObjects.length,
        nieuw: newCount,
        fout: errorCount,
        dubbel: duplicateCount
      }
    );

    return {
      success: true,
      runId: runId,
      sourceFileName: sourceFile.getName(),
      totalRows: rawObjects.length,
      newRows: newCount,
      duplicateRows: duplicateCount,
      errorRows: errorCount,
      message: 'Verbruiksimport voltooid.'
    };

  } catch (err) {
    finishConsumptionImportRun(runId, {
      Status: IMPORT_RUN_STATUS.FAILED,
      AantalRuweLijnen: rawObjects.length,
      AantalGeldigeLijnen: rawObjects.filter(x => safeText(x.ValidatieStatus) === IMPORT_VALIDATION_STATUS.NEW).length,
      AantalFoutLijnen: rawObjects.filter(x => safeText(x.ValidatieStatus) === IMPORT_VALIDATION_STATUS.ERROR).length,
      AantalDubbeleLijnen: rawObjects.filter(x => safeText(x.ValidatieStatus) === IMPORT_VALIDATION_STATUS.DUPLICATE).length,
      Opmerking: safeText(err && err.message ? err.message : err)
    });

    writeConsumptionImportLog(runId, 'ERROR', 'Import mislukt', {
      error: safeText(err && err.message ? err.message : err)
    });

    throw err;

  } finally {
    cleanupTemporaryImportedSpreadsheet(tempFileId);
  }
}

/* ---------------------------------------------------------
   Scheduler wrappers
   --------------------------------------------------------- */

function runNightlyConsumptionImport() {
  return importConsumptionFromDrive();
}

function testConsumptionImport() {
  return importConsumptionFromDrive();
}