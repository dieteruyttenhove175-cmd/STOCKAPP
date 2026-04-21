/* =========================================================
   40A_ReceiptImportService.gs — upload/import ontvangstlijsten
   Browser parsed rows -> Ontvangsten / OntvangstLijnen
   ========================================================= */

function parseReceiptDateFromFilename_(fileName) {
  const text = safeText(fileName);
  if (!text) return '';

  const match = text.match(/(\d{2})[-_](\d{2})[-_](\d{4})/);
  if (!match) return '';

  return `${match[3]}-${match[2]}-${match[1]}`;
}

function normalizeImportedReceiptRow_(row) {
  const artikelCode = safeText(
    row.artikelCode ||
    row.ArtikelCode ||
    row.Artikel ||
    row.artikel ||
    ''
  );

  const artikelOmschrijving = safeText(
    row.artikelOmschrijving ||
    row.ArtikelOmschrijving ||
    row.Omschrijving ||
    row.omschrijving ||
    ''
  );

  const eenheid = safeText(
    row.eenheid ||
    row.Eenheid ||
    ''
  );

  const aantal = safeNumber(
    row.aantal ||
    row.Aantal ||
    0,
    0
  );

  const palletNr = safeText(
    row.palletNr ||
    row.PalletNR ||
    row.PalletNr ||
    ''
  );

  const labelNr = safeText(
    row.labelNr ||
    row.LabelNR ||
    row.LabelNr ||
    ''
  );

  const notitie = safeText(
    row.notitie ||
    row.Notitie ||
    ''
  );

  return {
    artikelCode,
    artikelOmschrijving,
    eenheid,
    aantal,
    palletNr,
    labelNr,
    notitie
  };
}

function validateImportedReceiptRows_(rows) {
  if (!(rows || []).length) {
    throw new Error('Geen ontvangstlijnen ontvangen.');
  }

  const errors = [];

  (rows || []).forEach((row, idx) => {
    const rowNr = idx + 1;

    if (!safeText(row.artikelCode)) {
      errors.push(`Rij ${rowNr}: artikelcode ontbreekt.`);
    }

    if (!safeText(row.artikelOmschrijving)) {
      errors.push(`Rij ${rowNr}: omschrijving ontbreekt.`);
    }

    if (!safeText(row.eenheid)) {
      errors.push(`Rij ${rowNr}: eenheid ontbreekt.`);
    }

    if (safeNumber(row.aantal, 0) <= 0) {
      errors.push(`Rij ${rowNr}: aantal moet groter zijn dan 0.`);
    }
  });

  if (errors.length) {
    throw new Error(errors.join(' | '));
  }
}

function groupImportedReceiptRows_(rows) {
  const grouped = {};

  (rows || []).forEach(row => {
    const key = [
      safeText(row.artikelCode),
      safeText(row.eenheid)
    ].join('|');

    if (!grouped[key]) {
      grouped[key] = {
        artikelCode: safeText(row.artikelCode),
        artikelOmschrijving: safeText(row.artikelOmschrijving),
        eenheid: safeText(row.eenheid),
        besteldAantal: 0,
        ontvangenAantal: 0,
        redenDelta: '',
        opmerking: '',
        _pallets: [],
        _labels: [],
        _notes: []
      };
    }

    grouped[key].besteldAantal += safeNumber(row.aantal, 0);
    grouped[key].ontvangenAantal += safeNumber(row.aantal, 0);

    if (row.palletNr) grouped[key]._pallets.push(safeText(row.palletNr));
    if (row.labelNr) grouped[key]._labels.push(safeText(row.labelNr));
    if (row.notitie) grouped[key]._notes.push(safeText(row.notitie));
  });

  return Object.keys(grouped).map(key => {
    const item = grouped[key];

    const pallets = [...new Set(item._pallets)].filter(Boolean);
    const labels = [...new Set(item._labels)].filter(Boolean);
    const notes = [...new Set(item._notes)].filter(Boolean);

    const parts = ['Import uit ontvangstlijst'];

    if (pallets.length) parts.push(`Pallets: ${pallets.join(', ')}`);
    if (labels.length) parts.push(`Labels: ${labels.join(', ')}`);
    if (notes.length) parts.push(`Notities: ${notes.join(' | ')}`);

    return {
      artikelCode: item.artikelCode,
      artikelOmschrijving: item.artikelOmschrijving,
      eenheid: item.eenheid,
      besteldAantal: item.besteldAantal,
      ontvangenAantal: item.ontvangenAantal,
      redenDelta: '',
      opmerking: parts.join(' | ')
    };
  });
}

function importReceiptFromRows(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertWarehouseAccess(sessionId);

  const rawRows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!rawRows.length) {
    throw new Error('Geen ontvangstlijnen ontvangen.');
  }

  const fileName = safeText(payload.fileName);
  const documentDatum =
    safeText(payload.documentDatum) ||
    parseReceiptDateFromFilename_(fileName);

  if (!documentDatum) {
    throw new Error('Documentdatum ontbreekt.');
  }

  const normalizedRows = rawRows.map(normalizeImportedReceiptRow_);
  validateImportedReceiptRows_(normalizedRows);

  const groupedLines = groupImportedReceiptRows_(normalizedRows);
  if (!groupedLines.length) {
    throw new Error('Na groepering bleven geen geldige ontvangstlijnen over.');
  }

  const createResult = createManualReceipt({
    sessionId: sessionId,
    bestelbonNr: safeText(payload.bestelbonNr),
    externeReferentie: safeText(payload.externeReferentie),
    documentDatum: documentDatum,
    ontvangstdatum: safeText(payload.ontvangstdatum || documentDatum),
    bronType: 'Upload'
  });

  const ontvangstId = safeText(createResult.ontvangstId);
  if (!ontvangstId) {
    throw new Error('Kon geen ontvangst aanmaken.');
  }

  saveReceiptLines({
    sessionId: sessionId,
    ontvangstId: ontvangstId,
    lines: groupedLines
  });

  updateReceiptHeader(ontvangstId, {
    BronType: 'Upload',
    BronBestand: fileName
  });

  writeAudit(
    'Ontvangst geïmporteerd uit upload',
    user.rol,
    user.naam || user.email || 'Magazijn',
    'Ontvangst',
    ontvangstId,
    {
      bronBestand: fileName,
      documentDatum: documentDatum,
      ruweLijnen: rawRows.length,
      gegroepeerdeLijnen: groupedLines.length
    }
  );

  return {
    success: true,
    ontvangstId: ontvangstId,
    lineCount: groupedLines.length,
    message: 'Ontvangst succesvol geïmporteerd.'
  };
}