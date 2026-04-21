/* =========================================================
   04_SheetRepository.gs — centrale sheet repository helpers
   ========================================================= */

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetOrThrow(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Tab niet gevonden: ' + sheetName);
  return sheet;
}

function getAllValues(sheetName) {
  return getSheetOrThrow(sheetName).getDataRange().getValues();
}

function getHeaders(sheetName) {
  const values = getAllValues(sheetName);
  return values.length ? values[0].map(h => String(h || '').trim()) : [];
}

function getColMap(headers) {
  const map = {};
  (headers || []).forEach((header, index) => {
    map[String(header || '').trim()] = index;
  });
  return map;
}

function readObjects(sheetName) {
  const values = getAllValues(sheetName);
  if (values.length < 2) return [];

  const headers = values[0].map(v => String(v || '').trim());

  return values.slice(1)
    .filter(row => row.some(cell => String(cell || '').trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
}

function readObjectsSafe(sheetName) {
  try {
    return readObjects(sheetName);
  } catch (e) {
    return [];
  }
}

function buildRowFromHeaders(headers, object) {
  return (headers || []).map(header =>
    object && object[header] !== undefined ? object[header] : ''
  );
}

function appendRows(sheetName, rows) {
  if (!rows || !rows.length) return;
  const sheet = getSheetOrThrow(sheetName);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function appendObjects(sheetName, objects) {
  if (!objects || !objects.length) return;
  const headers = getHeaders(sheetName);
  const rows = objects.map(obj => buildRowFromHeaders(headers, obj));
  appendRows(sheetName, rows);
}

function writeFullTable(sheetName, headers, rows) {
  const sheet = getSheetOrThrow(sheetName);
  sheet.clearContents();

  if (headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  if (rows && rows.length) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function findFirstRowIndexByField(sheetName, fieldName, fieldValue) {
  const values = getAllValues(sheetName);
  if (!values.length) return -1;

  const headers = values[0].map(h => String(h || '').trim());
  const col = getColMap(headers);
  if (col[fieldName] === undefined) return -1;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][col[fieldName]] || '').trim() === String(fieldValue || '').trim()) {
      return i + 1;
    }
  }
  return -1;
}