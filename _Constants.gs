/* =========================================================
   01_Constants.gs — tabs / rollen / locaties / notificaties
   ========================================================= */

const TABS = {
  TECHNICIANS: 'Techniekers',
  USERS: 'Gebruikers',
  DELIVERIES: 'Beleveringen',
  STOCK: 'Grabbelstock',
  ORDERS: 'Bestellingen',
  TOTALS: 'TotaalMaterialen',

  SUPPLIER_ARTICLES: 'LeveranciersArtikelen',

  RECEIPTS: 'Ontvangsten',
  RECEIPT_LINES: 'OntvangstLijnen',

  RETURNS: 'Retouren',
  RETURN_LINES: 'RetourLijnen',

  NEED_ISSUES: 'BehoefteUitgiftes',
  NEED_ISSUE_LINES: 'BehoefteUitgifteLijnen',

  CONSUMPTIONS: 'Verbruiken',
  CONSUMPTION_LINES: 'VerbruikLijnen',

  BUS_COUNTS: 'BusStockTellingen',
  BUS_COUNT_LINES: 'BusStockTellingLijnen',

  TRANSFER_REQUESTS: 'TransferRequests',
  TRANSFER_REQUEST_LINES: 'TransferRequestLijnen',
  TRANSFERS: 'Transfers',
  TRANSFER_LINES: 'TransferLijnen',

  WAREHOUSE_MOVEMENTS: 'MagazijnMutaties',
  CENTRAL_WAREHOUSE: 'CentraalMagazijn',
  MOBILE_WAREHOUSE: 'MobielMagazijn',

  NOTIFICATIONS: 'Notificaties',
  AUDIT_LOG: 'AuditLog',
  SESSIONS: 'Sessies',
  LOGIN_FAILURES: 'LoginFouten',

  CONSUMPTION_IMPORT_RUNS: 'VerbruikImportRuns',
  CONSUMPTION_IMPORT_RAW: 'VerbruikImportRaw',
  CONSUMPTION_IMPORT_CONFIG: 'VerbruikImportConfig',
  CONSUMPTION_IMPORT_LOG: 'VerbruikImportLog',

  CENTRAL_COUNTS : 'CentralStockTellingen',
  CENTRAL_COUNT_LINES : 'CentralStockTellingLijnen',

  MOBILE_REQUESTS : 'MobielMagazijnAanvragen',
  MOBILE_REQUEST_LINES : 'MobielMagazijnAanvraagLijnen',

  

  MOBILE_WAREHOUSES : 'MobieleMagazijnen'
};

const ROLE = {
  TECHNICIAN: 'Technieker',
  WAREHOUSE: 'Magazijn',
  MOBILE_WAREHOUSE: 'MobielMagazijn',
  MANAGER: 'Manager',
  ANALYSIS: 'Analyse',
  ADMIN: 'Admin'
};

const LOCATION = {
  CENTRAL: 'CentraalMagazijn',
  MOBILE: 'MobielMagazijn',
  SITE: 'Werf'
};

const NOTIFICATION_ROLE = {
  TECHNICIAN: 'Technieker',
  WAREHOUSE: 'Magazijn',
  MOBILE_WAREHOUSE: 'MobielMagazijn',
  MANAGER: 'Manager'
};

const MATERIAL_TYPE = {
  GRABBEL: 'Grabbel',
  NEED: 'Behoefte',
  MIXED: 'Gemengd'
};

const IMPORT_RUN_STATUS = {
  STARTED: 'Gestart',
  IMPORTED: 'Geïmporteerd',
  FAILED: 'Mislukt'
};

const IMPORT_VALIDATION_STATUS = {
  NEW: 'Nieuw',
  DUPLICATE: 'Dubbel',
  ERROR: 'Fout',
  VALIDATED: 'Gevalideerd',
  SKIPPED: 'Overgeslagen'
};

const IMPORT_BOOK_STATUS = {
  NOT_BOOKED: 'Niet geboekt',
  BOOKED: 'Geboekt',
  BOOK_ERROR: 'Boekfout'
};

const CENTRAL_COUNT_STATUS = {
  OPEN: 'Open',
  SUBMITTED: 'Ingediend',
  APPROVED: 'Goedgekeurd'
};

function getBusLocationCode(techniekerCode) {
  return 'Bus:' + String(techniekerCode || '').trim();
}

function parseBusLocation(location) {
  const text = String(location || '').trim();
  if (!/^Bus:/i.test(text)) return '';
  return text.split(':').slice(1).join(':').trim();
}