/* =========================================================
   00_Config.gs — globale config
   ========================================================= */

const APP_CONFIG = {
  TIMEZONE: Session.getScriptTimeZone() || 'Europe/Brussels',
  SESSION_HOURS: 12,
  DEFAULT_NOTIFICATION_STATUS: 'Open',
  DEFAULT_PAGE_TITLE: 'DigiQS Warehouse',
  MAX_AUDIT_ROWS: 10000
};

function getAppTimeZone() {
  return APP_CONFIG.TIMEZONE;
}