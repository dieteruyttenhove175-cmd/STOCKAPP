/* =========================================================
   03_Helpers.gs — generieke helpers
   ========================================================= */

function nowDate() {
  return new Date();
}

function nowStamp() {
  return Utilities.formatDate(nowDate(), getAppTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function formatDateValue(date, format) {
  return Utilities.formatDate(date, getAppTimeZone(), format);
}

function makeStampedId(prefix) {
  return String(prefix || 'ID') + '-' +
    Utilities.formatDate(nowDate(), getAppTimeZone(), 'yyyyMMddHHmmss') +
    '-' + Math.floor(Math.random() * 900 + 100);
}

function makeUuidId(prefix) {
  return String(prefix || 'ID') + '-' + Utilities.getUuid();
}

function isTrue(value) {
  const text = String(value || '').trim().toLowerCase();
  return ['ja', 'true', '1', 'yes', 'y', 'actief', 'active'].includes(text);
}

function safeText(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function safeNumber(value, fallback) {
  const num = Number(value);
  return isNaN(num) ? (fallback === undefined ? 0 : fallback) : num;
}

function normalizeRef(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[@._]/g, '-')
    .replace(/\s+/g, '-')
    .replace(/[^a-z0-9\-]/g, '')
    .replace(/\-+/g, '-')
    .replace(/^\-|\-$/g, '');
}

function normalizeSearchValue(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeLoginEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function isLikelyEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function toIsoDate(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, getAppTimeZone(), 'yyyy-MM-dd');
  }

  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  const dmy = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) {
    return `${dmy[3]}-${String(dmy[2]).padStart(2, '0')}-${String(dmy[1]).padStart(2, '0')}`;
  }

  const parsed = new Date(text);
  if (!isNaN(parsed)) {
    return Utilities.formatDate(parsed, getAppTimeZone(), 'yyyy-MM-dd');
  }

  return '';
}

function toDisplayDate(value) {
  const iso = toIsoDate(value);
  if (!iso) return '';
  const parts = iso.split('-');
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

function toDisplayDateTime(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, getAppTimeZone(), 'dd/MM/yyyy HH:mm');
  }

  const parsed = new Date(value);
  if (!isNaN(parsed)) {
    return Utilities.formatDate(parsed, getAppTimeZone(), 'dd/MM/yyyy HH:mm');
  }

  return String(value || '');
}

function normalizeTime(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, getAppTimeZone(), 'HH:mm');
  }

  const text = String(value).trim().replace('u', ':').replace('.', ':');
  const match = text.match(/(\d{1,2}):(\d{2})/);
  if (match) return `${String(match[1]).padStart(2, '0')}:${match[2]}`;

  return text;
}

function addDaysToIsoDate(isoDate, days) {
  if (!isoDate) return '';
  const d = new Date(isoDate + 'T00:00:00');
  d.setDate(d.getDate() + Number(days || 0));
  return Utilities.formatDate(d, getAppTimeZone(), 'yyyy-MM-dd');
}

function dayNameFromDate(value) {
  const iso = toIsoDate(value);
  if (!iso) return '';
  const date = new Date(iso + 'T00:00:00');
  const names = ['Zondag', 'Maandag', 'Dinsdag', 'Woensdag', 'Donderdag', 'Vrijdag', 'Zaterdag'];
  return names[date.getDay()] || '';
}

function getPayloadSessionId(payload) {
  return String((payload && payload.sessionId) || '').trim();
}

function runAction(options) {
  const button = options.buttonId ? document.getElementById(options.buttonId) : null;
  const originalText = button ? button.textContent : '';

  clearError();

  if (button) {
    button.disabled = true;
    button.textContent = options.busyButtonText || 'Backend bezig...';
  }

  showBusy(options.busyMessage || 'Backend bezig...');

  const runner = google.script.run
    .withSuccessHandler(result => {
      hideBusy();
      if (button) {
        button.disabled = false;
        button.textContent = originalText;
      }
      if (options.onSuccess) options.onSuccess(result);
    })
    .withFailureHandler(err => {
      hideBusy();
      if (button) {
        button.disabled = false;
        button.textContent = originalText;
      }
      showError(err && err.message ? err.message : 'Onbekende fout.', true);
      if (options.onFailure) options.onFailure(err);
    });

  runner[options.method](options.payload || {});
}

function safeJson(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
}