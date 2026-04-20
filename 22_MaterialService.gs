/* =========================================================
   22_MaterialService.gs — materiaaltype / artikelmaster
   ========================================================= */

function getGrabbelArticleCodeSet() {
  const set = {};

  readObjectsSafe(TABS.STOCK)
    .map(mapStockItem)
    .filter(item => item.active)
    .forEach(item => {
      const code = safeText(item.artikelCode);
      if (code) set[code] = true;
    });

  return set;
}

function isGrabbelArticle(articleCode) {
  const code = safeText(articleCode);
  if (!code) return false;

  const grabbelSet = getGrabbelArticleCodeSet();
  return !!grabbelSet[code];
}

function isBehoefteArticle(articleCode) {
  return determineMaterialTypeFromArticle(articleCode) === MATERIAL_TYPE.NEED;
}

function determineMaterialTypeFromArticle(articleCode) {
  const code = safeText(articleCode);
  if (!code) return MATERIAL_TYPE.NEED;

  return isGrabbelArticle(code) ? MATERIAL_TYPE.GRABBEL : MATERIAL_TYPE.NEED;
}

function determineMaterialTypeFromLines(lines, articleCodeFieldName) {
  const fieldName = safeText(articleCodeFieldName) || 'artikelCode';
  const foundTypes = {};

  (lines || []).forEach(line => {
    const code = safeText(line && line[fieldName]);
    if (!code) return;

    const type = determineMaterialTypeFromArticle(code);
    foundTypes[type] = true;
  });

  const keys = Object.keys(foundTypes);
  if (!keys.length) return '';
  if (keys.length === 1) return keys[0];
  return MATERIAL_TYPE.MIXED;
}

function determineReceiptMaterialType(lines) {
  return determineMaterialTypeFromLines(lines, 'artikelCode');
}

function determineNeedIssueMaterialType(lines) {
  return determineMaterialTypeFromLines(lines, 'artikelCode');
}

function determineReturnMaterialType(lines) {
  return determineMaterialTypeFromLines(lines, 'artikelCode');
}

function buildArticleMasterMap() {
  const master = {};

  readObjectsSafe(TABS.SUPPLIER_ARTICLES)
    .map(mapSupplierArticle)
    .filter(item => item.actief !== false)
    .forEach(item => {
      const code = safeText(item.artikelCode);
      if (!code) return;

      master[code] = {
        artikelCode: code,
        artikelOmschrijving: safeText(item.artikelOmschrijving),
        eenheid: safeText(item.eenheid),
        leverancier: safeText(item.leverancier),
        min: item.min === '' ? '' : safeNumber(item.min, 0),
        max: item.max === '' ? '' : safeNumber(item.max, 0),
        materiaalType: determineMaterialTypeFromArticle(code),
        bron: 'LeveranciersArtikelen'
      };
    });

  readObjectsSafe(TABS.STOCK)
    .map(mapStockItem)
    .filter(item => item.active)
    .forEach(item => {
      const code = safeText(item.artikelCode);
      if (!code) return;

      if (!master[code]) {
        master[code] = {
          artikelCode: code,
          artikelOmschrijving: safeText(item.omschrijving),
          eenheid: safeText(item.eenheid),
          leverancier: '',
          min: '',
          max: '',
          materiaalType: MATERIAL_TYPE.GRABBEL,
          bron: 'Grabbelstock'
        };
      } else {
        if (!master[code].artikelOmschrijving) master[code].artikelOmschrijving = safeText(item.omschrijving);
        if (!master[code].eenheid) master[code].eenheid = safeText(item.eenheid);
        if (!master[code].materiaalType) master[code].materiaalType = MATERIAL_TYPE.GRABBEL;
      }
    });

  return master;
}

function getArticleMaster(articleCode) {
  const code = safeText(articleCode);
  if (!code) return null;

  const map = buildArticleMasterMap();
  return map[code] || null;
}

function getAllowedReturnReasonsByActor(actorRole) {
  const role = safeText(actorRole);

  if (role === ROLE.TECHNICIAN) {
    return RETURN_REASONS_TECHNICIAN.slice();
  }

  if (role === ROLE.WAREHOUSE || role === ROLE.MOBILE_WAREHOUSE || role === ROLE.MANAGER) {
    return RETURN_REASONS_WAREHOUSE.slice();
  }

  return RETURN_REASONS_WAREHOUSE.slice();
}