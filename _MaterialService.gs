/* =========================================================
   22_MaterialService.gs
   Refactor: material master service
   Doel:
   - centrale artikellijst en leveranciersartikels
   - materiaaltype bepalen op één plaats
   - zoeken in volledige artikellijst
   - lookup- en cataloglaag voor andere services
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getSupplierArticlesTab_() {
  return TABS.SUPPLIER_ARTICLES || 'LeveranciersArtikelen';
}

function getMaterialMasterTab_() {
  return TABS.MATERIAL_MASTER || 'MateriaalMaster';
}

function getDefaultMaterialTypeUnknown_() {
  return 'Onbekend';
}

function getDefaultSupplierName_() {
  return 'Fluvius';
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapSupplierArticle(row) {
  var artikelCode = safeText(
    row.ArtikelCode ||
    row.ArtikelNr ||
    row.Code ||
    row.SupplierArticleCode
  );

  return {
    artikelCode: artikelCode,
    artikelOmschrijving: safeText(
      row.ArtikelOmschrijving ||
      row.Artikel ||
      row.Omschrijving ||
      row.Description
    ),
    typeMateriaal: safeText(
      row.TypeMateriaal ||
      row.MaterialType ||
      ''
    ),
    eenheid: safeText(
      row.Eenheid ||
      row.Unit ||
      row.UOM ||
      ''
    ),
    leverancier: safeText(
      row.Leverancier ||
      row.Supplier ||
      getDefaultSupplierName_()
    ),
    leverancierArtikelCode: safeText(
      row.LeverancierArtikelCode ||
      row.SupplierArticleCode ||
      artikelCode
    ),
    actief: row.Actief === undefined ? true : isTrue(row.Actief),
    minStock: safeNumber(row.MinStock || row.Min || 0, 0),
    maxStock: safeNumber(row.MaxStock || row.Max || 0, 0),
    categorie: safeText(row.Categorie || row.Category || ''),
    zoektekst: safeText(
      row.Zoektekst ||
      row.SearchText ||
      [
        artikelCode,
        row.ArtikelOmschrijving || row.Artikel || row.Omschrijving || '',
        row.Leverancier || row.Supplier || '',
        row.Categorie || row.Category || ''
      ].join(' ')
    ),
  };
}

function mapMaterialMasterArticle(row) {
  var artikelCode = safeText(
    row.ArtikelCode ||
    row.ArtikelNr ||
    row.Code ||
    row.MaterialCode
  );

  return {
    artikelCode: artikelCode,
    artikelOmschrijving: safeText(
      row.ArtikelOmschrijving ||
      row.Artikel ||
      row.Omschrijving ||
      row.Description
    ),
    typeMateriaal: safeText(
      row.TypeMateriaal ||
      row.MaterialType ||
      ''
    ),
    eenheid: safeText(
      row.Eenheid ||
      row.Unit ||
      row.UOM ||
      ''
    ),
    actief: row.Actief === undefined ? true : isTrue(row.Actief),
    categorie: safeText(row.Categorie || row.Category || ''),
    minStock: safeNumber(row.MinStock || row.Min || 0, 0),
    maxStock: safeNumber(row.MaxStock || row.Max || 0, 0),
    bron: 'MASTER',
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllSupplierArticles() {
  return readObjectsSafe(getSupplierArticlesTab_())
    .map(mapSupplierArticle)
    .filter(function (item) {
      return !!safeText(item.artikelCode);
    })
    .sort(function (a, b) {
      return (
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function getActiveSupplierArticles() {
  return getAllSupplierArticles().filter(function (item) {
    return item.actief;
  });
}

function getAllMaterialMasterArticles() {
  return readObjectsSafe(getMaterialMasterTab_())
    .map(mapMaterialMasterArticle)
    .filter(function (item) {
      return !!safeText(item.artikelCode);
    })
    .sort(function (a, b) {
      return (
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function getActiveMaterialMasterArticles() {
  return getAllMaterialMasterArticles().filter(function (item) {
    return item.actief;
  });
}

/* ---------------------------------------------------------
   Lookup maps
   --------------------------------------------------------- */

function buildSupplierArticleMap() {
  var map = {};
  getActiveSupplierArticles().forEach(function (item) {
    map[safeText(item.artikelCode)] = item;
  });
  return map;
}

function buildMaterialMasterMap() {
  var map = {};
  getActiveMaterialMasterArticles().forEach(function (item) {
    map[safeText(item.artikelCode)] = item;
  });
  return map;
}

function buildFullArticleCatalog() {
  var supplierMap = buildSupplierArticleMap();
  var masterMap = buildMaterialMasterMap();
  var keys = {};

  Object.keys(masterMap).forEach(function (key) { keys[key] = true; });
  Object.keys(supplierMap).forEach(function (key) { keys[key] = true; });

  return Object.keys(keys)
    .map(function (artikelCode) {
      var master = masterMap[artikelCode] || {};
      var supplier = supplierMap[artikelCode] || {};

      return {
        artikelCode: artikelCode,
        artikelOmschrijving: safeText(master.artikelOmschrijving || supplier.artikelOmschrijving),
        typeMateriaal: safeText(master.typeMateriaal || supplier.typeMateriaal || getDefaultMaterialTypeUnknown_()),
        eenheid: safeText(master.eenheid || supplier.eenheid),
        leverancier: safeText(supplier.leverancier || getDefaultSupplierName_()),
        leverancierArtikelCode: safeText(supplier.leverancierArtikelCode || artikelCode),
        categorie: safeText(master.categorie || supplier.categorie),
        minStock: safeNumber(master.minStock || supplier.minStock, 0),
        maxStock: safeNumber(master.maxStock || supplier.maxStock, 0),
        bron: master.artikelCode && supplier.artikelCode
          ? 'MASTER+SUPPLIER'
          : master.artikelCode
            ? 'MASTER'
            : 'SUPPLIER',
        zoektekst: safeText(
          [
            artikelCode,
            master.artikelOmschrijving || supplier.artikelOmschrijving || '',
            master.categorie || supplier.categorie || '',
            supplier.leverancier || ''
          ].join(' ')
        ),
      };
    })
    .sort(function (a, b) {
      return (
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving)) ||
        safeText(a.artikelCode).localeCompare(safeText(b.artikelCode))
      );
    });
}

function buildFullArticleCatalogMap() {
  var map = {};
  buildFullArticleCatalog().forEach(function (item) {
    map[safeText(item.artikelCode)] = item;
  });
  return map;
}

/* ---------------------------------------------------------
   Type determination
   --------------------------------------------------------- */

function determineMaterialTypeFromPrefix_(artikelCode) {
  var code = safeText(artikelCode).toUpperCase();
  if (!code) return '';

  if (/^(GB|GRAB|G-)/.test(code)) {
    return 'Grabbel';
  }
  if (/^(BH|BEH|B-)/.test(code)) {
    return 'Behoefte';
  }
  return '';
}

function determineMaterialTypeFromArticle(artikelCode) {
  var code = safeText(artikelCode);
  if (!code) {
    return getDefaultMaterialTypeUnknown_();
  }

  var masterMap = buildMaterialMasterMap();
  if (masterMap[code] && safeText(masterMap[code].typeMateriaal)) {
    return safeText(masterMap[code].typeMateriaal);
  }

  var supplierMap = buildSupplierArticleMap();
  if (supplierMap[code] && safeText(supplierMap[code].typeMateriaal)) {
    return safeText(supplierMap[code].typeMateriaal);
  }

  var prefixType = determineMaterialTypeFromPrefix_(code);
  if (prefixType) {
    return prefixType;
  }

  return getDefaultMaterialTypeUnknown_();
}

/* ---------------------------------------------------------
   Single item lookup
   --------------------------------------------------------- */

function getSupplierArticleByCode(artikelCode) {
  var code = safeText(artikelCode);
  if (!code) return null;

  var map = buildSupplierArticleMap();
  return map[code] || null;
}

function getMaterialMasterArticleByCode(artikelCode) {
  var code = safeText(artikelCode);
  if (!code) return null;

  var map = buildMaterialMasterMap();
  return map[code] || null;
}

function getCatalogArticleByCode(artikelCode) {
  var code = safeText(artikelCode);
  if (!code) return null;

  var map = buildFullArticleCatalogMap();
  return map[code] || null;
}

/* ---------------------------------------------------------
   Search / filters
   --------------------------------------------------------- */

function matchesMaterialSearch_(item, query) {
  var q = safeText(query).toLowerCase();
  if (!q) return true;

  return [
    item.artikelCode,
    item.artikelOmschrijving,
    item.typeMateriaal,
    item.eenheid,
    item.leverancier,
    item.leverancierArtikelCode,
    item.categorie,
    item.zoektekst,
  ].some(function (value) {
    return safeText(value).toLowerCase().indexOf(q) >= 0;
  });
}

function filterCatalogArticles_(rows, filters) {
  var f = filters || {};
  var query = safeText(f.query || f.search).toLowerCase();
  var typeMateriaal = safeText(f.typeMateriaal || f.materialType);
  var leverancier = safeText(f.leverancier || f.supplier);
  var bron = safeText(f.bron || f.source);

  return (rows || []).filter(function (item) {
    if (!matchesMaterialSearch_(item, query)) {
      return false;
    }
    if (typeMateriaal && safeText(item.typeMateriaal) !== typeMateriaal) {
      return false;
    }
    if (leverancier && safeText(item.leverancier) !== leverancier) {
      return false;
    }
    if (bron && safeText(item.bron) !== bron) {
      return false;
    }
    return true;
  });
}

function searchSupplierArticles(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  requireLoggedInUser(sessionId);

  var rows = filterCatalogArticles_(
    buildFullArticleCatalog().filter(function (item) {
      return safeText(item.bron) === 'SUPPLIER' || safeText(item.bron) === 'MASTER+SUPPLIER';
    }),
    payload.filters || payload
  );

  return {
    items: rows,
    summary: {
      totaal: rows.length,
    }
  };
}

function searchAllCatalogArticles(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  requireLoggedInUser(sessionId);

  var rows = filterCatalogArticles_(buildFullArticleCatalog(), payload.filters || payload);

  return {
    items: rows,
    summary: {
      totaal: rows.length,
      grabbel: rows.filter(function (x) { return safeText(x.typeMateriaal) === 'Grabbel'; }).length,
      behoefte: rows.filter(function (x) { return safeText(x.typeMateriaal) === 'Behoefte'; }).length,
      onbekend: rows.filter(function (x) { return safeText(x.typeMateriaal) === getDefaultMaterialTypeUnknown_(); }).length,
    }
  };
}

/* ---------------------------------------------------------
   Lightweight UI query
   --------------------------------------------------------- */

function getMaterialCatalogData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  assertRoleAllowed(
    actor,
    [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.MOBILE_WAREHOUSE, ROLE.TECHNICIAN],
    'Geen rechten om artikels te bekijken.'
  );

  var rows = filterCatalogArticles_(buildFullArticleCatalog(), payload.filters || payload);

  return {
    items: rows,
    catalog: rows,
    summary: {
      totaal: rows.length,
      grabbel: rows.filter(function (x) { return safeText(x.typeMateriaal) === 'Grabbel'; }).length,
      behoefte: rows.filter(function (x) { return safeText(x.typeMateriaal) === 'Behoefte'; }).length,
      leveranciers: Object.keys(
        rows.reduce(function (acc, row) {
          acc[safeText(row.leverancier)] = true;
          return acc;
        }, {})
      ).length,
      actorRol: safeText(actor.rol),
    }
  };
}
