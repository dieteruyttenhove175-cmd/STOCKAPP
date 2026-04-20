/* =========================================================
   05_Mappers.gs — centrale mappers
   ========================================================= */

function mapTechnician(row) {
  return {
    code: safeText(row.Code),
    naam: safeText(row.Naam),
    email: safeText(row.Email),
    gsm: safeText(row.GSM),
    patroon: safeText(row.Patroon),
    active: row.Actief === undefined ? true : isTrue(row.Actief)
  };
}

function mapUser(row) {
  return {
    email: normalizeLoginEmail(row.Email),
    naam: safeText(row.Naam),
    rol: safeText(row.Rol),
    techniekerCode: safeText(row.TechniekerCode),
    mobileWarehouseCode: safeText(row.MobielMagazijnCode),
    active: row.Actief === undefined ? true : isTrue(row.Actief),

    loginEmail: normalizeLoginEmail(row.LoginEmail),
    loginCode: safeText(row.LoginCode),
    codeGewijzigdOp: safeText(row.CodeGewijzigdOp),
    laatsteLoginOp: safeText(row.LaatsteLoginOp)
  };
}

function mapSession(row) {
  return {
    sessionId: safeText(row.SessionID),
    loginEmail: normalizeLoginEmail(row.LoginEmail),
    naam: safeText(row.Naam),
    rol: safeText(row.Rol),
    techniekerCode: safeText(row.TechniekerCode),
    aangemaaktOp: safeText(row.AangemaaktOp),
    verlooptOp: safeText(row.VerlooptOp),
    actief: row.Actief === undefined ? true : isTrue(row.Actief)
  };
}

function mapLoginFailure(row) {
  return {
    foutId: safeText(row.FoutID),
    tijdstip: toDisplayDateTime(row.Tijdstip),
    tijdstipRaw: safeText(row.Tijdstip),
    loginEmail: normalizeLoginEmail(row.LoginEmail),
    ingevoerdeCode: safeText(row.IngevoerdeCode),
    reden: safeText(row.Reden),
    matchGebruiker: safeText(row.MatchGebruiker)
  };
}

function mapDelivery(row) {
  const datumIso = toIsoDate(row.Datum);
  const tijdslot = normalizeTime(row.Tijdslot);

  return {
    beleveringId: safeText(row.BeleveringID),
    techniekerCode: safeText(row.TechniekerCode),
    technieker: safeText(row.Technieker),
    datumIso,
    datumDisplay: toDisplayDate(row.Datum),
    dag: safeText(row.Dag || dayNameFromDate(row.Datum)),
    tijdslot,
    patroon: safeText(row.Patroon),
    status: safeText(row.Status),
    sortKey: `${datumIso} ${tijdslot}`.trim(),
    eersteSlotId: safeText(row.EersteSlotID),
    laatsteSlotId: safeText(row.LaatsteSlotID),
    actiefSlotId: safeText(row.ActiefSlotID),
    grabbelKlaar: isTrue(row.GrabbelKlaar),
    behoefteKlaar: isTrue(row.BehoefteKlaar),
    volledigKlaar: isTrue(row.VolledigKlaar)
  };
}

function mapStockItem(row) {
  return {
    artikelCode: safeText(row.Artikelnummer || row.ArtikelCode),
    omschrijving: safeText(row.Artikelomschrijving || row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    pick: safeNumber(row.Pick || row.BasisehInPickeh, 0),
    active: row.Actief === undefined ? true : isTrue(row.Actief)
  };
}

function mapSupplierArticle(row) {
  return {
    artikelsoort: safeText(row.Artikelsoort),
    artikelCode: safeText(row.ArtikelCode || row.Artikel),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikelomschrijving),
    vestSpecifArtStatus: safeText(row.VestSpecifArtStatus || row['VestSpecif ArtStatus']),
    eenheid: safeText(row.Eenheid || row['Basis-HE']),
    pickEenheid: safeNumber(row.PickEenheid || row.BasisehInPickeh, 0),
    packEenheid: safeNumber(row.PackEenheid || row.basisEHinPackEH, 0),
    palletEenheid: safeNumber(row.PalletEenheid || row.basisEHinPalEH, 0),
    leverancier: safeText(row.Leverancier),
    actief: String(row.Actief || '').trim() === '' ? true : isTrue(row.Actief),
    min: row.Min === '' || row.Min == null ? '' : safeNumber(row.Min, 0),
    max: row.Max === '' || row.Max == null ? '' : safeNumber(row.Max, 0)
  };
}

function mapWarehouseOrder(row) {
  const beleveringDatumIso = toIsoDate(row.BeleveringDatum);
  const beleveringUur = normalizeTime(row.BeleveringUur);

  return {
    bestellingId: safeText(row.BestellingID),
    timestamp: toDisplayDateTime(row.Timestamp),
    timestampRaw: safeText(row.Timestamp),
    techniekerCode: safeText(row.TechniekerCode),
    techniekerNaam: safeText(row.TechniekerNaam),
    email: safeText(row.Email),
    gsm: safeText(row.GSM),
    beleveringId: safeText(row.BeleveringID),
    beleveringDatumIso,
    beleveringDag: safeText(row.BeleveringDag),
    beleveringUur,
    patroon: safeText(row.Patroon),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    pick: safeNumber(row.Pick, 0),
    aantalDozen: safeNumber(row.AantalDozen, 0),
    totaalStuks: safeNumber(row.TotaalStuks, 0),
    opmerking: safeText(row.Opmerking),
    status: safeText(row.Status),
    aantalDozenVoorzien: safeNumber(row.AantalDozenVoorzien, 0),
    totaalStuksVoorzien: safeNumber(row.TotaalStuksVoorzien, 0),
    deltaDozen: safeNumber(row.DeltaDozen, 0),
    deltaStuks: safeNumber(row.DeltaStuks, 0),
    redenDelta: safeText(row.RedenDelta),
    inKarretje: safeText(row.InKarretje),
    notitieMagazijn: safeText(row.NotitieMagazijn),
    ontvangenDoorTechnieker: safeText(row.OntvangenDoorTechnieker),
    ontvangenOp: toDisplayDateTime(row.OntvangenOp),
    ontvangenType: safeText(row.OntvangenType),
    managerGoedkeuringStatus: safeText(row.ManagerGoedkeuringStatus),
    managerGoedgekeurdDoor: safeText(row.ManagerGoedgekeurdDoor),
    managerGoedgekeurdOp: toDisplayDateTime(row.ManagerGoedgekeurdOp),
    managerOpmerking: safeText(row.ManagerOpmerking),
    techniekerLijnOntvangen: safeText(row.TechniekerLijnOntvangen),
    techniekerOntvangenDozen: safeNumber(row.TechniekerOntvangenDozen, 0),
    techniekerVerschilReden: safeText(row.TechniekerVerschilReden),
    deliverySortKey: `${beleveringDatumIso} ${beleveringUur}`.trim(),
    klaarTegenLabel: `${toDisplayDate(row.BeleveringDatum)} ${beleveringUur}`.trim()
  };
}

function mapReceipt(row) {
  return {
    ontvangstId: safeText(row.OntvangstID),
    typeMateriaal: safeText(row.TypeMateriaal),
    leverancier: safeText(row.Leverancier),
    bestelbonNr: safeText(row.BestelbonNr),
    externeReferentie: safeText(row.ExterneReferentie),
    bronType: safeText(row.BronType),
    bronBestand: safeText(row.BronBestand),
    documentDatumIso: toIsoDate(row.DocumentDatum),
    documentDatum: toDisplayDate(row.DocumentDatum),
    ontvangstdatumIso: toIsoDate(row.Ontvangstdatum),
    ontvangstdatum: toDisplayDate(row.Ontvangstdatum),
    status: safeText(row.Status),
    ingediendDoor: safeText(row.IngediendDoor),
    ingediendOp: toDisplayDateTime(row.IngediendOp),
    goedgekeurdDoor: safeText(row.GoedgekeurdDoor),
    goedgekeurdOp: toDisplayDateTime(row.GoedgekeurdOp),
    managerOpmerking: safeText(row.ManagerOpmerking)
  };
}

function mapReceiptLine(row) {
  return {
    ontvangstId: safeText(row.OntvangstID),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    besteldAantal: safeNumber(row.BesteldAantal, 0),
    ontvangenAantal: safeNumber(row.OntvangenAantal, 0),
    deltaAantal: safeNumber(row.DeltaAantal, 0),
    redenDelta: safeText(row.RedenDelta),
    opmerking: safeText(row.Opmerking),
    actief: String(row.Actief || '').trim() !== 'Nee'
  };
}

function mapReturn(row) {
  return {
    retourId: safeText(row.RetourID),
    typeMateriaal: safeText(row.TypeMateriaal),
    leverancier: safeText(row.Leverancier),
    techniekerCode: safeText(row.TechniekerCode),
    techniekerNaam: safeText(row.TechniekerNaam),
    documentDatumIso: toIsoDate(row.DocumentDatum),
    documentDatum: toDisplayDate(row.DocumentDatum),
    status: safeText(row.Status),
    ingediendDoor: safeText(row.IngediendDoor),
    ingediendOp: toDisplayDateTime(row.IngediendOp),
    goedgekeurdDoor: safeText(row.GoedgekeurdDoor),
    goedgekeurdOp: toDisplayDateTime(row.GoedgekeurdOp),
    managerOpmerking: safeText(row.ManagerOpmerking)
  };
}

function mapReturnLine(row) {
  return {
    retourId: safeText(row.RetourID),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    aantal: safeNumber(row.Aantal, 0),
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),
    actief: String(row.Actief || '').trim() !== 'Nee'
  };
}

function mapNeedIssue(row) {
  return {
    uitgifteId: safeText(row.UitgifteID),
    beleveringId: safeText(row.BeleveringID),
    bron: safeText(row.Bron),
    techniekerCode: safeText(row.TechniekerCode),
    techniekerNaam: safeText(row.TechniekerNaam),
    documentDatumIso: toIsoDate(row.DocumentDatum),
    documentDatum: toDisplayDate(row.DocumentDatum),
    status: safeText(row.Status),
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),
    geboektDoor: safeText(row.GeboektDoor),
    geboektOp: toDisplayDateTime(row.GeboektOp)
  };
}

function mapNeedIssueLine(row) {
  return {
    uitgifteId: safeText(row.UitgifteID),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    aantal: safeNumber(row.Aantal, 0),
    actief: String(row.Actief || '').trim() !== 'Nee'
  };
}

function mapConsumption(row) {
  return {
    verbruikId: safeText(row.VerbruikID),
    techniekerCode: safeText(row.TechniekerCode),
    techniekerNaam: safeText(row.TechniekerNaam),
    documentDatumIso: toIsoDate(row.DocumentDatum),
    documentDatum: toDisplayDate(row.DocumentDatum),
    werfRef: safeText(row.WerfRef),
    status: safeText(row.Status),
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),
    geboektDoor: safeText(row.GeboektDoor),
    geboektOp: toDisplayDateTime(row.GeboektOp)
  };
}

function mapConsumptionLine(row) {
  return {
    verbruikId: safeText(row.VerbruikID),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    aantal: safeNumber(row.Aantal, 0),
    actief: String(row.Actief || '').trim() !== 'Nee'
  };
}

function mapBusCount(row) {
  return {
    tellingId: safeText(row.TellingID),
    techniekerCode: safeText(row.TechniekerCode),
    techniekerNaam: safeText(row.TechniekerNaam),
    documentDatumIso: toIsoDate(row.DocumentDatum),
    documentDatum: toDisplayDate(row.DocumentDatum),
    status: safeText(row.Status),
    scopeType: safeText(row.ScopeType || 'Volledig'),
    scopeArtikelCodes: safeText(row.ScopeArtikelCodes),
    reden: safeText(row.Reden),
    aangemaaktDoor: safeText(row.AangemaaktDoor),
    aangemaaktOp: toDisplayDateTime(row.AangemaaktOp),
    ingediendDoor: safeText(row.IngediendDoor),
    ingediendOp: toDisplayDateTime(row.IngediendOp),
    goedgekeurdDoor: safeText(row.GoedgekeurdDoor),
    goedgekeurdOp: toDisplayDateTime(row.GoedgekeurdOp),
    managerOpmerking: safeText(row.ManagerOpmerking)
  };
}

function mapBusCountLine(row) {
  return {
    tellingId: safeText(row.TellingID),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    systeemAantal: safeNumber(row.SysteemAantal, 0),
    geteldAantal: row.GeteldAantal === '' || row.GeteldAantal == null ? '' : safeNumber(row.GeteldAantal, 0),
    deltaAantal: row.DeltaAantal === '' || row.DeltaAantal == null ? '' : safeNumber(row.DeltaAantal, 0),
    actief: String(row.Actief || '').trim() !== 'Nee'
  };
}

function mapWarehouseMovement(row) {
  return {
    mutatieId: safeText(row.MutatieID),
    datumBoeking: toDisplayDateTime(row.DatumBoeking),
    datumBoekingRaw: safeText(row.DatumBoeking),
    datumDocument: toIsoDate(row.DatumDocument),
    typeMutatie: safeText(row.TypeMutatie),
    bronId: safeText(row.BronID),
    typeMateriaal: safeText(row.TypeMateriaal),
    artikelCode: safeText(row.ArtikelCode),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving),
    eenheid: safeText(row.Eenheid),
    aantalIn: safeNumber(row.AantalIn, 0),
    aantalUit: safeNumber(row.AantalUit, 0),
    nettoAantal: safeNumber(row.NettoAantal, 0),
    locatieVan: safeText(row.LocatieVan),
    locatieNaar: safeText(row.LocatieNaar),
    reden: safeText(row.Reden),
    opmerking: safeText(row.Opmerking),
    goedgekeurdDoor: safeText(row.GoedgekeurdDoor),
    goedgekeurdOp: toDisplayDateTime(row.GoedgekeurdOp)
  };
}

function mapNotification(row) {
  return {
    notificatieId: safeText(row.NotificatieID),
    rol: safeText(row.Rol),
    ontvangerCode: safeText(row.OntvangerCode),
    ontvangerNaam: safeText(row.OntvangerNaam),
    type: safeText(row.Type),
    titel: safeText(row.Titel),
    bericht: safeText(row.Bericht),
    bronType: safeText(row.BronType),
    bronId: safeText(row.BronID),
    status: safeText(row.Status || NOTIFICATION_STATUS.OPEN),
    aangemaaktOp: toDisplayDateTime(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOp),
    gelezenOp: toDisplayDateTime(row.GelezenOp)
  };
}