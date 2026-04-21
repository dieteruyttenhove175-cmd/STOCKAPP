/* =========================================================
   31_GrabbelOrderService.gs — technieker/magazijn/manager
   grabbelstockflow
   ========================================================= */

const PREPARATION_DELTA_REASONS = [
  'Niet genoeg in stock',
  'Teveel besteld',
  'Andere reden'
];

const TECHNICIAN_RECEIPT_DIFF_REASONS = [
  '',
  'Te weinig ontvangen',
  'Niet ontvangen',
  'Beschadigd',
  'Andere reden'
];

function buildOrderRows(technician, delivery, stockMap, selectedItems, opmerking, timestamp, baseBestellingId) {
  const headers = getHeaders(TABS.ORDERS);

  return (selectedItems || []).map((selected, index) => {
    const stockItem = stockMap[selected.articleCode];
    if (!stockItem) {
      throw new Error('Artikel niet gevonden: ' + selected.articleCode);
    }

    const aantalDozen = safeNumber(selected.aantalDozen, 0);
    const totaalStuks = safeNumber(stockItem.pick, 0) * aantalDozen;
    const bestellingId = `${baseBestellingId}-${index + 1}`;

    return buildRowFromHeaders(headers, {
      BestellingID: bestellingId,
      Timestamp: timestamp,
      TechniekerCode: technician.code,
      TechniekerNaam: technician.naam,
      Email: technician.email,
      GSM: technician.gsm,
      BeleveringID: delivery.beleveringId,
      BeleveringDatum: delivery.datumIso,
      BeleveringDag: delivery.dag,
      BeleveringUur: delivery.tijdslot,
      Patroon: technician.patroon,
      ArtikelCode: stockItem.artikelCode,
      ArtikelOmschrijving: stockItem.omschrijving,
      Eenheid: stockItem.eenheid,
      Pick: stockItem.pick,
      AantalDozen: aantalDozen,
      TotaalStuks: totaalStuks,
      Opmerking: safeText(opmerking),
      Status: STATUS.REQUESTED,
      LaatsteUpdate: timestamp,

      AantalDozenVoorzien: aantalDozen,
      TotaalStuksVoorzien: totaalStuks,
      DeltaDozen: 0,
      DeltaStuks: 0,
      RedenDelta: '',

      InKarretje: 'Nee',
      InKarretjeDoor: '',
      InKarretjeOp: '',

      NotitieMagazijn: '',
      KlaargezetDoor: '',
      KlaargezetOp: '',
      MeegegevenDoor: '',
      MeegegevenOp: '',

      OntvangenDoorTechnieker: 'Nee',
      OntvangenOp: '',
      OntvangenType: '',

      ManagerGoedkeuringStatus: MANAGER_STATUS.NONE,
      ManagerGoedgekeurdDoor: '',
      ManagerGoedgekeurdOp: '',
      ManagerOpmerking: '',

      TechniekerLijnOntvangen: '',
      TechniekerOntvangenDozen: '',
      TechniekerVerschilReden: ''
    });
  });
}

function buildStockMapForOrdering() {
  const stockMap = {};
  readObjectsSafe(TABS.STOCK)
    .map(mapStockItem)
    .filter(item => item.active)
    .forEach(item => {
      stockMap[item.artikelCode] = item;
    });

  return stockMap;
}

function getAppData(techRef, sessionId) {
  assertTechnicianAccessToRef(techRef, sessionId);

  const technicians = readObjectsSafe(TABS.TECHNICIANS).map(mapTechnician);
  const technician = findTechnicianByRef(technicians, techRef);
  if (!technician || !technician.active) {
    throw new Error('Technieker niet gevonden of niet actief.');
  }

  const allOrders = getAllGrabbelOrders();
  const techOrders = allOrders
    .filter(item => normalizeRef(item.techniekerCode) === normalizeRef(technician.code))
    .sort((a, b) => String(a.deliverySortKey || '').localeCompare(String(b.deliverySortKey || '')));

  const articlePopularity = buildArticlePopularityMap(allOrders);

  const stockItems = readObjectsSafe(TABS.STOCK)
    .map(mapStockItem)
    .filter(item => item.active)
    .sort((a, b) => {
      const aCount = articlePopularity[a.artikelCode] || 0;
      const bCount = articlePopularity[b.artikelCode] || 0;
      if (bCount !== aCount) return bCount - aCount;
      return String(a.omschrijving || '').localeCompare(String(b.omschrijving || ''));
    });

  const todayIso = Utilities.formatDate(nowDate(), getAppTimeZone(), 'yyyy-MM-dd');

  const planningDeliveries = getDeliveriesForTechnician(technician)
    .map(enrichDeliveryWithTimingState)
    .filter(item => item.datumIso && item.datumIso >= todayIso)
    .sort((a, b) => String(a.sortKey || '').localeCompare(String(b.sortKey || '')))
    .slice(0, 4)
    .map(item => {
      const hasOrder = hasExistingOrderForDelivery(techOrders, item);
      return {
        ...item,
        hasOrder,
        canCreateOrder: !hasOrder && !item.cutoffPassed
      };
    });

  const deliveryGroups = buildDeliveryGroupsFromOrders(techOrders)
    .filter(group => !!group.beleveringDatumIso)
    .filter(group => {
      if (group.groupStatus !== STATUS.RECEIVED) return true;
      const visibleUntil = addDaysToIsoDate(group.beleveringDatumIso, 2);
      return visibleUntil >= todayIso;
    })
    .sort((a, b) => String(a.klaarTegenSort || '').localeCompare(String(b.klaarTegenSort || '')))
    .map(group => ({
      ...group,
      isEditable: group.groupStatus === STATUS.REQUESTED,
      isFrozen: group.groupStatus !== STATUS.REQUESTED,
      isClosed: group.groupStatus === STATUS.RECEIVED
    }));

  return {
    technician,
    planningDeliveries,
    stockItems,
    deliveryGroups,
    notifications: getNotificationsForTechnician(techRef),
    generatedAt: Utilities.formatDate(nowDate(), getAppTimeZone(), 'dd/MM/yyyy HH:mm')
  };
}

function submitOrder(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const access = assertTechnicianAccessToRef(safeText(payload.techRef), sessionId);

  const techRef = safeText(payload.techRef);
  const deliveryId = safeText(payload.deliveryId);
  const opmerking = safeText(payload.opmerking);
  const items = Array.isArray(payload.items) ? payload.items : [];

  if (!techRef) throw new Error('Technieker ontbreekt.');
  if (!deliveryId) throw new Error('Beleveringsmoment ontbreekt.');

  const selectedItems = items
    .map(item => ({
      articleCode: safeText(item.articleCode),
      aantalDozen: safeNumber(item.aantalDozen, 0)
    }))
    .filter(item => item.articleCode && item.aantalDozen > 0);

  if (!selectedItems.length) {
    throw new Error('Vul minstens voor 1 artikel een aantal dozen in.');
  }

  const technicians = readObjectsSafe(TABS.TECHNICIANS).map(mapTechnician);
  const technician = findTechnicianByRef(technicians, techRef);
  if (!technician) throw new Error('Technieker niet gevonden.');

  const delivery = getAllDeliveries()
    .find(d => d.beleveringId === deliveryId && deliveryMatchesTechnician(d, technician));

  if (!delivery) throw new Error('Beleveringsmoment niet gevonden.');
  if (isDeliveryCutoffPassed(delivery)) {
    throw new Error('Bestellen voor dit beleveringsmoment kan niet meer.');
  }

  const existingOrders = getOrdersForTechnician(technician.code);
  if (hasExistingOrderForDelivery(existingOrders, delivery)) {
    throw new Error('Voor dit beleveringsmoment bestaat al een bestelling.');
  }

  const stockMap = buildStockMapForOrdering();
  const timestamp = nowStamp();
  const baseBestellingId = makeOrderId();

  const rows = buildOrderRows(
    technician,
    delivery,
    stockMap,
    selectedItems,
    opmerking,
    timestamp,
    baseBestellingId
  );

  appendRows(TABS.ORDERS, rows);
  rebuildTotalsSheet();

  writeAudit(
    'Bestelling aangemaakt',
    access.user.rol,
    access.user.naam || access.user.techniekerCode || access.user.email,
    'BestellingGroep',
    baseBestellingId,
    {
      techniekerCode: technician.code,
      beleveringId: delivery.beleveringId,
      lijnen: rows.length,
      opmerking: opmerking
    }
  );

  pushWarehouseNotification(
    'NieuweGrabbelstockBestelling',
    'Nieuwe grabbelstock bestelling',
    `${technician.naam} plaatste een grabbelstock bestelling voor ${delivery.datumDisplay} ${delivery.tijdslot}.`,
    'BestellingGroep',
    delivery.beleveringId || baseBestellingId
  );

  return {
    success: true,
    bestellingId: baseBestellingId,
    lines: rows.length,
    message: 'Bestelling opgeslagen.'
  };
}

function updateTechnicianOrderGroup(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const access = assertTechnicianAccessToRef(safeText(payload.techRef), sessionId);

  const techRef = safeText(payload.techRef);
  const deliveryId = safeText(payload.deliveryId);
  const opmerking = safeText(payload.opmerking);
  const items = Array.isArray(payload.items) ? payload.items : [];

  const selectedItems = items
    .map(item => ({
      articleCode: safeText(item.articleCode),
      aantalDozen: safeNumber(item.aantalDozen, 0)
    }))
    .filter(item => item.articleCode && item.aantalDozen > 0);

  if (!selectedItems.length) {
    throw new Error('De aangepaste bestelling moet minstens 1 artikellijn bevatten.');
  }

  const technicians = readObjectsSafe(TABS.TECHNICIANS).map(mapTechnician);
  const technician = findTechnicianByRef(technicians, techRef);
  if (!technician) throw new Error('Technieker niet gevonden.');

  const delivery = getAllDeliveries()
    .find(d => d.beleveringId === deliveryId && deliveryMatchesTechnician(d, technician));

  if (!delivery) throw new Error('Beleveringsmoment niet gevonden.');
  if (isDeliveryCutoffPassed(delivery)) {
    throw new Error('Bestellen voor dit beleveringsmoment kan niet meer.');
  }

  const sheet = getSheetOrThrow(TABS.ORDERS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab Bestellingen is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);
  const rows = values.slice(1);

  const targetRows = rows.filter(row =>
    normalizeRef(safeText(row[col['TechniekerCode']])) === normalizeRef(technician.code) &&
    safeText(row[col['BeleveringID']]) === deliveryId
  );

  if (!targetRows.length) {
    throw new Error('Er bestaat nog geen bestelling om aan te passen.');
  }

  if (targetRows.some(row => safeText(row[col['Status']]) !== STATUS.REQUESTED)) {
    throw new Error('Deze bestelling is bevroren.');
  }

  const stockMap = buildStockMapForOrdering();

  const remainingRows = rows.filter(row => !(
    normalizeRef(safeText(row[col['TechniekerCode']])) === normalizeRef(technician.code) &&
    safeText(row[col['BeleveringID']]) === deliveryId
  ));

  const newRows = buildOrderRows(
    technician,
    delivery,
    stockMap,
    selectedItems,
    opmerking,
    nowStamp(),
    makeOrderId()
  );

  writeFullTable(TABS.ORDERS, headers, remainingRows.concat(newRows));
  rebuildTotalsSheet();

  writeAudit(
    'Bestelling aangepast',
    access.user.rol,
    access.user.naam || access.user.techniekerCode || access.user.email,
    'BestellingGroep',
    deliveryId,
    {
      techniekerCode: technician.code,
      beleveringId: deliveryId,
      lijnen: newRows.length,
      opmerking: opmerking
    }
  );

  return {
    success: true,
    lines: newRows.length,
    message: 'Bestelling aangepast.'
  };
}

function confirmTechnicianReceipt(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const access = assertTechnicianAccessToRef(safeText(payload.techRef), sessionId);

  const orderIds = Array.isArray(payload.orderIds) ? payload.orderIds.map(safeText).filter(Boolean) : [];
  const techRef = safeText(payload.techRef);
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!orderIds.length) throw new Error('Geen orderIds ontvangen.');
  if (!techRef) throw new Error('Technieker ontbreekt.');

  const technicians = readObjectsSafe(TABS.TECHNICIANS).map(mapTechnician);
  const technician = findTechnicianByRef(technicians, techRef);
  if (!technician) throw new Error('Technieker niet gevonden.');

  const lineMap = {};
  lines.forEach(line => {
    const id = safeText(line.orderId);
    if (!id) return;

    lineMap[id] = {
      receivedOk: !!line.receivedOk,
      ontvangenDozen: safeNumber(line.ontvangenDozen, 0),
      verschilReden: safeText(line.verschilReden)
    };
  });

  const sheet = getSheetOrThrow(TABS.ORDERS);
  const data = sheet.getDataRange().getValues();
  if (!data.length) throw new Error('Tab Bestellingen is leeg.');

  const headers = data[0].map(h => safeText(h));
  const col = getColMap(headers);

  const idSet = {};
  orderIds.forEach(id => { idSet[id] = true; });

  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentId = safeText(row[col['BestellingID']]);

    if (!idSet[currentId]) continue;
    if (normalizeRef(safeText(row[col['TechniekerCode']])) !== normalizeRef(technician.code)) continue;

    const currentStatus = safeText(row[col['Status']]);
    if (!(currentStatus === STATUS.DISPATCHED || currentStatus === STATUS.AUTO_RECEIVED)) continue;

    const providedDozen = safeNumber(row[col['AantalDozenVoorzien']], 0);
    const input = lineMap[currentId] || null;

    let exactReceived = true;
    let ontvangenDozen = providedDozen;
    let verschilReden = '';

    if (providedDozen > 0 && input) {
      exactReceived = !!input.receivedOk;
      ontvangenDozen = exactReceived ? providedDozen : safeNumber(input.ontvangenDozen, 0);
      verschilReden = safeText(input.verschilReden);

      if (!exactReceived) {
        if (ontvangenDozen < 0) {
          throw new Error('Ontvangen dozen mag niet negatief zijn.');
        }

        if (ontvangenDozen >= providedDozen) {
          throw new Error('Bij een verschil moet ontvangen lager zijn dan voorzien.');
        }

        if (!verschilReden) {
          throw new Error('Verschilreden ontbreekt.');
        }

        if (!TECHNICIAN_RECEIPT_DIFF_REASONS.includes(verschilReden)) {
          throw new Error('Ongeldige verschilreden.');
        }
      }
    }

    row[col['TechniekerLijnOntvangen']] = exactReceived ? 'Ja' : 'Nee';
    row[col['TechniekerOntvangenDozen']] = ontvangenDozen;
    row[col['TechniekerVerschilReden']] = verschilReden;
    row[col['Status']] = STATUS.RECEIVED;
    row[col['OntvangenDoorTechnieker']] = 'Ja';
    row[col['OntvangenOp']] = nowStamp();
    row[col['OntvangenType']] = 'Bevestigd door technieker';
    row[col['ManagerGoedkeuringStatus']] = MANAGER_STATUS.PENDING;
    if (col['LaatsteUpdate'] !== undefined) row[col['LaatsteUpdate']] = nowStamp();

    data[i] = row;
    updated++;
  }

  if (!updated) {
    throw new Error('Geen lijnen gevonden voor ontvangst.');
  }

  if (data.length > 1) {
    sheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
  }

  const firstOrder = getAllGrabbelOrders().find(x => orderIds.includes(x.bestellingId));

  if (firstOrder) {
    pushManagerNotification(
      'OntvangstTeControleren',
      'Technieker bevestigde ontvangst',
      `${firstOrder.techniekerNaam} bevestigde ontvangst voor belevering ${firstOrder.klaarTegenLabel}. Managercontrole is nodig.`,
      'BestellingGroep',
      firstOrder.beleveringId || firstOrder.bestellingId
    );
  }

  writeAudit(
    'Ontvangst bevestigd door technieker',
    access.user.rol,
    access.user.naam || access.user.techniekerCode || access.user.email,
    'BestellingGroep',
    firstOrder ? (firstOrder.beleveringId || firstOrder.bestellingId) : orderIds.join(', '),
    { lijnen: updated }
  );

  return {
    success: true,
    updated,
    message: 'Ontvangst bevestigd.'
  };
}

function savePreparationBatch(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertWarehouseAccess(sessionId);

  const actor = safeText(payload.actor || user.naam || 'Magazijn');
  const lines = Array.isArray(payload.lines) ? payload.lines : [];

  if (!lines.length) {
    throw new Error('Geen lijnen ontvangen.');
  }

  const sheet = getSheetOrThrow(TABS.ORDERS);
  const data = sheet.getDataRange().getValues();
  if (!data.length) throw new Error('Tab Bestellingen is leeg.');

  const headers = data[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = 0;

  lines.forEach(line => {
    const orderId = safeText(line.orderId);
    const rowIndex = data.findIndex((row, i) => i > 0 && safeText(row[col['BestellingID']]) === orderId);
    if (rowIndex === -1) return;

    const row = data[rowIndex];
    const currentStatus = safeText(row[col['Status']]);

    if (currentStatus !== STATUS.REQUESTED && currentStatus !== STATUS.READY) {
      return;
    }

    const gevraagdDozen = safeNumber(row[col['AantalDozen']], 0);
    const gevraagdStuks = safeNumber(row[col['TotaalStuks']], 0);
    const pick = safeNumber(row[col['Pick']], 0);

    const inKarretje = line.inKarretje ? 'Ja' : 'Nee';
    let aantalDozenVoorzien = safeNumber(line.aantalDozenVoorzien, 0);
    let redenDelta = safeText(line.redenDelta);

    if (inKarretje === 'Ja') {
      aantalDozenVoorzien = gevraagdDozen;
      redenDelta = '';
    } else {
      if (aantalDozenVoorzien < 0) {
        throw new Error(`Negatief voorzien aantal bij order ${orderId}.`);
      }

      if (aantalDozenVoorzien >= gevraagdDozen) {
        throw new Error(`Bij niet-afgevinkt moet voorzien lager zijn dan gevraagd bij order ${orderId}.`);
      }

      if (!PREPARATION_DELTA_REASONS.includes(redenDelta)) {
        throw new Error(`Ongeldige reden bij order ${orderId}.`);
      }
    }

    const voorzienStuks = pick * aantalDozenVoorzien;

    row[col['AantalDozenVoorzien']] = aantalDozenVoorzien;
    row[col['TotaalStuksVoorzien']] = voorzienStuks;
    row[col['DeltaDozen']] = gevraagdDozen - aantalDozenVoorzien;
    row[col['DeltaStuks']] = gevraagdStuks - voorzienStuks;
    row[col['RedenDelta']] = redenDelta;
    row[col['InKarretje']] = inKarretje;

    if (col['LaatsteUpdate'] !== undefined) row[col['LaatsteUpdate']] = nowStamp();
    if (col['InKarretjeDoor'] !== undefined) row[col['InKarretjeDoor']] = inKarretje === 'Ja' ? actor : '';
    if (col['InKarretjeOp'] !== undefined) row[col['InKarretjeOp']] = inKarretje === 'Ja' ? nowStamp() : '';

    data[rowIndex] = row;
    updated++;
  });

  if (!updated) {
    throw new Error('Geen lijnen bijgewerkt.');
  }

  if (data.length > 1) {
    sheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
  }

  writeAudit(
    'Voorbereidingslijnen opgeslagen',
    user.rol,
    actor,
    'BestellingLijnen',
    lines.map(x => safeText(x.orderId)).join(', '),
    { lijnen: lines.length }
  );

  return {
    success: true,
    updated,
    message: 'Voorbereidingslijnen opgeslagen.'
  };
}

function updateWarehouseGroupStatus(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertWarehouseAccess(sessionId);

  const orderIds = Array.isArray(payload.orderIds) ? payload.orderIds.map(safeText).filter(Boolean) : [];
  const status = safeText(payload.status);
  const actor = safeText(payload.actor || user.naam || 'Magazijn');
  const note = safeText(payload.note);

  if (!orderIds.length) throw new Error('Geen orderIds ontvangen.');
  if (![STATUS.READY, STATUS.DISPATCHED, STATUS.NOT_PICKED].includes(status)) {
    throw new Error('Ongeldige status.');
  }

  const sheet = getSheetOrThrow(TABS.ORDERS);
  const data = sheet.getDataRange().getValues();
  if (!data.length) throw new Error('Tab Bestellingen is leeg.');

  const headers = data[0].map(h => safeText(h));
  const col = getColMap(headers);

  const idSet = {};
  orderIds.forEach(id => { idSet[id] = true; });

  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentId = safeText(row[col['BestellingID']]);
    if (!idSet[currentId]) continue;

    row[col['Status']] = status;
    if (col['LaatsteUpdate'] !== undefined) row[col['LaatsteUpdate']] = nowStamp();
    if (col['NotitieMagazijn'] !== undefined) row[col['NotitieMagazijn']] = note;

    if (status === STATUS.READY) {
      if (col['KlaargezetDoor'] !== undefined) row[col['KlaargezetDoor']] = actor;
      if (col['KlaargezetOp'] !== undefined) row[col['KlaargezetOp']] = nowStamp();
    }

    if (status === STATUS.DISPATCHED) {
      if (col['MeegegevenDoor'] !== undefined) row[col['MeegegevenDoor']] = actor;
      if (col['MeegegevenOp'] !== undefined) row[col['MeegegevenOp']] = nowStamp();
    }

    data[i] = row;
    updated++;
  }

  if (!updated) {
    throw new Error('Geen lijnen gevonden.');
  }

  if (data.length > 1) {
    sheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
  }

  const firstOrder = getAllGrabbelOrders().find(x => orderIds.includes(x.bestellingId));

  if (updated && status === STATUS.READY && firstOrder) {
    pushTechnicianNotification(
      firstOrder.techniekerCode,
      firstOrder.techniekerNaam,
      'LeveringKlaargezet',
      'Je levering staat klaar',
      `Je belevering van ${firstOrder.klaarTegenLabel} staat op Klaargezet.`,
      'BestellingGroep',
      firstOrder.beleveringId || firstOrder.bestellingId
    );
  }

  if (updated && status === STATUS.DISPATCHED && firstOrder) {
    pushTechnicianNotification(
      firstOrder.techniekerCode,
      firstOrder.techniekerNaam,
      'LeveringMeegegeven',
      'Je levering is meegegeven',
      `Je belevering van ${firstOrder.klaarTegenLabel} is meegegeven. Bevestig ontvangst zodra je alles ontvangen hebt.`,
      'BestellingGroep',
      firstOrder.beleveringId || firstOrder.bestellingId
    );
  }

  writeAudit(
    'Beleveringsstatus gewijzigd',
    user.rol,
    actor,
    'BestellingGroep',
    firstOrder ? (firstOrder.beleveringId || firstOrder.bestellingId) : orderIds.join(', '),
    {
      status: status,
      lijnen: updated,
      note: note
    }
  );

  return {
    success: true,
    updated,
    status,
    message: 'Beleveringsmoment bijgewerkt.'
  };
}

function approveManagerGroup(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const user = assertManagerAccess(sessionId);

  const orderIds = Array.isArray(payload.orderIds) ? payload.orderIds.map(safeText).filter(Boolean) : [];
  const actor = safeText(payload.actor || user.naam || 'Manager');
  const note = safeText(payload.note);

  if (!orderIds.length) {
    throw new Error('Geen orderIds ontvangen.');
  }

  const sheet = getSheetOrThrow(TABS.ORDERS);
  const data = sheet.getDataRange().getValues();
  if (!data.length) throw new Error('Tab Bestellingen is leeg.');

  const headers = data[0].map(h => safeText(h));
  const col = getColMap(headers);

  const idSet = {};
  orderIds.forEach(id => { idSet[id] = true; });

  let updated = 0;
  let bronId = '';
  let techniekerNaam = '';
  let klaarTegenLabel = '';

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentId = safeText(row[col['BestellingID']]);
    if (!idSet[currentId]) continue;

    if (!bronId) bronId = safeText(row[col['BeleveringID']]) || currentId;
    if (!techniekerNaam) techniekerNaam = safeText(row[col['TechniekerNaam']]);
    if (!klaarTegenLabel) {
      klaarTegenLabel = `${toDisplayDate(row[col['BeleveringDatum']])} ${normalizeTime(row[col['BeleveringUur']])}`.trim();
    }

    row[col['ManagerGoedkeuringStatus']] = MANAGER_STATUS.APPROVED;
    row[col['ManagerGoedgekeurdDoor']] = actor;
    row[col['ManagerGoedgekeurdOp']] = nowStamp();
    row[col['ManagerOpmerking']] = note;
    row[col['Status']] = STATUS.APPROVED;
    if (col['LaatsteUpdate'] !== undefined) row[col['LaatsteUpdate']] = nowStamp();

    data[i] = row;
    updated++;
  }

  if (!updated) {
    throw new Error('Geen lijnen gevonden voor goedkeuring.');
  }

  if (data.length > 1) {
    sheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
  }

  if (bronId) {
    markManagerNotificationsBySource('BestellingGroep', bronId);

    pushWarehouseNotification(
      'LeveringGoedgekeurd',
      'Grabbelstock levering goedgekeurd',
      `${techniekerNaam || 'Technieker'} - belevering ${klaarTegenLabel || bronId} is door de manager goedgekeurd.`,
      'BestellingGroep',
      bronId
    );
  }

  writeAudit(
    'Grabbelstock levering goedgekeurd',
    user.rol,
    actor,
    'BestellingGroep',
    bronId || orderIds.join(', '),
    {
      lijnen: updated,
      note: note
    }
  );

  return {
    success: true,
    updated,
    message: 'Goedgekeurd.'
  };
}

function autoConfirmAfter3Hours() {
  const sheet = getSheetOrThrow(TABS.ORDERS);
  const data = sheet.getDataRange().getValues();
  if (!data.length) return { success: true, updated: 0, message: 'Geen bestellingen.' };

  const headers = data[0].map(h => safeText(h));
  const col = getColMap(headers);

  const now = nowDate();
  let updated = 0;
  const affectedGroupIds = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = safeText(row[col['Status']]);
    const meegegevenOp = row[col['MeegegevenOp']];
    const ontvangen = safeText(row[col['OntvangenDoorTechnieker']]);

    if (status !== STATUS.DISPATCHED || !meegegevenOp || ontvangen === 'Ja') continue;

    const meegegevenDate = new Date(meegegevenOp);
    if (isNaN(meegegevenDate)) continue;

    if (now.getTime() - meegegevenDate.getTime() < 3 * 60 * 60 * 1000) continue;

    row[col['Status']] = STATUS.AUTO_RECEIVED;
    row[col['OntvangenDoorTechnieker']] = 'Nee';
    row[col['OntvangenOp']] = nowStamp();
    row[col['OntvangenType']] = 'Automatisch na 3u';
    row[col['ManagerGoedkeuringStatus']] = MANAGER_STATUS.PENDING;
    if (col['LaatsteUpdate'] !== undefined) row[col['LaatsteUpdate']] = nowStamp();

    data[i] = row;

    const bronId = safeText(row[col['BeleveringID']]) || safeText(row[col['BestellingID']]);
    if (bronId) affectedGroupIds[bronId] = true;
    updated++;
  }

  if (updated && data.length > 1) {
    sheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
  }

  Object.keys(affectedGroupIds).forEach(bronId => {
    pushManagerNotification(
      'AutoOntvangst',
      'Automatische ontvangst na 3 uur',
      `Belevering ${bronId} is automatisch ontvangen gemarkeerd omdat de technieker niet tijdig bevestigde.`,
      'BestellingGroep',
      bronId
    );
  });

  if (updated) {
    writeSystemAudit(
      'Automatische ontvangst na 3 uur',
      'BestellingGroep',
      Object.keys(affectedGroupIds).join(', '),
      {
        groepen: Object.keys(affectedGroupIds).length,
        lijnen: updated
      }
    );
  }

  return {
    success: true,
    updated,
    message: 'Automatische ontvangstcontrole uitgevoerd.'
  };
}