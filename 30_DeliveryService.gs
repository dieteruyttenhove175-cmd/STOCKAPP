/* =========================================================
   30_DeliveryService.gs — leveringen / groepen / totalen
   ========================================================= */

function makeOrderId() {
  return makeStampedId('B');
}

function getAllGrabbelOrders() {
  return readObjectsSafe(TABS.ORDERS)
    .map(mapWarehouseOrder);
}

function getOrdersForTechnician(techniekerCode) {
  const code = safeText(techniekerCode);

  return getAllGrabbelOrders()
    .filter(item => normalizeRef(item.techniekerCode) === normalizeRef(code))
    .sort((a, b) => String(a.deliverySortKey || '').localeCompare(String(b.deliverySortKey || '')));
}

function getAllDeliveries() {
  return readObjectsSafe(TABS.DELIVERIES)
    .map(mapDelivery)
    .sort((a, b) => String(a.sortKey || '').localeCompare(String(b.sortKey || '')));
}

function getDeliveriesForTechnician(technician) {
  return getAllDeliveries().filter(item => deliveryMatchesTechnician(item, technician));
}

function hasExistingOrderForDelivery(techOrders, delivery) {
  const targetDeliveryId = safeText(delivery && delivery.beleveringId);
  const targetDate = safeText(delivery && delivery.datumIso);
  const targetTime = safeText(delivery && delivery.tijdslot);

  return (techOrders || []).some(order => {
    const sameDeliveryId = safeText(order.beleveringId) === targetDeliveryId;
    const sameMoment =
      safeText(order.beleveringDatumIso) === targetDate &&
      safeText(order.beleveringUur) === targetTime;

    return sameDeliveryId || sameMoment;
  });
}

function isDeliveryCutoffPassed(delivery) {
  const datumIso = safeText(delivery && delivery.datumIso);
  const tijdslot = safeText(delivery && delivery.tijdslot);
  if (!datumIso || !tijdslot) return false;

  const deliveryDateTime = new Date(`${datumIso}T${tijdslot}:00`);
  if (isNaN(deliveryDateTime)) return false;

  const cutoff = new Date(deliveryDateTime.getTime() - 24 * 60 * 60 * 1000);
  return nowDate().getTime() > cutoff.getTime();
}

function enrichDeliveryWithTimingState(delivery) {
  const datumIso = safeText(delivery && delivery.datumIso);
  const tijdslot = safeText(delivery && delivery.tijdslot);
  const now = nowDate();

  let isUpcoming = false;
  let cutoffPassed = false;
  let cutoffDisplay = '';

  if (datumIso && tijdslot) {
    const deliveryDateTime = new Date(`${datumIso}T${tijdslot}:00`);
    if (!isNaN(deliveryDateTime)) {
      isUpcoming = deliveryDateTime >= now;

      const cutoff = new Date(deliveryDateTime.getTime() - 24 * 60 * 60 * 1000);
      cutoffPassed = now > cutoff;
      cutoffDisplay = Utilities.formatDate(cutoff, getAppTimeZone(), 'dd/MM/yyyy HH:mm');
    }
  }

  return {
    ...delivery,
    isUpcoming,
    cutoffPassed,
    cutoffDisplay
  };
}

function calculateGroupStatus(lineItems) {
  const items = lineItems || [];
  if (!items.length) return STATUS.REQUESTED;

  const statuses = items.map(item => safeText(item.status));
  const all = allowed => statuses.every(s => allowed.includes(s));
  const some = value => statuses.some(s => s === value);

  if (all([STATUS.APPROVED])) {
    return STATUS.APPROVED;
  }

  if (all([STATUS.RECEIVED, STATUS.APPROVED]) && some(STATUS.RECEIVED)) {
    return STATUS.RECEIVED;
  }

  if (
    all([STATUS.AUTO_RECEIVED, STATUS.RECEIVED, STATUS.APPROVED]) &&
    some(STATUS.AUTO_RECEIVED) &&
    !some(STATUS.DISPATCHED)
  ) {
    return STATUS.AUTO_RECEIVED;
  }

  if (
    all([STATUS.DISPATCHED, STATUS.RECEIVED, STATUS.AUTO_RECEIVED, STATUS.APPROVED]) &&
    some(STATUS.DISPATCHED)
  ) {
    return STATUS.DISPATCHED;
  }

  if (
    all([STATUS.READY, STATUS.DISPATCHED, STATUS.RECEIVED, STATUS.AUTO_RECEIVED, STATUS.APPROVED]) &&
    some(STATUS.READY)
  ) {
    return STATUS.READY;
  }

  if (some(STATUS.NOT_PICKED)) {
    return STATUS.NOT_PICKED;
  }

  return STATUS.REQUESTED;
}

function buildDeliveryGroupsFromOrders(orders) {
  const grouped = {};

  (orders || []).forEach(order => {
    const key = safeText(order.beleveringId) || `${safeText(order.techniekerCode)}|${safeText(order.deliverySortKey)}`;

    if (!grouped[key]) {
      grouped[key] = {
        key,
        beleveringId: safeText(order.beleveringId),
        techniekerCode: safeText(order.techniekerCode),
        techniekerNaam: safeText(order.techniekerNaam),
        beleveringDatumIso: safeText(order.beleveringDatumIso),
        klaarTegenLabel: safeText(order.klaarTegenLabel),
        klaarTegenSort: safeText(order.deliverySortKey),
        orderIds: [],
        lineItems: [],
        opmerking: safeText(order.opmerking)
      };
    }

    grouped[key].orderIds.push(order.bestellingId);
    grouped[key].lineItems.push(order);

    if (!grouped[key].opmerking && order.opmerking) {
      grouped[key].opmerking = safeText(order.opmerking);
    }
  });

  return Object.keys(grouped).map(key => {
    const group = grouped[key];
    group.lineItems.sort((a, b) =>
      String(a.artikelOmschrijving || '').localeCompare(String(b.artikelOmschrijving || '')) ||
      String(a.artikelCode || '').localeCompare(String(b.artikelCode || ''))
    );
    group.groupStatus = calculateGroupStatus(group.lineItems);
    group.aantalLijnen = group.lineItems.length;
    group.orderIds = [...new Set(group.orderIds)];
    group.isEditable = group.groupStatus === STATUS.REQUESTED;
    group.isFrozen = group.groupStatus !== STATUS.REQUESTED;
    group.isClosed = group.groupStatus === STATUS.RECEIVED || group.groupStatus === STATUS.APPROVED;
    return group;
  });
}

function buildUpcomingDeliveryGroupsWithinDays(orders, daysAhead) {
  const todayIso = Utilities.formatDate(nowDate(), getAppTimeZone(), 'yyyy-MM-dd');
  const maxDate = new Date();
  maxDate.setDate(maxDate.getDate() + Number(daysAhead || 0));
  const maxIso = Utilities.formatDate(maxDate, getAppTimeZone(), 'yyyy-MM-dd');

  return buildDeliveryGroupsFromOrders(orders)
    .filter(group => !!group.beleveringDatumIso)
    .filter(group => group.beleveringDatumIso >= todayIso && group.beleveringDatumIso <= maxIso)
    .sort((a, b) => String(a.klaarTegenSort || '').localeCompare(String(b.klaarTegenSort || '')));
}

function buildManagerApprovalGroups(allOrders) {
  return buildDeliveryGroupsFromOrders(
    (allOrders || []).filter(order => safeText(order.managerGoedkeuringStatus) === MANAGER_STATUS.PENDING)
  ).sort((a, b) => String(a.klaarTegenSort || '').localeCompare(String(b.klaarTegenSort || '')));
}

function buildArticlePopularityMap(orders) {
  const map = {};

  (orders || []).forEach(order => {
    const code = safeText(order.artikelCode);
    if (!code) return;

    map[code] = (map[code] || 0) + safeNumber(order.aantalDozen, 0);
  });

  return map;
}

function rebuildTotalsSheet() {
  const orders = getAllGrabbelOrders().map(order => ({
    artikelCode: safeText(order.artikelCode),
    artikelOmschrijving: safeText(order.artikelOmschrijving),
    eenheid: safeText(order.eenheid),
    aantalDozen: safeNumber(order.aantalDozen, 0),
    totaalStuks: safeNumber(order.totaalStuks, 0),
    timestamp: safeText(order.timestampRaw)
  }));

  const grouped = {};

  orders.forEach(order => {
    if (!order.artikelCode) return;

    if (!grouped[order.artikelCode]) {
      grouped[order.artikelCode] = {
        artikelCode: order.artikelCode,
        artikelOmschrijving: order.artikelOmschrijving,
        eenheid: order.eenheid,
        totaalDozen: 0,
        totaalStuks: 0,
        bestellijnen: 0,
        laatsteBesteldatum: order.timestamp || ''
      };
    }

    grouped[order.artikelCode].totaalDozen += order.aantalDozen;
    grouped[order.artikelCode].totaalStuks += order.totaalStuks;
    grouped[order.artikelCode].bestellijnen += 1;

    if (String(order.timestamp || '') > String(grouped[order.artikelCode].laatsteBesteldatum || '')) {
      grouped[order.artikelCode].laatsteBesteldatum = order.timestamp || '';
    }
  });

  const rows = Object.keys(grouped)
    .map(key => grouped[key])
    .sort((a, b) => b.totaalDozen - a.totaalDozen || String(a.artikelOmschrijving || '').localeCompare(String(b.artikelOmschrijving || '')))
    .map(item => [
      item.artikelCode,
      item.artikelOmschrijving,
      item.eenheid,
      item.totaalDozen,
      item.totaalStuks,
      item.bestellijnen,
      item.laatsteBesteldatum
    ]);

  writeFullTable(
    TABS.TOTALS,
    ['ArtikelCode', 'ArtikelOmschrijving', 'Eenheid', 'TotaalDozen', 'TotaalStuks', 'Bestellijnen', 'LaatsteBesteldatum'],
    rows
  );

  return {
    success: true,
    lines: rows.length,
    message: 'TotaalMaterialen vernieuwd.'
  };
}