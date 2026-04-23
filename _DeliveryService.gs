/* =========================================================
   30_DeliveryService.gs
   Refactor: delivery core service
   Doel:
   - levermomenten centraal lezen
   - cutoff- en timinglogica op één plaats
   - ordergroepen per belevering opbouwen
   - read models voor magazijn / technieker / manager
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getDeliveryTab_() {
  return TABS.DELIVERIES || 'Beleveringen';
}

function getOrderTabForDelivery_() {
  return TABS.ORDERS || 'Orders';
}

function getDeliveryDefaultCutoffHours_() {
  return safeNumber((APP_CONFIG && APP_CONFIG.DELIVERY_CUTOFF_HOURS) || 24, 24);
}

function getDeliveryStatusPlanned_() {
  if (typeof DELIVERY_STATUS !== 'undefined' && DELIVERY_STATUS && DELIVERY_STATUS.PLANNED) {
    return DELIVERY_STATUS.PLANNED;
  }
  return 'Gepland';
}

function getDeliveryStatusOpen_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.OPEN) {
    return ORDER_STATUS.OPEN;
  }
  return 'Open';
}

function getDeliveryStatusPrepared_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.PREPARED) {
    return ORDER_STATUS.PREPARED;
  }
  return 'Klaargezet';
}

function getDeliveryStatusGiven_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.GIVEN) {
    return ORDER_STATUS.GIVEN;
  }
  return 'Meegegeven';
}

function getDeliveryStatusNotPickedUp_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.NOT_PICKED_UP) {
    return ORDER_STATUS.NOT_PICKED_UP;
  }
  return 'Niet afgehaald';
}

function getDeliveryStatusReceived_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.RECEIVED) {
    return ORDER_STATUS.RECEIVED;
  }
  return 'Ontvangen';
}

function getDeliveryStatusClosed_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.CLOSED) {
    return ORDER_STATUS.CLOSED;
  }
  return 'Gesloten';
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapDeliverySlot(row) {
  var dateIso = safeText(
    row.DatumIso ||
    row.DeliveryDateIso ||
    row.DocumentDatumIso ||
    row.Datum ||
    row.Date
  );

  var startTime = safeText(row.StartTijd || row.StartTime || '');
  var endTime = safeText(row.EindTijd || row.EndTime || '');
  var techCode = safeText(row.TechniekerCode || row.TechnicianCode || row.TechCode || '');
  var slotCode = safeText(row.BeleveringID || row.DeliveryID || row.DeliverySlotID || row.ID);

  return {
    deliveryId: slotCode,
    deliveryCode: slotCode,
    techniekerCode: techCode,
    techniekerNaam: safeText(row.TechniekerNaam || row.TechnicianName || ''),
    datum: safeText(row.Datum || row.Date || dateIso),
    datumIso: dateIso,
    startTijd: startTime,
    eindTijd: endTime,
    title: safeText(row.Titel || row.Title || ''),
    status: safeText(row.Status || getDeliveryStatusPlanned_()),
    locatie: safeText(row.Locatie || row.Location || ''),
    opmerking: safeText(row.Opmerking || row.Remark || ''),
  };
}

function mapDeliveryOrderRow_(row) {
  return {
    orderId: safeText(row.OrderID || row.OrderId || row.ID),
    deliveryId: safeText(row.DeliveryID || row.BeleveringID || row.DeliverySlotID || ''),
    techniekerCode: safeText(row.TechniekerCode || row.TechnicianCode || row.TechCode || ''),
    status: safeText(row.Status),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    gevraagdAantal: safeNumber(row.GevraagdAantal || row.Gevraagd || row.RequestedQty, 0),
    voorzienAantal: safeNumber(row.VoorzienAantal || row.Voorzien || row.PreparedQty, 0),
    ontvangenAantal: safeNumber(row.OntvangenAantal || row.ReceivedQty, 0),
    redenDelta: safeText(row.RedenDelta || row.DeltaReason || ''),
    opmerking: safeText(row.Opmerking || ''),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum || row.Datum || ''),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllDeliverySlots() {
  return readObjectsSafe(getDeliveryTab_())
    .map(mapDeliverySlot)
    .sort(function (a, b) {
      return (
        safeText(a.datumIso).localeCompare(safeText(b.datumIso)) ||
        safeText(a.startTijd).localeCompare(safeText(b.startTijd)) ||
        safeText(a.deliveryId).localeCompare(safeText(b.deliveryId))
      );
    });
}

function getDeliverySlotById(deliveryId) {
  var id = safeText(deliveryId);
  if (!id) return null;

  return getAllDeliverySlots().find(function (item) {
    return safeText(item.deliveryId) === id;
  }) || null;
}

function getAllOrdersForDeliveryLayer_() {
  return readObjectsSafe(getOrderTabForDelivery_()).map(mapDeliveryOrderRow_);
}

function getOrdersByDeliveryId(deliveryId) {
  var id = safeText(deliveryId);
  return getAllOrdersForDeliveryLayer_().filter(function (item) {
    return safeText(item.deliveryId) === id;
  });
}

function getOrdersByTechnicianCodeForDelivery_(techniekerCode) {
  var code = safeText(techniekerCode);
  return getAllOrdersForDeliveryLayer_().filter(function (item) {
    return safeText(item.techniekerCode) === code;
  });
}

/* ---------------------------------------------------------
   Time / cutoff helpers
   --------------------------------------------------------- */

function buildDeliveryDateTimeRaw_(datumIso, tijd) {
  var datePart = safeText(datumIso);
  var timePart = safeText(tijd || '00:00');
  if (!datePart) return '';
  if (timePart.length === 5) {
    timePart += ':00';
  }
  return datePart + 'T' + timePart;
}

function getNowDeliveryRaw_() {
  return Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function getDeliveryCutoffRaw_(delivery) {
  var deliveryRaw = buildDeliveryDateTimeRaw_(delivery.datumIso, delivery.startTijd);
  if (!deliveryRaw) return '';

  var dt = new Date(deliveryRaw);
  dt.setHours(dt.getHours() - getDeliveryDefaultCutoffHours_());
  return Utilities.formatDate(dt, TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function isDeliveryCutoffPassed(delivery) {
  var cutoffRaw = getDeliveryCutoffRaw_(delivery || {});
  if (!cutoffRaw) return false;
  return cutoffRaw <= getNowDeliveryRaw_();
}

function isDeliveryPast(delivery) {
  var endRaw = buildDeliveryDateTimeRaw_(delivery.datumIso, delivery.eindTijd || delivery.startTijd);
  if (!endRaw) return false;
  return endRaw < getNowDeliveryRaw_();
}

function enrichDeliveryWithTiming(delivery) {
  var item = Object.assign({}, delivery || {});
  var cutoffRaw = getDeliveryCutoffRaw_(item);

  item.cutoffRaw = cutoffRaw;
  item.cutoffPassed = isDeliveryCutoffPassed(item);
  item.isPast = isDeliveryPast(item);
  item.isEditable = !item.cutoffPassed && !item.isPast;
  item.windowLabel = [safeText(item.startTijd), safeText(item.eindTijd)].filter(Boolean).join(' - ');

  return item;
}

/* ---------------------------------------------------------
   Group status helpers
   --------------------------------------------------------- */

function calculateDeliveryGroupStatus(orderLines) {
  var rows = Array.isArray(orderLines) ? orderLines : [];
  if (!rows.length) {
    return getDeliveryStatusOpen_();
  }

  var statuses = rows.map(function (row) {
    return safeText(row.status);
  });

  if (statuses.every(function (status) { return status === getDeliveryStatusReceived_(); })) {
    return getDeliveryStatusReceived_();
  }

  if (statuses.every(function (status) { return status === getDeliveryStatusClosed_(); })) {
    return getDeliveryStatusClosed_();
  }

  if (statuses.every(function (status) { return status === getDeliveryStatusGiven_(); })) {
    return getDeliveryStatusGiven_();
  }

  if (statuses.some(function (status) { return status === getDeliveryStatusNotPickedUp_(); })) {
    return getDeliveryStatusNotPickedUp_();
  }

  if (statuses.some(function (status) { return status === getDeliveryStatusPrepared_(); })) {
    return getDeliveryStatusPrepared_();
  }

  return getDeliveryStatusOpen_();
}

function hasOrdersForDelivery(orderLines) {
  return Array.isArray(orderLines) && orderLines.length > 0;
}

function buildDeliveryGroup(delivery, orderLines) {
  var slot = enrichDeliveryWithTiming(delivery || {});
  var lines = Array.isArray(orderLines) ? orderLines.slice() : [];
  var totalRequested = lines.reduce(function (sum, line) {
    return sum + safeNumber(line.gevraagdAantal, 0);
  }, 0);
  var totalPrepared = lines.reduce(function (sum, line) {
    return sum + safeNumber(line.voorzienAantal, 0);
  }, 0);

  return Object.assign({}, slot, {
    lines: lines,
    lineCount: lines.length,
    hasOrders: hasOrdersForDelivery(lines),
    totalRequested: totalRequested,
    totalPrepared: totalPrepared,
    status: calculateDeliveryGroupStatus(lines),
    hasDelta: lines.some(function (line) {
      return safeNumber(line.gevraagdAantal, 0) !== safeNumber(line.voorzienAantal, 0);
    }),
  });
}

/* ---------------------------------------------------------
   Projections
   --------------------------------------------------------- */

function buildDeliveryGroupsFromOrders(deliveries, orderRows) {
  var lineMap = {};

  (orderRows || []).forEach(function (row) {
    var deliveryId = safeText(row.deliveryId);
    if (!deliveryId) return;
    if (!lineMap[deliveryId]) lineMap[deliveryId] = [];
    lineMap[deliveryId].push(row);
  });

  return (deliveries || []).map(function (delivery) {
    var lines = lineMap[safeText(delivery.deliveryId)] || [];
    return buildDeliveryGroup(delivery, lines);
  });
}

function getDeliveryGroups() {
  return buildDeliveryGroupsFromOrders(
    getAllDeliverySlots(),
    getAllOrdersForDeliveryLayer_()
  );
}

function getDeliveryGroupById(deliveryId) {
  var slot = getDeliverySlotById(deliveryId);
  if (!slot) return null;
  return buildDeliveryGroup(slot, getOrdersByDeliveryId(deliveryId));
}

function getUpcomingDeliveriesForTechnician(techniekerCode, daysAhead) {
  var code = safeText(techniekerCode);
  var maxDays = safeNumber(daysAhead, 7);
  var now = new Date();
  var end = new Date(now);
  end.setDate(end.getDate() + maxDays);

  return getAllDeliverySlots()
    .filter(function (slot) {
      if (safeText(slot.techniekerCode) !== code) return false;
      var raw = buildDeliveryDateTimeRaw_(slot.datumIso, slot.startTijd);
      if (!raw) return false;
      var dt = new Date(raw);
      return dt >= now && dt <= end;
    })
    .map(enrichDeliveryWithTiming);
}

function getUpcomingDeliveryGroups(daysAhead) {
  var maxDays = safeNumber(daysAhead, 7);
  var now = new Date();
  var end = new Date(now);
  end.setDate(end.getDate() + maxDays);

  return getDeliveryGroups().filter(function (group) {
    var raw = buildDeliveryDateTimeRaw_(group.datumIso, group.startTijd);
    if (!raw) return false;
    var dt = new Date(raw);
    return dt >= now && dt <= end;
  });
}

function hasExistingOrderForDelivery(techniekerCode, deliveryId) {
  var techCode = safeText(techniekerCode);
  var id = safeText(deliveryId);

  return getAllOrdersForDeliveryLayer_().some(function (row) {
    return safeText(row.techniekerCode) === techCode &&
      safeText(row.deliveryId) === id;
  });
}

/* ---------------------------------------------------------
   Manager / warehouse summary
   --------------------------------------------------------- */

function buildDeliverySummary(groups) {
  var rows = Array.isArray(groups) ? groups : [];

  return {
    totaal: rows.length,
    open: rows.filter(function (x) { return safeText(x.status) === getDeliveryStatusOpen_(); }).length,
    klaargezet: rows.filter(function (x) { return safeText(x.status) === getDeliveryStatusPrepared_(); }).length,
    meegegeven: rows.filter(function (x) { return safeText(x.status) === getDeliveryStatusGiven_(); }).length,
    nietAfgehaald: rows.filter(function (x) { return safeText(x.status) === getDeliveryStatusNotPickedUp_(); }).length,
    ontvangen: rows.filter(function (x) { return safeText(x.status) === getDeliveryStatusReceived_(); }).length,
    gesloten: rows.filter(function (x) { return safeText(x.status) === getDeliveryStatusClosed_(); }).length,
    vandaag: rows.filter(function (x) {
      return safeText(x.datumIso) === Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
    }).length,
  };
}

function getDeliveryGroupsData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  if (!roleAllowed(actor, [ROLE.WAREHOUSE, ROLE.MANAGER, ROLE.TECHNICIAN])) {
    throw new Error('Geen rechten om leveringen te bekijken.');
  }

  var daysAhead = safeNumber(payload.daysAhead, 7);
  var groups;

  if (roleAllowed(actor, [ROLE.TECHNICIAN])) {
    var techCode = safeText(
      payload.techniekerCode ||
      payload.technicianCode ||
      actor.techniekerCode ||
      actor.technicianCode ||
      actor.code
    );

    groups = buildDeliveryGroupsFromOrders(
      getUpcomingDeliveriesForTechnician(techCode, daysAhead),
      getOrdersByTechnicianCodeForDelivery_(techCode)
    );
  } else {
    groups = getUpcomingDeliveryGroups(daysAhead);
  }

  return {
    items: groups,
    deliveryGroups: groups,
    summary: buildDeliverySummary(groups),
  };
}
