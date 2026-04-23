/* =========================================================
   31_GrabbelOrderService.gs
   Refactor: grabbel order core service
   Doel:
   - grabbelbestellingen per belevering en technieker
   - lijnlaag in één ordersheet
   - technieker submit/update
   - magazijn klaargezet / meegegeven / niet afgehaald
   - technieker ontvangst
   - managerafsluiting
   ========================================================= */

/* ---------------------------------------------------------
   Tabs / fallbacks
   --------------------------------------------------------- */

function getGrabbelOrderTab_() {
  return TABS.ORDERS || 'Orders';
}

function getGrabbelOrderStatusOpen_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.OPEN) {
    return ORDER_STATUS.OPEN;
  }
  return 'Open';
}

function getGrabbelOrderStatusPrepared_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.PREPARED) {
    return ORDER_STATUS.PREPARED;
  }
  return 'Klaargezet';
}

function getGrabbelOrderStatusGiven_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.GIVEN) {
    return ORDER_STATUS.GIVEN;
  }
  return 'Meegegeven';
}

function getGrabbelOrderStatusNotPickedUp_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.NOT_PICKED_UP) {
    return ORDER_STATUS.NOT_PICKED_UP;
  }
  return 'Niet afgehaald';
}

function getGrabbelOrderStatusReceived_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.RECEIVED) {
    return ORDER_STATUS.RECEIVED;
  }
  return 'Ontvangen';
}

function getGrabbelOrderStatusClosed_() {
  if (typeof ORDER_STATUS !== 'undefined' && ORDER_STATUS && ORDER_STATUS.CLOSED) {
    return ORDER_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function makeGrabbelOrderLineId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'ORD-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

/* ---------------------------------------------------------
   Mapping
   --------------------------------------------------------- */

function mapGrabbelOrderRow(row) {
  var gevraagd = safeNumber(row.GevraagdAantal || row.Gevraagd || row.RequestedQty, 0);
  var voorzien = safeNumber(row.VoorzienAantal || row.Voorzien || row.PreparedQty, gevraagd);

  return {
    orderId: safeText(row.OrderID || row.OrderId || row.ID),
    deliveryId: safeText(row.DeliveryID || row.BeleveringID || row.DeliverySlotID),
    techniekerCode: safeText(row.TechniekerCode || row.TechnicianCode || row.TechCode),
    techniekerNaam: safeText(row.TechniekerNaam || row.TechnicianName),
    artikelCode: safeText(row.ArtikelCode || row.ArtikelNr),
    artikelOmschrijving: safeText(row.ArtikelOmschrijving || row.Artikel),
    typeMateriaal: safeText(row.TypeMateriaal || determineMaterialTypeFromArticle(safeText(row.ArtikelCode || row.ArtikelNr))),
    eenheid: safeText(row.Eenheid || row.Unit),
    gevraagdAantal: gevraagd,
    voorzienAantal: voorzien,
    ontvangenAantal: safeNumber(row.OntvangenAantal || row.ReceivedQty, 0),
    deltaAantal: safeNumber(row.DeltaAantal, voorzien - gevraagd),
    redenDelta: safeText(row.RedenDelta || row.DeltaReason),
    opmerking: safeText(row.Opmerking),
    status: safeText(row.Status || getGrabbelOrderStatusOpen_()),
    actor: safeText(row.Actor),
    documentDatum: safeText(row.DocumentDatum),
    documentDatumIso: safeText(row.DocumentDatumIso || row.DocumentDatum),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    klaargezetOp: safeText(row.KlaargezetOp),
    meegegevenOp: safeText(row.MeegegevenOp),
    ontvangenOp: safeText(row.OntvangenOp),
    geslotenOp: safeText(row.GeslotenOp),
  };
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllGrabbelOrders() {
  return readObjectsSafe(getGrabbelOrderTab_())
    .map(mapGrabbelOrderRow)
    .sort(function (a, b) {
      return (
        safeText(b.documentDatumIso).localeCompare(safeText(a.documentDatumIso)) ||
        safeText(a.deliveryId).localeCompare(safeText(b.deliveryId)) ||
        safeText(a.artikelOmschrijving).localeCompare(safeText(b.artikelOmschrijving))
      );
    });
}

function getGrabbelOrdersByDeliveryId(deliveryId) {
  var id = safeText(deliveryId);
  return getAllGrabbelOrders().filter(function (item) {
    return safeText(item.deliveryId) === id;
  });
}

function getGrabbelOrdersByTechnicianCode(techniekerCode) {
  var code = safeText(techniekerCode);
  return getAllGrabbelOrders().filter(function (item) {
    return safeText(item.techniekerCode) === code;
  });
}

function buildGrabbelOrderGroups(orderRows) {
  var grouped = {};

  (orderRows || []).forEach(function (row) {
    var key = safeText(row.deliveryId) + '|' + safeText(row.techniekerCode);
    if (!grouped[key]) {
      grouped[key] = {
        deliveryId: safeText(row.deliveryId),
        techniekerCode: safeText(row.techniekerCode),
        techniekerNaam: safeText(row.techniekerNaam),
        lines: [],
      };
    }
    grouped[key].lines.push(row);
  });

  return Object.keys(grouped)
    .map(function (key) {
      var group = grouped[key];
      var status = calculateDeliveryGroupStatus(group.lines);
      var totalRequested = group.lines.reduce(function (sum, line) {
        return sum + safeNumber(line.gevraagdAantal, 0);
      }, 0);
      var totalPrepared = group.lines.reduce(function (sum, line) {
        return sum + safeNumber(line.voorzienAantal, 0);
      }, 0);

      return Object.assign({}, group, {
        status: status,
        totalRequested: totalRequested,
        totalPrepared: totalPrepared,
        lineCount: group.lines.length,
        hasDelta: group.lines.some(function (line) {
          return safeNumber(line.gevraagdAantal, 0) !== safeNumber(line.voorzienAantal, 0);
        }),
      });
    })
    .sort(function (a, b) {
      return (
        safeText(a.deliveryId).localeCompare(safeText(b.deliveryId)) ||
        safeText(a.techniekerNaam).localeCompare(safeText(b.techniekerNaam))
      );
    });
}

function getGrabbelOrderGroup(deliveryId, techniekerCode) {
  var did = safeText(deliveryId);
  var tcode = safeText(techniekerCode);

  return buildGrabbelOrderGroups(getAllGrabbelOrders()).find(function (item) {
    return safeText(item.deliveryId) === did &&
      safeText(item.techniekerCode) === tcode;
  }) || null;
}

function hasGrabbelOrderGroup(deliveryId, techniekerCode) {
  return !!getGrabbelOrderGroup(deliveryId, techniekerCode);
}

/* ---------------------------------------------------------
   Normalization / validation
   --------------------------------------------------------- */

function normalizeGrabbelOrderLines_(lines) {
  return (Array.isArray(lines) ? lines : []).map(function (line) {
    var gevraagd = safeNumber(line.gevraagdAantal || line.gevraagd || line.requestedQty, 0);
    var voorzien =
      line.voorzienAantal === undefined || line.voorzienAantal === null || line.voorzienAantal === ''
        ? gevraagd
        : safeNumber(line.voorzienAantal || line.voorzien || line.preparedQty, 0);

    return {
      artikelCode: safeText(line.artikelCode || line.artikelNr),
      artikelOmschrijving: safeText(line.artikelOmschrijving || line.artikel),
      typeMateriaal: safeText(line.typeMateriaal || determineMaterialTypeFromArticle(safeText(line.artikelCode || line.artikelNr))),
      eenheid: safeText(line.eenheid || line.unit || 'Stuk'),
      gevraagdAantal: gevraagd,
      voorzienAantal: voorzien,
      ontvangenAantal: safeNumber(line.ontvangenAantal || line.receivedQty, 0),
      deltaAantal: safeNumber(line.deltaAantal, voorzien - gevraagd),
      redenDelta: safeText(line.redenDelta || line.deltaReason),
      opmerking: safeText(line.opmerking),
    };
  });
}

function validateGrabbelOrderLinesForTechnician_(lines) {
  if (!lines.length) {
    throw new Error('Geen bestellijnen ontvangen.');
  }

  lines.forEach(function (line, index) {
    var rowNr = index + 1;
    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.gevraagdAantal, 0) <= 0) {
      throw new Error('Gevraagd aantal moet groter zijn dan 0 op lijn ' + rowNr + '.');
    }
  });

  return true;
}

function validateGrabbelOrderLinesForWarehouse_(lines) {
  if (!lines.length) {
    throw new Error('Geen bestellijnen ontvangen.');
  }

  lines.forEach(function (line, index) {
    var rowNr = index + 1;
    if (!safeText(line.artikelCode)) {
      throw new Error('Artikelcode ontbreekt op lijn ' + rowNr + '.');
    }
    if (safeNumber(line.voorzienAantal, 0) < 0) {
      throw new Error('Voorzien aantal mag niet negatief zijn op lijn ' + rowNr + '.');
    }
    if (
      safeNumber(line.voorzienAantal, 0) !== safeNumber(line.gevraagdAantal, 0) &&
      !safeText(line.redenDelta)
    ) {
      throw new Error('Reden delta is verplicht op lijn ' + rowNr + '.');
    }
  });

  return true;
}

function assertDeliveryEditableForOrder_(delivery, techniekerCode, sessionId) {
  if (!delivery) {
    throw new Error('Belevering niet gevonden.');
  }

  var access = assertTechnicianAccessToRef(sessionId, techniekerCode);
  if (safeText(access.technician.code || access.technician.ref) !== safeText(techniekerCode)) {
    throw new Error('Geen toegang tot deze technieker.');
  }

  var enriched = enrichDeliveryWithTiming(delivery);
  if (enriched.cutoffPassed) {
    throw new Error('Cutoff voor deze belevering is verstreken.');
  }
  if (enriched.isPast) {
    throw new Error('Deze belevering ligt in het verleden.');
  }

  return enriched;
}

/* ---------------------------------------------------------
   Internal write helper
   --------------------------------------------------------- */

function replaceGrabbelOrderGroup_(deliveryId, techniekerCode, lines, sharedFields) {
  var table = getAllValues(getGrabbelOrderTab_());
  var headerRow = table.length ? table[0] : null;
  var currentRows = readObjectsSafe(getGrabbelOrderTab_());

  var did = safeText(deliveryId);
  var tcode = safeText(techniekerCode);

  var kept = currentRows.filter(function (row) {
    var mapped = mapGrabbelOrderRow(row);
    return !(
      safeText(mapped.deliveryId) === did &&
      safeText(mapped.techniekerCode) === tcode
    );
  });

  var newRows = lines.map(function (line) {
    return {
      OrderID: makeGrabbelOrderLineId_(),
      DeliveryID: did,
      TechniekerCode: tcode,
      TechniekerNaam: safeText(sharedFields.techniekerNaam),
      ArtikelCode: line.artikelCode,
      ArtikelOmschrijving: line.artikelOmschrijving,
      TypeMateriaal: line.typeMateriaal,
      Eenheid: line.eenheid,
      GevraagdAantal: line.gevraagdAantal,
      VoorzienAantal: line.voorzienAantal,
      OntvangenAantal: line.ontvangenAantal,
      DeltaAantal: line.deltaAantal,
      RedenDelta: line.redenDelta,
      Opmerking: line.opmerking,
      Status: safeText(sharedFields.status || getGrabbelOrderStatusOpen_()),
      Actor: safeText(sharedFields.actor),
      DocumentDatum: safeText(sharedFields.documentDatum),
      DocumentDatumIso: safeText(sharedFields.documentDatumIso || sharedFields.documentDatum),
      AangemaaktOp: safeText(sharedFields.aangemaaktOp),
      AangemaaktOpRaw: safeText(sharedFields.aangemaaktOpRaw),
      KlaargezetOp: safeText(sharedFields.klaargezetOp),
      MeegegevenOp: safeText(sharedFields.meegegevenOp),
      OntvangenOp: safeText(sharedFields.ontvangenOp),
      GeslotenOp: safeText(sharedFields.geslotenOp),
    };
  });

  var finalObjects = kept.concat(newRows);

  if (!headerRow && finalObjects.length) {
    appendObjects(getGrabbelOrderTab_(), newRows);
  } else if (headerRow) {
    writeFullTable(
      getGrabbelOrderTab_(),
      headerRow,
      finalObjects.map(function (obj) {
        return buildRowFromHeaders(headerRow, obj);
      })
    );
  } else if (newRows.length) {
    appendObjects(getGrabbelOrderTab_(), newRows);
  }

  return getGrabbelOrderGroup(did, tcode);
}

/* ---------------------------------------------------------
   Technician flow
   --------------------------------------------------------- */

function submitOrder(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var techniekerCode = safeText(
    payload.techniekerCode ||
    payload.technicianCode ||
    payload.techCode
  );
  var deliveryId = safeText(payload.deliveryId);
  if (!techniekerCode) throw new Error('TechniekerCode ontbreekt.');
  if (!deliveryId) throw new Error('DeliveryId ontbreekt.');

  var actorAccess = assertTechnicianAccessToRef(sessionId, techniekerCode);
  var delivery = assertDeliveryEditableForOrder_(
    getDeliverySlotById(deliveryId),
    techniekerCode,
    sessionId
  );

  var lines = normalizeGrabbelOrderLines_(payload.lines);
  validateGrabbelOrderLinesForTechnician_(lines);

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var group = replaceGrabbelOrderGroup_(
    deliveryId,
    techniekerCode,
    lines.map(function (line) {
      return Object.assign({}, line, {
        voorzienAantal: line.gevraagdAantal,
        ontvangenAantal: 0,
        deltaAantal: 0,
        redenDelta: '',
      });
    }),
    {
      techniekerNaam: safeText(delivery.techniekerNaam || actorAccess.technician.naam),
      status: getGrabbelOrderStatusOpen_(),
      actor: safeText(actorAccess.user.naam || actorAccess.user.email),
      documentDatum: safeText(delivery.datumIso),
      documentDatumIso: safeText(delivery.datumIso),
      aangemaaktOp: toDisplayDateTime(nowRaw),
      aangemaaktOpRaw: nowRaw,
      klaargezetOp: '',
      meegegevenOp: '',
      ontvangenOp: '',
      geslotenOp: '',
    }
  );

  writeAudit({
    actie: 'SUBMIT_GRABBEL_ORDER',
    actor: actorAccess.user,
    documentType: 'GrabbelOrderGroup',
    documentId: deliveryId + ':' + techniekerCode,
    details: {
      lineCount: lines.length,
      deliveryId: deliveryId,
      techniekerCode: techniekerCode,
    },
  });

  if (typeof pushWarehouseNotification === 'function') {
    pushWarehouseNotification(
      'GrabbelOrder',
      'Nieuwe grabbelbestelling',
      'Er werd een nieuwe grabbelbestelling ingediend.',
      'GrabbelOrder',
      deliveryId + ':' + techniekerCode,
      'WAREHOUSE'
    );
  }

  return group;
}

function updateTechnicianOrderGroup(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var techniekerCode = safeText(
    payload.techniekerCode ||
    payload.technicianCode ||
    payload.techCode
  );
  var deliveryId = safeText(payload.deliveryId);
  if (!techniekerCode) throw new Error('TechniekerCode ontbreekt.');
  if (!deliveryId) throw new Error('DeliveryId ontbreekt.');

  var access = assertTechnicianAccessToRef(sessionId, techniekerCode);
  var existing = getGrabbelOrderGroup(deliveryId, techniekerCode);
  if (!existing) throw new Error('Bestelgroep niet gevonden.');

  var status = safeText(existing.status);
  if ([getGrabbelOrderStatusPrepared_(), getGrabbelOrderStatusGiven_(), getGrabbelOrderStatusReceived_(), getGrabbelOrderStatusClosed_()].indexOf(status) >= 0) {
    throw new Error('Bestelgroep is niet meer wijzigbaar.');
  }

  assertDeliveryEditableForOrder_(getDeliverySlotById(deliveryId), techniekerCode, sessionId);

  var lines = normalizeGrabbelOrderLines_(payload.lines);
  validateGrabbelOrderLinesForTechnician_(lines);

  var preservedFirst = existing.lines[0] || {};
  var group = replaceGrabbelOrderGroup_(
    deliveryId,
    techniekerCode,
    lines.map(function (line) {
      return Object.assign({}, line, {
        voorzienAantal: line.gevraagdAantal,
        ontvangenAantal: 0,
        deltaAantal: 0,
        redenDelta: '',
      });
    }),
    {
      techniekerNaam: safeText(existing.techniekerNaam || access.technician.naam),
      status: getGrabbelOrderStatusOpen_(),
      actor: safeText(access.user.naam || access.user.email),
      documentDatum: safeText(preservedFirst.documentDatum),
      documentDatumIso: safeText(preservedFirst.documentDatumIso),
      aangemaaktOp: safeText(preservedFirst.aangemaaktOp),
      aangemaaktOpRaw: safeText(preservedFirst.aangemaaktOpRaw),
      klaargezetOp: '',
      meegegevenOp: '',
      ontvangenOp: '',
      geslotenOp: '',
    }
  );

  writeAudit({
    actie: 'UPDATE_GRABBEL_ORDER',
    actor: access.user,
    documentType: 'GrabbelOrderGroup',
    documentId: deliveryId + ':' + techniekerCode,
    details: {
      lineCount: lines.length,
    },
  });

  return group;
}

/* ---------------------------------------------------------
   Warehouse flow
   --------------------------------------------------------- */

function prepareOrderGroup(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertWarehouseAccess(sessionId);
  var deliveryId = safeText(payload.deliveryId);
  var techniekerCode = safeText(payload.techniekerCode || payload.technicianCode || payload.techCode);
  if (!deliveryId) throw new Error('DeliveryId ontbreekt.');
  if (!techniekerCode) throw new Error('TechniekerCode ontbreekt.');

  var existing = getGrabbelOrderGroup(deliveryId, techniekerCode);
  if (!existing) throw new Error('Bestelgroep niet gevonden.');

  var status = safeText(existing.status);
  if ([getGrabbelOrderStatusGiven_(), getGrabbelOrderStatusReceived_(), getGrabbelOrderStatusClosed_()].indexOf(status) >= 0) {
    throw new Error('Bestelgroep is niet meer wijzigbaar voor magazijn.');
  }

  var lines = normalizeGrabbelOrderLines_(payload.lines);
  validateGrabbelOrderLinesForWarehouse_(lines);

  var first = existing.lines[0] || {};
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  var group = replaceGrabbelOrderGroup_(
    deliveryId,
    techniekerCode,
    lines.map(function (line) {
      return Object.assign({}, line, {
        ontvangenAantal: 0,
        deltaAantal: safeNumber(line.voorzienAantal, 0) - safeNumber(line.gevraagdAantal, 0),
      });
    }),
    {
      techniekerNaam: safeText(existing.techniekerNaam),
      status: getGrabbelOrderStatusPrepared_(),
      actor: safeText(actor.naam || actor.email),
      documentDatum: safeText(first.documentDatum),
      documentDatumIso: safeText(first.documentDatumIso),
      aangemaaktOp: safeText(first.aangemaaktOp),
      aangemaaktOpRaw: safeText(first.aangemaaktOpRaw),
      klaargezetOp: toDisplayDateTime(nowRaw),
      meegegevenOp: '',
      ontvangenOp: '',
      geslotenOp: '',
    }
  );

  writeAudit({
    actie: 'PREPARE_GRABBEL_ORDER',
    actor: actor,
    documentType: 'GrabbelOrderGroup',
    documentId: deliveryId + ':' + techniekerCode,
    details: {
      lineCount: lines.length,
    },
  });

  if (typeof pushTechnicianNotification === 'function') {
    pushTechnicianNotification(
      'GrabbelOrder',
      'Bestelling klaargezet',
      'Je grabbelbestelling werd klaargezet.',
      'GrabbelOrder',
      deliveryId + ':' + techniekerCode,
      techniekerCode,
      existing.techniekerNaam
    );
  }

  return group;
}

function markOrderGroupGiven(payload) {
  return updateGrabbelOrderGroupStatus_(payload, getGrabbelOrderStatusGiven_(), 'MeegegevenOp', 'MARK_GRABBEL_ORDER_GIVEN');
}

function markOrderGroupNotPickedUp(payload) {
  return updateGrabbelOrderGroupStatus_(payload, getGrabbelOrderStatusNotPickedUp_(), '', 'MARK_GRABBEL_ORDER_NOT_PICKED_UP');
}

function updateGrabbelOrderGroupStatus_(payload, nextStatus, timestampField, auditAction) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertWarehouseAccess(sessionId);
  var deliveryId = safeText(payload.deliveryId);
  var techniekerCode = safeText(payload.techniekerCode || payload.technicianCode || payload.techCode);
  if (!deliveryId) throw new Error('DeliveryId ontbreekt.');
  if (!techniekerCode) throw new Error('TechniekerCode ontbreekt.');

  var existing = getGrabbelOrderGroup(deliveryId, techniekerCode);
  if (!existing) throw new Error('Bestelgroep niet gevonden.');

  var first = existing.lines[0] || {};
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var stamp = toDisplayDateTime(nowRaw);

  var group = replaceGrabbelOrderGroup_(
    deliveryId,
    techniekerCode,
    existing.lines.map(function (line) {
      return Object.assign({}, line);
    }),
    {
      techniekerNaam: safeText(existing.techniekerNaam),
      status: nextStatus,
      actor: safeText(actor.naam || actor.email),
      documentDatum: safeText(first.documentDatum),
      documentDatumIso: safeText(first.documentDatumIso),
      aangemaaktOp: safeText(first.aangemaaktOp),
      aangemaaktOpRaw: safeText(first.aangemaaktOpRaw),
      klaargezetOp: safeText(first.klaargezetOp),
      meegegevenOp: timestampField === 'MeegegevenOp' ? stamp : safeText(first.meegegevenOp),
      ontvangenOp: safeText(first.ontvangenOp),
      geslotenOp: safeText(first.geslotenOp),
    }
  );

  writeAudit({
    actie: auditAction,
    actor: actor,
    documentType: 'GrabbelOrderGroup',
    documentId: deliveryId + ':' + techniekerCode,
    details: {
      nextStatus: nextStatus,
    },
  });

  return group;
}

/* ---------------------------------------------------------
   Technician receipt / manager close
   --------------------------------------------------------- */

function confirmTechnicianReceipt(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var techniekerCode = safeText(
    payload.techniekerCode ||
    payload.technicianCode ||
    payload.techCode
  );
  var deliveryId = safeText(payload.deliveryId);
  if (!techniekerCode) throw new Error('TechniekerCode ontbreekt.');
  if (!deliveryId) throw new Error('DeliveryId ontbreekt.');

  var access = assertTechnicianAccessToRef(sessionId, techniekerCode);
  var existing = getGrabbelOrderGroup(deliveryId, techniekerCode);
  if (!existing) throw new Error('Bestelgroep niet gevonden.');

  var allowedStatuses = [getGrabbelOrderStatusGiven_(), getGrabbelOrderStatusPrepared_()];
  if (allowedStatuses.indexOf(safeText(existing.status)) < 0) {
    throw new Error('Bestelgroep kan niet ontvangen worden vanuit status "' + safeText(existing.status) + '".');
  }

  var first = existing.lines[0] || {};
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var stamp = toDisplayDateTime(nowRaw);

  var group = replaceGrabbelOrderGroup_(
    deliveryId,
    techniekerCode,
    existing.lines.map(function (line) {
      return Object.assign({}, line, {
        ontvangenAantal: safeNumber(line.voorzienAantal, 0),
      });
    }),
    {
      techniekerNaam: safeText(existing.techniekerNaam || access.technician.naam),
      status: getGrabbelOrderStatusReceived_(),
      actor: safeText(access.user.naam || access.user.email),
      documentDatum: safeText(first.documentDatum),
      documentDatumIso: safeText(first.documentDatumIso),
      aangemaaktOp: safeText(first.aangemaaktOp),
      aangemaaktOpRaw: safeText(first.aangemaaktOpRaw),
      klaargezetOp: safeText(first.klaargezetOp),
      meegegevenOp: safeText(first.meegegevenOp),
      ontvangenOp: stamp,
      geslotenOp: '',
    }
  );

  writeAudit({
    actie: 'CONFIRM_GRABBEL_ORDER_RECEIPT',
    actor: access.user,
    documentType: 'GrabbelOrderGroup',
    documentId: deliveryId + ':' + techniekerCode,
    details: {},
  });

  if (typeof pushManagerNotification === 'function') {
    pushManagerNotification(
      'GrabbelOrder',
      'Bestelling ontvangen door technieker',
      'Een grabbelbestelling werd ontvangen en wacht op afsluiting.',
      'GrabbelOrder',
      deliveryId + ':' + techniekerCode,
      'MANAGER'
    );
  }

  return group;
}

function approveManagerGroup(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertManagerAccess(sessionId);
  var deliveryId = safeText(payload.deliveryId);
  var techniekerCode = safeText(payload.techniekerCode || payload.technicianCode || payload.techCode);
  if (!deliveryId) throw new Error('DeliveryId ontbreekt.');
  if (!techniekerCode) throw new Error('TechniekerCode ontbreekt.');

  var existing = getGrabbelOrderGroup(deliveryId, techniekerCode);
  if (!existing) throw new Error('Bestelgroep niet gevonden.');

  var allowedStatuses = [getGrabbelOrderStatusReceived_(), getGrabbelOrderStatusNotPickedUp_()];
  if (allowedStatuses.indexOf(safeText(existing.status)) < 0) {
    throw new Error('Manager kan deze bestelgroep nog niet afsluiten.');
  }

  var first = existing.lines[0] || {};
  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
  var stamp = toDisplayDateTime(nowRaw);

  var group = replaceGrabbelOrderGroup_(
    deliveryId,
    techniekerCode,
    existing.lines.map(function (line) {
      return Object.assign({}, line);
    }),
    {
      techniekerNaam: safeText(existing.techniekerNaam),
      status: getGrabbelOrderStatusClosed_(),
      actor: safeText(actor.naam || actor.email),
      documentDatum: safeText(first.documentDatum),
      documentDatumIso: safeText(first.documentDatumIso),
      aangemaaktOp: safeText(first.aangemaaktOp),
      aangemaaktOpRaw: safeText(first.aangemaaktOpRaw),
      klaargezetOp: safeText(first.klaargezetOp),
      meegegevenOp: safeText(first.meegegevenOp),
      ontvangenOp: safeText(first.ontvangenOp),
      geslotenOp: stamp,
    }
  );

  if (typeof markNotificationsProcessedBySource === 'function') {
    markNotificationsProcessedBySource({
      sessionId: sessionId,
      bronType: 'GrabbelOrder',
      bronId: deliveryId + ':' + techniekerCode,
    });
  }

  writeAudit({
    actie: 'CLOSE_GRABBEL_ORDER_GROUP',
    actor: actor,
    documentType: 'GrabbelOrderGroup',
    documentId: deliveryId + ':' + techniekerCode,
    details: {},
  });

  return group;
}

/* ---------------------------------------------------------
   Auto confirm helper
   --------------------------------------------------------- */

function autoConfirmAfter3Hours(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertWarehouseAccess(sessionId);
  var now = new Date();
  var groups = buildGrabbelOrderGroups(getAllGrabbelOrders());
  var updated = [];

  groups.forEach(function (group) {
    if (safeText(group.status) !== getGrabbelOrderStatusGiven_()) {
      return;
    }

    var first = group.lines[0] || {};
    var meegegevenOp = safeText(first.meegegevenOp);
    if (!meegegevenOp) {
      return;
    }

    var dt = new Date(meegegevenOp);
    if (isNaN(dt.getTime())) {
      return;
    }

    var diffMs = now.getTime() - dt.getTime();
    if (diffMs < 3 * 60 * 60 * 1000) {
      return;
    }

    var result = replaceGrabbelOrderGroup_(
      group.deliveryId,
      group.techniekerCode,
      group.lines.map(function (line) {
        return Object.assign({}, line, {
          ontvangenAantal: safeNumber(line.voorzienAantal, 0),
        });
      }),
      {
        techniekerNaam: safeText(group.techniekerNaam),
        status: getGrabbelOrderStatusReceived_(),
        actor: safeText(actor.naam || actor.email),
        documentDatum: safeText(first.documentDatum),
        documentDatumIso: safeText(first.documentDatumIso),
        aangemaaktOp: safeText(first.aangemaaktOp),
        aangemaaktOpRaw: safeText(first.aangemaaktOpRaw),
        klaargezetOp: safeText(first.klaargezetOp),
        meegegevenOp: safeText(first.meegegevenOp),
        ontvangenOp: Utilities.formatDate(now, TIMEZONE, 'dd/MM/yyyy HH:mm:ss'),
        geslotenOp: '',
      }
    );

    updated.push({
      deliveryId: group.deliveryId,
      techniekerCode: group.techniekerCode,
      status: result.status,
    });
  });

  if (updated.length) {
    writeAudit({
      actie: 'AUTO_CONFIRM_GRABBEL_AFTER_3_HOURS',
      actor: actor,
      documentType: 'GrabbelOrderGroup',
      documentId: 'AUTO',
      details: {
        updatedCount: updated.length,
      },
    });
  }

  return {
    updatedCount: updated.length,
    items: updated,
  };
}

/* ---------------------------------------------------------
   Screen queries
   --------------------------------------------------------- */

function buildTechnicianGrabbelView_(techniekerCode) {
  var deliveries = getUpcomingDeliveriesForTechnician(techniekerCode, 7);
  var groups = buildGrabbelOrderGroups(getGrabbelOrdersByTechnicianCode(techniekerCode));

  return {
    plannedDeliveries: deliveries.map(function (delivery) {
      return Object.assign({}, delivery, {
        hasExistingOrder: hasExistingOrderForDelivery(techniekerCode, delivery.deliveryId),
      });
    }),
    deliveryGroups: groups,
  };
}

function getAppData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var techniekerCode = safeText(
    payload.techniekerCode ||
    payload.technicianCode ||
    payload.techCode
  );

  var access = assertTechnicianAccessToRef(sessionId, techniekerCode);
  var techCode = safeText(access.technician.code || access.technician.ref);

  var techView = buildTechnicianGrabbelView_(techCode);
  var notifications =
    typeof getNotificationsData === 'function'
      ? getNotificationsData({ sessionId: sessionId }).items
      : [];

  return {
    techniekerCode: techCode,
    techniekerNaam: safeText(access.technician.naam),
    plannedDeliveries: techView.plannedDeliveries,
    deliveryGroups: techView.deliveryGroups,
    notifications: notifications,
    busCounts: [],
    mobileRequests: [],
  };
}

function getWarehouseData(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = assertWarehouseAccess(sessionId);
  var groups = getUpcomingDeliveryGroups(7);
  var notifications =
    typeof getNotificationsData === 'function'
      ? getNotificationsData({ sessionId: sessionId }).items
      : [];

  return {
    actor: {
      naam: safeText(actor.naam),
      rol: safeText(actor.rol),
    },
    deliveryGroups: groups,
    summary: buildDeliverySummary(groups),
    notifications: notifications,
  };
}
