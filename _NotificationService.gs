/* =========================================================
   21_NotificationService.gs
   Refactor: notification core service
   Doel:
   - één centrale notificatielaag
   - generieke create/push helpers
   - read models per gebruiker/rol
   - markeren als gelezen / verwerkt
   ========================================================= */

/* ---------------------------------------------------------
   Status / role fallbacks
   --------------------------------------------------------- */

function getNotificationStatusOpen_() {
  if (typeof NOTIFICATION_STATUS !== 'undefined' && NOTIFICATION_STATUS && NOTIFICATION_STATUS.OPEN) {
    return NOTIFICATION_STATUS.OPEN;
  }
  return 'Open';
}

function getNotificationStatusRead_() {
  if (typeof NOTIFICATION_STATUS !== 'undefined' && NOTIFICATION_STATUS && NOTIFICATION_STATUS.READ) {
    return NOTIFICATION_STATUS.READ;
  }
  return 'Gelezen';
}

function getNotificationStatusProcessed_() {
  if (typeof NOTIFICATION_STATUS !== 'undefined' && NOTIFICATION_STATUS && NOTIFICATION_STATUS.PROCESSED) {
    return NOTIFICATION_STATUS.PROCESSED;
  }
  return 'Verwerkt';
}

function getNotificationStatusClosed_() {
  if (typeof NOTIFICATION_STATUS !== 'undefined' && NOTIFICATION_STATUS && NOTIFICATION_STATUS.CLOSED) {
    return NOTIFICATION_STATUS.CLOSED;
  }
  return 'Gesloten';
}

function getNotificationRoleManager_() {
  if (typeof NOTIFICATION_ROLE !== 'undefined' && NOTIFICATION_ROLE && NOTIFICATION_ROLE.MANAGER) {
    return NOTIFICATION_ROLE.MANAGER;
  }
  return 'Manager';
}

function getNotificationRoleWarehouse_() {
  if (typeof NOTIFICATION_ROLE !== 'undefined' && NOTIFICATION_ROLE && NOTIFICATION_ROLE.WAREHOUSE) {
    return NOTIFICATION_ROLE.WAREHOUSE;
  }
  return 'Magazijn';
}

function getNotificationRoleMobileWarehouse_() {
  if (typeof NOTIFICATION_ROLE !== 'undefined' && NOTIFICATION_ROLE && NOTIFICATION_ROLE.MOBILE_WAREHOUSE) {
    return NOTIFICATION_ROLE.MOBILE_WAREHOUSE;
  }
  return 'MobielMagazijn';
}

function getNotificationRoleTechnician_() {
  if (typeof NOTIFICATION_ROLE !== 'undefined' && NOTIFICATION_ROLE && NOTIFICATION_ROLE.TECHNICIAN) {
    return NOTIFICATION_ROLE.TECHNICIAN;
  }
  return 'Technieker';
}

/* ---------------------------------------------------------
   IDs / mapping
   --------------------------------------------------------- */

function makeNotificationId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return 'NTF-' + stamp + '-' + makeUuidId().slice(0, 6).toUpperCase();
}

function mapNotification(row) {
  return {
    notificationId: safeText(row.NotificationID || row.NotificationId || row.ID),
    rol: safeText(row.Rol),
    ontvangerCode: safeText(row.OntvangerCode),
    ontvangerNaam: safeText(row.OntvangerNaam),
    type: safeText(row.Type),
    titel: safeText(row.Titel),
    bericht: safeText(row.Bericht),
    bronType: safeText(row.BronType),
    bronId: safeText(row.BronID || row.BronId),
    status: safeText(row.Status),
    aangemaaktOp: safeText(row.AangemaaktOp),
    aangemaaktOpRaw: safeText(row.AangemaaktOpRaw || row.AangemaaktOp),
    gelezenOp: safeText(row.GelezenOp),
    verwerktOp: safeText(row.VerwerktOp),
    geslotenOp: safeText(row.GeslotenOp),
    extraJson: safeText(row.ExtraJson || row.ExtraJSON),
  };
}

/* ---------------------------------------------------------
   Builders
   --------------------------------------------------------- */

function buildNotificationObject(payload) {
  payload = payload || {};

  var nowRaw = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  return {
    NotificationID: makeNotificationId_(),
    Rol: safeText(payload.rol),
    OntvangerCode: safeText(payload.ontvangerCode),
    OntvangerNaam: safeText(payload.ontvangerNaam),
    Type: safeText(payload.type),
    Titel: safeText(payload.titel),
    Bericht: safeText(payload.bericht),
    BronType: safeText(payload.bronType),
    BronID: safeText(payload.bronId),
    Status: getNotificationStatusOpen_(),
    AangemaaktOp: toDisplayDateTime(nowRaw),
    AangemaaktOpRaw: nowRaw,
    GelezenOp: '',
    VerwerktOp: '',
    GeslotenOp: '',
    ExtraJson: safeJson(payload.extra || {}),
  };
}

/* ---------------------------------------------------------
   Create / push
   --------------------------------------------------------- */

function hasOpenNotification(payload) {
  payload = payload || {};

  var rol = safeText(payload.rol);
  var ontvangerCode = safeText(payload.ontvangerCode);
  var type = safeText(payload.type);
  var bronType = safeText(payload.bronType);
  var bronId = safeText(payload.bronId);

  return readObjectsSafe(TABS.NOTIFICATIONS)
    .map(mapNotification)
    .some(function (item) {
      return safeText(item.rol) === rol &&
        safeText(item.ontvangerCode) === ontvangerCode &&
        safeText(item.type) === type &&
        safeText(item.bronType) === bronType &&
        safeText(item.bronId) === bronId &&
        safeText(item.status) === getNotificationStatusOpen_();
    });
}

function pushNotification(payload) {
  payload = payload || {};

  if (!safeText(payload.rol)) throw new Error('Rol is verplicht voor notificatie.');
  if (!safeText(payload.type)) throw new Error('Type is verplicht voor notificatie.');
  if (!safeText(payload.titel)) throw new Error('Titel is verplicht voor notificatie.');
  if (!safeText(payload.bericht)) throw new Error('Bericht is verplicht voor notificatie.');

  if (isTrue(payload.skipIfOpenExists) !== false) {
    if (hasOpenNotification(payload)) {
      return {
        created: false,
        duplicateOpenNotification: true,
      };
    }
  }

  var obj = buildNotificationObject(payload);
  appendObjects(TABS.NOTIFICATIONS, [obj]);

  return {
    created: true,
    notification: mapNotification(obj),
  };
}

function pushManagerNotification(type, title, message, bronType, bronId, ontvangerCode) {
  return pushNotification({
    rol: getNotificationRoleManager_(),
    ontvangerCode: safeText(ontvangerCode) || 'MANAGER',
    ontvangerNaam: 'Manager',
    type: type,
    titel: title,
    bericht: message,
    bronType: bronType,
    bronId: bronId,
  });
}

function pushWarehouseNotification(type, title, message, bronType, bronId, ontvangerCode) {
  return pushNotification({
    rol: getNotificationRoleWarehouse_(),
    ontvangerCode: safeText(ontvangerCode) || 'WAREHOUSE',
    ontvangerNaam: 'Magazijn',
    type: type,
    titel: title,
    bericht: message,
    bronType: bronType,
    bronId: bronId,
  });
}

function pushMobileWarehouseNotification(type, title, message, bronType, bronId, ontvangerCode) {
  return pushNotification({
    rol: getNotificationRoleMobileWarehouse_(),
    ontvangerCode: safeText(ontvangerCode) || 'MOBIEL',
    ontvangerNaam: 'Mobiel magazijn',
    type: type,
    titel: title,
    bericht: message,
    bronType: bronType,
    bronId: bronId,
  });
}

function pushTechnicianNotification(type, title, message, bronType, bronId, ontvangerCode, ontvangerNaam) {
  return pushNotification({
    rol: getNotificationRoleTechnician_(),
    ontvangerCode: safeText(ontvangerCode),
    ontvangerNaam: safeText(ontvangerNaam) || safeText(ontvangerCode),
    type: type,
    titel: title,
    bericht: message,
    bronType: bronType,
    bronId: bronId,
  });
}

/* ---------------------------------------------------------
   Read layer
   --------------------------------------------------------- */

function getAllNotifications() {
  return readObjectsSafe(TABS.NOTIFICATIONS)
    .map(mapNotification)
    .sort(function (a, b) {
      return (
        safeText(b.aangemaaktOpRaw).localeCompare(safeText(a.aangemaaktOpRaw)) ||
        safeText(b.notificationId).localeCompare(safeText(a.notificationId))
      );
    });
}

function getNotificationById(notificationId) {
  var id = safeText(notificationId);
  if (!id) return null;
  return getAllNotifications().find(function (item) {
    return safeText(item.notificationId) === id;
  }) || null;
}

function resolveNotificationContextForActor_(actor) {
  if (roleAllowed(actor, [ROLE.MANAGER])) {
    return {
      allowedRole: getNotificationRoleManager_(),
      recipientCode: 'MANAGER',
      actorType: 'manager',
    };
  }

  if (roleAllowed(actor, [ROLE.WAREHOUSE])) {
    return {
      allowedRole: getNotificationRoleWarehouse_(),
      recipientCode: 'WAREHOUSE',
      actorType: 'warehouse',
    };
  }

  if (roleAllowed(actor, [ROLE.MOBILE_WAREHOUSE])) {
    return {
      allowedRole: getNotificationRoleMobileWarehouse_(),
      recipientCode: safeText(actor.mobileWarehouseCode) || 'MOBIEL',
      actorType: 'mobilewarehouse',
    };
  }

  if (roleAllowed(actor, [ROLE.TECHNICIAN])) {
    return {
      allowedRole: getNotificationRoleTechnician_(),
      recipientCode: safeText(actor.code || actor.techniekerCode || actor.technicianCode),
      actorType: 'technician',
    };
  }

  return {
    allowedRole: '',
    recipientCode: '',
    actorType: '',
  };
}

function filterNotificationsForActor_(rows, actor) {
  if (roleAllowed(actor, [ROLE.ADMIN])) {
    return rows;
  }

  var ctx = resolveNotificationContextForActor_(actor);
  if (!ctx.allowedRole) return [];

  return rows.filter(function (item) {
    if (safeText(item.rol) !== safeText(ctx.allowedRole)) {
      return false;
    }

    if (ctx.actorType === 'manager' || ctx.actorType === 'warehouse') {
      return true;
    }

    if (ctx.actorType === 'mobilewarehouse') {
      return !ctx.recipientCode ||
        safeText(item.ontvangerCode) === ctx.recipientCode ||
        safeText(item.ontvangerCode) === 'MOBIEL';
    }

    if (ctx.actorType === 'technician') {
      return safeText(item.ontvangerCode) === safeText(ctx.recipientCode);
    }

    return false;
  });
}

function getNotificationsData(payload) {
  payload = payload || {};
  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);

  var rows = filterNotificationsForActor_(getAllNotifications(), actor);

  return {
    items: rows,
    summary: {
      totaal: rows.length,
      open: rows.filter(function (x) { return safeText(x.status) === getNotificationStatusOpen_(); }).length,
      gelezen: rows.filter(function (x) { return safeText(x.status) === getNotificationStatusRead_(); }).length,
      verwerkt: rows.filter(function (x) { return safeText(x.status) === getNotificationStatusProcessed_(); }).length,
      gesloten: rows.filter(function (x) { return safeText(x.status) === getNotificationStatusClosed_(); }).length,
    }
  };
}

function getUnreadNotificationCountForActor(actor) {
  return filterNotificationsForActor_(getAllNotifications(), actor)
    .filter(function (item) {
      return safeText(item.status) === getNotificationStatusOpen_();
    })
    .length;
}

/* ---------------------------------------------------------
   Access / ownership
   --------------------------------------------------------- */

function assertNotificationOwnership(actor, notification) {
  if (!notification) {
    throw new Error('Notificatie niet gevonden.');
  }

  if (roleAllowed(actor, [ROLE.ADMIN])) {
    return true;
  }

  var ctx = resolveNotificationContextForActor_(actor);
  if (!ctx.allowedRole) {
    throw new Error('Geen rechten voor notificaties.');
  }

  if (safeText(notification.rol) !== safeText(ctx.allowedRole)) {
    throw new Error('Geen toegang tot deze notificatie.');
  }

  if (ctx.actorType === 'manager' || ctx.actorType === 'warehouse') {
    return true;
  }

  if (ctx.actorType === 'mobilewarehouse') {
    if (
      !ctx.recipientCode ||
      safeText(notification.ontvangerCode) === ctx.recipientCode ||
      safeText(notification.ontvangerCode) === 'MOBIEL'
    ) {
      return true;
    }
  }

  if (ctx.actorType === 'technician') {
    if (safeText(notification.ontvangerCode) === safeText(ctx.recipientCode)) {
      return true;
    }
  }

  throw new Error('Geen toegang tot deze notificatie.');
}

/* ---------------------------------------------------------
   Row update helpers
   --------------------------------------------------------- */

function updateNotificationsByPredicate_(predicate, mutator) {
  var table = getAllValues(TABS.NOTIFICATIONS);
  if (!table.length) {
    return { updatedCount: 0 };
  }

  var headerRow = table[0];
  var dataRows = table.slice(1);
  var updatedCount = 0;

  var newRows = dataRows.map(function (row) {
    var obj = rowToObject(headerRow, row);
    var mapped = mapNotification(obj);

    if (!predicate(mapped, obj)) {
      return row;
    }

    mutator(obj, mapped);
    updatedCount += 1;
    return buildRowFromHeaders(headerRow, obj);
  });

  if (updatedCount) {
    writeFullTable(TABS.NOTIFICATIONS, headerRow, newRows);
  }

  return { updatedCount: updatedCount };
}

/* ---------------------------------------------------------
   Markeren
   --------------------------------------------------------- */

function markNotificationRead(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);
  var notificationId = safeText(payload.notificationId);
  if (!notificationId) throw new Error('NotificationId ontbreekt.');

  var notification = getNotificationById(notificationId);
  assertNotificationOwnership(actor, notification);

  var result = updateNotificationsByPredicate_(
    function (mapped) {
      return safeText(mapped.notificationId) === notificationId &&
        safeText(mapped.status) === getNotificationStatusOpen_();
    },
    function (obj) {
      obj.Status = getNotificationStatusRead_();
      obj.GelezenOp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
    }
  );

  if (result.updatedCount) {
    writeAudit({
      actie: 'MARK_NOTIFICATION_READ',
      actor: actor,
      documentType: 'Notificatie',
      documentId: notificationId,
      details: {},
    });
  }

  return {
    updatedCount: result.updatedCount,
    notificationId: notificationId,
  };
}

function markAllNotificationsRead(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);
  var rows = filterNotificationsForActor_(getAllNotifications(), actor);
  var allowedIds = {};
  rows.forEach(function (item) {
    allowedIds[item.notificationId] = true;
  });

  var result = updateNotificationsByPredicate_(
    function (mapped) {
      return allowedIds[safeText(mapped.notificationId)] &&
        safeText(mapped.status) === getNotificationStatusOpen_();
    },
    function (obj) {
      obj.Status = getNotificationStatusRead_();
      obj.GelezenOp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
    }
  );

  if (result.updatedCount) {
    writeAudit({
      actie: 'MARK_ALL_NOTIFICATIONS_READ',
      actor: actor,
      documentType: 'Notificaties',
      documentId: 'ALL',
      details: {
        updatedCount: result.updatedCount,
      },
    });
  }

  return {
    updatedCount: result.updatedCount,
  };
}

function markNotificationsProcessedBySource(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);
  var bronType = safeText(payload.bronType);
  var bronId = safeText(payload.bronId);

  if (!bronType || !bronId) {
    throw new Error('BronType en BronId zijn verplicht.');
  }

  var result = updateNotificationsByPredicate_(
    function (mapped) {
      return safeText(mapped.bronType) === bronType &&
        safeText(mapped.bronId) === bronId &&
        [getNotificationStatusOpen_(), getNotificationStatusRead_()].indexOf(safeText(mapped.status)) >= 0;
    },
    function (obj) {
      obj.Status = getNotificationStatusProcessed_();
      obj.VerwerktOp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
    }
  );

  if (result.updatedCount) {
    writeAudit({
      actie: 'PROCESS_NOTIFICATIONS_BY_SOURCE',
      actor: actor,
      documentType: 'Notificaties',
      documentId: bronType + ':' + bronId,
      details: {
        updatedCount: result.updatedCount,
      },
    });
  }

  return {
    updatedCount: result.updatedCount,
    bronType: bronType,
    bronId: bronId,
  };
}

function closeNotificationsBySource(payload) {
  payload = payload || {};

  var sessionId = getPayloadSessionId(payload);
  var actor = requireLoggedInUser(sessionId);
  var bronType = safeText(payload.bronType);
  var bronId = safeText(payload.bronId);

  if (!bronType || !bronId) {
    throw new Error('BronType en BronId zijn verplicht.');
  }

  var result = updateNotificationsByPredicate_(
    function (mapped) {
      return safeText(mapped.bronType) === bronType &&
        safeText(mapped.bronId) === bronId &&
        safeText(mapped.status) !== getNotificationStatusClosed_();
    },
    function (obj) {
      obj.Status = getNotificationStatusClosed_();
      obj.GeslotenOp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
    }
  );

  if (result.updatedCount) {
    writeAudit({
      actie: 'CLOSE_NOTIFICATIONS_BY_SOURCE',
      actor: actor,
      documentType: 'Notificaties',
      documentId: bronType + ':' + bronId,
      details: {
        updatedCount: result.updatedCount,
      },
    });
  }

  return {
    updatedCount: result.updatedCount,
    bronType: bronType,
    bronId: bronId,
  };
}
