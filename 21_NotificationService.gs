/* =========================================================
   21_NotificationService.gs — meldingen / routing / gelezen
   ========================================================= */

function makeNotificationId() {
  return makeStampedId('N');
}

function buildNotificationObject(payload) {
  return {
    NotificatieID: makeNotificationId(),
    Rol: safeText(payload.rol),
    OntvangerCode: safeText(payload.ontvangerCode),
    OntvangerNaam: safeText(payload.ontvangerNaam),
    Type: safeText(payload.type),
    Titel: safeText(payload.titel),
    Bericht: safeText(payload.bericht),
    BronType: safeText(payload.bronType),
    BronID: safeText(payload.bronId),
    Status: NOTIFICATION_STATUS.OPEN,
    AangemaaktOp: nowStamp(),
    GelezenOp: ''
  };
}

function getAllNotifications() {
  return readObjectsSafe(TABS.NOTIFICATIONS)
    .map(mapNotification)
    .sort((a, b) => String(b.aangemaaktOpRaw || '').localeCompare(String(a.aangemaaktOpRaw || '')));
}

function hasOpenNotification(payload) {
  const rol = safeText(payload.rol);
  const ontvangerCode = safeText(payload.ontvangerCode);
  const type = safeText(payload.type);
  const bronType = safeText(payload.bronType);
  const bronId = safeText(payload.bronId);

  return getAllNotifications().some(item =>
    item.status === NOTIFICATION_STATUS.OPEN &&
    safeText(item.rol) === rol &&
    safeText(item.ontvangerCode) === ontvangerCode &&
    safeText(item.type) === type &&
    safeText(item.bronType) === bronType &&
    safeText(item.bronId) === bronId
  );
}

function pushNotification(payload) {
  if (!payload) throw new Error('Geen notificatiepayload ontvangen.');

  if (hasOpenNotification(payload)) {
    return { success: true, skipped: true };
  }

  appendObjects(TABS.NOTIFICATIONS, [buildNotificationObject(payload)]);

  writeAudit(
    'Notificatie aangemaakt',
    'Systeem',
    'NotificatieService',
    'Notificatie',
    payload.bronId || '',
    {
      rol: payload.rol,
      ontvangerCode: payload.ontvangerCode,
      type: payload.type,
      bronType: payload.bronType,
      bronId: payload.bronId
    }
  );

  return { success: true, skipped: false };
}

function pushManagerNotification(type, title, message, bronType, bronId) {
  return pushNotification({
    rol: NOTIFICATION_ROLE.MANAGER,
    ontvangerCode: 'MANAGER',
    ontvangerNaam: 'Manager',
    type: type,
    titel: title,
    bericht: message,
    bronType: bronType,
    bronId: bronId
  });
}

function pushWarehouseNotification(type, title, message, bronType, bronId) {
  return pushNotification({
    rol: NOTIFICATION_ROLE.WAREHOUSE,
    ontvangerCode: 'MAGAZIJN',
    ontvangerNaam: 'Magazijn',
    type: type,
    titel: title,
    bericht: message,
    bronType: bronType,
    bronId: bronId
  });
}

function pushMobileWarehouseNotification(type, title, message, bronType, bronId) {
  return pushNotification({
    rol: NOTIFICATION_ROLE.MOBILE_WAREHOUSE,
    ontvangerCode: 'MOBIEL_MAGAZIJN',
    ontvangerNaam: 'Mobiel magazijn',
    type: type,
    titel: title,
    bericht: message,
    bronType: bronType,
    bronId: bronId
  });
}

function pushTechnicianNotification(techniekerCode, techniekerNaam, type, title, message, bronType, bronId) {
  return pushNotification({
    rol: NOTIFICATION_ROLE.TECHNICIAN,
    ontvangerCode: safeText(techniekerCode),
    ontvangerNaam: safeText(techniekerNaam || getTechnicianNameByCode(techniekerCode)),
    type: type,
    titel: title,
    bericht: message,
    bronType: bronType,
    bronId: bronId
  });
}

function getNotificationsForManager() {
  return getAllNotifications().filter(item => item.rol === NOTIFICATION_ROLE.MANAGER);
}

function getNotificationsForWarehouse() {
  return getAllNotifications().filter(item => item.rol === NOTIFICATION_ROLE.WAREHOUSE);
}

function getNotificationsForMobileWarehouse() {
  return getAllNotifications().filter(item => item.rol === NOTIFICATION_ROLE.MOBILE_WAREHOUSE);
}

function getNotificationsForTechnician(techRef) {
  const technicians = readObjectsSafe(TABS.TECHNICIANS).map(mapTechnician);
  const technician = findTechnicianByRef(technicians, techRef);

  if (!technician || !technician.active) {
    throw new Error('Technieker niet gevonden of niet actief.');
  }

  return getAllNotifications().filter(item =>
    item.rol === NOTIFICATION_ROLE.TECHNICIAN &&
    normalizeRef(item.ontvangerCode) === normalizeRef(technician.code)
  );
}

function assertNotificationOwnership(notification, sessionId) {
  const user = requireLoggedInUser(sessionId);
  if (user.rol === ROLE.ADMIN) return user;

  if (notification.rol === NOTIFICATION_ROLE.MANAGER) {
    if (!roleAllowed(user, [ROLE.MANAGER, ROLE.ANALYSIS])) {
      throw new Error('Geen rechten voor deze melding.');
    }
    return user;
  }

  if (notification.rol === NOTIFICATION_ROLE.WAREHOUSE) {
    if (!roleAllowed(user, [ROLE.WAREHOUSE, ROLE.MANAGER])) {
      throw new Error('Geen rechten voor deze melding.');
    }
    return user;
  }

  if (notification.rol === NOTIFICATION_ROLE.MOBILE_WAREHOUSE) {
    if (!roleAllowed(user, [ROLE.MOBILE_WAREHOUSE, ROLE.MANAGER])) {
      throw new Error('Geen rechten voor deze melding.');
    }
    return user;
  }

  if (notification.rol === NOTIFICATION_ROLE.TECHNICIAN) {
    if (!roleAllowed(user, [ROLE.TECHNICIAN])) {
      throw new Error('Geen rechten voor deze melding.');
    }

    if (normalizeRef(user.techniekerCode) !== normalizeRef(notification.ontvangerCode)) {
      throw new Error('Je kan enkel je eigen meldingen aanpassen.');
    }

    return user;
  }

  throw new Error('Geen rechten voor deze melding.');
}

function markNotificationRead(payload) {
  if (!payload) throw new Error('Geen payload ontvangen.');

  const sessionId = getPayloadSessionId(payload);
  const notificatieId = safeText(payload.notificatieId);

  if (!notificatieId) {
    throw new Error('NotificatieID ontbreekt.');
  }

  const sheet = getSheetOrThrow(TABS.NOTIFICATIONS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Tab Notificaties is leeg.');

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = false;
  let currentNotification = null;
  let actor = null;

  for (let i = 1; i < values.length; i++) {
    if (safeText(values[i][col['NotificatieID']]) !== notificatieId) continue;

    currentNotification = mapNotification({
      NotificatieID: values[i][col['NotificatieID']],
      Rol: values[i][col['Rol']],
      OntvangerCode: values[i][col['OntvangerCode']],
      OntvangerNaam: values[i][col['OntvangerNaam']],
      Type: values[i][col['Type']],
      Titel: values[i][col['Titel']],
      Bericht: values[i][col['Bericht']],
      BronType: values[i][col['BronType']],
      BronID: values[i][col['BronID']],
      Status: values[i][col['Status']],
      AangemaaktOp: values[i][col['AangemaaktOp']],
      GelezenOp: values[i][col['GelezenOp']]
    });

    actor = assertNotificationOwnership(currentNotification, sessionId);

    values[i][col['Status']] = NOTIFICATION_STATUS.READ;
    values[i][col['GelezenOp']] = nowStamp();
    updated = true;
    break;
  }

  if (!updated) {
    throw new Error('Notificatie niet gevonden.');
  }

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  writeAudit(
    'Notificatie gelezen',
    actor ? actor.rol : '',
    actor ? (actor.naam || actor.email || actor.techniekerCode) : '',
    'Notificatie',
    notificatieId,
    {
      bronType: currentNotification ? currentNotification.bronType : '',
      bronId: currentNotification ? currentNotification.bronId : ''
    }
  );

  return {
    success: true,
    message: 'Notificatie gemarkeerd als gelezen.'
  };
}

function markNotificationsReadBySource(filters) {
  filters = filters || {};

  const sheet = getSheetOrThrow(TABS.NOTIFICATIONS);
  const values = sheet.getDataRange().getValues();
  if (!values.length) return { success: true, updated: 0 };

  const headers = values[0].map(h => safeText(h));
  const col = getColMap(headers);

  let updated = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const rol = safeText(row[col['Rol']]);
    const ontvangerCode = safeText(row[col['OntvangerCode']]);
    const type = safeText(row[col['Type']]);
    const bronType = safeText(row[col['BronType']]);
    const bronId = safeText(row[col['BronID']]);
    const status = safeText(row[col['Status']]);

    if (filters.rol && rol !== safeText(filters.rol)) continue;
    if (filters.ontvangerCode && ontvangerCode !== safeText(filters.ontvangerCode)) continue;
    if (filters.type && type !== safeText(filters.type)) continue;
    if (filters.bronType && bronType !== safeText(filters.bronType)) continue;
    if (filters.bronId && bronId !== safeText(filters.bronId)) continue;
    if (status !== NOTIFICATION_STATUS.OPEN) continue;

    values[i][col['Status']] = NOTIFICATION_STATUS.READ;
    values[i][col['GelezenOp']] = nowStamp();
    updated++;
  }

  if (updated && values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));
  }

  return { success: true, updated: updated };
}

function markManagerNotificationsBySource(bronType, bronId) {
  return markNotificationsReadBySource({
    rol: NOTIFICATION_ROLE.MANAGER,
    bronType: bronType,
    bronId: bronId
  });
}

function markWarehouseNotificationsBySource(bronType, bronId) {
  return markNotificationsReadBySource({
    rol: NOTIFICATION_ROLE.WAREHOUSE,
    bronType: bronType,
    bronId: bronId
  });
}

function markMobileWarehouseNotificationsBySource(bronType, bronId) {
  return markNotificationsReadBySource({
    rol: NOTIFICATION_ROLE.MOBILE_WAREHOUSE,
    bronType: bronType,
    bronId: bronId
  });
}

function getUnreadNotificationCountForRole(notificationRole) {
  return getAllNotifications().filter(item =>
    item.rol === notificationRole &&
    item.status === NOTIFICATION_STATUS.OPEN
  ).length;
}