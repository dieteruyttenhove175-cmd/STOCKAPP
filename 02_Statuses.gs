/* =========================================================
   02_Statuses.gs — statussen / redenen / mutatietypes
   ========================================================= */

const STATUS = {
  REQUESTED: 'Aangevraagd',
  READY: 'Klaargezet',
  DISPATCHED: 'Meegegeven',
  RECEIVED: 'Ontvangen',
  AUTO_RECEIVED: 'Automatisch ontvangen',
  NOT_PICKED: 'Niet afgehaald',
  APPROVED: 'Goedgekeurd',
  CLOSED: 'Gesloten',
  CANCELLED: 'Geannuleerd'
};

const MANAGER_STATUS = {
  NONE: '',
  PENDING: 'Te controleren',
  APPROVED: 'Goedgekeurd'
};

const RECEIPT_STATUS = {
  EXPECTED: 'Verwacht',
  IN_PROGRESS: 'In behandeling',
  SUBMITTED: 'Ingediend',
  APPROVED: 'Goedgekeurd'
};

const RETURN_STATUS = {
  IN_PROGRESS: 'In behandeling',
  SUBMITTED: 'Ingediend',
  APPROVED: 'Goedgekeurd'
};

const NEED_ISSUE_STATUS = {
  OPEN: 'Open',
  BOOKED: 'Geboekt'
};

const CONSUMPTION_STATUS = {
  OPEN: 'Open',
  BOOKED: 'Geboekt'
};

const BUS_COUNT_STATUS = {
  OPEN: 'Open',
  SUBMITTED: 'Ingediend',
  APPROVED: 'Goedgekeurd'
};

const TRANSFER_REQUEST_STATUS = {
  OPEN: 'Open',
  APPROVED: 'Goedgekeurd',
  REJECTED: 'Geweigerd',
  CONVERTED: 'Omgezet'
};

const TRANSFER_STATUS = {
  OPEN: 'Open',
  READY: 'Klaargezet',
  BOOKED: 'Geboekt',
  CANCELLED: 'Geannuleerd'
};

const NOTIFICATION_STATUS = {
  OPEN: 'Open',
  READ: 'Gelezen'
};

const RECEIPT_DELTA_REASONS = [
  '',
  'Niet volledig geleverd',
  'Beschadigd ontvangen',
  'Verkeerd artikel / verpakking',
  'Extra geleverd',
  'Andere reden'
];

const RETURN_REASONS_TECHNICIAN = [
  'NPC / beschadigde artikelen',
  'Andere reden'
];

const RETURN_REASONS_WAREHOUSE = [
  'Te veel geleverd',
  'NPC / beschadigde artikelen',
  'Andere reden'
];

const MOVEMENT_TYPE = {
  RECEIPT: 'Ontvangst',
  GRABBEL_DELIVERY: 'GrabbelBelevering',
  NEED_REPLENISHMENT: 'BehoefteAanvulling',
  CONSUMPTION: 'Verbruik',
  RETURN_IN: 'RetourIn',
  RETURN_TO_FLUVIUS: 'RetourNaarFluvius',
  TRANSFER_CENTRAL_TO_MOBILE: 'TransferCentraalNaarMobiel',
  TRANSFER_MOBILE_TO_BUS: 'TransferMobielNaarBus',
  BUS_CORRECTION_IN: 'BusCorrectieIn',
  BUS_CORRECTION_OUT: 'BusCorrectieUit',
  CENTRAL_CORRECTION_IN: 'CentraalCorrectieIn',
  CENTRAL_CORRECTION_OUT: 'CentraalCorrectieUit',
  MOBILE_CORRECTION_IN: 'MobielCorrectieIn',
  MOBILE_CORRECTION_OUT: 'MobielCorrectieUit',
  COUNT_CORRECTION_CENTRAL: 'TellingCorrectieCentraal',
  COUNT_CORRECTION_MOBILE: 'TellingCorrectieMobiel',
  COUNT_CORRECTION_BUS: 'TellingCorrectieBus',
  CENTRAL_COUNT_IN : 'CentralCountIn',
  CENTRAL_COUNT_OUT : 'CentralCountOut'
};