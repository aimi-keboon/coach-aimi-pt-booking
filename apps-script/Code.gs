/**
 * Coach Aimi PT Booking — Google Apps Script backend
 * Deploy as a web app executing as the deployment owner.
 * Timezone: Asia/Kuala_Lumpur; set project time zone to GMT+08:00 in Project Settings.
 * Do not paste JSON into this file; Apps Script manages its own manifest.
 *
 * Admin-only editor actions end with _ so google.script.run can't invoke them.
 */
const APP = Object.freeze({
  TZ: 'Asia/Kuala_Lumpur',
  CLIENTS: 'Clients', BOOKINGS: 'Bookings', SCHEDULE: 'Schedule',
  CALENDARS: 'Calendars', SETTINGS: 'Settings',
  CLIENT_HEADERS: ['Name', 'Email', 'Phone', 'Active', 'CreditsGranted'],
  BOOKING_HEADERS: ['BookingId', 'Phone', 'ClientName', 'ClientEmail', 'SessionDate', 'StartHour', 'Status', 'EventRefsJSON', 'CreatedAt', 'UpdatedAt'],
  SESSION_MINUTES: 60, LOGIN_TTL: 300, SESSION_TTL: 21600
});

/**
 * Tiny API bridge for the GitHub Pages interface.
 * Google Apps Script HTML Service provides google.script.run inside this hidden iframe.
 */
function doGet() {
  const trustedOrigin = 'https://aimi-keboon.github.io';
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>Coach Aimi API</title></head><body><script>
(function(){
  'use strict';
  var origin = ${JSON.stringify(trustedOrigin)};
  var permitted = ['requestLoginCode','verifyLoginCode','getAvailability',
      'getMyCredits','getMyBookings','makeBooking','rescheduleBooking','cancelBooking',
      'requestAdminCode','verifyAdminCode','adminGetData','adminSaveClient','adminGrantCredits',
      'adminAddCalendar','adminSetCalendarEnabled'];
  function send(message) { window.top.postMessage(Object.assign({channel:'coach-aimi-api'},message),origin); }
  window.addEventListener('message',function(event){
    if (event.origin !== origin || event.source !== window.top) return;
    var data = event.data;
    if (!data || data.channel !== 'coach-aimi-site' || data.type !== 'CALL') return;
    if (typeof data.id !== 'string' || typeof data.action !== 'string' ||
        permitted.indexOf(data.action) < 0 || !Array.isArray(data.args)) return;
    try {
      var runner = google.script.run
        .withSuccessHandler(function(value){send({type:'RESULT',id:data.id,ok:true,value:value});})
        .withFailureHandler(function(error){send({type:'RESULT',id:data.id,ok:false,
          error:error && error.message ? error.message : 'Booking service error'});});
      runner[data.action].apply(runner,data.args);
    } catch(error) {send({type:'RESULT',id:data.id,ok:false,error:error.message || 'Booking service error'});}
  });
  send({type:'READY'});
})();
</script></body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('Coach Aimi API bridge')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Editor Run menu entry point. The setup helper stays private to the web app. */
function setupProject() {
  return setupProject_();
}

/** Editor Run menu entry point to verify calendar access after setup. */
function checkSetup() {
  return checkSetup_();
}

/** RUN ONCE IN APPS SCRIPT EDITOR. A spreadsheet link appears in execution logs. */
function setupProject_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    upgradeSchema_();
    Logger.log('Existing booking spreadsheet: https://docs.google.com/spreadsheets/d/' + id + '/edit');
    return;
  }
  const ss = SpreadsheetApp.create('Coach Aimi — PT Booking Control Panel');
  const first = ss.getSheets()[0];
  first.setName(APP.CLIENTS);
  first.getRange(1, 1, 1, APP.CLIENT_HEADERS.length).setValues([APP.CLIENT_HEADERS]);
  first.getRange('C:C').setNumberFormat('@');
  addTab_(ss, APP.BOOKINGS, APP.BOOKING_HEADERS);
  addTab_(ss, APP.SCHEDULE, ['Day', 'Open', 'FirstHour', 'LastStartHour']);
  ss.getSheetByName(APP.SCHEDULE).getRange(2, 1, 7, 4).setValues([
    ['Sunday', false, 7, 22], ['Monday', true, 7, 22],
    ['Tuesday', true, 7, 22], ['Wednesday', false, 7, 22],
    ['Thursday', true, 7, 22], ['Friday', true, 7, 22], ['Saturday', true, 7, 22]
  ]);
  addTab_(ss, APP.CALENDARS, ['CalendarId', 'Label', 'Enabled']);
  ss.getSheetByName(APP.CALENDARS).getRange(2, 1, 2, 3).setValues([
    ['', 'Other calendar — enter ID to check its busy times', false],
    ['', 'Additional availability calendar (optional)', false]
  ]);
  addTab_(ss, APP.SETTINGS, ['Key', 'Value']);
  ss.getSheetByName(APP.SETTINGS).getRange(2, 1, 5, 2).setValues([
    ['CoachEmail', 'aimi@keboon.net'], ['ReplyToEmail', 'aimi@keboon.net'],
    ['BookingCalendarId', 'aimi@keboon.net'],
    ['BookingHorizonDays', 60], ['RescheduleCutoffHours', 24]
  ]);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  Logger.log('Created control panel: ' + ss.getUrl());
  Logger.log('Add clients to Clients; set BookingCalendarId in Settings; add other calendars in Calendars (Enabled=TRUE).');
}

/** Safe migration for projects created before credit balances were introduced. */
function upgradeSchema_() {
  const sheet = tab_(APP.CLIENTS);
  const header = sheet.getRange(1, 1, 1, Math.max(5, sheet.getLastColumn())).getValues()[0];
  if (header[4] !== 'CreditsGranted') {
    if (header[4] && String(header[4]).trim()) throw new Error('Clients column E is in use. Please make E the CreditsGranted column manually.');
    sheet.getRange(1, 5).setValue('CreditsGranted');
    sheet.getRange(1, 5).setBackground('#16332b').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange('E:E').setNumberFormat('0');
  }
}

/** RUN IN EDITOR AFTER SETUP TO VERIFY CALENDAR ACCESS. */
function checkSetup_() {
  upgradeSchema_();
  const calendars = calendars_();
  const schedule = schedule_();
  Logger.log(JSON.stringify({spreadsheet: spreadsheet_().getUrl(), availabilityCalendarIds: calendars.map(c => c.id), bookingDestinationId: bookingCalendar_().id, workingDays: schedule}, null, 2));
  return true;
}

/** Optional daily integrity audit. A time-based trigger runs as its CREATOR. */
function installAuditTrigger_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'auditBookings_').forEach(ScriptApp.deleteTrigger);
  ScriptApp.newTrigger('auditBookings_').timeBased().everyDays(1).atHour(8).create();
  Logger.log('Daily 8AM MYT audit trigger installed for current Google account.');
}

function auditBookings_() {
  const rows = bookingRows_().filter(r => r.status === 'CONFIRMED');
  const now = Date.now(), max = now + 14 * 86400000;
  const issues = [];
  rows.forEach(b => {
    const start = dateHour_(b.date, b.hour);
    if (start.getTime() < now || start.getTime() > max) return;
    let refs;
    try { refs = JSON.parse(b.refs); } catch (_) { refs = []; }
    for (const ref of refs) {
      try {
        const cal = CalendarApp.getCalendarById(ref.calendarId);
        if (!cal || !cal.getEventById(ref.eventId)) issues.push(b.id + ' missing event in ' + ref.calendarId);
      } catch (e) { issues.push(b.id + ' error in ' + ref.calendarId + ': ' + e.message); }
    }
    // Validate stored event references, rather than today's configured destination: a changed
    // BookingCalendarId must never invalidate or move previously confirmed bookings.
    if (!refs.length) issues.push(b.id + ' has no saved Google Calendar event reference.');
  });
  if (issues.length) {
    MailApp.sendEmail({to: settings_().CoachEmail, subject: 'Coach Aimi: calendar sync audit needs attention',
      body: issues.slice(0, 100).join('\n'), name: 'Coach Aimi Booking'});
  }
}


/**
 * Coach administration. Visitors never choose an admin email address: the OTP is sent
 * only to CoachEmail from the private Settings sheet. No admin passwords/tokens go
 * into public GitHub Pages files.
 */
function requestAdminCode() {
  const email = String(settings_().CoachEmail || '').trim();
  if (!/^\S+@\S+\.\S+$/.test(email)) throw new Error('Set CoachEmail in the Settings sheet first.');
  const cache = CacheService.getScriptCache();
  const limiter = 'admin-rate';
  const hits = Number(cache.get(limiter) || 0);
  if (hits >= 3) return {message: 'If coach login is available, a code has been requested. Try again later if needed.'};
  cache.put(limiter, String(hits + 1), 900);
  const code = String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000).padStart(6, '0');
  cache.put('admin-otp', JSON.stringify({hash: hash_('coach-admin:' + code), attempts: 0}), 300);
  MailApp.sendEmail({
    to: email,
    subject: 'Coach Aimi — admin login code',
    body: 'Your Coach Admin sign-in code is ' + code + '. It expires in 5 minutes. Never share it.\n\nIf you did not request this, ignore this email.',
    name: 'Coach Aimi Booking', replyTo: String(settings_().ReplyToEmail || email)
  });
  return {message: 'A 6-digit login code was sent to your coach notification email.'};
}

function verifyAdminCode(inputCode) {
  const cache = CacheService.getScriptCache();
  const stored = cache.get('admin-otp');
  if (!stored || !/^\d{6}$/.test(String(inputCode))) throw new Error('Invalid or expired admin code.');
  const record = JSON.parse(stored);
  if (record.attempts >= 5) {cache.remove('admin-otp'); throw new Error('Too many attempts. Request a new code later.');}
  if (record.hash !== hash_('coach-admin:' + String(inputCode))) {
    record.attempts++;
    if (record.attempts >= 5) cache.remove('admin-otp');
    else cache.put('admin-otp', JSON.stringify(record), 300);
    throw new Error('Invalid or expired admin code.');
  }
  cache.remove('admin-otp');
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('admin-session-' + hash_(token).slice(0, 48), 'coach', 21600);
  return {token, email: settings_().CoachEmail};
}

function requireAdmin_(token) {
  if (typeof token !== 'string' || !/^[0-9a-f]{64}$/i.test(token) ||
      CacheService.getScriptCache().get('admin-session-' + hash_(token).slice(0, 48)) !== 'coach')
    throw new Error('Admin session expired. Please sign in again.');
}

function adminGetData(token) {
  requireAdmin_(token);
  const bookedByPhone = {};
  bookingRows_().forEach(b => {
    if (b.status === 'CONFIRMED') bookedByPhone[b.phone] = (bookedByPhone[b.phone] || 0) + 1;
  });
  const clients = values_(APP.CLIENTS).filter(r => r[2]).map(r => {
    const phone = normalizePhone_(r[2]);
    const granted = Number(r[4]) || 0;
    const used = bookedByPhone[phone] || 0;
    return {name: String(r[0]), email: String(r[1]), phone,
      active: enabled_(r[3]), creditsGranted: granted,
      creditsUsed: used, remaining: Math.max(0, granted - used)};
  });
  const destinationId = String(settings_().BookingCalendarId || '');
  const calendars = values_(APP.CALENDARS).filter(r => r[0]).map(r => ({
    id: String(r[0]).trim(), label: String(r[1] || ''), enabled: enabled_(r[2])
  }));
  return {clients, calendars, destinationId, sheetUrl: spreadsheet_().getUrl(),
    adminEmail: settings_().CoachEmail};
}

function adminSaveClient(token, data) {
  requireAdmin_(token);
  if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('Invalid client details.');
  const name = String(data.name || '').trim();
  const email = String(data.email || '').trim();
  const phone = normalizePhone_(data.phone);
  const active = data.active === true;
  const initialCredits = Number(data.initialCredits || 0);
  if (!name || name.length > 100 || /^[=+@-]/.test(name)) throw new Error('Enter a valid client name.');
  if (email.length > 200 || !/^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(email) || /^[=+@-]/.test(email))
    throw new Error('Enter a valid client email.');
  if (!Number.isSafeInteger(initialCredits) || initialCredits < 0 || initialCredits > 10000)
    throw new Error('Starting credits must be between 0 and 10,000.');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sheet = tab_(APP.CLIENTS);
    const rows = values_(APP.CLIENTS);
    const indices = rows.map((r,i) => {try{return normalizePhone_(r[2]) === phone ? i + 2 : 0;}catch(_){return 0;}}).filter(Boolean);
    if (indices.length > 1) throw new Error('Duplicate phone number in Clients sheet. Fix duplicates before editing.');
    if (indices.length) {
      const current = sheet.getRange(indices[0], 5).getValue();
      sheet.getRange(indices[0], 1, 1, 4).setValues([[name,email,phone,active]]);
      SpreadsheetApp.flush();
      return {message:'Client details saved. Existing credits (' + current + ') were preserved.'};
    }
    sheet.appendRow([name,email,phone,active,initialCredits]);
    SpreadsheetApp.flush();
    return {message:'Client added with ' + initialCredits + ' starting credits.'};
  } finally {lock.releaseLock();}
}

function adminGrantCredits(token, phoneInput, amountInput) {
  requireAdmin_(token);
  const phone = normalizePhone_(phoneInput);
  const amount = Number(amountInput);
  if (!Number.isSafeInteger(amount) || amount < 1 || amount > 1000)
    throw new Error('Credit top-up must be 1–1,000 sessions.');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sheet = tab_(APP.CLIENTS);
    const rows = values_(APP.CLIENTS);
    const matches = rows.map((r,i) => {try{return normalizePhone_(r[2]) === phone ? i + 2 : 0;}catch(_){return 0;}}).filter(Boolean);
    if (matches.length !== 1) throw new Error('Client not found or duplicate phone numbers detected.');
    const old = Number(sheet.getRange(matches[0],5).getValue() || 0);
    if (!Number.isSafeInteger(old) || old < 0 || old + amount > 100000)
      throw new Error('Invalid existing credit value.');
    sheet.getRange(matches[0],5).setValue(old + amount);
    SpreadsheetApp.flush();
    return {message: amount + ' credits added. Total granted: ' + (old + amount) + '.'};
  } finally {lock.releaseLock();}
}

function adminAddCalendar(token, inputId, inputLabel) {
  requireAdmin_(token);
  const id = String(inputId || '').trim();
  const label = String(inputLabel || '').trim();
  if (!id || id.length > 254 || /\s/.test(id) || !id.includes('@'))
    throw new Error('Enter a valid Calendar ID. An email works for an accessible primary Google Calendar.');
  if (label.length > 80 || /^[=+@-]/.test(label)) throw new Error('Enter a valid calendar label.');
  const cal = CalendarApp.getCalendarById(id);
  if (!cal) throw new Error('Cannot access this calendar. Share it with the Apps Script account and subscribe to it if needed. Secondary calendars require their exact Calendar ID.');
  if (id === String(settings_().BookingCalendarId || '').trim())
    return {message:'This is your booking destination and is already checked automatically.'};
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sheet = tab_(APP.CALENDARS);
    const rows = values_(APP.CALENDARS);
    const existing = rows.findIndex(r => String(r[0]).trim() === id);
    const name = label || cal.getName();
    if (existing >= 0) sheet.getRange(existing+2,1,1,3).setValues([[id,name,true]]);
    else sheet.appendRow([id,name,true]);
    SpreadsheetApp.flush();
    return {message:'Calendar added and enabled for busy-time checking: ' + name};
  } finally {lock.releaseLock();}
}

function adminSetCalendarEnabled(token, inputId, enabledInput) {
  requireAdmin_(token);
  const id = String(inputId || '').trim();
  if (!id || id === String(settings_().BookingCalendarId || '').trim())
    throw new Error('The booking destination calendar cannot be disabled.');
  const enabled = enabledInput === true;
  if (enabled && !CalendarApp.getCalendarById(id))
    throw new Error('This calendar is no longer accessible to the Apps Script account.');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sheet = tab_(APP.CALENDARS);
    const rows = values_(APP.CALENDARS);
    const index = rows.findIndex(r => String(r[0]).trim() === id);
    if (index < 0) throw new Error('Calendar not found.');
    sheet.getRange(index+2,3).setValue(enabled);
    SpreadsheetApp.flush();
    return {message: enabled ? 'Calendar enabled.' : 'Calendar disabled. Existing bookings remain unchanged.'};
  } finally {lock.releaseLock();}
}

/** Public API: anonymous availability returns only available hours, never client data. */
function getAvailability(dateString) {
  validateDate_(dateString);
  const policy = dayPolicy_(dateString);
  const today = new Date();
  const horizon = Number(settings_().BookingHorizonDays) || 60;
  const dayStart = dateHour_(dateString, 0);
  if (!policy.open || dayStart.getTime() > today.getTime() + horizon * 86400000) {
    return {date: dateString, open: false, slots: [], reason: policy.open ? 'Outside booking window' : 'Coach Aimi is off on this day'};
  }
  const calendars = calendars_(); // include all enabled calendars and booking destination
  const taken = bookingRows_().filter(b => b.status === 'CONFIRMED' && b.date === dateString).map(b => b.hour);
  const slots = [];
  for (let h = policy.first; h <= policy.last; h++) {
    const start = dateHour_(dateString, h), end = new Date(start.getTime() + 60 * 60000);
    if (start <= today) continue;
    const busy = taken.includes(h) || calendars.some(c => hasBusyEvent_(c.cal, start, end));
    if (!busy) slots.push({hour: h, label: fmtHour_(h) + ' – ' + fmtHour_(h + 1)});
  }
  return {date: dateString, open: true, slots, reason: slots.length ? '' : 'No available sessions on this date'};
}

/** Send a code only to the email already stored for this phone in Clients. */
function requestLoginCode(phoneInput) {
  const phone = normalizePhone_(phoneInput);
  const cache = CacheService.getScriptCache();
  const rateKey = 'rate-' + hash_(phone).slice(0, 28);
  const count = Number(cache.get(rateKey) || 0);
  if (count >= 3) return {message: 'If your number is registered, a code will be sent to your saved email.'};
  cache.put(rateKey, String(count + 1), 900);
  const client = findClient_(phone);
  if (client && client.active) {
    const code = String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000).padStart(6, '0');
    const codeKey = 'otp-' + hash_(phone).slice(0, 28);
    cache.put(codeKey, JSON.stringify({hash: hash_(code + ':' + phone), attempts: 0}), APP.LOGIN_TTL);
    MailApp.sendEmail({to: client.email, subject: 'Your Coach Aimi booking sign-in code',
      body: 'Your one-time sign-in code is ' + code + '. It expires in 5 minutes. Do not share this code.\n\nIf you did not request this, ignore this email.',
      name: 'Coach Aimi Booking', replyTo: settings_().ReplyToEmail});
  }
  return {message: 'If your number is registered, a code will be sent to your saved email.'};
}

function verifyLoginCode(phoneInput, inputCode) {
  const phone = normalizePhone_(phoneInput);
  const client = findClient_(phone);
  const cache = CacheService.getScriptCache();
  const codeKey = 'otp-' + hash_(phone).slice(0, 28);
  const raw = cache.get(codeKey);
  if (!client || !client.active || !raw || !/^\d{6}$/.test(String(inputCode))) throw new Error('Invalid or expired verification code.');
  const rec = JSON.parse(raw);
  if (rec.attempts >= 5) { cache.remove(codeKey); throw new Error('Too many attempts. Request a new code.'); }
  if (rec.hash !== hash_(String(inputCode) + ':' + phone)) {
    rec.attempts++;
    cache.put(codeKey, JSON.stringify(rec), Math.max(1, APP.LOGIN_TTL - 30));
    throw new Error('Invalid or expired verification code.');
  }
  cache.remove(codeKey);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('session-' + hash_(token).slice(0, 48), phone, APP.SESSION_TTL);
  return {token, client: {name: client.name, email: maskEmail_(client.email), creditsRemaining: creditBalance_(client).remaining}};
}

/** Client credit balance: total granted minus all confirmed (including past/no-show) sessions.
 * Cancellation more than 24 hours before the start marks a row CANCELLED, restoring its credit.
 * Rescheduling doesn't add or deduct a credit. Updating CreditsGranted in the private Sheet
 * is the only administrative top-up required. Never manually delete booking history rows.
 */
function creditBalance_(client) {
  const granted = client.creditsGranted;
  const used = bookingRows_().filter(b => b.phone === client.phone && b.status === 'CONFIRMED').length;
  return {granted, used, remaining: Math.max(0, granted - used)};
}
function getMyCredits(token) {
  return creditBalance_(authenticatedClient_(token));
}

function getMyBookings(token) {
  const client = authenticatedClient_(token);
  return bookingRows_().filter(b => b.phone === client.phone).map(publicBooking_).sort((a, b) =>
    (b.date + '-' + b.hour.toString().padStart(2, '0')).localeCompare(a.date + '-' + a.hour.toString().padStart(2, '0')));
}

function makeBooking(token, dateString, hour) {
  const client = authenticatedClient_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  let result;
  try {
    const h = validateSlot_(dateString, hour, false);
    if (creditBalance_(client).remaining < 1) throw new Error('You have no session credits remaining. Contact Coach Aimi to top up.');
    ensureSlotFree_(dateString, h);
    const id = Utilities.getUuid();
    const refs = createEvents_(client, dateString, h, id);
    const now = new Date().toISOString();
    try {
      tab_(APP.BOOKINGS).appendRow([id, client.phone, client.name, client.email, dateString, h, 'CONFIRMED', JSON.stringify(refs), now, now]);
      SpreadsheetApp.flush();
    } catch (e) { deleteEventRefs_(refs); throw e; }
    result = {id, date: dateString, hour: h, label: fmtHour_(h) + ' – ' + fmtHour_(h + 1), creditsRemaining: creditBalance_(client).remaining};
  } finally { lock.releaseLock(); }
  notifyBoth_(client, 'BOOKED', result);
  return {message: 'Your session is confirmed!', booking: result};
}

function rescheduleBooking(token, bookingId, dateString, hour) {
  const client = authenticatedClient_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  let result;
  try {
    const b = ownConfirmedBooking_(bookingId, client.phone);
    enforceCutoff_(b);
    const h = validateSlot_(dateString, hour, true);
    if (dateString === b.date && h === b.hour) throw new Error('Choose a different session time.');
    ensureSlotFree_(dateString, h);
    const newRefs = createEvents_(client, dateString, h, b.id);
    const sheet = tab_(APP.BOOKINGS);
    try {
      sheet.getRange(b.row, 5, 1, 4).setValues([[dateString, h, 'CONFIRMED', JSON.stringify(newRefs)]]);
      sheet.getRange(b.row, 10).setValue(new Date().toISOString());
      SpreadsheetApp.flush();
    } catch (e) { deleteEventRefs_(newRefs); throw e; }
    // If deleting old events fails, retain the new booking and notify the coach to fix duplication.
    const deletionErrors = deleteEventRefs_(parseRefs_(b.refs));
    if (deletionErrors.length) alertCoach_('Manual cleanup needed for reschedule ' + b.id + ': ' + deletionErrors.join('; '));
    result = {id: b.id, date: dateString, hour: h, previousDate: b.date, previousHour: b.hour,
      label: fmtHour_(h) + ' – ' + fmtHour_(h + 1)};
  } finally { lock.releaseLock(); }
  notifyBoth_(client, 'RESCHEDULED', result);
  return {message: 'Your session was rescheduled.', booking: result};
}

function cancelBooking(token, bookingId) {
  const client = authenticatedClient_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  let result;
  try {
    const b = ownConfirmedBooking_(bookingId, client.phone);
    enforceCutoff_(b);
    const sheet = tab_(APP.BOOKINGS);
    sheet.getRange(b.row, 7).setValue('CANCELLED');
    sheet.getRange(b.row, 10).setValue(new Date().toISOString());
    SpreadsheetApp.flush();
    const deletionErrors = deleteEventRefs_(parseRefs_(b.refs));
    if (deletionErrors.length) alertCoach_('Manual calendar cleanup needed for cancellation ' + b.id + ': ' + deletionErrors.join('; '));
    result = {id: b.id, date: b.date, hour: b.hour, label: fmtHour_(b.hour) + ' – ' + fmtHour_(b.hour + 1), creditsRemaining: creditBalance_(client).remaining};
  } finally { lock.releaseLock(); }
  notifyBoth_(client, 'CANCELLED', result);
  return {message: 'Your session was cancelled.', booking: result};
}

function addTab_(ss, name, header) {
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, header.length).setBackground('#16332b').setFontColor('#ffffff').setFontWeight('bold');
  sheet.autoResizeColumns(1, header.length);
  return sheet;
}
function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Booking system not set up yet. Coach: run setupProject_ in Apps Script editor.');
  return SpreadsheetApp.openById(id);
}
function tab_(name) { return spreadsheet_().getSheetByName(name); }
function values_(name) { return tab_(name).getDataRange().getValues().slice(1); }
function enabled_(v) { return v === true || String(v).trim().toLowerCase() === 'true' || String(v).trim().toLowerCase() === 'yes'; }
function settings_() { const obj = {}; values_(APP.SETTINGS).forEach(r => { if (r[0]) obj[String(r[0]).trim()] = String(r[1]).trim(); }); return obj; }
function schedule_() {
  const rows = values_(APP.SCHEDULE);
  const map = {};
  rows.forEach(r => { if (r[0]) map[String(r[0]).toLowerCase()] = {open: enabled_(r[1]), first: Number(r[2]), last: Number(r[3])}; });
  return map;
}
/** Booking destination is writable; all other enabled calendars are read-only busy blockers. */
function bookingCalendar_() {
  const id = String(settings_().BookingCalendarId || '').trim();
  if (!id) throw new Error('Coach must set BookingCalendarId in Settings.');
  const cal = CalendarApp.getCalendarById(id);
  if (!cal) throw new Error('The booking destination calendar is not accessible. Check sharing permissions.');
  return {id, cal};
}
/** Availability includes every enabled calendar AND the booking destination, without duplicates. */
function calendars_() {
  const destination = bookingCalendar_();
  const rows = values_(APP.CALENDARS).filter(r => enabled_(r[2]));
  const ids = rows.map(r => String(r[0]).trim());
  if (ids.some(id => !id) || new Set(ids).size !== ids.length)
    throw new Error('Each enabled availability calendar needs a unique, nonempty Calendar ID.');
  const additional = ids.filter(id => id !== destination.id).map(id => {
    const cal = CalendarApp.getCalendarById(id);
    if (!cal) throw new Error('An availability calendar is not accessible. Check its ID or sharing permissions.');
    return {id, cal};
  });
  return [destination].concat(additional);
}
function findClient_(phone) {
  const rows = values_(APP.CLIENTS);
  for (const r of rows) {
    let rowPhone;
    try { rowPhone = normalizePhone_(r[2]); } catch (_) { continue; }
    if (rowPhone !== phone) continue;
    const email = String(r[1]).trim();
    if (!email || !/^\S+@\S+\.\S+$/.test(email)) return null;
    // Blank credits mean zero: no free sessions when an admin forgets to grant credits.
    const rawCredits = r[4];
    const creditsGranted = rawCredits === '' ? 0 : Number(rawCredits);
    if (!Number.isSafeInteger(creditsGranted) || creditsGranted < 0)
      throw new Error('Coach must set CreditsGranted to a non-negative whole number for this client.');
    return {name: String(r[0]).trim(), email, phone, active: enabled_(r[3]), creditsGranted};
  }
  return null;
}
function normalizePhone_(phone) {
  let s = String(phone == null ? '' : phone).replace(/[^\d+]/g, '').replace(/^00/, '+');
  if (s.startsWith('+')) s = s.slice(1);
  if (s.startsWith('0')) s = '60' + s.slice(1);
  if (!/^\d{9,15}$/.test(s)) throw new Error('Enter a valid phone number, e.g. 0123456789.');
  return s;
}
function hash_(s) {
  const pepper = PropertiesService.getScriptProperties().getProperty('OTP_PEPPER') || initPepper_();
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pepper + ':' + s)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function initPepper_() {
  const key = Utilities.getUuid() + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('OTP_PEPPER', key);
  return key;
}
function authenticatedClient_(token) {
  if (typeof token !== 'string' || !/^[0-9a-f]{64}$/i.test(token)) throw new Error('Your session expired. Sign in again.');
  const phone = CacheService.getScriptCache().get('session-' + hash_(token).slice(0, 48));
  const client = phone && findClient_(phone);
  if (!client || !client.active) throw new Error('Your session expired. Sign in again.');
  return client;
}
function maskEmail_(email) { const parts = email.split('@'); return parts[0].slice(0, 2) + '***@' + parts[1]; }
function bookingRows_() {
  return values_(APP.BOOKINGS).map((r, i) => ({row: i + 2, id: String(r[0]), phone: String(r[1]),
    name: String(r[2]), email: String(r[3]), date: String(r[4]), hour: Number(r[5]),
    status: String(r[6]), refs: String(r[7] || ''), createdAt: String(r[8] || '')})).filter(b => b.id);
}
function publicBooking_(b) {
  return {id: b.id, date: b.date, hour: b.hour, label: fmtHour_(b.hour) + ' – ' + fmtHour_(b.hour + 1),
    status: b.status, canEdit: b.status === 'CONFIRMED' && dateHour_(b.date, b.hour).getTime() - Date.now() >= Number(settings_().RescheduleCutoffHours || 24) * 3600000};
}
function ownConfirmedBooking_(id, phone) {
  const b = bookingRows_().find(b => b.id === String(id) && b.phone === phone);
  if (!b || b.status !== 'CONFIRMED') throw new Error('Booking not found or not editable.');
  return b;
}
function enforceCutoff_(b) {
  if (dateHour_(b.date, b.hour).getTime() - Date.now() < Number(settings_().RescheduleCutoffHours || 24) * 3600000)
    throw new Error('Changes must be made at least 24 hours before your booked session.');
}
function validateDate_(s) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('Invalid date.');
  const [y, m, d] = s.split('-').map(Number);
  const candidate = new Date(y, m - 1, d);
  if (candidate.getFullYear() !== y || candidate.getMonth() !== m - 1 || candidate.getDate() !== d) throw new Error('Invalid date.');
}
function dateHour_(dateString, h) {
  const [y, m, d] = dateString.split('-').map(Number);
  return new Date(y, m - 1, d, Number(h), 0, 0, 0);
}
function dayPolicy_(dateString) {
  const weekday = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'][dateHour_(dateString, 12).getDay()];
  const policy = schedule_()[weekday];
  if (!policy) throw new Error('Coach must complete the Schedule sheet.');
  return policy;
}
function validateSlot_(dateString, hour, isChange) {
  validateDate_(dateString);
  const h = Number(hour);
  if (!Number.isInteger(h)) throw new Error('Choose a valid hour.');
  const policy = dayPolicy_(dateString);
  if (!policy.open || !Number.isInteger(policy.first) || !Number.isInteger(policy.last) || h < policy.first || h > policy.last || h < 0 || h > 22)
    throw new Error('This session is outside Coach Aimi’s current working hours.');
  const now = Date.now(), start = dateHour_(dateString, h).getTime();
  const horizon = Number(settings_().BookingHorizonDays) || 60;
  if (start <= now || start > now + horizon * 86400000) throw new Error('This booking date is not available.');
  if (isChange && start - now < Number(settings_().RescheduleCutoffHours || 24) * 3600000)
    throw new Error('Please choose a new session at least 24 hours from now.');
  return h;
}
function hasBusyEvent_(cal, start, end) {
  return cal.getEvents(start, end).some(ev => ev.getTransparency() !== CalendarApp.EventTransparency.TRANSPARENT);
}
function ensureSlotFree_(date, hour) {
  if (bookingRows_().some(b => b.status === 'CONFIRMED' && b.date === date && b.hour === hour))
    throw new Error('Sorry, that hour has just been booked. Choose another slot.');
  const start = dateHour_(date, hour), end = new Date(start.getTime() + APP.SESSION_MINUTES * 60000);
  if (calendars_().some(c => hasBusyEvent_(c.cal, start, end)))
    throw new Error('Sorry, that hour is now busy on the coach’s calendar. Choose another slot.');
}
function createEvents_(client, date, hour, id) {
  const start = dateHour_(date, hour), end = new Date(start.getTime() + APP.SESSION_MINUTES * 60000);
  const destination = bookingCalendar_();
  try {
    const event = destination.cal.createEvent('PT • ' + client.name + ' with Coach Aimi', start, end, {
      description: 'Coach Aimi Personal Training\nBooking ID: ' + id + '\nClient: ' + client.name +
        '\nEmail: ' + client.email + '\nPhone: +' + client.phone + '\nManaged from Coach Aimi booking app.'
    });
    return [{calendarId: destination.id, eventId: event.getId()}];
  } catch (e) { throw new Error('Could not save the booking to ' + destination.id + '. ' + e.message); }
}
function parseRefs_(raw) { try { return JSON.parse(raw) || []; } catch (_) { return []; } }
function deleteEventRefs_(refs) {
  const errors = [];
  refs.forEach(ref => {
    try {
      const cal = CalendarApp.getCalendarById(ref.calendarId);
      const ev = cal && cal.getEventById(ref.eventId);
      if (ev) ev.deleteEvent();
    } catch (e) { errors.push(ref.calendarId + ': ' + e.message); }
  });
  return errors;
}
function fmtHour_(hour) {
  const h = Number(hour) % 24;
  return (h % 12 || 12) + ':00 ' + (h < 12 ? 'AM' : 'PM');
}
function alertCoach_(details) {
  try { MailApp.sendEmail({to: settings_().CoachEmail, subject: 'Coach Aimi booking needs manual attention', body: details}); }
  catch (e) { console.error('Could not send coach alert: ' + e.message); }
}
function notifyBoth_(client, action, booking) {
  const cfg = settings_();
  const when = booking.date + ', ' + booking.label + ' (Malaysia time)';
  const previous = booking.previousDate ? '\nPrevious session: ' + booking.previousDate + ' at ' + fmtHour_(booking.previousHour) + '\n' : '';
  const details = 'Client: ' + client.name + '\nEmail: ' + client.email + '\nPhone: +' + client.phone +
    '\nBooking reference: ' + booking.id + '\nSession: ' + when + '\n' + previous;
  const subject = 'PT session ' + action.toLowerCase() + ' — ' + booking.date;
  for (const email of [...new Set([client.email, cfg.CoachEmail].filter(Boolean))]) {
    try {
      MailApp.sendEmail({to: email, subject, body: 'Coach Aimi Personal Training\n\nStatus: ' + action + '\n\n' + details +
        '\nYou can manage eligible bookings via the booking website. Changes require at least 24 hours notice.\n',
        name: 'Coach Aimi Booking', replyTo: cfg.ReplyToEmail});
    } catch (e) { console.error('Booking saved but notification failed to ' + email + ': ' + e.message); }
  }
}