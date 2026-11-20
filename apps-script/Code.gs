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
/**
 * Top-level HTML Service pages work without cross-origin iframe embedding.
 * GitHub Pages remains the public landing page and the source of these templates.
 * Backend operations run through Google's supported google.script.run API.
 */
function doGet(e) {
  // Coach login is the default /exec page so it does not depend on URL parameters.
  const page = e && e.parameter && (e.parameter.page || e.parameter.view);
  if (page === 'client') {
    return HtmlService.createHtmlOutput(COACH_CLIENT_PAGE_)
      .setTitle('Coach Aimi — Client Booking');
  }
  return HtmlService.createHtmlOutput(COACH_ADMIN_PAGE_)
    .setTitle('Coach Aimi — Coach Admin');
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
/**
 * Email-only coach access, as requested. The email must match CoachEmail in Settings.
 * CAUTION: This is NOT identity verification: anyone who knows this address can get access.
 * Use only if you knowingly accept this privacy risk.
 */
function adminLogin(emailInput) {
  const supplied = String(emailInput || '').trim().toLowerCase();
  const coach = String(settings_().CoachEmail || '').trim().toLowerCase();
  if (!/^\S+@\S+\.\S+$/.test(coach)) throw new Error('CoachEmail is not configured in Settings.');
  if (supplied !== coach) throw new Error('Email not recognized.');
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('admin-session-' + hash_(token).slice(0, 48), 'coach', APP.SESSION_TTL);
  return {token, email: coach};
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
  const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const config = schedule_();
  const schedule = days.map(day => ({
    day, open: !!config[day.toLowerCase()].open,
    first: config[day.toLowerCase()].first, last: config[day.toLowerCase()].last
  }));
  const upcoming = bookingRows_()
    .filter(b => b.status === 'CONFIRMED' && dateHour_(b.date,b.hour).getTime() > Date.now())
    .sort((a,b) => a.date.localeCompare(b.date) || a.hour-b.hour).slice(0,25)
    .map(b => ({name:b.name, date:b.date, time:fmtHour_(b.hour), hour:b.hour}));
  return {clients, calendars, destinationId, schedule, upcoming, sheetUrl: spreadsheet_().getUrl(),
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

/**
 * Update weekly availability in one operation, without altering existing bookings.
 * Day names and hours are validated on the server.
 */
function adminSaveSchedule(token, submitted) {
  requireAdmin_(token);
  if (!Array.isArray(submitted) || submitted.length !== 7)
    throw new Error('Please provide all seven days of the week.');
  const names = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const values = names.map(day => {
    const entry = submitted.find(x => x && x.day === day);
    if (!entry) throw new Error('Missing schedule for ' + day);
    const first = Number(entry.first), last = Number(entry.last);
    if (!Number.isInteger(first) || !Number.isInteger(last) || first < 0 || first > 22 ||
        last < first || last > 22)
      throw new Error('Invalid working hours for ' + day + '. Latest start must be 10 PM or earlier.');
    return [day, entry.open === true, first, last];
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    tab_(APP.SCHEDULE).getRange(2, 1, 7, 4).setValues(values);
    SpreadsheetApp.flush();
    return {message:'Schedule updated. Existing confirmed bookings remain valid.'};
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

// GitHub UI snapshots for top-level HTML Service (no separate Apps Script HTML files).
const COACH_ADMIN_PAGE_ = "<!doctype html>\n<html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<title>Coach Admin | Coach Aimi</title>\n<meta name=\"robots\" content=\"noindex,nofollow\">\n<style>\n:root{font-family:Inter,system-ui,-apple-system,\"Segoe UI\",Arial,sans-serif;color:#edf8ec;background:#0b1715}\n*{box-sizing:border-box}body{margin:0;background:radial-gradient(circle at 90% 0%,#255740 0%,transparent 38%),#0b1715;min-height:100vh}\na{color:#b9f774}button,input{font:inherit}button{cursor:pointer}button:disabled{opacity:.55;cursor:wait}\nmain{max-width:1040px;margin:auto;padding:28px 18px 80px}\nheader{display:flex;justify-content:space-between;align-items:center;gap:16px;margin-bottom:55px}\n.brand{font-size:19px;font-weight:900;letter-spacing:-1px}.brand span{color:#b9f774}\nnav a{font-size:13px}.hero{margin-bottom:24px}.kicker{color:#b9f774;text-transform:uppercase;letter-spacing:2px;font-size:11px;font-weight:800}\nh1{font-size:clamp(34px,6vw,60px);letter-spacing:-.06em;margin:8px 0}p{line-height:1.6}\n.panel{background:#f7fbf5;color:#17291d;padding:30px;border-radius:20px;margin:18px 0;box-shadow:0 10px 40px #0003}\n.panel h2{margin:0 0 8px;letter-spacing:-.03em}.muted{color:#67806b;font-size:13px}\nform{display:grid;gap:14px;margin-top:16px}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}\nlabel{font-size:12px;font-weight:700;color:#45604b;display:block}\ninput{margin-top:7px;display:block;width:100%;padding:13px;border:1px solid #d4e3d1;border-radius:10px;background:#fff;color:#17291d}\ninput[type=checkbox]{width:auto;display:inline-block;margin:0 8px 0 0}\nbutton{border:0;border-radius:10px;padding:13px 16px;font-weight:800}\n.primary{background:#b9f774;color:#19331b}.secondary{background:#e7f1e1;color:#254a2c}.danger{background:#fce7e7;color:#8d2020}\n.flex{display:flex;align-items:center;gap:9px;flex-wrap:wrap}.between{justify-content:space-between}\n.row{border-top:1px solid #dce8da;padding:15px 0;display:flex;justify-content:space-between;gap:18px;align-items:center;flex-wrap:wrap}\n.row strong{display:block}.row small{color:#647b6b}.row .actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}\n.row input[type=number]{width:85px;margin:0}\n.tabs{display:flex;gap:8px;flex-wrap:wrap;margin:25px 0 5px}\n.tabs button{background:#233f31;color:#ddf5de}.tabs button.active{background:#b9f774;color:#18311c}\n.notice{background:#e6f4dc;color:#285b30;padding:14px 18px;border-radius:12px;margin:12px 0;font-size:13px}\n.notice.error{background:#ffeded;color:#8d2929}.hidden{display:none!important}\n@media(max-width:650px){.grid{grid-template-columns:1fr}.panel{padding:20px}header{margin-bottom:32px}.row{align-items:start}}\n</style></head><body><main>\n<header><div class=\"brand\">COACH <span>AIMI.</span> / ADMIN</div><nav><a href=\"https://aimi-keboon.github.io/coach-aimi-pt-booking/\">Client booking page</a></nav></header>\n<div class=\"hero\"><div class=\"kicker\">PRIVATE MANAGEMENT</div><h1>Your coaching dashboard.</h1><p>Manage clients, session credits, and calendars that affect your availability.</p></div>\n<div id=\"notice\" class=\"notice hidden\" role=\"status\" aria-live=\"polite\"></div>\n<section class=\"panel\" id=\"loginPanel\">\n<h2>Sign in as Coach Aimi</h2><p class=\"muted\">We'll send a one-time login code to the coach email configured in your private Google Sheet. This page does not use a public admin password.</p>\n<div class=\"flex\"><button id=\"sendCode\" class=\"primary\">Email me a sign-in code</button></div>\n<form id=\"loginForm\" class=\"hidden\"><label>6-digit email code<input id=\"code\" inputmode=\"numeric\" autocomplete=\"one-time-code\" pattern=\"[0-9]{6}\" maxlength=\"6\" required></label><div><button class=\"primary\">Verify and enter dashboard</button></div></form>\n</section>\n<section id=\"dashboard\" class=\"hidden\">\n<div class=\"flex between\"><p id=\"adminLabel\"></p><button id=\"logout\" class=\"secondary\">Sign out</button></div>\n<div class=\"tabs\"><button id=\"clientsTab\" class=\"active\">Clients & credits</button><button id=\"calendarsTab\">Calendars</button></div>\n<div id=\"clientsPanel\">\n<section class=\"panel\"><h2>Add or update a client</h2><p class=\"muted\">Use the same phone number to update an existing client. Existing credit balances will not be reset by an update.</p>\n<form id=\"clientForm\"><div class=\"grid\">\n<label>Client name<input id=\"clientName\" maxlength=\"100\" required placeholder=\"Client name\"></label>\n<label>Client email<input id=\"clientEmail\" type=\"email\" required placeholder=\"client@example.com\"></label>\n<label>Phone number<input id=\"clientPhone\" type=\"tel\" required placeholder=\"0123456789\"></label>\n<label>Starting credits (new clients only)<input id=\"initialCredits\" type=\"number\" min=\"0\" max=\"10000\" step=\"1\" value=\"0\" required></label>\n</div><label><input type=\"checkbox\" id=\"clientActive\" checked> Active client — can log in and book</label>\n<div><button class=\"primary\">Save client</button></div></form></section>\n<section class=\"panel\"><div class=\"flex between\"><h2>Current clients</h2><button id=\"refreshClients\" class=\"secondary\">Refresh</button></div>\n<p class=\"muted\">Credits remaining = total granted minus confirmed sessions. Add credits using the + button for each client.</p>\n<div id=\"clientRows\">Loading...</div></section>\n</div>\n<div id=\"calendarsPanel\" class=\"hidden\">\n<section class=\"panel\"><h2>Availability calendars</h2><p class=\"muted\">All enabled calendars block busy times. New sessions are saved only in your fixed booking destination calendar.</p><p id=\"destination\"></p>\n<form id=\"calendarForm\"><div class=\"grid\">\n<label>Google Calendar ID or primary-calendar email<input id=\"calendarId\" required maxlength=\"254\" placeholder=\"e.g. another.name@gmail.com\"></label>\n<label>Label (optional)<input id=\"calendarLabel\" maxlength=\"80\" placeholder=\"Personal calendar\"></label>\n</div><div><button class=\"primary\">Add calendar & enable</button></div></form>\n<p class=\"muted\">An email address usually works for an accessible primary Google Calendar. Other calendars often have IDs ending in @group.calendar.google.com. The Google account running Apps Script must already have access; entering an email does not share a calendar automatically.</p></section>\n<section class=\"panel\"><h2>Connected calendars</h2><div id=\"calendarRows\">Loading...</div></section>\n</div>\n</section>\n<script>\n'use strict';\nconst $=id=>document.getElementById(id);\nlet token='';\nfunction notice(s,bad=false){$('notice').textContent=s;$('notice').classList.remove('hidden');$('notice').classList.toggle('error',bad);}\nfunction safeRun(fn){return async e=>{if(e)e.preventDefault();try{await fn(e);}catch(err){notice(err.message||String(err),true);}}}\nfunction call(name,...args) { return new Promise((resolve,reject) => {\n    const runner=google.script.run.withSuccessHandler(resolve).withFailureHandler(err=>reject(new Error(err&&err.message?err.message:'Booking service error')));\n    try {switch(name) {\n        case 'requestAdminCode': runner.requestAdminCode(); break;\n        case 'verifyAdminCode': runner.verifyAdminCode(args[0]); break;\n        case 'adminGetData': runner.adminGetData(args[0]); break;\n        case 'adminSaveClient': runner.adminSaveClient(args[0],args[1]); break;\n        case 'adminGrantCredits': runner.adminGrantCredits(args[0],args[1],args[2]); break;\n        case 'adminAddCalendar': runner.adminAddCalendar(args[0],args[1],args[2]); break;\n        case 'adminSetCalendarEnabled': runner.adminSetCalendarEnabled(args[0],args[1],args[2]); break;\n        default: reject(new Error('Unsupported operation.'));\n      }} catch (e) {reject(e);}\n  }); }\nasync function busy(button,fn){const old=button.textContent;button.disabled=true;button.textContent='Working...';try{await fn();}finally{button.disabled=false;button.textContent=old;}}\n$('sendCode').addEventListener('click',safeRun(async()=>busy($('sendCode'),async()=>{const r=await call('requestAdminCode');notice(r.message);$('loginForm').classList.remove('hidden');})));\n$('loginForm').addEventListener('submit',safeRun(async()=>{const b=$('loginForm').querySelector('button');await busy(b,async()=>{const r=await call('verifyAdminCode',$('code').value.trim());token=r.token;$('loginPanel').classList.add('hidden');$('dashboard').classList.remove('hidden');$('adminLabel').textContent='Signed in as '+r.email;notice('Welcome, Coach Aimi.');await refresh();});}));\n$('logout').onclick=()=>{token='';$('dashboard').classList.add('hidden');$('loginPanel').classList.remove('hidden');$('code').value='';notice('Signed out of this browser.');};\nfunction textNode(tag,s){const n=document.createElement(tag);n.textContent=s;return n;}\nfunction choose(tab){$('clientsPanel').classList.toggle('hidden',tab!=='clients');$('calendarsPanel').classList.toggle('hidden',tab!=='calendars');$('clientsTab').classList.toggle('active',tab==='clients');$('calendarsTab').classList.toggle('active',tab==='calendars');}\n$('clientsTab').onclick=()=>choose('clients');$('calendarsTab').onclick=()=>choose('calendars');\nasync function refresh(){const d=await call('adminGetData',token);renderClients(d.clients);renderCalendars(d.calendars,d.destinationId);}\nfunction renderClients(clients){const root=$('clientRows');root.replaceChildren();if(!clients.length){root.append(textNode('p','No clients yet. Add one above.'));return;}\n for(const c of clients){const row=document.createElement('div');row.className='row';const info=document.createElement('div');info.append(textNode('strong',c.name),textNode('small',c.email+' · +'+c.phone+' · '+(c.active?'Active':'Inactive')),textNode('div',c.remaining+' remaining · '+c.creditsUsed+' used · '+c.creditsGranted+' granted'));const actions=document.createElement('div');actions.className='actions';const qty=document.createElement('input');qty.type='number';qty.min='1';qty.max='1000';qty.step='1';qty.value='1';qty.setAttribute('aria-label','Credits to add for '+c.name);const add=document.createElement('button');add.className='primary';add.textContent='+ Add credits';add.onclick=safeRun(async()=>busy(add,async()=>{if(!confirm('Grant '+qty.value+' additional credits to '+c.name+'?'))return;const r=await call('adminGrantCredits',token,c.phone,Number(qty.value));notice(r.message);await refresh();}));const edit=document.createElement('button');edit.className='secondary';edit.textContent='Edit details';edit.onclick=()=>{$('clientName').value=c.name;$('clientEmail').value=c.email;$('clientPhone').value='+'+c.phone;$('clientActive').checked=c.active;$('initialCredits').value=0;$('clientName').focus();notice('Editing '+c.name+'. Save client to update details; existing credits stay unchanged.');};actions.append(qty,add,edit);row.append(info,actions);root.append(row);}}\nfunction renderCalendars(calendars,destinationId){$('destination').textContent='Booking destination (always checked): '+destinationId;const root=$('calendarRows');root.replaceChildren();if(!calendars.length)root.append(textNode('p','No extra availability calendars yet.'));for(const c of calendars){const row=document.createElement('div');row.className='row';const info=document.createElement('div');info.append(textNode('strong',c.label||'Availability calendar'),textNode('small',c.id+' · '+(c.enabled?'Enabled':'Disabled')));const actions=document.createElement('div');actions.className='actions';const b=document.createElement('button');b.className=c.enabled?'danger':'primary';b.textContent=c.enabled?'Disable':'Enable';b.onclick=safeRun(async()=>busy(b,async()=>{const r=await call('adminSetCalendarEnabled',token,c.id,!c.enabled);notice(r.message);await refresh();}));actions.append(b);row.append(info,actions);root.append(row);}}\n$('clientForm').addEventListener('submit',safeRun(async()=>{const b=$('clientForm').querySelector('button');await busy(b,async()=>{const r=await call('adminSaveClient',token,{name:$('clientName').value,email:$('clientEmail').value,phone:$('clientPhone').value,active:$('clientActive').checked,initialCredits:Number($('initialCredits').value)});notice(r.message);$('clientForm').reset();await refresh();});}));\n$('calendarForm').addEventListener('submit',safeRun(async()=>{const b=$('calendarForm').querySelector('button');await busy(b,async()=>{const r=await call('adminAddCalendar',token,$('calendarId').value,$('calendarLabel').value);notice(r.message);$('calendarForm').reset();await refresh();});}));\n$('refreshClients').addEventListener('click',safeRun(refresh));\n</script></main></body></html>";
const COACH_CLIENT_PAGE_ = "<!doctype html>\n<html lang=\"en\">\n<head>\n  <meta charset=\"utf-8\">\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n  <meta name=\"theme-color\" content=\"#0b1715\">\n  <title>Coach Aimi — Book a Personal Training Session</title>\n  <style>\n    :root{font-family:Inter,ui-sans-serif,system-ui,-apple-system,\"Segoe UI\",Arial,sans-serif;color:#eff6ed;background:#0b1715;font-synthesis:none}\n    *{box-sizing:border-box}body{margin:0;min-height:100vh;background:radial-gradient(ellipse 75% 45% at 85% 4%,#225b45 0%,transparent 65%),#0b1715}\n    button,input{font:inherit}button{cursor:pointer}button:disabled{cursor:not-allowed;opacity:.5}\n    .wrap{max-width:1060px;margin:0 auto;padding:26px 22px 76px}.top{display:flex;align-items:center;justify-content:space-between;gap:14px}\n    .logo{font-size:17px;font-weight:800;letter-spacing:-.7px}.logo span{color:#b8fb69}.tag{border:1px solid #3c5b4a;border-radius:100px;padding:9px 12px;font-size:11px;letter-spacing:1.1px;text-transform:uppercase;color:#bee1cb}\n    .hero{padding:66px 0 36px;display:grid;grid-template-columns:1.4fr .6fr;gap:20px;align-items:end}.eyebrow{font-size:11px;letter-spacing:2.8px;text-transform:uppercase;color:#b8fb69;font-weight:800;margin:0 0 18px}\n    h1{font-size:clamp(44px,8vw,86px);letter-spacing:-.085em;line-height:.97;max-width:740px;margin:0;font-weight:850}h1 em{font-style:normal;color:#b8fb69}\n    .desc{line-height:1.65;color:#abc2b4;font-size:14px;max-width:280px;margin-bottom:9px}.shell{background:#f8faf4;color:#18261f;border-radius:24px;box-shadow:0 18px 80px #0005;overflow:hidden}\n    .shellhead{display:flex;align-items:center;justify-content:space-between;padding:26px 30px;border-bottom:1px solid #e2e8df}.shellhead strong{font-size:17px;letter-spacing:-.5px}.shellhead small{display:block;color:#788579;margin-top:4px;font-size:12px}\n    .step{font-size:11px;border-radius:100px;background:#e8f1dc;color:#416624;padding:9px 12px;font-weight:700}.shellbody{padding:30px}\n    .grid{display:grid;grid-template-columns:1fr 1fr;gap:22px}.field{display:block;margin:0 0 16px;font-size:12px;color:#4f6053;font-weight:700}.field input{display:block;width:100%;height:52px;border:1px solid #d9e2d7;border-radius:12px;margin-top:9px;padding:0 14px;background:#fff;color:#172a20;outline:none}.field input:focus{border-color:#759f4d;box-shadow:0 0 0 3px #cdeba955}\n    .primary{display:inline-flex;align-items:center;justify-content:center;gap:9px;border:0;background:#b9f774;color:#142817;border-radius:12px;padding:16px 22px;font-weight:800;min-height:50px}.primary:hover:enabled{background:#a4eb5b}.secondary{background:#f0f4ed;border:1px solid #e1e8dd;color:#294332;border-radius:12px;padding:12px 17px;font-weight:700}\n    .hint{color:#768476;font-size:12px;line-height:1.6}.notice{border:1px solid #cce6ac;background:#f0fbe5;color:#3b5d23;border-radius:12px;padding:14px 16px;font-size:12px;line-height:1.6;margin:18px 0}\n    .error{border-color:#efc3c3;background:#fff0f0;color:#9c3838}.tabnav{display:flex;border-bottom:1px solid #dce5db;gap:7px;padding:0 30px}.tab{border:0;background:transparent;padding:19px 13px;font-weight:700;color:#7a877b;border-bottom:3px solid transparent}.tab.active{color:#284d2a;border-color:#84be52}.panel{padding:29px 30px 33px}\n    h2{font-size:24px;letter-spacing:-.9px;margin:0 0 7px}.sub{font-size:12px;line-height:1.6;color:#758578;margin:0 0 22px}.daybar{display:flex;gap:7px;overflow-x:auto;padding-bottom:7px;margin:0 0 20px}\n    .day{min-width:81px;border:1px solid #dfe7dc;border-radius:13px;background:white;padding:11px 7px;display:flex;align-items:center;flex-direction:column;gap:6px;color:#354b38}.day.active{border-color:#4b8f51;background:#e8f7db}.day b{font-size:18px}.day span{font-size:11px}\n    .slotgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:18px 0 22px}.slot{background:white;border:1px solid #dbe6d8;padding:15px 7px;font-weight:700;color:#294531;border-radius:12px;font-size:12px}.slot.active{border:2px solid #479547;background:#e9f9d6;padding:14px 6px}\n    .split{display:grid;grid-template-columns:1fr 260px;gap:25px}.summary{background:#eaf3e4;padding:23px;border-radius:16px;align-self:start}.summary small{font-size:10px;letter-spacing:1.4px;color:#658265;font-weight:800;text-transform:uppercase}.summary b{display:block;font-size:20px;letter-spacing:-.7px;margin:12px 0}.summary p{font-size:12px;color:#647563;line-height:1.6}.summary .primary{width:100%;margin-top:12px;font-size:12px;padding:14px}\n    .bookcard{border:1px solid #dfe7db;background:#fff;border-radius:15px;padding:19px;margin-bottom:11px;display:flex;justify-content:space-between;align-items:center;gap:15px}.bookcard strong{display:block;font-size:15px;margin:6px 0}.bookcard small{color:#6b806e}.pill{display:inline-block;font-size:10px;padding:5px 8px;border-radius:50px;color:#476f34;background:#edf8e5;font-weight:800;letter-spacing:.3px}.pill.cancelled{color:#8e5454;background:#f9eaea}.actions{display:flex;gap:7px;flex-wrap:wrap}\n    .muted{color:#8d9f8d}.hidden{display:none!important}.footer{display:flex;align-items:center;justify-content:space-between;gap:15px;color:#7e9987;font-size:11px;margin:27px 4px 0}.footer span{color:#c0e5bd}\n    /* Simple login-first homepage. Full-width session dashboard only after client sign-in. */\n    .compact-hero{padding:55px 0 24px;display:block;text-align:center}\n    .compact-hero h1{font-size:clamp(31px,5vw,46px);line-height:1.15;letter-spacing:-.05em;max-width:none}\n    .compact-hero .desc{font-size:14px;max-width:none;margin:11px 0 0;color:#b4cabb}\n    .shell{max-width:680px;margin:0 auto}\n    body.client-active .shell{max-width:1060px}\n    .loginselect{padding:34px 30px 36px}\n    .loginselect h2{margin:0 0 8px}\n    .loginselect .sub{margin-bottom:22px}\n    .loginchoices{display:grid;grid-template-columns:1fr 1fr;gap:13px}\n    .loginchoice{display:flex;flex-direction:column;align-items:flex-start;justify-content:center;text-decoration:none;text-align:left;min-height:118px;padding:22px;border-radius:15px;border:1px solid #dce6d7;background:#fff;color:#223a2b}\n    .loginchoice strong{font-size:18px;letter-spacing:-.3px}\n    .loginchoice span{font-size:12px;line-height:1.5;color:#6c806f;margin-top:9px;font-weight:400}\n    .loginchoice.client{background:#dff5cd;border-color:#a9dc86}\n    .loginchoice:hover,.loginchoice:focus-visible{outline:2px solid #82ba5c;outline-offset:1px}\n    .backlogin{margin-top:16px;background:transparent;color:#365d3b;padding:8px 0;text-decoration:underline;font-size:12px}\n    @media(max-width:580px){.compact-hero{padding:35px 8px 22px}.compact-hero h1{font-size:34px}.loginselect{padding:24px 20px}.loginchoices{grid-template-columns:1fr}.loginchoice{min-height:98px}}\n    @media(max-width:740px){.wrap{padding:18px 13px 50px}.hero{display:block;padding:53px 8px 26px}.desc{max-width:100%;margin-top:22px}.tag{font-size:9px}.shellhead{padding:21px}.shellbody,.panel{padding:20px}.tabnav{padding:0 16px}.split,.grid{grid-template-columns:1fr}.slotgrid{grid-template-columns:repeat(2,1fr)}.bookcard{align-items:flex-start;flex-direction:column}.footer{align-items:flex-start;flex-direction:column}h1{font-size:clamp(48px,13vw,65px)}}\n  </style>\n</head>\n<body>\n<div class=\"wrap\">\n  <header class=\"top\"><div class=\"logo\">COACH<span> AIMI.</span></div><div class=\"tag\">PT BOOKING</div></header>\n  <section class=\"hero compact-hero\"><h1>Personal training bookings</h1><p class=\"desc\">Book or manage your 1-hour sessions with Coach Aimi.</p></section>\n  <main class=\"shell\">\n    <section id=\"loginSelect\" class=\"loginselect hidden\">\n      <h2>Log in</h2><p class=\"sub\">Choose how you'd like to use the app.</p>\n      <div class=\"loginchoices\">\n        <button class=\"loginchoice client\" id=\"clientLogin\" type=\"button\"><strong>Client login →</strong><span>Book a session, see credits and manage bookings.</span></button>\n        <a class=\"loginchoice\" href=\"./admin.html\"><strong>Coach login →</strong><span>Manage clients, credits and calendars.</span></a>\n      </div>\n    </section>\n    <div class=\"shellhead\" id=\"clientShellHead\"><div><strong>Book your PT session</strong><small>1 hour · 7:00 AM – 11:00 PM · Malaysia time</small></div><span class=\"step\" id=\"stepLabel\">CLIENT SIGN IN</span></div>\n    <div id=\"message\" class=\"notice hidden\" role=\"status\" aria-live=\"polite\"></div>\n    <section class=\"shellbody\" id=\"phoneView\">\n      <div class=\"grid\"><div><h2>Welcome back.</h2><p class=\"sub\">Enter the phone number Coach Aimi has on file. We'll send a one-time verification code to your registered email.</p></div>\n      <form id=\"phoneForm\"><label class=\"field\" for=\"phone\">Your mobile number<input id=\"phone\" type=\"tel\" autocomplete=\"tel\" placeholder=\"e.g. 0123456789\" required></label><button class=\"primary\" id=\"sendCode\" type=\"submit\">Send email code →</button></form></div>\n      <p class=\"hint\">New client? Ask Coach Aimi to add your details first.</p>\n      <button id=\"backToLogins\" type=\"button\" class=\"backlogin\">← Back to login options</button>\n    </section>\n    <section class=\"shellbody hidden\" id=\"codeView\">\n      <h2>Check your inbox.</h2><p class=\"sub\">Enter the 6-digit code sent to the email registered with your number. It expires after 5 minutes.</p>\n      <form id=\"codeForm\"><label class=\"field\" for=\"code\">Verification code<input id=\"code\" inputmode=\"numeric\" pattern=\"[0-9]{6}\" maxlength=\"6\" autocomplete=\"one-time-code\" placeholder=\"••••••\" required></label>\n      <button id=\"verifyCode\" class=\"primary\" type=\"submit\">Verify & continue →</button> <button id=\"backPhone\" type=\"button\" class=\"secondary\">Change number</button></form>\n    </section>\n    <section class=\"hidden\" id=\"dashboard\">\n      <div id=\"creditBanner\" style=\"padding:17px 30px;background:#e9f7db;color:#25452d;font-size:14px\" role=\"status\" aria-live=\"polite\"><strong id=\"creditCount\">—</strong> session credits remaining <span style=\"font-size:12px;color:#56705c\">· 1 credit per confirmed session. Cancellations at least 24h ahead return the credit.</span></div>\n      <nav class=\"tabnav\"><button class=\"tab active\" id=\"bookTab\" type=\"button\">Book a session</button><button class=\"tab\" id=\"myTab\" type=\"button\">My bookings</button><button class=\"tab\" id=\"signout\" type=\"button\" style=\"margin-left:auto\">Sign out</button></nav>\n      <div class=\"panel\" id=\"bookPanel\">\n        <h2 id=\"welcome\">Choose your training hour.</h2><p class=\"sub\">Wednesday and Sunday are off by default. Your coach can change working days anytime.</p>\n        <div class=\"split\"><div><label class=\"field\" for=\"date\">Training date<input id=\"date\" type=\"date\" required></label>\n          <div class=\"daybar\" id=\"daybar\" aria-label=\"Upcoming days\"></div>\n          <p class=\"sub\" style=\"margin:0 0 6px\">Available one-hour sessions</p><div class=\"slotgrid\" id=\"slots\"></div><p id=\"slotStatus\" class=\"hint\"></p></div>\n          <aside class=\"summary\"><small id=\"summaryEyebrow\">YOUR SESSION</small><b id=\"selectedTime\">Select a time</b><p id=\"selectedDate\">Choose a date and an available hour to confirm.</p><p>Coach Aimi checks all enabled calendars for busy times. Your booking is saved only to her PT calendar.</p><p>Each confirmed session uses 1 credit.</p><button class=\"primary\" id=\"confirmBooking\" type=\"button\" disabled>Confirm booking →</button><button class=\"secondary hidden\" id=\"stopEdit\" style=\"width:100%;margin-top:9px\" type=\"button\">Back to new booking</button></aside></div>\n      </div>\n      <div class=\"panel hidden\" id=\"myPanel\"><h2>Your sessions.</h2><p class=\"sub\">You may reschedule or cancel only when the session starts at least 24 hours from now.</p><div id=\"myBookings\"></div></div>\n    </section>\n  </main>\n  <footer class=\"footer\"><div>Coach Aimi · PT bookings</div><div>Malaysia time (MYT)</div></footer>\n</div>\n<script>\n  'use strict';\n  const $ = id => document.getElementById(id);\n  const state = {phone:'', token:'', client:null, date:'', hour:null, editingId:null, creditsRemaining:0, working:false};\n  // GitHub Pages owns the UI; Apps Script only runs the business logic.\n  // After deployment, paste your Apps Script /exec URL here. No secrets or client data.\n  const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwtPCdBcTFna4QuoP9e0Z78lIVwkvBGy8zqaGSdKXM8Hv-1edBAgCwAdwh-QWyuU1Ix/exec';\n  function call(name,...args) { return new Promise((resolve,reject) => {\n    const runner=google.script.run.withSuccessHandler(resolve).withFailureHandler(err=>reject(new Error(err&&err.message?err.message:'Booking service error')));\n    try {switch(name) {\n        case 'requestLoginCode': runner.requestLoginCode(args[0]); break;\n        case 'verifyLoginCode': runner.verifyLoginCode(args[0],args[1]); break;\n        case 'getAvailability': runner.getAvailability(args[0]); break;\n        case 'getMyCredits': runner.getMyCredits(args[0]); break;\n        case 'getMyBookings': runner.getMyBookings(args[0]); break;\n        case 'makeBooking': runner.makeBooking(args[0],args[1],args[2]); break;\n        case 'rescheduleBooking': runner.rescheduleBooking(args[0],args[1],args[2],args[3]); break;\n        case 'cancelBooking': runner.cancelBooking(args[0],args[1]); break;\n        default: reject(new Error('Unsupported operation.'));\n      }} catch (e) {reject(e);}\n  }); }\n  function msg(text, error=false){const el=$('message');el.textContent=text;el.classList.remove('hidden');el.classList.toggle('error',error);el.scrollIntoView({block:'nearest',behavior:'smooth'});}\n  function clearMsg(){$('message').classList.add('hidden');}\n  function show(view){['loginSelect','phoneView','codeView','dashboard'].forEach(id=>$(id).classList.toggle('hidden',id!==view));$('clientShellHead').classList.toggle('hidden',view==='loginSelect');document.body.classList.toggle('client-active',view==='dashboard');$('stepLabel').textContent=view==='dashboard'?'MY TRAINING':view==='codeView'?'EMAIL VERIFICATION':'CLIENT SIGN IN';clearMsg();}\n  $('clientLogin').addEventListener('click',()=>{show('phoneView');$('phone').focus();});\n  $('backToLogins').addEventListener('click',()=>location.assign('https://aimi-keboon.github.io/coach-aimi-pt-booking/'));\n  function loading(bool){state.working=bool;['sendCode','verifyCode'].forEach(id=>$(id).disabled=bool);refreshSummary();}\n  function todayMY(){const parts=new Intl.DateTimeFormat('en-GB',{timeZone:'Asia/Kuala_Lumpur',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date());const get=t=>parts.find(x=>x.type===t).value;return get('year')+'-'+get('month')+'-'+get('day');}\n  function dateAdd(d,n){const x=new Date(d+'T12:00:00Z');x.setUTCDate(x.getUTCDate()+n);return x.toISOString().slice(0,10);}\n  function niceDate(d){return new Date(d+'T12:00:00Z').toLocaleDateString('en-MY',{weekday:'long',day:'numeric',month:'long',year:'numeric',timeZone:'UTC'});}\n  function fmtHour(h){return (h%12||12)+':00 '+(h<12?'AM':'PM');}\n  $('phoneForm').addEventListener('submit',async e=>{e.preventDefault();loading(true);state.phone=$('phone').value.trim();try{const r=await call('requestLoginCode',state.phone);show('codeView');msg(r.message);}catch(err){msg(err.message,true);}finally{loading(false);}});\n  $('backPhone').addEventListener('click',()=>show('phoneView'));\n  $('codeForm').addEventListener('submit',async e=>{e.preventDefault();loading(true);try{const r=await call('verifyLoginCode',state.phone,$('code').value.trim());state.token=r.token;state.client=r.client;state.creditsRemaining=r.client.creditsRemaining;show('dashboard');renderCredits();$('welcome').textContent='Let’s train, '+r.client.name+'.';state.date=todayMY();$('date').min=state.date;$('date').max=dateAdd(state.date,60);$('date').value=state.date;await Promise.all([loadSlots(),loadBookings(),loadCredits()]);}catch(err){msg(err.message,true);}finally{loading(false);}});\n  $('signout').addEventListener('click',()=>{state.token='';state.client=null;state.editingId=null;state.hour=null;state.creditsRemaining=0;$('code').value='';show('phoneView');});\n  async function loadCredits(){const r=await call('getMyCredits',state.token);state.creditsRemaining=r.remaining;renderCredits();refreshSummary();}\n  function renderCredits(){$('creditCount').textContent=state.creditsRemaining;}\n  function activateTab(which){$('bookTab').classList.toggle('active',which==='book');$('myTab').classList.toggle('active',which==='my');$('bookPanel').classList.toggle('hidden',which!=='book');$('myPanel').classList.toggle('hidden',which!=='my');clearMsg();}\n  $('bookTab').addEventListener('click',()=>activateTab('book'));\n  $('myTab').addEventListener('click',async()=>{activateTab('my');await loadBookings();});\n  function drawDays(){const host=$('daybar');host.replaceChildren();for(let i=0;i<7;i++){const d=dateAdd(state.date,i);const btn=document.createElement('button');btn.type='button';btn.className='day'+(d===state.date?' active':'');const weekday=new Date(d+'T12:00:00Z').toLocaleDateString('en-MY',{timeZone:'UTC',weekday:'short'});btn.innerHTML='<span>'+weekday+'</span><b>'+Number(d.slice(-2))+'</b>';btn.addEventListener('click',()=>{$('date').value=d;state.date=d;loadSlots();});host.appendChild(btn);}}\n  $('date').addEventListener('change',()=>{state.date=$('date').value;loadSlots();});\n  async function loadSlots(){state.hour=null;refreshSummary();drawDays();$('slots').replaceChildren();$('slotStatus').textContent='Checking calendar availability…';try{const availability=await call('getAvailability',state.date);$('slotStatus').textContent=availability.reason||'Select your preferred hour.';for(const s of availability.slots){const b=document.createElement('button');b.type='button';b.className='slot';b.textContent=s.label;b.addEventListener('click',()=>{state.hour=s.hour;[...$('slots').children].forEach(x=>x.classList.remove('active'));b.classList.add('active');refreshSummary();});$('slots').appendChild(b);}}catch(err){$('slotStatus').textContent=err.message;msg(err.message,true);}}\n  function refreshSummary(){const editing=!!state.editingId;$('summaryEyebrow').textContent=editing?'RESCHEDULE SESSION':'YOUR SESSION';$('selectedTime').textContent=state.hour===null?'Select a time':fmtHour(state.hour)+' – '+fmtHour(state.hour+1);$('selectedDate').textContent=state.hour===null?'Choose an available hour.':niceDate(state.date);$('confirmBooking').textContent=editing?'Confirm new time →':'Confirm booking →';$('confirmBooking').disabled=state.working||state.hour===null||(!editing&&state.creditsRemaining<1);$('stopEdit').classList.toggle('hidden',!editing);}\n  $('stopEdit').addEventListener('click',()=>{state.editingId=null;refreshSummary();});\n  $('confirmBooking').addEventListener('click',async()=>{if(state.hour===null||state.working)return;loading(true);try{const result=state.editingId?await call('rescheduleBooking',state.token,state.editingId,state.date,state.hour):await call('makeBooking',state.token,state.date,state.hour);state.editingId=null;state.hour=null;refreshSummary();await Promise.all([loadSlots(),loadBookings(),loadCredits()]);activateTab('my');msg(result.message+' Email notifications have been requested for you and Coach Aimi.');}catch(err){msg(err.message,true);await loadSlots();}finally{loading(false);}});\n  async function loadBookings(){const host=$('myBookings');host.textContent='Loading sessions…';try{const rows=await call('getMyBookings',state.token);host.replaceChildren();if(!rows.length){const p=document.createElement('p');p.className='hint';p.textContent='No sessions yet. Book your first training hour!';host.appendChild(p);return;}for(const b of rows){const card=document.createElement('article');card.className='bookcard';const left=document.createElement('div');const badge=document.createElement('span');badge.className='pill'+(b.status==='CANCELLED'?' cancelled':'');badge.textContent=b.status;const label=document.createElement('strong');label.textContent=niceDate(b.date);const time=document.createElement('small');time.textContent=b.label+' · Coach Aimi';left.append(badge,label,time);const actions=document.createElement('div');actions.className='actions';if(b.canEdit){const edit=document.createElement('button');edit.className='secondary';edit.textContent='Reschedule';edit.onclick=()=>{state.editingId=b.id;state.date=b.date;$('date').value=b.date;activateTab('book');loadSlots();refreshSummary();};const cancel=document.createElement('button');cancel.className='secondary';cancel.textContent='Cancel';cancel.onclick=async()=>{if(!confirm('Cancel your session on '+niceDate(b.date)+' at '+b.label+'?'))return;try{await call('cancelBooking',state.token,b.id);await Promise.all([loadBookings(),loadCredits()]);msg('Session cancelled. One credit is available again. Notification emails have been requested.');}catch(err){msg(err.message,true);}};actions.append(edit,cancel);}else if(b.status==='CONFIRMED'){const note=document.createElement('small');note.textContent='Changes closed (under 24 hours)';actions.appendChild(note);}card.append(left,actions);host.appendChild(card);}}catch(err){host.textContent='Could not load bookings.';msg(err.message,true);}}\n</script>\n</body>\n</html>";
