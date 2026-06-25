/**
 * SAT Doctor — Booking API (v1)
 * Bind this script to your SAT Doctor Google Spreadsheet.
 * Deploy: Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone
 */

var SHEET_SLOTS = 'Slots';
var SHEET_LOG = 'Log';
var SHEET_BOOKINGS = 'Bookings';
var SHEET_ENROLLMENTS = 'Enrollments';

// Set PRIVATE_SPREADSHEET_ID in Script Properties to write Bookings,
// Enrollments, and Logs to a separate private sheet, keeping the public
// sheet (used for calendar display) limited to Slots only.
function getPrivateSpreadsheet_() {
  var privateId = PropertiesService.getScriptProperties().getProperty('PRIVATE_SPREADSHEET_ID');
  if (privateId) return SpreadsheetApp.openById(privateId);
  return getSpreadsheet_(); // fall back to same sheet if not configured
}

function getPrivateSheet_(name) {
  var ss = getPrivateSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var headers = name === SHEET_BOOKINGS ? BOOKINGS_HEADERS
                : name === SHEET_ENROLLMENTS ? ENROLLMENTS_HEADERS
                : LOG_HEADERS;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

var SLOTS_HEADERS = [
  'slot_id', 'start_at', 'end_at', 'lesson_type', 'title', 'tutor',
  'capacity', 'booked_count', 'status', 'zoom_join_url', 'lesson_number', 'subject'
];

var BOOKINGS_HEADERS = [
  'booking_id', 'slot_id', 'student_name', 'student_email', 'student_phone',
  'course', 'status', 'created_at', 'zoom_join_url'
];

var ENROLLMENTS_HEADERS = [
  'timestamp', 'name', 'email', 'phone', 'course', 'message'
];

var LOG_HEADERS = [
  'timestamp', 'source', 'action', 'status', 'detail'
];

// ── HTTP handlers ───────────────────────────────────────────────────────────

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};

  if (params.page === 'calendar') {
    return HtmlService.createHtmlOutputFromFile('Calendar')
      .setTitle('Book a Group Class — SAT Doctor')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var action = (params.action || '').toLowerCase();
  var result;

  try {
    log_('doGet', action || 'ping', 'received', 'params: ' + JSON.stringify(params));

    if (action === 'ping' || action === '') {
      result = {
        ok: true,
        message: 'SAT Doctor booking API is running.',
        spreadsheetConfigured: !!getSpreadsheetId_(),
        hint: 'Use ?action=slots&from=2026-01-01&to=2027-12-31 to list sessions.'
      };
    } else if (action === 'slots') {
      ensureSheets_();
      result = { ok: true, slots: listSlots_(params.from, params.to, params.level) };
    } else if (action === 'enroll') {
      ensureSheets_();
      result = logEnrollment_({
        name: params.name || '',
        email: params.email || '',
        phone: params.phone || '',
        course: params.course || '',
        message: params.message || '',
        timestamp: params.timestamp || ''
      });
      log_('doGet', 'enroll', 'ok', 'name=' + (params.name || '') + ' email=' + (params.email || ''));
    } else if (action === 'book') {
      ensureSheets_();
      log_('doGet', 'book', 'attempting', 'slotId=' + params.slotId + ' name=' + params.name + ' email=' + params.email);
      result = createBooking_({
        slotId: params.slotId,
        name: params.name,
        email: params.email,
        phone: params.phone,
        course: params.course || '',
        requestId: params.requestId || ''
      });
      log_('doGet', 'book', result.ok ? 'ok' : 'fail', JSON.stringify(result));
    } else {
      result = {
        ok: false,
        error: 'Unknown action "' + action + '". Use action=ping, action=slots, or action=book.'
      };
      log_('doGet', action, 'unknown', 'unrecognized action');
    }
  } catch (err) {
    result = { ok: false, error: String(err.message || err) };
    log_('doGet', action, 'exception', String(err.message || err));
  }

  return jsonpResponse_(result, params.callback);
}

/** Called from Calendar.html via google.script.run (reliable; no JSONP/CORS). */
function apiListSlots(fromIso, toIso, level) {
  try {
    ensureSheets_();
    return { ok: true, slots: listSlots_(fromIso, toIso, level || '') };
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

/** Called from Calendar.html via google.script.run */
function apiCreateBooking(payload) {
  try {
    ensureSheets_();
    return createBooking_(payload || {});
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

function doPost(e) {
  var result = { ok: false, error: 'Invalid request' };

  try {
    ensureSheets_();
    var body = parseRequestBody_(e);
    var action = (body.action || '').toLowerCase();

    log_('doPost', action || 'no-action', 'received',
      'name=' + (body.name || '') + ' email=' + (body.email || '') +
      ' slotId=' + (body.slotId || '') + ' action=' + action);

    if (action === 'book') {
      result = createBooking_(body);
    } else if (action === 'enroll') {
      result = logEnrollment_(body);
      log_('doPost', 'enroll', 'ok', 'name=' + body.name + ' email=' + body.email);
    } else {
      // Legacy enrollment form (no action field)
      result = logEnrollment_(body);
    }
  } catch (err) {
    result = { ok: false, error: String(err.message || err) };
    log_('doPost', 'exception', 'fail', String(err.message || err));
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Slots ───────────────────────────────────────────────────────────────────

function listSlots_(fromIso, toIso, levelFilter) {
  var sheet = getSheet_(SHEET_SLOTS);
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];

  var headers = rows[0].map(normalizeHeader_);
  var fromMs = fromIso ? Date.parse(fromIso) : Date.now();
  var toMs = toIso ? Date.parse(toIso) : fromMs + 90 * 24 * 60 * 60 * 1000;
  var level = levelFilter ? String(levelFilter).trim() : '';

  var slots = [];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue;

    var slot = rowToSlot_(headers, row);
    if (!slot.slot_id) continue;
    if (slot.status !== 'open' && slot.status !== 'full') continue;

    var startMs = Date.parse(slot.start_at);
    if (isNaN(startMs) || startMs < fromMs || startMs > toMs) continue;

    if (level && slot.lesson_type !== level) continue;

    var capacity = parseInt(slot.capacity, 10) || 0;
    var booked = parseInt(slot.booked_count, 10) || 0;
    var spotsLeft = capacity - booked;
    var isFull = spotsLeft <= 0;

    slots.push({
      id: slot.slot_id,
      title: slot.title || slot.lesson_type,
      start: slot.start_at,
      end: slot.end_at,
      lessonType: slot.lesson_type,
      tutor: slot.tutor,
      capacity: capacity,
      booked: booked,
      spotsLeft: isFull ? 0 : spotsLeft,
      isFull: isFull,
      extendedProps: {
        lessonType: slot.lesson_type,
        tutor: slot.tutor,
        spotsLeft: isFull ? 0 : spotsLeft,
        isFull: isFull,
        capacity: capacity
      }
    });
  }

  slots.sort(function (a, b) {
    return Date.parse(a.start) - Date.parse(b.start);
  });

  return slots;
}

// ── Booking ─────────────────────────────────────────────────────────────────

function createBooking_(data) {
  var slotId = String(data.slotId || '').trim();
  var name = String(data.name || '').trim();
  var email = String(data.email || '').trim().toLowerCase();
  var phone = String(data.phone || '').trim();
  var course = String(data.course || '').trim();
  var requestId = String(data.requestId || '').trim();

  log_('createBooking', 'start', 'info', 'slotId=' + slotId + ' email=' + email);

  if (!slotId || !name || !email || !phone) {
    log_('createBooking', 'validation', 'fail', 'missing fields: slotId=' + slotId + ' name=' + name + ' email=' + email + ' phone=' + phone);
    return { ok: false, error: 'Missing required fields (slot, name, email, phone).' };
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    log_('createBooking', 'validation', 'fail', 'invalid email: ' + email);
    return { ok: false, error: 'Invalid email address.' };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    log_('createBooking', 'lock', 'fail', 'could not acquire lock');
    return { ok: false, error: 'Server busy — please try again.' };
  }

  try {
    if (requestId && isDuplicateRequest_(requestId)) {
      log_('createBooking', 'duplicate', 'fail', 'requestId=' + requestId);
      return { ok: false, error: 'This booking was already submitted.' };
    }

    var slotsSheet = getSheet_(SHEET_SLOTS);
    var slotRowIndex = findSlotRowIndex_(slotsSheet, slotId);
    if (slotRowIndex < 0) {
      log_('createBooking', 'findSlot', 'fail', 'slotId not found: ' + slotId);
      return { ok: false, error: 'Session not found.' };
    }
    log_('createBooking', 'findSlot', 'ok', 'found at row ' + slotRowIndex);

    var headers = slotsSheet.getRange(1, 1, 1, slotsSheet.getLastColumn()).getValues()[0].map(normalizeHeader_);
    var slotRow = slotsSheet.getRange(slotRowIndex, 1, 1, headers.length).getValues()[0];
    var slot = rowToSlot_(headers, slotRow);
    log_('createBooking', 'slotData', 'info', 'status=' + slot.status + ' capacity=' + slot.capacity + ' booked=' + slot.booked_count);

    if (slot.status !== 'open') {
      log_('createBooking', 'slotStatus', 'fail', 'status is: ' + slot.status);
      return { ok: false, error: 'This session is no longer available.' };
    }

    var capacity = parseInt(slot.capacity, 10) || 0;
    var booked = parseInt(slot.booked_count, 10) || 0;
    if (booked >= capacity) {
      log_('createBooking', 'capacity', 'fail', 'booked=' + booked + ' capacity=' + capacity);
      return { ok: false, error: 'This session is full.' };
    }

    if (hasExistingBooking_(slotId, email)) {
      log_('createBooking', 'duplicate', 'fail', 'already booked: ' + email);
      return { ok: false, error: 'You are already registered for this session.' };
    }

    var bookingId = 'BK-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    var now = new Date();
    var zoomUrl = slot.zoom_join_url || getDefaultZoomUrl_(slot.lesson_type);
    log_('createBooking', 'appendBooking', 'attempting', 'bookingId=' + bookingId + ' zoomUrl=' + zoomUrl);
    appendBooking_({
      booking_id: bookingId,
      slot_id: slotId,
      student_name: name,
      student_email: email,
      student_phone: phone,
      course: course,
      status: 'confirmed',
      created_at: now.toISOString(),
      zoom_join_url: zoomUrl
    });
    log_('createBooking', 'appendBooking', 'ok', 'row written to Bookings sheet');

    var newBooked = booked + 1;
    var bookedCol = headers.indexOf('booked_count') + 1;
    var statusCol = headers.indexOf('status') + 1;
    slotsSheet.getRange(slotRowIndex, bookedCol).setValue(newBooked);
    if (newBooked >= capacity && statusCol > 0) {
      slotsSheet.getRange(slotRowIndex, statusCol).setValue('full');
    }

    if (requestId) markRequestId_(requestId);

    log_('createBooking', 'sendEmail', 'attempting', 'to=' + email);
    try {
      sendBookingConfirmationEmail_(name, email, slot, zoomUrl, bookingId);
      log_('createBooking', 'sendEmail', 'ok', 'email sent to ' + email);
    } catch(emailErr) {
      log_('createBooking', 'sendEmail', 'fail', String(emailErr.message || emailErr));
    }

    // Schedule reminder (24h before) and follow-up (1h after session ends)
    try {
      scheduleFollowUpEmails_(bookingId, name, email, slot, zoomUrl);
      log_('createBooking', 'triggers', 'ok', 'reminder and follow-up triggers scheduled');
    } catch(triggerErr) {
      log_('createBooking', 'triggers', 'fail', String(triggerErr.message || triggerErr));
    }

    return {
      ok: true,
      bookingId: bookingId,
      slotId: slotId,
      title: slot.title,
      start: slot.start_at,
      end: slot.end_at,
      lessonType: slot.lesson_type,
      tutor: slot.tutor,
      zoomUrl: zoomUrl
    };
  } finally {
    lock.releaseLock();
  }
}

function appendBooking_(booking) {
  var sheet = getPrivateSheet_(SHEET_BOOKINGS);
  sheet.appendRow([
    booking.booking_id,
    booking.slot_id,
    booking.student_name,
    booking.student_email,
    booking.student_phone,
    booking.course,
    booking.status,
    booking.created_at,
    booking.zoom_join_url
  ]);
}

function hasExistingBooking_(slotId, email) {
  var sheet = getPrivateSheet_(SHEET_BOOKINGS);
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return false;

  var headers = rows[0].map(normalizeHeader_);
  var slotCol = headers.indexOf('slot_id');
  var emailCol = headers.indexOf('student_email');
  var statusCol = headers.indexOf('status');

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (String(row[slotCol]).trim() !== slotId) continue;
    if (String(row[emailCol]).trim().toLowerCase() !== email) continue;
    var status = statusCol >= 0 ? String(row[statusCol]).toLowerCase() : 'confirmed';
    if (status !== 'cancelled') return true;
  }
  return false;
}

// ── Enrollment (legacy) ─────────────────────────────────────────────────────

function logEnrollment_(body) {
  var sheet = getPrivateSheet_(SHEET_ENROLLMENTS);
  sheet.appendRow([
    body.timestamp || new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }),
    body.name || '',
    body.email || '',
    body.phone || '',
    body.course || '',
    body.message || ''
  ]);
  return { ok: true };
}

// ── Scheduled Email Triggers ────────────────────────────────────────────────

function scheduleFollowUpEmails_(bookingId, name, email, slot, zoomUrl) {
  var startMs  = Date.parse(slot.start_at);
  var endMs    = Date.parse(slot.end_at);
  if (isNaN(startMs) || startMs <= Date.now()) return;

  var reminderTime = new Date(startMs - 24 * 60 * 60 * 1000);
  var followUpTime = new Date(endMs   +  1 * 60 * 60 * 1000);

  var props = PropertiesService.getScriptProperties();
  props.setProperty('booking_' + bookingId, JSON.stringify({
    bookingId:    bookingId,
    name:         name,
    email:        email,
    lessonType:   slot.lesson_type,
    lessonNumber: String(slot.lesson_number || '1'),
    subject:      slot.subject || 'Math',
    title:        slot.title || slot.lesson_type,
    tutor:        slot.tutor || 'SAT Doctor',
    startAt:      slot.start_at,
    endAt:        slot.end_at,
    zoomUrl:      zoomUrl || ''
  }));

  var reminderTrigger = ScriptApp.newTrigger('sendReminderEmail')
    .timeBased().at(reminderTime).create();
  props.setProperty('trigger_reminder_' + reminderTrigger.getUniqueId(), bookingId);

  var followUpTrigger = ScriptApp.newTrigger('sendFollowUpEmail')
    .timeBased().at(followUpTime).create();
  props.setProperty('trigger_followup_' + followUpTrigger.getUniqueId(), bookingId);
}

function sendReminderEmail(e) {
  var triggerId = e.triggerUid;
  var props     = PropertiesService.getScriptProperties();
  var bookingId = props.getProperty('trigger_reminder_' + triggerId);
  if (!bookingId) { log_('sendReminderEmail', 'lookup', 'fail', 'no bookingId for ' + triggerId); return; }

  var data = JSON.parse(props.getProperty('booking_' + bookingId) || '{}');
  if (!data.email) { log_('sendReminderEmail', 'lookup', 'fail', 'no data for ' + bookingId); return; }

  try {
    var tz      = 'America/New_York';
    var start   = new Date(data.startAt);
    var end     = new Date(data.endAt);
    var dateStr = Utilities.formatDate(start, tz, 'EEEE, MMMM d, yyyy');
    var timeStr = Utilities.formatDate(start, tz, 'h:mm a') + ' \u2013 ' + Utilities.formatDate(end, tz, 'h:mm a') + ' ET';
    var subject = 'Your SAT Doctor class is tomorrow \u2014 ' + dateStr;
    var attachments = getPreClassAttachments_(data.lessonType, data.lessonNumber, data.subject);

    var htmlBody = [
      '<div style="font-family:Georgia,serif;max-width:580px;margin:0 auto;">',
      '<div style="background:#1a2a5e;padding:24px 36px;border-radius:8px 8px 0 0;">',
        '<p style="margin:0;font-size:20px;font-weight:bold;color:#fff;">SAT Doctor</p>',
        '<p style="margin:6px 0 0;font-size:12px;color:rgba(255,255,255,0.5);letter-spacing:0.08em;text-transform:uppercase;">Class Reminder</p>',
      '</div>',
      '<div style="background:#fff;padding:28px 36px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 8px 8px;">',
        '<p style="font-size:16px;margin:0 0 8px;">Hi ' + escapeHtml_(data.name) + ',</p>',
        '<p style="font-size:14px;color:#374151;line-height:1.7;margin:0 0 20px;">Just a friendly reminder that your SAT group class is <strong>tomorrow</strong>. We\'re looking forward to seeing you!</p>',
        '<div style="background:#f8f9fc;border:1px solid #e0e4f0;border-radius:6px;padding:16px 20px;margin-bottom:20px;">',
          '<p style="margin:0 0 10px;font-size:11px;font-weight:600;letter-spacing:0.1em;text-transform:uppercase;color:#6b7280;">Session Details</p>',
          '<table style="width:100%;font-size:14px;border-collapse:collapse;">',
            '<tr><td style="color:#6b7280;padding:4px 0;width:36%;">Class</td><td style="color:#111827;font-weight:600;">' + escapeHtml_(data.title) + '</td></tr>',
            '<tr><td style="color:#6b7280;padding:4px 0;">Date</td><td style="color:#111827;font-weight:600;">' + dateStr + '</td></tr>',
            '<tr><td style="color:#6b7280;padding:4px 0;">Time</td><td style="color:#111827;font-weight:600;">' + timeStr + '</td></tr>',
            '<tr><td style="color:#6b7280;padding:4px 0;">Tutor</td><td style="color:#111827;font-weight:600;">' + escapeHtml_(data.tutor) + '</td></tr>',
          '</table>',
        '</div>',
        data.zoomUrl
          ? '<p style="margin:0 0 6px;"><a href="' + escapeHtml_(data.zoomUrl) + '" style="display:inline-block;background:#1a2a5e;color:#fff;padding:11px 24px;text-decoration:none;border-radius:4px;font-size:14px;font-weight:600;">Join Zoom Session</a></p><p style="font-size:12px;color:#9ca3af;margin:4px 0 20px;">Or copy: ' + escapeHtml_(data.zoomUrl) + '</p>'
          : '<p style="font-size:14px;color:#374151;margin:0 0 20px;">Your Zoom link will be shared shortly before the session.</p>',
        attachments.length ? '<p style="font-size:14px;color:#374151;margin:0 0 20px;">We\'ve attached your <strong>pre-class materials</strong>. Please review them before the session.</p>' : '',
        '<p style="font-size:14px;color:#374151;margin:0 0 20px;">If you have any questions, reply to this email and we\'ll get back to you right away.</p>',
        '<p style="font-size:14px;color:#374151;margin:0 0 4px;">See you tomorrow,</p>',
        '<p style="font-size:15px;font-weight:700;color:#1a2a5e;margin:0 0 2px;">' + escapeHtml_(data.tutor) + '</p>',
        '<p style="font-size:13px;color:#6b7280;margin:0;">SAT Doctor</p>',
      '</div></div>'
    ].join('');

    GmailApp.sendEmail(data.email, subject, stripHtml_(htmlBody), {
      htmlBody: htmlBody, name: 'SAT Doctor', attachments: attachments
    });
    log_('sendReminderEmail', 'send', 'ok', 'sent to ' + data.email);
  } catch(err) {
    log_('sendReminderEmail', 'send', 'fail', String(err.message || err));
  } finally {
    deleteTriggerById_(triggerId);
    props.deleteProperty('trigger_reminder_' + triggerId);
  }
}

function sendFollowUpEmail(e) {
  var triggerId = e.triggerUid;
  var props     = PropertiesService.getScriptProperties();
  var bookingId = props.getProperty('trigger_followup_' + triggerId);
  if (!bookingId) { log_('sendFollowUpEmail', 'lookup', 'fail', 'no bookingId for ' + triggerId); return; }

  var data = JSON.parse(props.getProperty('booking_' + bookingId) || '{}');
  if (!data.email) { log_('sendFollowUpEmail', 'lookup', 'fail', 'no data for ' + bookingId); return; }

  try {
    var firstName   = data.name.split(' ')[0];
    var subject     = 'Great work today, ' + firstName + '! Your homework is attached';
    var attachments = getHomeworkAttachments_(data.lessonType, data.lessonNumber, data.subject);

    var htmlBody = [
      '<div style="font-family:Georgia,serif;max-width:580px;margin:0 auto;">',
      '<div style="background:#1a2a5e;padding:24px 36px;border-radius:8px 8px 0 0;">',
        '<p style="margin:0;font-size:20px;font-weight:bold;color:#fff;">SAT Doctor</p>',
        '<p style="margin:6px 0 0;font-size:12px;color:rgba(255,255,255,0.5);letter-spacing:0.08em;text-transform:uppercase;">Post-Class Follow-Up</p>',
      '</div>',
      '<div style="background:#fff;padding:28px 36px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 8px 8px;">',
        '<p style="font-size:16px;margin:0 0 8px;">Hi ' + escapeHtml_(data.name) + ',</p>',
        '<p style="font-size:14px;color:#374151;line-height:1.7;margin:0 0 20px;">Thank you for attending today\'s session \u2014 it was a pleasure working with you! To reinforce what we covered, please complete the homework assignment attached to this email.</p>',
        '<div style="background:#fffbeb;border:1px solid #f59e0b;border-radius:6px;padding:16px 20px;margin-bottom:20px;">',
          '<p style="margin:0 0 10px;font-size:12px;font-weight:600;letter-spacing:0.08em;text-transform:uppercase;color:#92400e;">Homework Instructions</p>',
          '<ol style="margin:0;padding-left:18px;font-size:14px;color:#374151;line-height:1.9;">',
            '<li>Complete the attached homework assignment at your own pace.</li>',
            '<li>Show your work clearly where applicable.</li>',
            '<li>Scan or photograph your completed work and save it as a <strong>PDF</strong>.</li>',
            '<li>Email your PDF to <a href="mailto:thesatdoctor1600@gmail.com" style="color:#1a2a5e;">thesatdoctor1600@gmail.com</a> with the subject:<br/><strong style="color:#1a2a5e;">Homework \u2014 ' + escapeHtml_(data.title) + ' \u2014 ' + escapeHtml_(data.name) + '</strong></li>',
          '</ol>',
        '</div>',
        '<p style="font-size:14px;color:#374151;margin:0 0 20px;">If you have any questions about the material, don\'t hesitate to reply to this email.</p>',
        '<p style="font-size:14px;color:#374151;margin:0 0 4px;">Keep up the great work,</p>',
        '<p style="font-size:15px;font-weight:700;color:#1a2a5e;margin:0 0 2px;">' + escapeHtml_(data.tutor) + '</p>',
        '<p style="font-size:13px;color:#6b7280;margin:0;">SAT Doctor</p>',
      '</div></div>'
    ].join('');

    GmailApp.sendEmail(data.email, subject, stripHtml_(htmlBody), {
      htmlBody: htmlBody, name: 'SAT Doctor', attachments: attachments
    });
    log_('sendFollowUpEmail', 'send', 'ok', 'sent to ' + data.email);
  } catch(err) {
    log_('sendFollowUpEmail', 'send', 'fail', String(err.message || err));
  } finally {
    deleteTriggerById_(triggerId);
    props.deleteProperty('trigger_followup_' + triggerId);
    props.deleteProperty('booking_' + bookingId);
  }
}

function getPreClassAttachments_(lessonType, lessonNumber, subject) {
  var levelKey = lessonType === 'Level 1' ? 'L1' : 'L2';
  var subjKey  = (subject === 'Reading') ? 'Reading' : 'Math';
  var propKey  = 'PRE_' + levelKey + '_' + (parseInt(lessonNumber, 10) || 1) + '_' + subjKey;
  return getAttachmentBlobs_(propKey);
}

function getHomeworkAttachments_(lessonType, lessonNumber, subject) {
  var levelKey = lessonType === 'Level 1' ? 'L1' : 'L2';
  var subjKey  = (subject === 'Reading') ? 'Reading' : 'Math';
  var propKey  = 'HW_' + levelKey + '_' + (parseInt(lessonNumber, 10) || 1) + '_' + subjKey;
  return getAttachmentBlobs_(propKey);
}

function getAttachmentBlobs_(propKey) {
  var ids = (PropertiesService.getScriptProperties().getProperty(propKey) || '').trim();
  if (!ids) return [];
  var blobs = [];
  ids.split(',').forEach(function(id) {
    id = id.trim();
    if (!id) return;
    try {
      blobs.push(DriveApp.getFileById(id).getBlob());
    } catch(err) {
      log_('getAttachmentBlobs_', propKey, 'fail', 'file id=' + id + ': ' + err.message);
    }
  });
  return blobs;
}

function deleteTriggerById_(triggerId) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getUniqueId() === triggerId) ScriptApp.deleteTrigger(t);
  });
}

// ── Email ───────────────────────────────────────────────────────────────────

function sendBookingConfirmationEmail_(name, email, slot, zoomUrl, bookingId) {
  var tz = 'America/New_York';
  var start = new Date(slot.start_at);
  var end = new Date(slot.end_at);
  var dateStr = Utilities.formatDate(start, tz, 'EEEE, MMMM d, yyyy');
  var timeStr = Utilities.formatDate(start, tz, 'h:mm a') + ' – ' +
    Utilities.formatDate(end, tz, 'h:mm a') + ' ET';

  var subject = 'Confirmed: ' + (slot.title || 'SAT Doctor Group Class') + ' — ' +
    Utilities.formatDate(start, tz, 'MMM d');

  var tutor = escapeHtml_(slot.tutor || 'SAT Doctor');

  var htmlBody = [
    '<div style="font-family:Georgia,serif;max-width:580px;margin:0 auto;color:#1a1a2e;">',

    // Header banner
    '<div style="background:#1a2a5e;padding:28px 36px;border-radius:8px 8px 0 0;">',
      '<p style="margin:0;font-family:Georgia,serif;font-size:22px;font-weight:bold;color:#ffffff;letter-spacing:-0.01em;">SAT Doctor</p>',
      '<p style="margin:6px 0 0;font-size:13px;color:rgba(255,255,255,0.55);letter-spacing:0.06em;text-transform:uppercase;">Group Class Confirmation</p>',
    '</div>',

    // Body
    '<div style="background:#ffffff;padding:32px 36px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 8px 8px;">',

      '<p style="font-size:16px;color:#111827;margin:0 0 8px;">Hi ' + escapeHtml_(name) + ',</p>',
      '<p style="font-size:15px;color:#374151;line-height:1.7;margin:0 0 24px;">',
        'We\'re thrilled to have you joining us! Your spot in the upcoming SAT group class is confirmed, and we look forward to working with you. ',
        'Please keep an eye on your inbox — we\'ll be sending over additional study materials and prep resources ahead of the session to help you hit the ground running.',
      '</p>',

      // Session details card
      '<div style="background:#f8f9fc;border:1px solid #e0e4f0;border-radius:6px;padding:20px 24px;margin-bottom:24px;">',
        '<p style="margin:0 0 14px;font-size:11px;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:#6b7280;">Session Details</p>',
        '<table style="width:100%;border-collapse:collapse;font-size:14px;">',
          '<tr><td style="padding:5px 0;color:#6b7280;width:38%;">Class</td><td style="padding:5px 0;color:#111827;font-weight:600;">' + escapeHtml_(slot.title || slot.lesson_type) + '</td></tr>',
          '<tr><td style="padding:5px 0;color:#6b7280;">Date</td><td style="padding:5px 0;color:#111827;font-weight:600;">' + dateStr + '</td></tr>',
          '<tr><td style="padding:5px 0;color:#6b7280;">Time</td><td style="padding:5px 0;color:#111827;font-weight:600;">' + timeStr + '</td></tr>',
          '<tr><td style="padding:5px 0;color:#6b7280;">Tutor</td><td style="padding:5px 0;color:#111827;font-weight:600;">' + tutor + '</td></tr>',
          '<tr><td style="padding:5px 0;color:#6b7280;">Confirmation</td><td style="padding:5px 0;color:#111827;font-weight:600;">' + escapeHtml_(bookingId) + '</td></tr>',
        '</table>',
      '</div>',

      // Zoom button or placeholder
      zoomUrl
        ? '<p style="margin:0 0 8px;"><a href="' + escapeHtml_(zoomUrl) + '" style="display:inline-block;background:#1a2a5e;color:#ffffff;padding:13px 28px;text-decoration:none;border-radius:4px;font-size:14px;font-weight:600;letter-spacing:0.02em;">Join Zoom Session</a></p>' +
          '<p style="font-size:12px;color:#9ca3af;margin:6px 0 24px;">Or copy: ' + escapeHtml_(zoomUrl) + '</p>'
        : '<p style="font-size:14px;color:#374151;margin:0 0 24px;">Your Zoom link will be sent before the session.</p>',

      // Payment note
      '<div style="background:#fffbeb;border:1px solid #f59e0b;border-radius:6px;padding:14px 18px;margin-bottom:28px;">',
        '<p style="margin:0;font-size:13px;color:#92400e;line-height:1.65;">',
          '<strong>Payment note:</strong> Your spot is tentatively held. A member of our team will reach out shortly to confirm your registration and arrange payment. Your seat is not finalized until payment is received.',
        '</p>',
      '</div>',

      '<p style="font-size:14px;color:#374151;margin:0 0 28px;">If you have any questions in the meantime, don\'t hesitate to reply to this email — we\'re happy to help.</p>',

      // Signature
      '<p style="font-size:14px;color:#374151;margin:0 0 4px;">Warm regards,</p>',
      '<p style="font-size:15px;font-weight:700;color:#1a2a5e;margin:0 0 2px;">' + tutor + '</p>',
      '<p style="font-size:13px;color:#6b7280;margin:0;">SAT Doctor</p>',

    '</div>',
    '</div>'
  ].join('');

  var notifyEmail = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL') ||
    'thesatdoctor1600@gmail.com';

  GmailApp.sendEmail(email, subject, stripHtml_(htmlBody), {
    htmlBody: htmlBody,
    name: 'SAT Doctor'
  });

  if (notifyEmail && notifyEmail !== email) {
    GmailApp.sendEmail(
      notifyEmail,
      'New booking: ' + name + ' — ' + (slot.title || slot.lesson_type),
      name + ' booked ' + slot.title + ' on ' + dateStr + ' at ' + timeStr
    );
  }
}

// ── Admin helpers (run from Script Editor or custom menu) ───────────────────

function onOpen() {
  cacheSpreadsheetId_();
  SpreadsheetApp.getUi()
    .createMenu('SAT Doctor')
    .addItem('Initialize sheets', 'initializeSheets')
    .addItem('Save spreadsheet ID for web app', 'saveSpreadsheetId')
    .addItem('Add sample group slots', 'addSampleSlots')
    .addItem('Add June 14, 2026 session (10 AM – 12 PM ET)', 'addJune14_2026Slot')
    .addSeparator()
    .addItem('Add new session…', 'addNewSlotPrompt')
    .addToUi();
}

function addNewSlotPrompt() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:16px;margin:0;color:#222;}' +
    'h3{margin:0 0 14px;font-size:15px;color:#1a2a5e;}' +
    '.row{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px;}' +
    '.field{display:flex;flex-direction:column;margin-bottom:10px;}' +
    'label{font-weight:bold;margin-bottom:4px;font-size:12px;color:#444;}' +
    'input,select{padding:7px 9px;border:1px solid #ccc;border-radius:4px;font-size:13px;width:100%;box-sizing:border-box;}' +
    'input:focus,select:focus{outline:none;border-color:#1a2a5e;}' +
    '.btn{background:#1a2a5e;color:#fff;border:none;padding:10px 24px;border-radius:4px;font-size:13px;cursor:pointer;width:100%;margin-top:6px;}' +
    '.btn:hover{background:#111d42;}' +
    '.err{color:#b45309;font-size:12px;margin-top:8px;display:none;}' +
    '</style></head><body>' +
    '<h3>Add New Group Session</h3>' +
    '<div class="row">' +
      '<div class="field"><label>Date</label><input type="date" id="dt" /></div>' +
      '<div class="field"><label>Start Time (ET)</label><input type="time" id="tm" value="10:00" /></div>' +
    '</div>' +
    '<div class="row">' +
      '<div class="field"><label>Duration (minutes)</label><input type="number" id="dur" value="90" min="30" max="360" /></div>' +
      '<div class="field"><label>Capacity</label><input type="number" id="cap" value="6" min="1" max="50" /></div>' +
    '</div>' +
    '<div class="row">' +
      '<div class="field"><label>Level</label>' +
        '<select id="lvl">' +
          '<option value="Level 1">Level 1 — Beginner</option>' +
          '<option value="Level 2">Level 2 — Advanced</option>' +
        '</select>' +
      '</div>' +
      '<div class="field"><label>Tutor</label>' +
        '<select id="tutor">' +
          '<option value="Krutant Mehta">Krutant Mehta</option>' +
          '<option value="Abhi Nallamalli">Abhi Nallamalli</option>' +
        '</select>' +
      '</div>' +
    '</div>' +
    '<div class="row">' +
      '<div class="field"><label>Lesson number (1–6)</label><input type="number" id="lessonnum" value="1" min="1" max="6" /></div>' +
      '<div class="field"><label>Subject</label>' +
        '<select id="subject">' +
          '<option value="Math">Math</option>' +
          '<option value="Reading">Reading</option>' +
        '</select>' +
      '</div>' +
    '</div>' +
    '<div class="field"><label>Zoom URL (optional — leave blank to use default)</label><input type="url" id="zoom" placeholder="https://zoom.us/j/..." /></div>' +
    '<p class="err" id="err"></p>' +
    '<button class="btn" onclick="submit()">Add session</button>' +
    '<script>' +
    'function submit(){' +
      'var dt=document.getElementById("dt").value;' +
      'var tm=document.getElementById("tm").value;' +
      'var dur=parseInt(document.getElementById("dur").value);' +
      'var cap=parseInt(document.getElementById("cap").value);' +
      'var lvl=document.getElementById("lvl").value;' +
      'var tutor=document.getElementById("tutor").value;' +
      'var zoom=document.getElementById("zoom").value.trim();' +
      'var lessonnum=parseInt(document.getElementById("lessonnum").value)||1;' +
      'var subject=document.getElementById("subject").value;' +
      'var err=document.getElementById("err");' +
      'if(!dt||!tm){err.textContent="Please enter a date and time.";err.style.display="block";return;}' +
      'err.style.display="none";' +
      'google.script.run' +
        '.withSuccessHandler(function(msg){google.script.host.close();google.script.run.showAlert(msg);})' +
        '.withFailureHandler(function(e){err.textContent=e.message;err.style.display="block";})' +
        '.addSlotFromDialog(dt,tm,dur,cap,lvl,tutor,zoom,lessonnum,subject);' +
    '}' +
    '<\/script>' +
    '</body></html>'
  ).setWidth(480).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Group Session');
}

function showAlert(msg) {
  SpreadsheetApp.getUi().alert(msg);
}

function addSlotFromDialog(dateStr, timeStr, durationMin, capacity, level, tutor, zoomUrl, lessonNumber, subject) {
  try {
    var combined = dateStr + ' ' + timeStr;
    var start = Utilities.parseDate(combined, 'America/New_York', 'yyyy-MM-dd HH:mm');
    if (!start || isNaN(start.getTime())) {
      throw new Error('Could not parse date/time: "' + combined + '"');
    }

    if (!zoomUrl) zoomUrl = getDefaultZoomUrl_(level);
    lessonNumber = parseInt(lessonNumber, 10) || 1;
    subject = (subject === 'Reading') ? 'Reading' : 'Math';

    var title = level === 'Level 1'
      ? 'SAT Group \u2014 Beginner (Level 1)'
      : 'SAT Group \u2014 Advanced (Level 2)';

    ensureSheets_(true);
    var sheet = getSheet_(SHEET_SLOTS);
    var row = makeSlotRow_(level, title, tutor, start, durationMin, capacity, zoomUrl, lessonNumber, subject);
    sheet.appendRow(row);

    var end = new Date(start.getTime() + durationMin * 60 * 1000);
    return (
      'Session added!\n' +
      'Level: ' + level + ' (Lesson ' + lessonNumber + ', ' + subject + ')\n' +
      'Tutor: ' + tutor + '\n' +
      'Date:  ' + Utilities.formatDate(start, 'America/New_York', 'EEE, MMM d yyyy') + '\n' +
      'Time:  ' + Utilities.formatDate(start, 'America/New_York', 'h:mm a') +
               ' \u2013 ' + Utilities.formatDate(end, 'America/New_York', 'h:mm a') + ' ET\n' +
      'Capacity: ' + capacity + '\n' +
      'Zoom: ' + (zoomUrl || '(none set)')
    );
  } catch(err) {
    throw new Error(err.message || String(err));
  }
}

function saveSpreadsheetId() {
  cacheSpreadsheetId_();
  var id = getSpreadsheetId_();
  if (id) {
    SpreadsheetApp.getUi().alert('Saved SPREADSHEET_ID:\n' + id);
  } else {
    SpreadsheetApp.getUi().alert('Open this script from your booking spreadsheet (Extensions → Apps Script).');
  }
}

function cacheSpreadsheetId_() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', active.getId());
    }
  } catch (e) { /* not in sheet context */ }
}

function getSpreadsheetId_() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
}

function getSpreadsheet_() {
  var id = getSpreadsheetId_();
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    cacheSpreadsheetId_();
    return active;
  }
  throw new Error(
    'Set SPREADSHEET_ID: open the booking spreadsheet → SAT Doctor → Save spreadsheet ID for web app, then redeploy the web app.'
  );
}

function initializeSheets() {
  ensureSheets_(true);
  SpreadsheetApp.getUi().alert('Sheets ready: Slots, Bookings, Enrollments.');
}

function addSampleSlots() {
  ensureSheets_(true);
  var sheet = getSheet_(SHEET_SLOTS);
  var existing = sheet.getLastRow();
  if (existing > 1) {
    var ui = SpreadsheetApp.getUi();
    var confirm = ui.alert(
      'Slots sheet already has data. Add 5 sample sessions anyway?',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }

  var zoomL1 = PropertiesService.getScriptProperties().getProperty('ZOOM_LEVEL1') || '';
  var zoomL2 = PropertiesService.getScriptProperties().getProperty('ZOOM_LEVEL2') || '';

  var now = new Date();
  var nextSat = nextWeekday_(6, 10, 0); // next Saturday 10:00
  var nextSun = new Date(nextSat.getTime() + 24 * 60 * 60 * 1000);
  nextSun.setHours(14, 0, 0, 0);

  var june14 = easternDate_(2026, 6, 14, 10, 0);

  var samples = [
    makeSlotRow_('Level 1', 'SAT Group — Beginner (Level 1)', 'Krutant Mehta', june14, 120, 6, zoomL1),
    makeSlotRow_('Level 1', 'SAT Group — Beginner (Level 1)', 'Krutant Mehta', nextSat, 90, 6, zoomL1),
    makeSlotRow_('Level 1', 'SAT Group — Beginner (Level 1)', 'Krutant Mehta', addDays_(nextSat, 7), 90, 6, zoomL1),
    makeSlotRow_('Level 2', 'SAT Group — Advanced (Level 2)', 'Abhi Nallamalli', nextSun, 90, 6, zoomL2),
    makeSlotRow_('Level 2', 'SAT Group — Advanced (Level 2)', 'Abhi Nallamalli', addDays_(nextSun, 7), 90, 6, zoomL2)
  ];

  samples.forEach(function (row) {
    sheet.appendRow(row);
  });

  SpreadsheetApp.getUi().alert('Added ' + samples.length + ' sample slots. Set ZOOM_LEVEL1 / ZOOM_LEVEL2 in Script Properties.');
}

/**
 * Sunday, June 14, 2026 — 10:00 AM to 12:00 PM Eastern (America/New_York).
 */
function addJune14_2026Slot() {
  ensureSheets_(true);
  var sheet = getSheet_(SHEET_SLOTS);
  var start = easternDate_(2026, 6, 14, 10, 0);
  var startIso = start.toISOString();

  if (slotExistsAtStart_(sheet, startIso)) {
    SpreadsheetApp.getUi().alert('A slot at June 14, 2026 10:00 AM ET already exists.');
    return;
  }

  var zoomL1 = PropertiesService.getScriptProperties().getProperty('ZOOM_LEVEL1') || '';
  var row = makeSlotRow_(
    'Level 1',
    'SAT Group — Beginner (Level 1)',
    'Krutant Mehta',
    start,
    120,
    6,
    zoomL1
  );

  sheet.appendRow(row);
  SpreadsheetApp.getUi().alert(
    'Added: Sunday, June 14, 2026\n10:00 AM – 12:00 PM Eastern\nLevel 1 — Krutant Mehta'
  );
}

function slotExistsAtStart_(sheet, startIso) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return false;
  var headers = rows[0].map(normalizeHeader_);
  var startCol = headers.indexOf('start_at');
  if (startCol < 0) return false;

  for (var i = 1; i < rows.length; i++) {
    var existing = rows[i][startCol];
    if (!existing) continue;
    var existingIso = existing instanceof Date ? existing.toISOString() : String(existing);
    if (existingIso === startIso) return true;
  }
  return false;
}

/** Build a Date in US Eastern (EST/EDT) for calendar storage as UTC via toISOString(). */
function easternDate_(year, month, day, hour, minute) {
  var s = year + '-' + pad2_(month) + '-' + pad2_(day) + ' ' +
    pad2_(hour) + ':' + pad2_(minute) + ':00';
  return Utilities.parseDate(s, 'America/New_York', 'yyyy-MM-dd HH:mm:ss');
}

function pad2_(n) {
  return n < 10 ? '0' + n : String(n);
}

function makeSlotRow_(lessonType, title, tutor, start, durationMin, capacity, zoomUrl, lessonNumber, subject) {
  var end = new Date(start.getTime() + durationMin * 60 * 1000);
  return [
    'SLOT-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
    start.toISOString(),
    end.toISOString(),
    lessonType,
    title,
    tutor,
    capacity,
    0,
    'open',
    zoomUrl || '',
    lessonNumber || 1,
    subject || 'Math'
  ];
}

// ── Google Form → Bookings sheet trigger ────────────────────────────────────
// Set this up: Form editor → ... menu → Script editor (or use this same script)
// Triggers → Add trigger → onFormSubmit → From form → On form submit

function onFormSubmit(e) {
  try {
    ensureSheets_();

    // When triggered from spreadsheet "On form submit", e.namedValues has the data.
    // Keys match the form field labels exactly.
    var r = e.namedValues || {};

    var name   = (r['Name']    || r['name']    || [''])[0].trim();
    var email  = (r['Email']   || r['email']   || [''])[0].trim();
    var phone  = (r['Phone']   || r['phone']   || [''])[0].trim();
    var course = (r['Course']  || r['course']  || [''])[0].trim();
    var slotId = (r['Slot ID'] || r['slot_id'] || r['SlotID'] || [''])[0].trim();

    log_('onFormSubmit', 'received', 'info',
      'name=' + name + ' email=' + email + ' slotId=' + slotId);

    if (!slotId || !name || !email || !phone) {
      log_('onFormSubmit', 'validation', 'fail',
        'missing fields — got keys: ' + Object.keys(r).join(', '));
      return;
    }

    var result = createBooking_({
      slotId: slotId,
      name: name,
      email: email,
      phone: phone,
      course: course,
      requestId: 'form-' + Date.now()
    });

    log_('onFormSubmit', 'createBooking', result.ok ? 'ok' : 'fail', JSON.stringify(result));
  } catch(err) {
    log_('onFormSubmit', 'exception', 'fail', String(err.message || err));
  }
}

// ── Sheet setup ─────────────────────────────────────────────────────────────

function ensureSheets_(silent) {
  var ss = getSpreadsheet_();
  ensureSheetWithHeaders_(ss, SHEET_SLOTS, SLOTS_HEADERS);
  ensureSheetWithHeaders_(ss, SHEET_BOOKINGS, BOOKINGS_HEADERS);
  ensureSheetWithHeaders_(ss, SHEET_ENROLLMENTS, ENROLLMENTS_HEADERS);
  ensureSheetWithHeaders_(ss, SHEET_LOG, LOG_HEADERS);
}

function initializeSheets() {
  ensureSheets_(true);
}

function ensureSheetWithHeaders_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function getSheet_(name) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    ensureSheets_(true);
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

// ── Utilities ───────────────────────────────────────────────────────────────

function parseRequestBody_(e) {
  var body = {};
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(function (key) {
      body[key] = e.parameter[key];
    });
  }
  if (e && e.postData && e.postData.contents) {
    var contents = e.postData.contents;
    // Try JSON parse first (handles application/json and text/plain with JSON body)
    try {
      var parsed = JSON.parse(contents);
      Object.keys(parsed).forEach(function (key) {
        body[key] = parsed[key];
      });
    } catch (jsonErr) {
      // Fall back to URL-encoded form parsing
      var parts = contents.split('&');
      parts.forEach(function (part) {
        var pair = part.split('=');
        if (pair.length >= 2) {
          body[decodeURIComponent(pair[0])] = decodeURIComponent(pair.slice(1).join('=').replace(/\+/g, ' '));
        }
      });
    }
  }
  return body;
}

function jsonpResponse_(result, callback) {
  var json = JSON.stringify(result);
  if (callback) {
    var safeCallback = String(callback).replace(/[^\w$.]/g, '');
    return ContentService
      .createTextOutput(safeCallback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeHeader_(h) {
  return String(h).trim().toLowerCase().replace(/\s+/g, '_');
}

function rowToSlot_(headers, row) {
  var slot = {};
  headers.forEach(function (h, i) {
    slot[h] = row[i] !== undefined && row[i] !== null ? String(row[i]) : '';
  });
  return slot;
}

function findSlotRowIndex_(sheet, slotId) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(normalizeHeader_);
  var idCol = headers.indexOf('slot_id');
  if (idCol < 0) return -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === slotId) return i + 1;
  }
  return -1;
}

function getDefaultZoomUrl_(lessonType) {
  var props = PropertiesService.getScriptProperties();
  if (lessonType === 'Level 1') return props.getProperty('ZOOM_LEVEL1') || '';
  if (lessonType === 'Level 2') return props.getProperty('ZOOM_LEVEL2') || '';
  return props.getProperty('ZOOM_DEFAULT') || '';
}

function isDuplicateRequest_(requestId) {
  var cache = CacheService.getScriptCache();
  return cache.get('req_' + requestId) !== null;
}

function markRequestId_(requestId) {
  CacheService.getScriptCache().put('req_' + requestId, '1', 600);
}

function nextWeekday_(dayOfWeek, hour, minute) {
  var d = new Date();
  d.setHours(hour, minute, 0, 0);
  var diff = (dayOfWeek + 7 - d.getDay()) % 7;
  if (diff === 0 && d < new Date()) diff = 7;
  d.setDate(d.getDate() + diff);
  return d;
}

function addDays_(date, days) {
  var d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function log_(source, action, status, detail) {
  try {
    var ss = getSpreadsheet_();
    var privateSs = getPrivateSpreadsheet_();
    var sheet = privateSs.getSheetByName(SHEET_LOG);
    if (!sheet) {
      sheet = privateSs.insertSheet(SHEET_LOG);
      sheet.getRange(1, 1, 1, 5).setValues([LOG_HEADERS]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(4, 80);
      sheet.setColumnWidth(5, 500);
    }
    var ts = new Date().toLocaleString('en-US', { timeZone: 'America/New_York' });
    sheet.appendRow([ts, source, action, status, String(detail || '').slice(0, 1000)]);
  } catch(e) { /* never let logging crash the main flow */ }
}

function stripHtml_(html) {
  return String(html).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}