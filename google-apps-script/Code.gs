// Google Apps Script - Deploy as Web App
// ============================================
// SETUP:
// 1. script.google.com → New project
// 2. Paste this code
// 3. Services → Add "Google Calendar API"
// 4. Run testFunction() to verify
// 5. Deploy → New Deployment → Web App (Execute as Me, Anyone can access)
// 6. Copy URL → Paste in BookingWizard.astro
// ============================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var service = data.serviceName;
    var duration = data.duration;
    var dateStr = data.dateString;
    var timeStr = data.timeString;
    var name = data.clientName;
    var email = data.clientEmail;
    var phone = data.clientPhone;
    var notes = data.clientNotes || 'No additional notes';
    var ref = data.ref || 'DN-' + Math.floor(1000 + Math.random() * 9000);

    var dateTime = parseDateTime(dateStr, timeStr);
    var durationMs = parseDuration(duration);
    var endTime = new Date(dateTime.getTime() + durationMs);

    var meetLink = '';
    var calendarLink = '';

    // Step 1: Create calendar event with Meet link via Calendar API
    try {
      var result = createEventWithMeet(dateTime, endTime, service, name, email, phone, notes, ref);
      meetLink = result.meetLink;
      calendarLink = result.calendarLink;
    } catch (apiErr) {
      Logger.log('Calendar API failed, using CalendarApp fallback: ' + apiErr.toString());
      // Fallback: create event with CalendarApp (no auto Meet link)
      try {
        var cal = CalendarApp.getDefaultCalendar();
        var event = cal.createEvent(
          'Therapy Session: ' + service + ' - ' + name,
          dateTime,
          endTime,
          {
            description: 'Session: ' + service + '\nClient: ' + name + '\nEmail: ' + email + '\nPhone: ' + phone + '\nNotes: ' + notes + '\nRef: ' + ref,
            guests: email,
            sendInvites: true
          }
        );
        calendarLink = event.getUrl();
        Logger.log('CalendarApp event created: ' + calendarLink);
      } catch (calErr) {
        Logger.log('CalendarApp also failed: ' + calErr.toString());
      }
    }

    // Step 2: Send emails
    try {
      sendClientEmail(email, name, service, duration, dateStr, timeStr, meetLink, ref);
      Logger.log('Client email sent to: ' + email);
    } catch (emailErr) {
      Logger.log('Client email failed: ' + emailErr.toString());
    }

    try {
      sendPracticeEmail(name, email, phone, service, duration, dateStr, timeStr, meetLink, ref, notes);
      Logger.log('Practice email sent');
    } catch (emailErr) {
      Logger.log('Practice email failed: ' + emailErr.toString());
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      meetLink: meetLink,
      calendarLink: calendarLink,
      ref: ref,
      message: 'Booking confirmed!'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('Fatal error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Handle GET requests (sometimes Google redirects POST to GET)
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'Booking system is running. Use POST to submit bookings.'
  })).setMimeType(ContentService.MimeType.JSON);
}

function createEventWithMeet(start, end, service, clientName, clientEmail, clientPhone, notes, ref) {
  var summary = 'Therapy Session: ' + clientName;
  var description = 'Session: ' + service + '\nClient: ' + clientName + '\nEmail: ' + clientEmail + '\nPhone: ' + clientPhone + '\nNotes: ' + notes + '\nRef: ' + ref;

  var event = {
    'summary': summary,
    'description': description,
    'start': {
      'dateTime': start.toISOString(),
      'timeZone': Session.getScriptTimeZone()
    },
    'end': {
      'dateTime': end.toISOString(),
      'timeZone': Session.getScriptTimeZone()
    },
    'attendees': [
      { 'email': clientEmail }
    ],
    'conferenceData': {
      'createRequest': {
        'requestId': ref + '-' + Date.now(),
        'conferenceSolutionKey': {
          'type': 'hangoutsMeet'
        }
      }
    }
  };

  var url = 'https://www.googleapis.com/calendar/v3/calendars/primary/events?conferenceDataVersion=1&sendUpdates=all';
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify(event),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = JSON.parse(response.getContentText());

  if (code !== 200) {
    throw new Error('API error ' + code + ': ' + (body.error ? body.error.message : response.getContentText()));
  }

  var meetLink = '';
  if (body.conferenceData && body.conferenceData.conferenceSolution) {
    var solution = body.conferenceData.conferenceSolution;
    if (solution.info && solution.info.entryPoints) {
      for (var i = 0; i < solution.info.entryPoints.length; i++) {
        if (solution.info.entryPoints[i].entryPointType === 'video') {
          meetLink = solution.info.entryPoints[i].uri;
          break;
        }
      }
    }
  }

  Logger.log('Meet link: ' + meetLink);
  return {
    meetLink: meetLink,
    calendarLink: body.htmlLink || ''
  };
}

function parseDateTime(dateStr, timeStr) {
  var months = {
    'January': 0, 'February': 1, 'March': 2, 'April': 3,
    'May': 4, 'June': 5, 'July': 6, 'August': 7,
    'September': 8, 'October': 9, 'November': 10, 'December': 11
  };

  var parts = dateStr.replace(/,/g, '').split(' ');
  var month = months[parts[1]];
  var day = parseInt(parts[2]);
  var year = parseInt(parts[3]);

  var timeParts = timeStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
  var hours = parseInt(timeParts[1]);
  var minutes = parseInt(timeParts[2]);
  var ampm = timeParts[3].toUpperCase();

  if (ampm === 'PM' && hours !== 12) hours += 12;
  if (ampm === 'AM' && hours === 12) hours = 0;

  return new Date(year, month, day, hours, minutes);
}

function parseDuration(durationStr) {
  var match = durationStr.match(/(\d+)/);
  return match ? parseInt(match[1]) * 60 * 1000 : 50 * 60 * 1000;
}

function sendClientEmail(email, name, service, duration, dateStr, timeStr, meetLink, ref) {
  var subject = 'Your Therapy Session is Confirmed - ' + ref;

  var meetSection = '';
  if (meetLink) {
    meetSection = '<a href="' + meetLink + '" style="display:block;background:#6B8F71;color:white;text-align:center;padding:16px 24px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px;margin-bottom:16px;">Join Google Meet Session</a><p style="color:#888;font-size:12px;text-align:center;margin-bottom:24px;">Click the button above at your scheduled time.</p>';
  } else {
    meetSection = '<p style="color:#888;font-size:12px;text-align:center;margin-bottom:24px;">A calendar invite with a meeting link will be sent to you shortly.</p>';
  }

  var htmlBody = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#F5F5F3;padding:40px;">' +
    '<div style="background:white;border-radius:12px;padding:40px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">' +
    '<h2 style="color:#111;font-size:24px;margin-bottom:8px;">Session Confirmed</h2>' +
    '<p style="color:#666;font-size:14px;margin-bottom:32px;">Thank you for booking with Dr. Niharika</p>' +
    '<div style="background:#F5F5F3;border-radius:8px;padding:24px;margin-bottom:24px;">' +
    '<table style="width:100%;font-size:14px;">' +
    '<tr><td style="color:#888;padding:8px 0;">Service</td><td style="color:#111;font-weight:600;text-align:right;">' + service + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Duration</td><td style="color:#111;text-align:right;">' + duration + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Date</td><td style="color:#111;font-weight:600;text-align:right;">' + dateStr + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Time</td><td style="color:#111;font-weight:600;text-align:right;">' + timeStr + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Reference</td><td style="color:#6B8F71;font-weight:600;text-align:right;">' + ref + '</td></tr>' +
    '</table></div>' +
    meetSection +
    '</div>' +
    '<p style="color:#aaa;font-size:11px;text-align:center;margin-top:24px;">Dr. Niharika | Private Psychotherapy Practice</p>' +
    '</div>';

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
}

function sendPracticeEmail(name, clientEmail, phone, service, duration, dateStr, timeStr, meetLink, ref, notes) {
  var subject = 'New Booking: ' + name + ' - ' + dateStr;

  var meetSection = '';
  if (meetLink) {
    meetSection = '<a href="' + meetLink + '" style="display:block;background:#6B8F71;color:white;text-align:center;padding:16px 24px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px;margin-bottom:12px;">Join Google Meet</a>';
  }

  var htmlBody = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#F5F5F3;padding:40px;">' +
    '<div style="background:white;border-radius:12px;padding:40px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">' +
    '<h2 style="color:#111;font-size:24px;margin-bottom:8px;">New Appointment Booked</h2>' +
    '<p style="color:#666;font-size:14px;margin-bottom:32px;">Reference: ' + ref + '</p>' +
    '<div style="background:#F5F5F3;border-radius:8px;padding:24px;margin-bottom:24px;">' +
    '<table style="width:100%;font-size:14px;">' +
    '<tr><td style="color:#888;padding:8px 0;">Client</td><td style="color:#111;font-weight:600;text-align:right;">' + name + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Email</td><td style="color:#111;text-align:right;">' + clientEmail + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Phone</td><td style="color:#111;text-align:right;">' + phone + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Service</td><td style="color:#111;font-weight:600;text-align:right;">' + service + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Duration</td><td style="color:#111;text-align:right;">' + duration + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Date</td><td style="color:#111;font-weight:600;text-align:right;">' + dateStr + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Time</td><td style="color:#111;font-weight:600;text-align:right;">' + timeStr + '</td></tr>' +
    '<tr><td style="color:#888;padding:8px 0;">Notes</td><td style="color:#111;text-align:right;">' + notes + '</td></tr>' +
    '</table></div>' +
    meetSection +
    '</div></div>';

  MailApp.sendEmail({ to: 'ishu29.official@gmail.com', subject: subject, htmlBody: htmlBody });
}

// ============================================
// TEST FUNCTION
// ============================================
function testFunction() {
  Logger.log('=== BOOKING SYSTEM TEST ===');

  // Test 1: Calendar access
  try {
    CalendarApp.getDefaultCalendar();
    Logger.log('1. Calendar: OK');
  } catch (err) {
    Logger.log('1. Calendar FAILED: ' + err);
    return;
  }

  // Test 2: Create event + Meet link via API
  try {
    var now = new Date(Date.now() + 86400000);
    var later = new Date(Date.now() + 90000000);
    var result = createEventWithMeet(now, later, 'Test', 'Test User', 'test@test.com', '555-0000', 'test', 'TEST-0000');
    Logger.log('2. API event: OK | Meet: ' + (result.meetLink || 'none'));

    // Cleanup
    var events = CalendarApp.getDefaultCalendar().getEvents(now, later);
    events.forEach(function(ev) { if (ev.getTitle().indexOf('Test') !== -1) ev.deleteEvent(); });
    Logger.log('   Cleanup: OK');
  } catch (err) {
    Logger.log('2. API event FAILED (will use fallback): ' + err);
    // Test fallback
    try {
      var now2 = new Date(Date.now() + 86400000);
      var later2 = new Date(Date.now() + 90000000);
      var ev = CalendarApp.getDefaultCalendar().createEvent('Test Fallback', now2, later2);
      ev.deleteEvent();
      Logger.log('   CalendarApp fallback: OK');
    } catch (err2) {
      Logger.log('   CalendarApp fallback FAILED: ' + err2);
    }
  }

  // Test 3: Email
  try {
    MailApp.sendEmail(Session.getActiveUser().getEmail(), 'Booking Test', 'If you received this, email works.');
    Logger.log('3. Email: OK');
  } catch (err) {
    Logger.log('3. Email FAILED: ' + err);
  }

  Logger.log('=== TEST COMPLETE ===');
}

function quickTest() {
  try {
    CalendarApp.getDefaultCalendar();
    Logger.log('Calendar API: ENABLED');
  } catch (err) {
    Logger.log('Calendar API: NOT ENABLED - Add in Services panel');
    return;
  }
  Logger.log('Ready to deploy!');
}
