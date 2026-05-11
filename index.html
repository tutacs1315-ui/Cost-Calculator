/**
 * ══════════════════════════════════════════════════════════
 * AKL CARPET CARE — Google Apps Script Backend
 * ══════════════════════════════════════════════════════════
 *
 * SETUP INSTRUCTIONS (one-time, takes ~5 minutes):
 *
 * 1. Go to script.google.com → click "New project"
 * 2. Delete any existing code and paste this entire file
 * 3. Fill in YOUR values in the CONFIGURATION section below
 * 4. Click "Save" (floppy disk icon)
 * 5. Click "Deploy" → "New deployment"
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Click "Deploy" → copy the Web App URL
 * 7. Paste that URL into AKL_Calculator_v2.html
 *    where it says: const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE'
 *
 * DONE. Every calculator submission will now:
 *   ✅ Log a row to your Google Sheet
 *   ✅ Create an AMBER calendar event in Xavier's calendar
 *   ✅ Send the client a confirmation email
 * ══════════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════
// CONFIGURATION — fill these in before deploying
// ══════════════════════════════════════════════

const CONFIG = {

  // Your Google Sheet ID
  // (the long string in the URL: docs.google.com/spreadsheets/d/THIS_PART/edit)
  SHEET_ID: 'YOUR_GOOGLE_SHEET_ID_HERE',

  // The tab name inside your Sheet to log jobs
  SHEET_TAB: 'Enquiries',

  // Xavier's Google Calendar ID
  // Go to Google Calendar → Settings → click your calendar → scroll to "Calendar ID"
  // Usually looks like: xavier@gmail.com  OR  a long string ending in @group.calendar.google.com
  CALENDAR_ID: 'YOUR_CALENDAR_ID_HERE',

  // Your business email — you'll get a BCC of every submission
  BUSINESS_EMAIL: 'admin@aucklandcarpetcare.co.nz',

  // Business name shown in emails
  BUSINESS_NAME: 'AKL Carpet Care',

  // Business phone shown in emails
  BUSINESS_PHONE: '021 075 0733',
};

// ══════════════════════════════════════════════
// MAIN HANDLER — runs on every form submission
// ══════════════════════════════════════════════

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Run all three actions
    logToSheet(data);
    createCalendarEvent(data);
    sendConfirmationEmail(data);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Error: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ══════════════════════════════════════════════
// 1. LOG TO GOOGLE SHEETS
// ══════════════════════════════════════════════

function logToSheet(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet   = ss.getSheetByName(CONFIG.SHEET_TAB);

  // Create the tab if it doesn't exist yet
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_TAB);
    // Add header row
    sheet.appendRow([
      'Timestamp', 'Name', 'Phone', 'Email', 'Address', 'Suburb',
      'Preferred Day', 'Preferred Time', 'Contact Method',
      'Carpet Type', 'Repeat Client', 'Travel Fee',
      'Estimate Total', 'Status', 'Job Summary'
    ]);
    // Style header row
    const headerRange = sheet.getRange(1, 1, 1, 15);
    headerRange.setBackground('#0D2654');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Append the new enquiry row
  sheet.appendRow([
    new Date(),                          // Timestamp
    data.name        || '',              // Name
    data.phone       || '',              // Phone
    data.email       || '',              // Email
    data.address     || '',              // Address
    data.suburb      || '',              // Suburb
    data.preferredDay  || '',            // Preferred Day
    data.preferredTime || '',            // Preferred Time
    data.contactMethod || '',            // Contact Method
    data.carpetType  || '',              // Carpet Type
    data.repeatClient|| '',              // Repeat Client
    data.travelFee   || 'None',          // Travel Fee
    data.estimateTotal || '',            // Estimate Total
    'AMBER — Pending Confirmation',      // Status (amber until manually confirmed)
    data.jobSummary  || '',              // Full Job Sheet
  ]);

  // Colour the new row amber
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 15).setBackground('#FFF3CD');
  sheet.getRange(lastRow, 14).setFontColor('#856404').setFontWeight('bold'); // Status cell
}

// ══════════════════════════════════════════════
// 2. CREATE AMBER CALENDAR EVENT
// ══════════════════════════════════════════════

function createCalendarEvent(data) {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) {
    Logger.log('Calendar not found: ' + CONFIG.CALENDAR_ID);
    return;
  }

  // Work out the event date
  // If they selected a specific day of the week, find the next occurrence
  // Otherwise use tomorrow as a placeholder
  const eventDate = resolvePreferredDate(data.preferredDay);

  // Work out start/end time from preference
  const times = resolvePreferredTime(data.preferredTime, eventDate);

  // Build event title
  const title = `⚠️ AMBER — ${data.name} | ${data.suburb || 'Suburb TBC'} | ${data.estimateTotal}`;

  // Build event description
  const description = [
    '⚠️ PENDING CONFIRMATION — Contact client to confirm time',
    '',
    `Client: ${data.name}`,
    `Phone: ${data.phone}`,
    `Email: ${data.email}`,
    `Address: ${data.address || 'TBC'}`,
    `Suburb: ${data.suburb || 'TBC'}`,
    `Contact via: ${data.contactMethod || 'TBC'}`,
    '',
    `Estimate: ${data.estimateTotal}`,
    `Travel fee: ${data.travelFee}`,
    `Repeat client: ${data.repeatClient}`,
    '',
    `Preferred day: ${data.preferredDay || 'Flexible'}`,
    `Preferred time: ${data.preferredTime || 'Flexible'}`,
    '',
    '─── JOB SHEET ───',
    data.jobSummary || '',
    '',
    '─────────────────',
    'Once confirmed: change event colour to GREEN and update status in Sheets.',
  ].join('\n');

  // Create the event
  const event = calendar.createEvent(title, times.start, times.end, {
    description: description,
    location: data.address || data.suburb || '',
  });

  // Set event colour to YELLOW (Banana) to represent amber/pending
  event.setColor(CalendarApp.EventColor.YELLOW);
}

function resolvePreferredDate(preferredDay) {
  const today = new Date();
  const dayMap = { 'Mon': 1, 'Tue': 2, 'Wed': 3, 'Thu': 4, 'Fri': 5, 'Sat AM': 6 };

  if (preferredDay && dayMap[preferredDay] !== undefined) {
    const targetDay = dayMap[preferredDay];
    const currentDay = today.getDay(); // 0=Sun, 1=Mon...
    let daysUntil = targetDay - currentDay;
    if (daysUntil <= 0) daysUntil += 7; // next week if already passed
    const result = new Date(today);
    result.setDate(today.getDate() + daysUntil);
    return result;
  }

  // Default: tomorrow
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  return tomorrow;
}

function resolvePreferredTime(preferredTime, baseDate) {
  const timeMap = {
    '8am–10am':  { startH: 8,  endH: 10 },
    '11am–1pm':  { startH: 11, endH: 13 },
    '2pm–4pm':   { startH: 14, endH: 16 },
    'Flexible':  { startH: 9,  endH: 11 },
  };

  const t = timeMap[preferredTime] || { startH: 9, endH: 11 };

  const start = new Date(baseDate);
  start.setHours(t.startH, 0, 0, 0);

  const end = new Date(baseDate);
  end.setHours(t.endH, 0, 0, 0);

  return { start, end };
}

// ══════════════════════════════════════════════
// 3. SEND CONFIRMATION EMAIL TO CLIENT
// ══════════════════════════════════════════════

function sendConfirmationEmail(data) {
  if (!data.email) return;

  const firstName = data.name ? data.name.split(' ')[0] : 'there';
  const contactMethodLabel = {
    'phone-call': 'phone call',
    'text':       'text message',
    'email':      'email',
  }[data.contactMethod] || 'phone';

  const subject = `AKL Carpet Care — We've received your booking request`;

  const body = `
Kia Ora ${firstName},

Thanks for reaching out to AKL Carpet Care!

We've received your booking request and someone from our team will be in touch with you on ${data.phone} via ${contactMethodLabel} to confirm a time.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
YOUR REQUEST SUMMARY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Suburb: ${data.suburb || 'TBC'}
Preferred day: ${data.preferredDay || 'Flexible'}
Preferred time: ${data.preferredTime || 'Flexible'}
Estimate total: ${data.estimateTotal}

Please note: this is an estimate only. Final pricing is confirmed once we've assessed the space on-site or reviewed photos.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

If you have any questions in the meantime, feel free to call or text us directly:
📞 ${CONFIG.BUSINESS_PHONE}
✉️ ${CONFIG.BUSINESS_EMAIL}

"We stand behind our work — or we'll make it right."

Ngā mihi,
The AKL Carpet Care Team
aucklandcarpetcare.co.nz
`.trim();

  // Send to client
  GmailApp.sendEmail(data.email, subject, body, {
    name: CONFIG.BUSINESS_NAME,
    replyTo: CONFIG.BUSINESS_EMAIL,
    // BCC yourself so you see every confirmation sent
    bcc: CONFIG.BUSINESS_EMAIL,
  });
}

// ══════════════════════════════════════════════
// TEST FUNCTION — run this manually to verify setup
// Select "testSetup" from the function dropdown and click Run
// ══════════════════════════════════════════════

function testSetup() {
  const testData = {
    name:          'Test Client',
    phone:         '021 000 0001',
    email:         CONFIG.BUSINESS_EMAIL, // sends test email to yourself
    address:       '34 Stoddard Road',
    suburb:        'Mt Roskill',
    preferredDay:  'Mon',
    preferredTime: '8am–10am',
    contactMethod: 'phone-call',
    carpetType:    'Synthetic (Nylon / Polyester)',
    repeatClient:  'No',
    travelFee:     'None',
    estimateTotal: '$195.00',
    jobSummary:    '=== TEST JOB SHEET ===\nThis is a test submission.',
    timestamp:     new Date().toISOString(),
  };

  logToSheet(testData);
  Logger.log('✅ Sheet row logged');

  createCalendarEvent(testData);
  Logger.log('✅ Calendar event created');

  sendConfirmationEmail(testData);
  Logger.log('✅ Confirmation email sent');

  Logger.log('All three actions completed successfully!');
}
