/**
 * JCC SITE CONTENT MANAGER — BACKEND
 * ------------------------------------------------
 * Paste this entire file into Extensions > Apps Script
 * on a blank Google Sheet, then deploy as a Web App.
 * See SETUP_INSTRUCTIONS.txt for the full walkthrough.
 * ------------------------------------------------
 */

var SHEET_NAME = 'SiteContent';
var DRIVE_FOLDER_NAME = 'MY WEBSITE FILES';
var NOTIFY_EMAIL = 'your-email@example.com'; // FALLBACK ONLY — set the real notification email from the admin panel instead (Notifications section). This is only used if that hasn't been saved yet.

// ---------- READ (called by the live site + admin panel on load) ----------
function doGet(e) {
  var type = (e && e.parameter && e.parameter.type) || 'content';

  if (type === 'testimonials') {
    return jsonResponse_({ success: true, testimonials: getTestimonials_(true) });
  }

  if (type === 'testimonials_all') {
    return jsonResponse_({ success: true, testimonials: getTestimonials_(false) });
  }

  var sheet = getSheet_();
  var json = sheet.getRange('A1').getValue();
  var content = json ? JSON.parse(json) : {};

  // Bundled in here too, so the live site only needs ONE request instead of
  // two — each separate request pays Apps Script's cold-start delay on its
  // own, so merging them roughly halves that tax on page load.
  content.approvedTestimonials = getTestimonials_(true);

  return jsonResponse_(content);
}

// ---------- WRITE (called by the admin panel: save content / upload photo) ----------
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ success: false, error: 'Bad request body' });
  }

  var action = body.action;

  if (action === 'saveContent') {
    var sheet = getSheet_();
    sheet.getRange('A1').setValue(JSON.stringify(body.data));
    sheet.getRange('B1').setValue(new Date());
    return jsonResponse_({ success: true });
  }

  if (action === 'uploadPhoto') {
    return handleUpload_(body);
  }

  if (action === 'submitLead') {
    return handleLeadSubmit_(body);
  }

  if (action === 'submitTestimonial') {
    return handleTestimonialSubmit_(body);
  }

  if (action === 'updateTestimonialStatus') {
    return handleTestimonialStatusUpdate_(body);
  }

  if (action === 'deleteTestimonial') {
    return handleTestimonialDelete_(body);
  }

  return jsonResponse_({ success: false, error: 'Unknown action: ' + action });
}

function handleLeadSubmit_(body) {
  try {
    var d = body.data || {};
    var sheet = getLeadsSheet_();
    sheet.appendRow([
      new Date(),
      d.productName || '',
      d.name || '',
      d.birthday || '',
      d.gender || '',
      d.mobile || '',
      d.email || ''
    ]);

    var recipient = getNotifyEmail_();
    var subject = 'New Sample Quotation Request — ' + (d.productName || 'Product');
    var messageBody =
      'You have a new sample quotation request from your website.\n\n' +
      'Product: ' + (d.productName || '') + '\n' +
      '-----------------------------------\n' +
      'Name: ' + (d.name || '') + '\n' +
      'Birthday: ' + (d.birthday || '') + '\n' +
      'Gender: ' + (d.gender || '') + '\n' +
      'Mobile No: ' + (d.mobile || '') + '\n' +
      'Email: ' + (d.email || '') + '\n' +
      '-----------------------------------\n\n' +
      'Submitted: ' + new Date().toString();

    try {
      MailApp.sendEmail(recipient, subject, messageBody);
    } catch (mailErr) {
      // The lead is already saved above — don't let an email hiccup report
      // the whole submission as failed. Log it so it's visible in Executions.
      console.error('Lead saved, but notification email failed: ' + mailErr.message);
    }

    return jsonResponse_({ success: true });
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

function handleUpload_(body) {
  try {
    var dataUrl = body.dataUrl || '';
    var match = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
    if (!match) {
      return jsonResponse_({ success: false, error: 'Invalid image data' });
    }
    var mimeType = match[1];
    var base64 = match[2];
    var bytes = Utilities.base64Decode(base64);

    var safeName = (body.filename || 'photo').replace(/[^a-zA-Z0-9-_]/g, '_');
    var ext = mimeType.split('/')[1] || 'jpg';
    var blob = Utilities.newBlob(bytes, mimeType, safeName + '_' + Date.now() + '.' + ext);

    var folder = getFolder_();
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileId = file.getId();
    // This URL format renders reliably as an <img src> for publicly shared Drive files.
    var url = 'https://lh3.googleusercontent.com/d/' + fileId;

    return jsonResponse_({ success: true, url: url, fileId: fileId });
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

// ---------- HELPERS ----------
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1').setValue('{}');
    sheet.getRange('A1:B1').setNote('A1 = full site content (JSON). B1 = last saved timestamp. Do not edit A1 by hand.');
  }
  return sheet;
}

// Reads the notification email from the admin panel's saved settings.
// Falls back to the NOTIFY_EMAIL constant above if it hasn't been set yet
// (e.g. on a brand new deployment before the owner has opened admin.html).
function getNotifyEmail_() {
  try {
    var sheet = getSheet_();
    var json = sheet.getRange('A1').getValue();
    if (json) {
      var content = JSON.parse(json);
      if (content.notifyEmail) return content.notifyEmail;
    }
  } catch (e) {
    // fall through to the constant below
  }
  return NOTIFY_EMAIL;
}

function getFolder_() {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function getLeadsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Leads');
  if (!sheet) {
    sheet = ss.insertSheet('Leads');
    sheet.appendRow(['Timestamp', 'Product', 'Name', 'Birthday', 'Gender', 'Mobile No', 'Email']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:G1').setFontWeight('bold');
  }
  return sheet;
}

// ---------- TESTIMONIALS ----------
function getTestimonialsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Testimonials');
  if (!sheet) {
    sheet = ss.insertSheet('Testimonials');
    sheet.appendRow(['ID', 'Timestamp', 'Name', 'Role', 'Rating', 'Message', 'Status']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:G1').setFontWeight('bold');
  }
  return sheet;
}

function handleTestimonialSubmit_(body) {
  try {
    var d = body.data || {};
    var sheet = getTestimonialsSheet_();
    var id = 't_' + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
    var rating = Math.max(1, Math.min(5, parseInt(d.rating, 10) || 5));

    sheet.appendRow([
      id,
      new Date(),
      d.name || '',
      d.role || '',
      rating,
      d.message || '',
      'pending'
    ]);

    // Let the site owner know a new review is waiting for approval
    try {
      MailApp.sendEmail(
        getNotifyEmail_(),
        'New client feedback awaiting approval',
        'A new testimonial was submitted on your website and is waiting for your approval.\n\n' +
        'Name: ' + (d.name || '') + '\n' +
        'Role: ' + (d.role || '') + '\n' +
        'Rating: ' + rating + ' / 5\n' +
        'Message: ' + (d.message || '') + '\n\n' +
        'Go to your admin panel to approve or reject it.'
      );
    } catch (mailErr) {
      // don't fail the whole submission if the email step has an issue
    }

    return jsonResponse_({ success: true });
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

function getTestimonials_(approvedOnly) {
  var sheet = getTestimonialsSheet_();
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var status = row[6] || 'pending';
    if (approvedOnly && status !== 'approved') continue;
    out.push({
      id: row[0],
      timestamp: row[1] instanceof Date ? row[1].toISOString() : String(row[1]),
      name: row[2],
      role: row[3],
      rating: row[4],
      message: row[5],
      status: status
    });
  }
  // newest first
  out.reverse();
  return out;
}

function handleTestimonialStatusUpdate_(body) {
  try {
    var id = body.id;
    var newStatus = body.status;
    if (!id || !newStatus) return jsonResponse_({ success: false, error: 'Missing id or status' });

    var sheet = getTestimonialsSheet_();
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.getRange(i + 1, 7).setValue(newStatus); // column G = Status
        return jsonResponse_({ success: true });
      }
    }
    return jsonResponse_({ success: false, error: 'Testimonial not found' });
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

function handleTestimonialDelete_(body) {
  try {
    var id = body.id;
    if (!id) return jsonResponse_({ success: false, error: 'Missing id' });

    var sheet = getTestimonialsSheet_();
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.deleteRow(i + 1);
        return jsonResponse_({ success: true });
      }
    }
    return jsonResponse_({ success: false, error: 'Testimonial not found' });
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * ONE-TIME AUTHORIZATION HELPER
 * ------------------------------------------------
 * Run this once manually (see below) to grant the mail-sending permission.
 * Web apps deployed as "Execute as: Me" don't automatically get new
 * permissions just from redeploying — Google requires you to explicitly
 * authorize each capability (like sending email) by running the script
 * yourself at least once.
 *
 * HOW TO RUN IT:
 *   1. In the Apps Script editor, use the function dropdown at the top
 *      (next to the Run button) and select "authorizeMe".
 *   2. Click "Run".
 *   3. A permissions popup will appear — click "Review permissions",
 *      choose your account, click "Advanced", then "Go to [project name]
 *      (unsafe)", then "Allow".
 *   4. Check your email — you should receive a test message titled
 *      "JCC Site — Authorization Test".
 *   5. That's it. The lead form on your live site will now be able to
 *      send you emails. You do NOT need to redeploy after this step.
 */
function authorizeMe() {
  MailApp.sendEmail(
    getNotifyEmail_(),
    'JCC Site — Authorization Test',
    'If you are reading this, email notifications are now authorized and working correctly.'
  );
}

/**
 * DIAGNOSTIC — run this to see exactly what's going on, no guessing.
 * ------------------------------------------------
 * HOW TO RUN IT AND SEE THE RESULT:
 *   1. Function dropdown (top, next to Run) → select "debugNotifyEmail".
 *   2. Click Run.
 *   3. Click "Executions" in the left sidebar (clock icon).
 *   4. Click the most recent run at the top of that list.
 *   5. You'll see exactly what email address is being used, and the
 *      raw saved content, right there in the log output.
 */
function debugNotifyEmail() {
  var resolved = getNotifyEmail_();
  console.log('=== RESOLVED NOTIFICATION EMAIL: ' + resolved + ' ===');

  var sheet = getSheet_();
  var rawJson = sheet.getRange('A1').getValue();
  console.log('=== RAW SAVED CONTENT (from admin.html Save Changes) ===');
  console.log(rawJson || '(A1 is empty — nothing has ever been saved from admin.html to this Sheet)');

  var quota = MailApp.getRemainingDailyQuota();
  console.log('=== REMAINING EMAIL QUOTA TODAY: ' + quota + ' ===');
}

/**
 * KEEP-WARM TRIGGER (optional, experimental)
 * ------------------------------------------------
 * Apps Script Web Apps can go "cold" after a period of no traffic, adding
 * a 1-3 second delay to the next real visitor's page load while Google
 * spins the script back up. This periodically pings your own site to try
 * to reduce how often that happens.
 *
 * IMPORTANT HONESTY NOTE: Google doesn't officially document or guarantee
 * that self-pinging actually prevents cold starts on Apps Script Web Apps
 * specifically (unlike some other platforms with an official "keep warm"
 * feature). Many people report it helping; it's not guaranteed. Consider
 * this a reasonable experiment, not a fix you can be 100% sure of.
 *
 * HOW TO TURN IT ON (run once):
 *   1. Function dropdown → select "setupKeepWarmTrigger".
 *   2. Click Run. Approve any permission prompt if asked.
 *   3. Done — it now pings your site automatically every 10 minutes.
 *
 * HOW TO TURN IT OFF LATER:
 *   1. Function dropdown → select "removeKeepWarmTrigger".
 *   2. Click Run.
 */
function keepWarm() {
  try {
    var url = ScriptApp.getService().getUrl();
    UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (err) {
    console.error('keepWarm ping failed: ' + err.message);
  }
}

function setupKeepWarmTrigger() {
  removeKeepWarmTrigger(); // avoid creating duplicates if run more than once
  ScriptApp.newTrigger('keepWarm')
    .timeBased()
    .everyMinutes(10)
    .create();
  console.log('Keep-warm trigger installed — pinging every 10 minutes.');
}

function removeKeepWarmTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'keepWarm') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  console.log('Removed ' + removed + ' existing keep-warm trigger(s).');
}
