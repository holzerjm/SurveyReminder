/**
 * Survey Reminder - Google Apps Script
 *
 * 1. Paste your SurveyReminder.csv data into a Google Sheet
 * 2. Open Extensions > Apps Script and paste this code
 * 3. Run sendTestEmail() first to preview one email to yourself
 * 4. Run sendAllReminders() to send to everyone
 *
 * Replae XXX with your event name
 * 
 * The format for the spreadsheet is
 * - Survey : A prefilled Google Form link pointing to the survey  
 * Email,Name,Customer,Topic,Meeting Date,Meeting time,Survey
 *
 *
 */

// ========== CONFIGURATION ==========
const SUBJECT = "Reminder: Please complete your XXX meeting survey(s)";
const SENDER_NAME = "XXX Survey Reminder";
// Set to your email to receive a test. Leave empty to skip test.
const TEST_EMAIL = "";  // e.g. "yourname@domain.com"
// ====================================

/**
 * Main function: groups meetings by email and sends one email per person.
 */
function sendAllReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1); // skip header row

  // Group rows by email address
  const grouped = {};
  rows.forEach(function(row) {
    const email = row[0].toString().trim().toLowerCase();
    if (!email) return; // skip empty rows

    if (!grouped[email]) {
      grouped[email] = {
        name: row[1].toString().trim(),
        meetings: []
      };
    }

    grouped[email].meetings.push({
      customer: row[2].toString().trim(),
      topic: row[3].toString().trim() || "(not specified)",
      date: formatDate(row[4]),
      time: row[5].toString().trim(),
      surveyUrl: row[6].toString().trim()
    });
  });

  // Send one email per person
  const emails = Object.keys(grouped);
  Logger.log("Total unique recipients: " + emails.length);

  let sentCount = 0;
  let errorCount = 0;

  emails.forEach(function(email) {
    const person = grouped[email];
    const htmlBody = buildEmailHtml(person.name, person.meetings);

    try {
      GmailApp.sendEmail(email, SUBJECT, "", {
        htmlBody: htmlBody,
        name: SENDER_NAME
      });
      sentCount++;
      Logger.log("Sent to: " + email + " (" + person.meetings.length + " surveys)");
    } catch (e) {
      errorCount++;
      Logger.log("ERROR sending to " + email + ": " + e.message);
    }
  });

  Logger.log("Done! Sent: " + sentCount + ", Errors: " + errorCount);
  SpreadsheetApp.getUi().alert(
    "Sending complete!\n\nSent: " + sentCount + "\nErrors: " + errorCount +
    "\n\nCheck View > Logs for details."
  );
}

/**
 * Test function: sends a sample email to TEST_EMAIL (or your own email).
 * Uses the first person's data as a preview.
 */
function sendTestEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert("No data found in the sheet.");
    return;
  }

  // Gather all meetings for the first email address
  const firstEmail = rows[0][0].toString().trim().toLowerCase();
  const firstName = rows[0][1].toString().trim();
  const meetings = [];

  rows.forEach(function(row) {
    if (row[0].toString().trim().toLowerCase() === firstEmail) {
      meetings.push({
        customer: row[2].toString().trim(),
        topic: row[3].toString().trim() || "(not specified)",
        date: formatDate(row[4]),
        time: row[5].toString().trim(),
        surveyUrl: row[6].toString().trim()
      });
    }
  });

  const recipient = TEST_EMAIL || Session.getActiveUser().getEmail();
  const htmlBody = buildEmailHtml(firstName, meetings);

  GmailApp.sendEmail(recipient, "[TEST] " + SUBJECT, "", {
    htmlBody: htmlBody,
    name: SENDER_NAME
  });

  SpreadsheetApp.getUi().alert(
    "Test email sent to: " + recipient +
    "\n\nPreview of email for: " + firstName + " (" + meetings.length + " surveys)" +
    "\n\nCheck your inbox to review before running sendAllReminders()."
  );
}

/**
 * Builds the HTML email body for one person.
 */
function buildEmailHtml(name, meetings) {
  let html = '<div style="font-family: Arial, sans-serif; max-width: 700px; margin: 0; text-align: left;">';

  html += '<p>Hi ' + escapeHtml(name) + ',</p>';
  html += '<p>Thank you for attending the MWC meetings. We would really appreciate it if you could take a few minutes to complete the survey for each of your meetings. Your feedback is valuable and helps us improve future engagements.</p>';
  html += '<p>You have <strong>' + meetings.length + ' survey' + (meetings.length > 1 ? 's' : '') + '</strong> to complete:</p>';

  // Meeting table
  html += '<table style="border-collapse: collapse; width: 100%; margin: 16px 0;">';
  html += '<tr style="background-color: #CC0000; color: white;">';
  html += '<th style="padding: 10px; text-align: left; border: 1px solid #ddd;">Customer</th>';
  html += '<th style="padding: 10px; text-align: left; border: 1px solid #ddd;">Topic</th>';
  html += '<th style="padding: 10px; text-align: left; border: 1px solid #ddd;">Date</th>';
  html += '<th style="padding: 10px; text-align: left; border: 1px solid #ddd;">Time</th>';
  html += '<th style="padding: 10px; text-align: center; border: 1px solid #ddd;">Survey</th>';
  html += '</tr>';

  meetings.forEach(function(m, i) {
    const bgColor = i % 2 === 0 ? '#ffffff' : '#f9f9f9';
    html += '<tr style="background-color: ' + bgColor + ';">';
    html += '<td style="padding: 8px 10px; border: 1px solid #ddd;">' + escapeHtml(m.customer) + '</td>';
    html += '<td style="padding: 8px 10px; border: 1px solid #ddd;">' + escapeHtml(m.topic) + '</td>';
    html += '<td style="padding: 8px 10px; border: 1px solid #ddd;">' + escapeHtml(m.date) + '</td>';
    html += '<td style="padding: 8px 10px; border: 1px solid #ddd;">' + escapeHtml(m.time) + '</td>';
    html += '<td style="padding: 8px 10px; border: 1px solid #ddd; text-align: center;">';
    html += '<a href="' + m.surveyUrl + '" style="display: inline-block; padding: 6px 16px; background-color: #CC0000; color: white; text-decoration: none; border-radius: 4px; font-size: 13px;">Complete Survey</a>';
    html += '</td>';
    html += '</tr>';
  });

  html += '</table>';

  html += '<p>Each survey takes only a couple of minutes to complete. The links above are pre-filled with the meeting details, so you just need to add your feedback.</p>';
  html += '<p>Thank you for your time!</p>';
  html += '<p style="color: #666; font-size: 12px; margin-top: 30px; border-top: 1px solid #ddd; padding-top: 10px;">This is an automated reminder from the XXX meeting survey team.</p>';

  html += '</div>';
  return html;
}

/**
 * Formats a date value from the spreadsheet into a readable string.
 */
function formatDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return value.toString().trim();
}

/**
 * Escapes HTML special characters to prevent injection.
 */
function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
