/**
 * Google Apps Script Email Relay for Work Reporter
 * ------------------------------------------------
 * This script runs completely FREE on Google Apps Script (port 443 HTTPS).
 * It bypasses cloud firewall blocks (like Render Free Tier port 587 block)
 * and sends emails directly through your own Gmail account with 0 hassle.
 * 
 * SETUP INSTRUCTIONS (Takes 2 minutes):
 * 1. Open https://script.google.com while signed in to your Gmail account.
 * 2. Click "New project" (+).
 * 3. Delete any default code in Code.gs, paste this entire script, and click the Save icon (💾).
 * 4. Click "Deploy" (top right) -> "New deployment".
 * 5. Click the gear icon next to "Select type" -> choose "Web app".
 * 6. Set:
 *    - Description: Work Reporter Email Relay
 *    - Execute as: Me (your Gmail address)
 *    - Who has access: Anyone
 * 7. Click "Deploy".
 * 8. If prompted, click "Authorize access" -> choose your Google Account -> "Advanced" -> "Go to Untitled project (unsafe)" -> "Allow".
 * 9. Copy the "Web app URL" (looks like: https://script.google.com/macros/s/.../exec).
 * 10. Open Work Reporter on your phone/browser -> Settings -> paste the URL into "Google Apps Script Relay URL" -> Save!
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var options = {
      htmlBody: data.body_html,
      name: "Work Reporter"
    };

    if (data.cc && data.cc.trim()) {
      options.cc = data.cc.trim();
    }
    if (data.reply_to && data.reply_to.trim()) {
      options.replyTo = data.reply_to.trim();
    }
    if (data.attachments && data.attachments.length > 0) {
      options.attachments = data.attachments.map(function(att) {
        var bytes = Utilities.base64Decode(att.data_base64);
        return Utilities.newBlob(bytes, att.mimetype || "application/octet-stream", att.filename);
      });
    }

    GmailApp.sendEmail(
      data.to,
      data.subject,
      data.body_text || "Please view this email in an HTML-compatible client.",
      options
    );

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", message: "Email sent successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Work Reporter Relay is active" }))
    .setMimeType(ContentService.MimeType.JSON);
}
