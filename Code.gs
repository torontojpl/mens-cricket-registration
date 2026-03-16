/**
 * Cricket Player Registration — Google Apps Script v3
 * =====================================================
 * Features:
 *   - 4 consent checkboxes tracked individually
 *   - Drawn signature captured as Base64 PNG
 *   - Signature embedded in the PDF registration card
 *   - Photo saved to Google Drive
 *   - PDF emailed to admin AND player as attachment
 *   - All data saved to Google Sheets
 *
 * UPDATE the CONFIG section before deploying.
 */

// ── CONFIG ──────────────────────────────────────────
var ADMIN_EMAIL  = "you@gmail.com";            // ← your email
var SHEET_NAME   = "Registrations";
var DRIVE_FOLDER = "Cricket Player Photos";
// ────────────────────────────────────────────────────

var HEADERS = [
  "ID", "Registered At",
  "First Name", "Last Name", "Date of Birth", "Gender",
  "Email", "Phone", "Address",
  "Emergency Contact", "Emergency Phone",
  "Batting Style", "Bowling Style", "Positions",
  "Experience", "Years Playing", "Shirt Number",
  "Previous Clubs", "Availability", "Notes",
  "Consent: Player/Code", "Consent: Parent/Guardian",
  "Consent: Medical", "Consent: Photo/Media",
  "Signature Provided", "Photo Link", "Status"
];


// ── MAIN POST HANDLER ───────────────────────────────
function doPost(e) {
  try {
    var data = e.parameter;

    // ── 1. Sheet setup ──────────────────────────────
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      var hr = sheet.getRange(1, 1, 1, HEADERS.length);
      hr.setBackground("#1a7c3e").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var now   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var newId = sheet.getLastRow();

    var positions    = data["position[]"]     || data["position"]     || "";
    var availability = data["availability[]"] || data["availability"] || "";

    // Consent values
    var cPlayer   = data["consent_player"]   ? "Yes" : "No";
    var cGuardian = data["consent_guardian"] ? "Yes" : "No";
    var cMedical  = data["consent_medical"]  ? "Yes" : "No";
    var cMedia    = data["consent_media"]    ? "Yes" : "No";
    var hasSig    = (data["signature_base64"] && data["signature_base64"].length > 100) ? "Yes" : "No";

    // ── 2. Save photo to Drive ──────────────────────
    var photoLink = "";
    if (data.photo_base64 && data.photo_base64.length > 100) {
      photoLink = savePhotoToDrive(data, newId);
    }

    // ── 3. Append row ───────────────────────────────
    sheet.appendRow([
      newId, now,
      data.first_name    || "", data.last_name      || "",
      data.dob           || "", data.gender         || "",
      data.email         || "", data.phone          || "",
      data.address       || "",
      data.emergency_name  || "", data.emergency_phone || "",
      data.batting_style || "", data.bowling_style  || "",
      positions,
      data.experience    || "", data.years_playing  || "",
      data.shirt_number  || "", data.previous_clubs || "",
      availability, data.notes || "",
      cPlayer, cGuardian, cMedical, cMedia,
      hasSig, photoLink, "Pending"
    ]);

    sheet.autoResizeColumns(1, HEADERS.length);

    // ── 4. Generate PDF with signature ─────────────
    var pdfBlob = generatePDF(data, newId, now, positions, availability, photoLink, cPlayer, cGuardian, cMedical, cMedia);

    // ── 5. Send emails ──────────────────────────────
    sendAdminEmail(data, newId, now, positions, availability, photoLink, cPlayer, cGuardian, cMedical, cMedia, pdfBlob);
    if (data.email) sendPlayerEmail(data, newId, pdfBlob);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", id: newId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("Error: " + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── SAVE PHOTO TO DRIVE ─────────────────────────────
function savePhotoToDrive(data, regId) {
  try {
    var decoded  = Utilities.base64Decode(data.photo_base64);
    var blob     = Utilities.newBlob(decoded, data.photo_mime || "image/jpeg", "photo.jpg");
    var folders  = DriveApp.getFoldersByName(DRIVE_FOLDER);
    var folder   = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER);
    var file     = folder.createFile(blob);
    file.setName("Player_" + regId + "_" + (data.first_name||"") + "_" + (data.last_name||"") + "_photo");
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(err) {
    Logger.log("Photo error: " + err.toString());
    return "";
  }
}


// ── GENERATE PDF ────────────────────────────────────
function generatePDF(data, regId, now, positions, availability, photoLink, cPlayer, cGuardian, cMedical, cMedia) {
  var fullName = (data.first_name||"") + " " + (data.last_name||"");

  // Build signature img tag — embed directly as Base64 data URI
  var sigImg = "";
  if (data.signature_base64 && data.signature_base64.length > 100) {
    sigImg = '<img src="data:image/png;base64,' + data.signature_base64 + '" style="height:70px;max-width:280px;object-fit:contain;display:block;"/>';
  } else {
    sigImg = '<span style="color:#aaa;font-size:13px;font-style:italic">No signature provided</span>';
  }

  function tick(v) { return v === "Yes" ? "&#10003;" : "&#10007;"; }
  function tickColor(v) { return v === "Yes" ? "#1a7c3e" : "#c0392b"; }

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    'body{font-family:Arial,sans-serif;font-size:12px;color:#1a1a1a;padding:32px}' +
    '.header{background:#0f5229;color:white;padding:22px 28px;border-radius:10px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center}' +
    '.header h1{font-size:20px;font-weight:800;margin-bottom:3px}' +
    '.header p{font-size:11px;opacity:.7}' +
    '.ref{background:#c9a227;color:#0f5229;font-weight:700;padding:6px 14px;border-radius:20px;font-size:13px;white-space:nowrap}' +
    '.section{margin-bottom:18px}' +
    '.stitle{font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#1a7c3e;border-bottom:1.5px solid #e8f5ed;padding-bottom:4px;margin-bottom:10px}' +
    '.grid{display:grid;grid-template-columns:1fr 1fr;gap:8px 20px}' +
    '.field label{font-size:10px;color:#888;font-weight:700;text-transform:uppercase;letter-spacing:.5px;display:block;margin-bottom:2px}' +
    '.field span{font-size:12px;color:#1a1a1a;font-weight:500}' +
    '.badge{display:inline-block;background:#e8f5ed;color:#0f5229;border:1px solid #9fe1cb;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;margin:2px}' +
    '.photo-row{display:flex;gap:16px;align-items:flex-start;margin-bottom:18px}' +
    '.photo-box{width:80px;height:96px;border:2px solid #d0d7de;border-radius:8px;overflow:hidden;flex-shrink:0;display:flex;align-items:center;justify-content:center;background:#f7f8fa;color:#aaa;font-size:10px;text-align:center}' +
    '.photo-box img{width:100%;height:100%;object-fit:cover}' +
    '.consent-row{display:flex;align-items:flex-start;gap:10px;margin-bottom:10px;padding:10px 12px;border-radius:8px;background:#fafafa}' +
    '.tick{font-size:16px;font-weight:700;flex-shrink:0;margin-top:-1px}' +
    '.consent-title{font-size:12px;font-weight:600;color:#1a1a1a}' +
    '.consent-desc{font-size:11px;color:#666;line-height:1.5;margin-top:2px}' +
    '.sig-section{margin-top:18px;padding:16px;border:1.5px solid #dde3e8;border-radius:10px;background:#fafafa}' +
    '.sig-section .stitle{margin-bottom:10px}' +
    '.sig-box{min-height:80px;display:flex;align-items:center;padding:10px 0}' +
    '.sig-meta{font-size:10px;color:#aaa;margin-top:6px}' +
    '.status-bar{background:#fff8e1;border:1.5px solid #e8d48a;border-radius:8px;padding:10px 14px;margin-bottom:18px;font-size:11px;color:#5a4a00;display:flex;justify-content:space-between;align-items:center}' +
    '.status-pill{background:#faeeda;color:#854f0b;font-weight:700;padding:2px 10px;border-radius:10px}' +
    '.footer{margin-top:24px;border-top:1px solid #e0e0e0;padding-top:10px;font-size:10px;color:#aaa;text-align:center}' +
    '</style></head><body>' +

    '<div class="header"><div><h1>Cricket Club</h1><p>Player Registration Card &nbsp;|&nbsp; ' + now + '</p></div><div class="ref">Ref #' + regId + '</div></div>' +
    '<div class="status-bar"><span>Registration received and pending review.</span><span class="status-pill">Pending</span></div>' +

    // Personal + photo
    '<div class="photo-row">' +
    '<div class="photo-box">' + (photoLink && photoLink.startsWith("http") ? '<p style="padding:6px;font-size:9px;color:#888">See Drive link</p>' : '<p>No photo</p>') + '</div>' +
    '<div style="flex:1"><div class="stitle">Personal information</div>' +
    '<div class="grid">' +
    '<div class="field"><label>Full name</label><span>' + fullName + '</span></div>' +
    '<div class="field"><label>Date of birth</label><span>' + (data.dob||"—") + '</span></div>' +
    '<div class="field"><label>Email</label><span>' + (data.email||"—") + '</span></div>' +
    '<div class="field"><label>Phone</label><span>' + (data.phone||"—") + '</span></div>' +
    '<div class="field"><label>Address</label><span>' + (data.address||"—") + '</span></div>' +
    '</div></div></div>' +

    // Cricket profile
    '<div class="section"><div class="stitle">Cricket profile</div>' +
    '<div class="grid">' +
    '<div class="field"><label>Batting</label><span>' + (data.batting_style||"—") + '</span></div>' +
    '<div class="field"><label>Bowling</label><span>' + (data.bowling_style||"—") + '</span></div>' +
    '<div class="field"><label>Experience</label><span>' + (data.experience||"—") + '</span></div>' +
    '<div class="field"><label>Years playing</label><span>' + (data.years_playing||"—") + '</span></div>' +
    '</div>' +
    '<div class="field" style="margin-top:8px"><label>Positions</label><div>' +
    (positions ? positions.split(",").map(function(p){return'<span class="badge">'+p.trim()+'</span>'}).join("") : "—") +
    '</div></div></div>' +

    // Availability
    '<div class="section"><div class="stitle">Availability</div>' +
    '<div>' + (availability ? availability.split(",").map(function(a){return'<span class="badge">'+a.trim()+'</span>'}).join("") : "—") + '</div></div>' +

    // Emergency
    '<div class="section"><div class="stitle">Emergency contact</div>' +
    '<div class="grid">' +
    '<div class="field"><label>Name</label><span>' + (data.emergency_name||"—") + '</span></div>' +
    '<div class="field"><label>Phone</label><span>' + (data.emergency_phone||"—") + '</span></div>' +
    '</div></div>' +

    // Consents
    '<div class="section"><div class="stitle">Consent declarations</div>' +

    '<div class="consent-row">' +
    '<div class="tick" style="color:' + tickColor(cPlayer) + '">' + tick(cPlayer) + '</div>' +
    '<div><div class="consent-title">Player consent &amp; code of conduct</div>' +
    '<div class="consent-desc">Agreed to abide by club rules, respect all players and officials.</div></div></div>' +

    '<div class="consent-row">' +
    '<div class="tick" style="color:' + tickColor(cGuardian) + '">' + tick(cGuardian) + '</div>' +
    '<div><div class="consent-title">Parent / guardian consent</div>' +
    '<div class="consent-desc">Consented to participation in all club activities (or confirmed N/A).</div></div></div>' +

    '<div class="consent-row">' +
    '<div class="tick" style="color:' + tickColor(cMedical) + '">' + tick(cMedical) + '</div>' +
    '<div><div class="consent-title">Medical disclosure consent</div>' +
    '<div class="consent-desc">Confirmed medically fit; consented to emergency treatment if required.</div></div></div>' +

    '<div class="consent-row">' +
    '<div class="tick" style="color:' + tickColor(cMedia) + '">' + tick(cMedia) + '</div>' +
    '<div><div class="consent-title">Photo &amp; media usage consent</div>' +
    '<div class="consent-desc">Consented to use of images/videos for club publications and social media.</div></div></div>' +
    '</div>' +

    // Signature
    '<div class="sig-section"><div class="stitle">Signature</div>' +
    '<div class="sig-box">' + sigImg + '</div>' +
    '<div class="sig-meta">Signed electronically on ' + now + ' &nbsp;|&nbsp; Ref #' + regId + '</div></div>' +

    (photoLink && photoLink.startsWith("http") ? '<div class="section" style="margin-top:16px"><div class="stitle">Player photo</div><p style="font-size:11px;color:#1a7c3e">' + photoLink + '</p></div>' : '') +

    '<div class="footer">Cricket Club &nbsp;|&nbsp; Registration #' + regId + ' &nbsp;|&nbsp; Generated ' + now + '<br>This document is confidential and for club use only.</div>' +
    '</body></html>';

  var blob = Utilities.newBlob(html, "text/html", "reg.html");
  var pdf  = blob.getAs("application/pdf");
  pdf.setName("CricketRegistration_" + regId + "_" + (data.first_name||"") + "_" + (data.last_name||"") + ".pdf");
  return pdf;
}


// ── ADMIN EMAIL ─────────────────────────────────────
function sendAdminEmail(data, regId, now, positions, availability, photoLink, cPlayer, cGuardian, cMedical, cMedia, pdfBlob) {
  var subject = "New Cricket Registration: " + data.first_name + " " + data.last_name + " [#" + regId + "]";
  var body =
    "New registration received.\n\n" +
    "Ref #       : " + regId + "\n" +
    "Name        : " + data.first_name + " " + data.last_name + "\n" +
    "Email       : " + (data.email||"—") + "\n" +
    "Phone       : " + (data.phone||"—") + "\n" +
    "Experience  : " + (data.experience||"—") + "\n" +
    "Batting     : " + (data.batting_style||"—") + "\n" +
    "Bowling     : " + (data.bowling_style||"—") + "\n" +
    "Positions   : " + (positions||"—") + "\n" +
    "Availability: " + (availability||"—") + "\n\n" +
    "--- Consents ---\n" +
    "Player/Code of conduct : " + cPlayer   + "\n" +
    "Parent/Guardian        : " + cGuardian + "\n" +
    "Medical                : " + cMedical  + "\n" +
    "Photo/Media            : " + cMedia    + "\n" +
    "Signature              : " + (data.signature_base64 ? "Yes — see PDF" : "Not provided") + "\n" +
    (photoLink ? "\nPhoto: " + photoLink + "\n" : "") +
    "\nRegistered: " + now + "\n\n" +
    "See the attached PDF for the full signed registration card.";

  MailApp.sendEmail({ to: ADMIN_EMAIL, subject: subject, body: body, attachments: [pdfBlob] });
}


// ── PLAYER EMAIL ────────────────────────────────────
function sendPlayerEmail(data, regId, pdfBlob) {
  var subject = "Your cricket registration is confirmed! [#" + regId + "]";
  var body =
    "Hi " + data.first_name + ",\n\n" +
    "Thanks for registering! Your signed registration card is attached as a PDF — please keep it for your records.\n\n" +
    "Your reference number: #" + regId + "\n\n" +
    "What you agreed to:\n" +
    "  Player consent & code of conduct\n" +
    "  Parent/guardian consent\n" +
    "  Medical disclosure consent\n" +
    "  Photo & media usage consent\n\n" +
    "We'll be in touch soon about training sessions and trials.\n\n" +
    "See you on the pitch!\n" +
    "The Cricket Club Team";

  MailApp.sendEmail({ to: data.email, subject: subject, body: body, attachments: [pdfBlob] });
}
