/**
 * =========================================================================
 * Google Apps Script: Portfolio Website Contact Form to Google Sheet
 * (No email notifications - saves inquiries directly to Sheet only)
 * =========================================================================
 */

// If this script is standalone (not created via Extensions > Apps Script inside your sheet),
// paste your Google Sheet ID below (the characters between /d/ and /edit in your sheet URL).
// If opened directly from inside the Google Sheet, you can leave it empty ("").
var SHEET_ID = ""; 

/**
 * Optional: Run this function once directly from the Apps Script editor to 
 * automatically format and style all your sheet headings & dropdowns!
 */
function setupSheetHeaders() {
  var sheet;
  if (SHEET_ID && SHEET_ID.trim() !== "") {
    sheet = SpreadsheetApp.openById(SHEET_ID.trim()).getActiveSheet();
  } else {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }

  // Define the column headers
  var headers = [
    "Timestamp",
    "Full Name",
    "Email Address",
    "Service Interest",
    "Project Details / Message",
    "Lead Status",
    "Follow-up Date",
    "Action Notes / Strategy"
  ];

  // Set the headers in Row 1
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Styling matching your portfolio theme
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#10161d");       // Dark slate matching portfolio
  headerRange.setFontColor("#6c97c6");        // Primary blue accent
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  // Set column widths for optimal readability
  sheet.setColumnWidth(1, 160); // Timestamp
  sheet.setColumnWidth(2, 180); // Full Name
  sheet.setColumnWidth(3, 230); // Email Address
  sheet.setColumnWidth(4, 210); // Service Interest
  sheet.setColumnWidth(5, 340); // Project Details / Message
  sheet.setColumnWidth(6, 150); // Lead Status
  sheet.setColumnWidth(7, 130); // Follow-up Date
  sheet.setColumnWidth(8, 250); // Action Notes

  // Add interactive dropdown data validation for 'Lead Status' (Column F)
  var statusOptions = [
    "New Lead",
    "Contacted",
    "Meeting Scheduled",
    "Proposal Sent",
    "Closed Won",
    "Closed Lost"
  ];
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("F2:F1000").setDataValidation(statusRule);

  SpreadsheetApp.flush();
  Logger.log("✓ Sheet headings and CRM workflow columns successfully set up!");
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000); // Prevent concurrent write race conditions

  try {
    var sheet;
    if (SHEET_ID && SHEET_ID.trim() !== "") {
      sheet = SpreadsheetApp.openById(SHEET_ID.trim()).getActiveSheet();
    } else {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    }

    // If the sheet is completely empty, auto-create headers
    if (sheet.getLastRow() === 0) {
      setupSheetHeaders();
    }

    // Extract incoming parameters from FormData or JSON
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var timestamp = new Date();
    var name = data.name || data.fullName || "";
    var email = data.email || "";
    var serviceInterest = data.service_interest || data.service || "";
    var message = data.message || data.projectDetails || "";
    var defaultStatus = "New Lead";

    // Append inquiry row to Google Sheet (No email sent)
    sheet.appendRow([
      timestamp,
      name,
      email,
      serviceInterest,
      message,
      defaultStatus,
      "", // Follow-up Date
      ""  // Action Notes
    ]);

    // Format timestamp column
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sheet.getRange(lastRow, 1, 1, 8).setVerticalAlignment("middle");

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", message: "Enquiry recorded successfully!" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ 
      status: "active", 
      message: "Abikumar Portfolio Contact Form API is running successfully." 
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
