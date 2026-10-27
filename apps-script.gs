// ======================================================================================
// ClimIntellio Form Data Handler & Auto-Formatter for Google Sheets
// ======================================================================================

const HEADERS = [
  "Timestamp",
  "Name",
  "Phone",
  "Email",
  "LinkedIn Profile",
  "Organization",
  "User Role (I am)",
  "Climate Hazards",
  "Metric Categories",
  "Selected Climate Metrics",
  "Location",
  "Description",
  "How did you reach this form"
];

// Optimal column widths in pixels for each column (A to M)
const COLUMN_WIDTHS = [
  160, // A: Timestamp
  150, // B: Name
  130, // C: Phone
  200, // D: Email
  170, // E: LinkedIn Profile
  180, // F: Organization
  160, // G: User Role
  220, // H: Climate Hazards (Multi-select)
  150, // I: Metric Categories
  280, // J: Selected Climate Metrics (Multi-select)
  200, // K: Location
  260, // L: Description
  160  // M: How did you reach
];

/**
 * Handles incoming POST requests from the web form
 */
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Ensure header row exists and is styled
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      applyHeaderStyling(sheet);
    }

    var p = (e && e.parameter) ? e.parameter : {};

    // Format multi-select hazards into clean bullet points
    var formattedHazards = formatListWithBullets(p.hazards, ", ");

    // Format multi-select categories into clean bullet points
    var formattedCategories = formatListWithBullets(p.metricCategories, ", ");

    // Format multi-select metrics into clean bullet points
    var formattedMetrics = formatListWithBullets(p.metrics, "; ");

    // Append the new row
    sheet.appendRow([
      new Date(),
      p.name || "",
      p.phone || "",
      p.email || "",
      p.linkedin || "",
      p.organization || "",
      p.userRole || "",
      formattedHazards,
      formattedCategories,
      formattedMetrics,
      p.location || "",
      p.description || "",
      p.referralSource || ""
    ]);

    var lastRow = sheet.getLastRow();

    // Style the newly appended data row so it never collapses
    styleDataRow(sheet, lastRow);

    return ContentService.createTextOutput("OK");
  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.message);
  }
}

/**
 * Helper: Converts comma or semicolon separated values into clean bulleted multi-line text
 */
function formatListWithBullets(str, delimiter) {
  if (!str) return "";
  var items = str.split(delimiter)
                 .map(function(item) { return item.trim(); })
                 .filter(function(item) { return item.length > 0; });
  
  if (items.length <= 1) return items.join("");
  return items.map(function(item) { return "• " + item; }).join("\n");
}

/**
 * Applies professional executive styling to the Header Row (Row 1)
 */
function applyHeaderStyling(sheet) {
  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  
  // ClimAgro Executive Navy Blue (#1B365D) with bold white text
  headerRange
    .setBackground("#1B365D")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.setRowHeight(1, 38);
  sheet.setFrozenRows(1);

  // Set predefined column widths
  for (var i = 0; i < COLUMN_WIDTHS.length; i++) {
    sheet.setColumnWidth(i + 1, COLUMN_WIDTHS[i]);
  }
}

/**
 * Styles a specific data row: enables text wrapping, top alignment, and clean font
 */
function styleDataRow(sheet, rowNum) {
  var rowRange = sheet.getRange(rowNum, 1, 1, HEADERS.length);
  
  // Alternating zebra row background (subtle light gray vs pure white)
  var bgColor = (rowNum % 2 === 0) ? "#FFFFFF" : "#F8FAFC";
  
  rowRange
    .setBackground(bgColor)
    .setFontFamily("Arial")
    .setFontSize(10)
    .setFontColor("#1F2937")
    .setVerticalAlignment("top")
    .setWrap(true);

  // Left-align text content, center-align timestamp & referral
  sheet.getRange(rowNum, 1).setHorizontalAlignment("center");
  sheet.getRange(rowNum, 2, 1, HEADERS.length - 1).setHorizontalAlignment("left");
}

/**
 * ONE-CLICK FIX FOR YOUR EXISTING SHEET:
 * Run this function inside Apps Script to immediately format all existing rows and columns!
 */
function formatEntireSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) return;

  // 1. Ensure header row is set and formatted
  applyHeaderStyling(sheet);

  // 2. Format all existing data rows (from row 2 down)
  if (lastRow >= 2) {
    for (var r = 2; r <= lastRow; r++) {
      // Re-format existing comma/semicolon lists in column H (Hazards) and J (Metrics)
      var hazardsCell = sheet.getRange(r, 8);
      var valH = hazardsCell.getValue().toString();
      if (valH && !valH.includes("•")) {
        hazardsCell.setValue(formatListWithBullets(valH, ","));
      }

      var categoriesCell = sheet.getRange(r, 9);
      var valC = categoriesCell.getValue().toString();
      if (valC && !valC.includes("•")) {
        categoriesCell.setValue(formatListWithBullets(valC, ","));
      }

      var metricsCell = sheet.getRange(r, 10);
      var valM = metricsCell.getValue().toString();
      if (valM && !valM.includes("•")) {
        // Handle both semicolon and comma if any
        var delim = valM.includes(";") ? ";" : ",";
        metricsCell.setValue(formatListWithBullets(valM, delim));
      }

      // Apply row styling
      styleDataRow(sheet, r);
    }
  }

  // 3. Format column A as readable timestamp
  sheet.getRange(2, 1, Math.max(1, lastRow - 1), 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
}

/**
 * Creates a convenient custom menu in Google Sheets top toolbar
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🌿 ClimIntellio")
    .addItem("✨ Format Sheet (Fix Columns & Colors)", "formatEntireSheet")
    .addToUi();
}
