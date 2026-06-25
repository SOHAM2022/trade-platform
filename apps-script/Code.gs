/**
 * Real Amount — Partner & Customer Onboarding
 * Google Apps Script Web App
 *
 * HOW TO DEPLOY:
 * 1. Go to script.google.com → New Project
 * 2. Paste this entire file into the editor
 * 3. Click "Deploy" → "New deployment"
 * 4. Type: Web App
 *    Execute as: Me
 *    Who has access: Anyone
 * 5. Click Deploy → copy the Web App URL
 * 6. Paste that URL into form.js → APPS_SCRIPT_URL
 *
 * NOTE: After updating the script, you MUST create a NEW deployment
 * (not re-deploy the existing one) for changes to take effect.
 */

// ── Config ──────────────────────────────────────────
// Paste your Google Sheet ID here.
// Get it from the Sheet URL:
// https://docs.google.com/spreadsheets/d/SHEET_ID_IS_HERE/edit

const SPREADSHEET_ID = "1bIO2okfM-wZNqWJtmQia2MyYv0BRPik6bLhESBC_DJI";

const PARTNER_SHEET_NAME = "Partner Submissions";
const CUSTOMER_SHEET_NAME = "Customer Submissions";
const FOLDER_NAME = "Real Amount — Uploads";

// ── Column headers ──────────────────────────────────
const PARTNER_HEADERS = [
  "Timestamp",
  "Reference ID",
  "Email",
  "Business Name",
  "Workplace Address",
  "GST Number",
  "Contact Person",
  "Mobile Number",
  "Business Type",
  "Products / Services",
  "Brands Dealt In",
  "Visiting Card URL",
];

const CUSTOMER_HEADERS = [
  "Timestamp",
  "Reference ID",
  "Account Type",
  "Customer Name",
  "Mobile Number",
  "Address",
  "Service Requested",
  "Items / Category",
  "Brands (if any)",
  "Payment Screenshot URL",
];

// ════════════════════════════════════════════════════
// POST handler — called when the form submits
// ════════════════════════════════════════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const accountType = (data.accountType || "").trim();

    if (accountType === "Customer") {
      return handleCustomerSubmission(data);
    } else {
      return handlePartnerSubmission(data);
    }
  } catch (err) {
    console.error("doPost error:", err.message);
    return jsonResponse({ success: false, error: err.message });
  }
}

// ════════════════════════════════════════════════════
// PARTNER submission handler
// ════════════════════════════════════════════════════
function handlePartnerSubmission(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PARTNER_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(PARTNER_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    appendHeaderRow(sheet, PARTNER_HEADERS);
  }

  // ── Handle file upload ──────────────────────────
  let fileUrl = "";
  if (data.fileData && data.fileName) {
    fileUrl = uploadFileToDrive(data, "partner_" + (data.referenceId || ""));
  }

  const referenceId = data.referenceId || getNextReferenceId(sheet);

  const row = [
    new Date(),
    referenceId,
    data.email || "",
    data.businessName || "",
    data.workplaceAddress || "",
    data.gstNumber || "",
    data.contactPerson || "",
    data.mobileNumber || "",
    data.businessType || "",
    Array.isArray(data.selectedItems)
      ? data.selectedItems.join(", ")
      : data.selectedItems || "",
    data.brands || "",
    fileUrl,
  ];

  sheet.appendRow(row);
  sheet.autoResizeColumns(1, PARTNER_HEADERS.length);

  return jsonResponse({ success: true, referenceId: referenceId });
}

// ════════════════════════════════════════════════════
// CUSTOMER submission handler
// ════════════════════════════════════════════════════
function handleCustomerSubmission(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CUSTOMER_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    appendHeaderRow(sheet, CUSTOMER_HEADERS);
  }

  // ── Handle payment screenshot upload ───────────
  let fileUrl = "";
  if (data.fileData && data.fileName) {
    fileUrl = uploadFileToDrive(data, "payment_" + (data.referenceId || ""));
  }

  const referenceId = data.referenceId || getNextReferenceId(sheet);

  const row = [
    new Date(),
    referenceId,
    "Customer",
    data.customerName || "",
    data.customerMobile || "",
    data.customerAddress || "",
    data.customerService || "",
    Array.isArray(data.selectedItems)
      ? data.selectedItems.join(", ")
      : data.selectedItems || "",
    data.brands || "",
    fileUrl,
  ];

  sheet.appendRow(row);
  sheet.autoResizeColumns(1, CUSTOMER_HEADERS.length);

  return jsonResponse({ success: true, referenceId: referenceId });
}

// ════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════

function appendHeaderRow(sheet, headers) {
  sheet.appendRow(headers);
  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1a3c5e")
    .setFontColor("#ffffff");
  sheet.setFrozenRows(1);
}

function uploadFileToDrive(data, prefix) {
  try {
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folders.hasNext()
      ? folders.next()
      : DriveApp.createFolder(FOLDER_NAME);

    const decoded = Utilities.base64Decode(data.fileData);
    const mimeType = data.fileType || "application/octet-stream";
    const safeName = (prefix + "_" + data.fileName).replace(
      /[^a-zA-Z0-9._\-]/g,
      "_",
    );
    const blob = Utilities.newBlob(decoded, mimeType, safeName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    console.error("File upload error:", err.message);
    return "";
  }
}

function getNextReferenceId(sheet) {
  const START = 11110;
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return "REF" + START;

  const lastId = sheet.getRange(lastRow, 2).getValue().toString().trim();

  if (lastId.startsWith("REF")) {
    const num = parseInt(lastId.replace("REF", ""), 10);
    if (!isNaN(num)) return "REF" + (num + 1);
  }

  return "REF" + (START + lastRow - 1);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

// ── Health check (GET) ───────────────────────────────
function doGet() {
  return ContentService.createTextOutput(
    "Real Amount Form API is running. Partner + Customer flows active.",
  ).setMimeType(ContentService.MimeType.TEXT);
}

// ════════════════════════════════════════════════════
// TEST FUNCTIONS — run from Apps Script editor
// ════════════════════════════════════════════════════

/** Test a Partner submission */
function testPartnerSubmission() {
  const fakeRequest = {
    postData: {
      contents: JSON.stringify({
        accountType: "Partner",
        email: "partner@example.com",
        businessName: "Test Electronics",
        workplaceAddress: "123 Market Street, Ludhiana",
        gstNumber: "03ABCDE1234F1Z5",
        contactPerson: "Rajesh Kumar",
        mobileNumber: "9876543210",
        businessType: "Product Sale (Retail)",
        selectedItems: ["Air Conditioner", "LED Tv", "Refrigerator"],
        brands: "Samsung, LG",
        referenceId: "REF11115",
      }),
    },
  };

  const result = doPost(fakeRequest);
  Logger.log("Partner Response: " + result.getContent());
}

/** Test a Customer submission */
function testCustomerSubmission() {
  const fakeRequest = {
    postData: {
      contents: JSON.stringify({
        accountType: "Customer",
        customerName: "Priya Sharma",
        customerMobile: "9988776655",
        customerAddress: "45 Rose Garden, Jalandhar",
        customerService: "Service/Repair",
        selectedItems: ["Ac Service/Repair", "Mobile Repair"],
        brands: "",
        referenceId: "REF11116",
      }),
    },
  };

  const result = doPost(fakeRequest);
  Logger.log("Customer Response: " + result.getContent());
}
