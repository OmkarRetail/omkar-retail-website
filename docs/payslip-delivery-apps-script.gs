/*
 * Deploy this as a NEW Google Apps Script web app when Drive access is available.
 * Do not replace the existing website-form Apps Script.
 *
 * Required Script Property:
 * FIREBASE_WEB_API_KEY    The Firebase browser API key from config.js
 *
 * The supplied Omkar payslip Drive folder is already set below. You may later
 * override it with a PAYSLIP_FOLDER_ID Script Property, if required.
 *
 * Deploy: Execute as Me | Who has access: Anyone.
 * The Firebase ID token check below restricts use to the Omkar owner account.
 */
const PAYSLIP_OWNER_EMAIL = "omkarretailventure@gmail.com";
const DEFAULT_PAYSLIP_FOLDER_ID = "1nWu-9ljfCpHA-YBLJu_B2npLGBUkdb8T";

function doPost(event) {
  try {
    const payload = JSON.parse(event.postData.contents || "{}");
    if (!isOwner_(payload.idToken)) throw new Error("Unauthorised request.");
    const deliveryMode = String(payload.deliveryMode || "both").toLowerCase();
    if (deliveryMode === "drive-batch") return savePayslipBatch_(payload);
    if (!["drive", "email", "both"].includes(deliveryMode)) throw new Error("Invalid delivery mode.");
    if (!payload.pdfBase64 || !payload.fileName) throw new Error("Missing payslip PDF data.");
    const pdf = Utilities.newBlob(Utilities.base64Decode(payload.pdfBase64), MimeType.PDF, safeFileName_(payload.fileName));
    if (deliveryMode === "drive" || deliveryMode === "both") {
      const rootFolder = getPayslipRootFolder_();
      const monthFolder = getOrCreateMonthFolder_(rootFolder, payload.salaryMonth);
      monthFolder.createFile(pdf);
    }

    const employeeEmail = String(payload.employeeEmail || "").trim();
    if (deliveryMode === "email" || deliveryMode === "both") {
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(employeeEmail)) throw new Error("Employee email address is missing or invalid.");
      if (MailApp.getRemainingDailyQuota() < 1) throw new Error("Daily email quota has been reached.");
      MailApp.sendEmail({
        to: employeeEmail,
        subject: `Payslip | ${payload.salaryCycle || "Omkar Retail Ventures"}`,
        body: `Dear ${payload.employeeName || "Employee"},\n\nPlease find your payslip attached.\n\nRegards,\nOmkar Retail Ventures`,
        attachments: [pdf],
        name: "Omkar Retail Ventures"
      });
    }
    return response_({ ok: true });
  } catch (error) {
    console.error(error);
    return response_({ ok: false, error: error.message });
  }
}

function savePayslipBatch_(payload) {
  const payslips = Array.isArray(payload.payslips) ? payload.payslips.slice(0, 100) : [];
  if (!payslips.length) throw new Error("No payslips were provided for Drive saving.");

  const monthFolder = getOrCreateMonthFolder_(getPayslipRootFolder_(), payload.salaryMonth);
  const failed = [];
  let saved = 0;
  payslips.forEach((payslip) => {
    try {
      if (!payslip.pdfBase64 || !payslip.fileName) throw new Error("Missing PDF data.");
      const pdf = Utilities.newBlob(Utilities.base64Decode(payslip.pdfBase64), MimeType.PDF, safeFileName_(payslip.fileName));
      monthFolder.createFile(pdf);
      saved += 1;
    } catch (error) {
      failed.push(`${payslip.employeeName || payslip.fileName || "Unknown employee"}: ${error.message}`);
    }
  });

  return response_({ ok: failed.length === 0, saved, failed });
}

function getPayslipRootFolder_() {
  const configuredFolderId = PropertiesService.getScriptProperties().getProperty("PAYSLIP_FOLDER_ID");
  return DriveApp.getFolderById(configuredFolderId || DEFAULT_PAYSLIP_FOLDER_ID);
}

function getOrCreateMonthFolder_(rootFolder, salaryMonth) {
  const folderName = monthFolderName_(salaryMonth);
  const existingFolders = rootFolder.getFoldersByName(folderName);
  return existingFolders.hasNext() ? existingFolders.next() : rootFolder.createFolder(folderName);
}

function monthFolderName_(salaryMonth) {
  const match = String(salaryMonth || "").match(/^(\d{4})-(\d{2})$/);
  if (!match) return "Unsorted Payslips";
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return `${months[Number(match[2]) - 1]} ${match[1]}`;
}

function isOwner_(idToken) {
  if (!idToken) return false;
  const apiKey = PropertiesService.getScriptProperties().getProperty("FIREBASE_WEB_API_KEY");
  if (!apiKey) throw new Error("FIREBASE_WEB_API_KEY has not been configured.");
  const request = UrlFetchApp.fetch(`https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=${encodeURIComponent(apiKey)}`, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ idToken }),
    muteHttpExceptions: true
  });
  if (request.getResponseCode() !== 200) return false;
  const account = JSON.parse(request.getContentText()).users?.[0] || {};
  return String(account.email || "").toLowerCase() === PAYSLIP_OWNER_EMAIL;
}

function safeFileName_(name) {
  return String(name).replace(/[^A-Za-z0-9._-]/g, "_").slice(0, 150) || "Omkar-Payslip.pdf";
}

function response_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
