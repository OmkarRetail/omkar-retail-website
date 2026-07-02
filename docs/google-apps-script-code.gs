const SPREADSHEET_ID = "135AltlhB-NsS5QmuaaRCHYko_RvW-UJGhrxi5fVM-2w";
const RESUME_FOLDER_ID = "1iXpaTpB87tNlQC1eBOQm_n39nvCK-6TZ";
const APPLICATIONS_SHEET_NAME = "Applications";
const SUBMIT_TOKEN = "63eda119fbe540b6b88301cdcafad154";

function doGet() {
  return jsonResponse_({
    status: "ok",
    message: "OMKAR application endpoint is running"
  });
}

function doPost(e) {
  try {
    const rawPayload = e.parameter.payload || "{}";
    const data = JSON.parse(rawPayload);

    if (data.formType !== "candidateApplication") {
      throw new Error("Unsupported form type");
    }

    if (data.submitToken !== SUBMIT_TOKEN) {
      throw new Error("Unauthorized submit request");
    }

    const resume = data.resume || {};
    let resumeFileName = data.resumeFileName || resume.fileName || "";
    let resumeUrl = data.resumeUrl || "";

    if (resume.base64 && resume.fileName) {
      const folder = DriveApp.getFolderById(RESUME_FOLDER_ID);
      const fileName = buildResumeFileName_(data.fullName, resume.fileName);
      const blob = Utilities.newBlob(
        Utilities.base64Decode(resume.base64),
        resume.mimeType || "application/octet-stream",
        fileName
      );
      const file = folder.createFile(blob);
      resumeFileName = file.getName();
      resumeUrl = file.getUrl();
    }

    const sheet = getApplicationsSheet_();
    ensureApplicationHeaders_(sheet);
    appendApplicationRow_(sheet, {
      receivedAt: new Date(),
      fullName: data.fullName || "",
      mobile: data.mobile || "",
      email: data.email || "",
      age: data.age || "",
      gender: data.gender || "",
      education: data.education || "",
      experience: data.experience || "",
      role: data.role || "",
      currentLocation: data.currentLocation || "",
      aadhaar: data.aadhaar || "",
      pan: data.pan || "",
      joinAvailability: data.joinAvailability || "",
      resumeFileName,
      resumeUrl,
      source: data.source || "",
      submittedAt: data.submittedAt || ""
    });

    return jsonResponse_({
      status: "success",
      resumeFileName,
      resumeUrl
    });
  } catch (error) {
    return jsonResponse_({
      status: "error",
      message: error && error.message ? error.message : String(error)
    });
  }
}

function getApplicationsSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(APPLICATIONS_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(APPLICATIONS_SHEET_NAME);
  }
  return sheet;
}

function ensureApplicationHeaders_(sheet) {
  const headers = [
    "Received At",
    "Full Name",
    "Mobile",
    "Email",
    "Age",
    "Gender",
    "Education",
    "Experience",
    "Role Applying For",
    "Current Location",
    "Aadhaar",
    "PAN",
    "Availability to Join",
    "Resume File Name",
    "Resume URL",
    "Source",
    "Submitted At"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach((header) => {
    if (!existingHeaders.includes(header)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
    }
  });
}

function appendApplicationRow_(sheet, values) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map((header) => {
    switch (header) {
      case "Received At": return values.receivedAt;
      case "Full Name": return values.fullName;
      case "Mobile": return values.mobile;
      case "Email": return values.email;
      case "Age": return values.age;
      case "Gender": return values.gender;
      case "Education": return values.education;
      case "Experience": return values.experience;
      case "Role Applying For": return values.role;
      case "Current Location": return values.currentLocation;
      case "Aadhaar": return values.aadhaar;
      case "PAN": return values.pan;
      case "Availability to Join": return values.joinAvailability;
      case "Resume File Name": return values.resumeFileName;
      case "Resume URL": return values.resumeUrl;
      case "Source": return values.source;
      case "Submitted At": return values.submittedAt;
      default: return "";
    }
  });
  sheet.appendRow(row);
}

function buildResumeFileName_(fullName, originalName) {
  const namePart = sanitizeFileName_(fullName || "candidate");
  const originalPart = sanitizeFileName_(originalName || "resume");
  return `${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss")}-${namePart}-${originalPart}`;
}

function sanitizeFileName_(value) {
  return String(value)
    .trim()
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 80);
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
