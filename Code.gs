/**
 * Department of Defence PNG external recruitment receiver.
 * Deploy this file as a Google Apps Script web app.
 *
 * Required deployment setting:
 *   Execute as: Me
 *   Who has access: Anyone
 */

const CONFIG = Object.freeze({
  spreadsheetName: 'DOD External Recruitment Applications',
  sheetName: 'Applications',
  uploadFolderName: 'DOD External Recruitment Uploads',
  senderName: 'Department of Defence Recruitment',
  maxPdfBytes: 8 * 1024 * 1024,
  maxPhotoBytes: 2 * 1024 * 1024,
  requiredFileKeys: ['coverLetter', 'qualificationCertificate', 'cvUpload', 'passportPhoto']
});

const HEADERS = Object.freeze([
  'Submission ID', 'Submitted At', 'Status', 'Division / Branch', 'Position Title',
  'Reference Number', 'Position Code', 'Cost Account', 'Class', 'First Name',
  'Middle Name', 'Last Name', 'Gender', 'Date of Birth', 'Citizenship', 'Phone',
  'Email', 'Address', 'Province of Residence', 'Employment Status',
  'Highest Qualification', 'Field of Study', 'Year Qualification Obtained',
  'Institution', 'Years Work Experience', 'Years Relevant Experience',
  'Training Attended', 'Training Year', 'Number of References',
  'Qualification Matches Position', 'EOI / Cover Letter', 'Training Certificates',
  'Qualification Certificates', 'CV', 'Passport Photo', 'Accuracy Declaration',
  'Privacy Consent', 'Acknowledgement Email'
]);

function doGet() {
  return jsonResponse_({
    ok: true,
    service: 'DOD External Recruitment Application Receiver',
    message: 'The recruitment receiver is online.'
  });
}

function doPost(e) {
  let lock;

  try {
    const data = parseRequest_(e);
    validateApplication_(data);

    lock = LockService.getScriptLock();
    lock.waitLock(30000);

    const resources = getOrCreateResources_();
    const sheet = resources.sheet;
    const duplicateRow = findSubmissionRow_(sheet, data.submissionId);

    if (duplicateRow) {
      return jsonResponse_({
        ok: true,
        duplicate: true,
        submissionId: data.submissionId,
        message: 'This application was already received.'
      });
    }

    const applicationFolder = createApplicationFolder_(resources.folder, data);
    const fileUrls = saveUploads_(applicationFolder, data);
    const row = buildApplicationRow_(data, fileUrls);

    sheet.appendRow(row);
    const rowNumber = sheet.getLastRow();
    SpreadsheetApp.flush();

    let emailStatus = 'Not sent';
    try {
      sendAcknowledgement_(data);
      sendRecruitmentNotification_(data, resources.spreadsheet.getUrl(), applicationFolder.getUrl());
      emailStatus = 'Sent';
    } catch (mailError) {
      emailStatus = 'Failed: ' + safeErrorMessage_(mailError);
      console.error(mailError);
    }

    sheet.getRange(rowNumber, HEADERS.indexOf('Acknowledgement Email') + 1).setValue(emailStatus);

    return jsonResponse_({
      ok: true,
      submissionId: data.submissionId,
      message: 'Application received successfully.'
    });
  } catch (error) {
    console.error(error);
    return jsonResponse_({
      ok: false,
      message: safeErrorMessage_(error)
    });
  } finally {
    if (lock && lock.hasLock()) lock.releaseLock();
  }
}

/**
 * Run once from the Apps Script editor before deploying.
 * Optionally provide the recruitment mailbox for internal notifications:
 * setupRecruitmentBackend('recruitment@example.gov.pg')
 */
function setupRecruitmentBackend(recruitmentEmail) {
  if (recruitmentEmail) {
    if (!isValidEmail_(recruitmentEmail)) throw new Error('Enter a valid recruitment email address.');
    PropertiesService.getScriptProperties().setProperty('RECRUITMENT_EMAIL', recruitmentEmail.trim());
  }

  const resources = getOrCreateResources_();
  const result = {
    spreadsheetId: resources.spreadsheet.getId(),
    spreadsheetUrl: resources.spreadsheet.getUrl(),
    uploadFolderId: resources.folder.getId(),
    uploadFolderUrl: resources.folder.getUrl(),
    recruitmentEmail: PropertiesService.getScriptProperties().getProperty('RECRUITMENT_EMAIL') || ''
  };

  console.log(JSON.stringify(result, null, 2));
  return result;
}

function parseRequest_(e) {
  if (!e || !e.postData || !e.postData.contents) throw new Error('The request body is empty.');

  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error('The application data is not valid JSON.');
  }

  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error('The application body is invalid.');
  }

  return data;
}

function validateApplication_(data) {
  const requiredText = {
    submissionId: 'Submission ID',
    division: 'Division',
    positionTitle: 'Position title',
    referenceNumber: 'Reference number',
    firstName: 'First name',
    lastName: 'Last name',
    gender: 'Gender',
    dateOfBirth: 'Date of birth',
    citizenship: 'Citizenship',
    phone: 'Phone number',
    email: 'Email',
    address: 'Address',
    provinceOfResidence: 'Province of residence',
    employmentStatus: 'Employment status',
    highestQualification: 'Highest qualification',
    fieldOfStudy: 'Field of study',
    qualificationObtained: 'Year qualification obtained',
    institution: 'Institution',
    workExperience: 'Work experience',
    relevantExperience: 'Relevant experience',
    references: 'Number of references',
    qualificationMatch: 'Qualification match'
  };

  Object.keys(requiredText).forEach(function (key) {
    if (data[key] === undefined || data[key] === null || String(data[key]).trim() === '') {
      throw new Error(requiredText[key] + ' is required.');
    }
  });

  if (!/^DOD-\d{10,}$/.test(String(data.submissionId))) throw new Error('The submission ID is invalid.');
  if (!isValidEmail_(data.email)) throw new Error('The email address is invalid.');
  if (data.accuracyDeclaration !== true || data.privacyConsent !== true) {
    throw new Error('The applicant declaration and privacy consent are required.');
  }

  CONFIG.requiredFileKeys.forEach(function (key) {
    if (!data[key]) throw new Error('A required application document is missing: ' + key + '.');
  });

  validateUpload_(data.coverLetter, 'EOI / cover letter', ['pdf'], CONFIG.maxPdfBytes);
  validateUpload_(data.trainingCertificate, 'Training certificate', ['pdf'], CONFIG.maxPdfBytes, true);
  validateUpload_(data.qualificationCertificate, 'Qualification certificate', ['pdf'], CONFIG.maxPdfBytes);
  validateUpload_(data.cvUpload, 'CV', ['pdf'], CONFIG.maxPdfBytes);
  validateUpload_(data.passportPhoto, 'Passport photo', ['jpg', 'jpeg', 'png'], CONFIG.maxPhotoBytes);
}

function validateUpload_(file, label, allowedExtensions, maxBytes, optional) {
  if (!file) {
    if (optional) return;
    throw new Error(label + ' is required.');
  }

  const fileName = String(file.fileName || '');
  const extension = fileName.indexOf('.') >= 0 ? fileName.split('.').pop().toLowerCase() : '';
  if (!allowedExtensions.includes(extension)) throw new Error(label + ' has an invalid file type.');
  if (!file.base64 || typeof file.base64 !== 'string') throw new Error(label + ' contains no file data.');

  const estimatedBytes = Math.floor(file.base64.length * 0.75);
  if (estimatedBytes > maxBytes) {
    throw new Error(label + ' exceeds the ' + Math.round(maxBytes / 1024 / 1024) + ' MB limit.');
  }
}

function getOrCreateResources_() {
  const properties = PropertiesService.getScriptProperties();
  let spreadsheet = getSpreadsheetByProperty_(properties.getProperty('SPREADSHEET_ID'));

  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create(CONFIG.spreadsheetName);
    properties.setProperty('SPREADSHEET_ID', spreadsheet.getId());
  }

  let sheet = spreadsheet.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    const sheets = spreadsheet.getSheets();
    sheet = sheets.length === 1 && sheets[0].getLastRow() === 0
      ? sheets[0].setName(CONFIG.sheetName)
      : spreadsheet.insertSheet(CONFIG.sheetName);
  }

  prepareSheet_(sheet);

  let folder = getFolderByProperty_(properties.getProperty('UPLOAD_FOLDER_ID'));
  if (!folder) {
    folder = DriveApp.createFolder(CONFIG.uploadFolderName);
    properties.setProperty('UPLOAD_FOLDER_ID', folder.getId());
  }

  return { spreadsheet: spreadsheet, sheet: sheet, folder: folder };
}

function getSpreadsheetByProperty_(id) {
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (error) { return null; }
}

function getFolderByProperty_(id) {
  if (!id) return null;
  try { return DriveApp.getFolderById(id); } catch (error) { return null; }
}

function prepareSheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground('#0f5132')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.autoResizeColumns(1, HEADERS.length);
  }
}

function findSubmissionRow_(sheet, submissionId) {
  if (sheet.getLastRow() < 2) return 0;
  const match = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(submissionId))
    .matchEntireCell(true)
    .findNext();
  return match ? match.getRow() : 0;
}

function createApplicationFolder_(rootFolder, data) {
  const folderName = [data.submissionId, data.referenceNumber, data.lastName + '_' + data.firstName]
    .map(sanitizeFilePart_)
    .filter(Boolean)
    .join(' - ');
  return rootFolder.createFolder(folderName.substring(0, 180));
}

function saveUploads_(folder, data) {
  return {
    coverLetter: saveUpload_(folder, data.coverLetter, '01_EOI'),
    trainingCertificate: saveUpload_(folder, data.trainingCertificate, '02_Training'),
    qualificationCertificate: saveUpload_(folder, data.qualificationCertificate, '03_Qualifications'),
    cvUpload: saveUpload_(folder, data.cvUpload, '04_CV'),
    passportPhoto: saveUpload_(folder, data.passportPhoto, '05_Photo')
  };
}

function saveUpload_(folder, file, prefix) {
  if (!file) return '';
  const bytes = Utilities.base64Decode(file.base64);
  const originalName = sanitizeFileName_(file.fileName || prefix);
  const mimeType = String(file.mimeType || 'application/octet-stream');
  const blob = Utilities.newBlob(bytes, mimeType, prefix + '_' + originalName);
  return folder.createFile(blob).getUrl();
}

function buildApplicationRow_(data, files) {
  return [
    data.submissionId, new Date(data.submittedAt || Date.now()), 'Received', data.division,
    data.positionTitle, data.referenceNumber, data.positionCode, data.costAccount,
    data.positionClass, data.firstName, data.middleName, data.lastName, data.gender,
    data.dateOfBirth, data.citizenship, data.phone, data.email, data.address,
    data.provinceOfResidence, data.employmentStatus, data.highestQualification,
    data.fieldOfStudy, data.qualificationObtained, data.institution, data.workExperience,
    data.relevantExperience, data.trainingAttended, data.trainingYear, data.references,
    data.qualificationMatch, files.coverLetter, files.trainingCertificate,
    files.qualificationCertificate, files.cvUpload, files.passportPhoto,
    data.accuracyDeclaration === true ? 'Yes' : 'No',
    data.privacyConsent === true ? 'Yes' : 'No', 'Pending'
  ].map(safeCell_);
}

function sendAcknowledgement_(data) {
  const subject = 'Application received – ' + data.referenceNumber + ' – ' + data.submissionId;
  const recruitmentEmail = PropertiesService.getScriptProperties().getProperty('RECRUITMENT_EMAIL') || '';
  const options = {
    to: String(data.email).trim(),
    subject: subject,
    name: CONFIG.senderName,
    htmlBody:
      '<p>Dear ' + escapeHtml_(data.firstName) + ' ' + escapeHtml_(data.lastName) + ',</p>' +
      '<p>Your application has been received by the Department of Defence recruitment team.</p>' +
      '<p><strong>Application reference:</strong> ' + escapeHtml_(data.submissionId) + '<br>' +
      '<strong>Position:</strong> ' + escapeHtml_(data.positionTitle) + '<br>' +
      '<strong>Position reference:</strong> ' + escapeHtml_(data.referenceNumber) + '<br>' +
      '<strong>Division:</strong> ' + escapeHtml_(data.division) + '</p>' +
      '<p>Please retain this email for your records. Only shortlisted applicants may be contacted.</p>' +
      '<p>Department of Defence<br>Papua New Guinea</p>'
  };

  if (recruitmentEmail) options.replyTo = recruitmentEmail;
  MailApp.sendEmail(options);
}

function sendRecruitmentNotification_(data, spreadsheetUrl, folderUrl) {
  const recruitmentEmail = PropertiesService.getScriptProperties().getProperty('RECRUITMENT_EMAIL') || '';
  if (!recruitmentEmail) return;

  MailApp.sendEmail({
    to: recruitmentEmail,
    subject: 'New application – ' + data.referenceNumber + ' – ' + data.lastName,
    name: CONFIG.senderName,
    htmlBody:
      '<p>A new external recruitment application has been received.</p>' +
      '<p><strong>Applicant:</strong> ' + escapeHtml_(data.firstName + ' ' + data.lastName) + '<br>' +
      '<strong>Position:</strong> ' + escapeHtml_(data.positionTitle) + '<br>' +
      '<strong>Reference:</strong> ' + escapeHtml_(data.referenceNumber) + '<br>' +
      '<strong>Application ID:</strong> ' + escapeHtml_(data.submissionId) + '</p>' +
      '<p><a href="' + escapeHtml_(spreadsheetUrl) + '">Open applications spreadsheet</a><br>' +
      '<a href="' + escapeHtml_(folderUrl) + '">Open applicant documents</a></p>'
  });
}

function safeCell_(value) {
  if (value instanceof Date) return value;
  const text = value === undefined || value === null ? '' : String(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function sanitizeFileName_(value) {
  return String(value || 'file')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
    .replace(/\s+/g, '_')
    .substring(0, 140);
}

function sanitizeFilePart_(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
}

function isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function escapeHtml_(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function safeErrorMessage_(error) {
  const message = error && error.message ? String(error.message) : 'The application could not be processed.';
  return message.substring(0, 300);
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
