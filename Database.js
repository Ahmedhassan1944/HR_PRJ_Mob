/**
 * =========================================================
 * DATABASE MODULE (Database.gs)
 * =========================================================
 * All Google Sheets read/write operations happen here.
 * The Spreadsheet must have these sheets (tabs):
 *  - tbl_Candidates
 *  - tbl_Documents
 *  - tbl_Users
 *  - tbl_SystemLogs
 *
 * To link: Open your Google Sheet, copy its ID from the URL
 * (the long string between /d/ and /edit), and paste below.
 * =========================================================
 */

// ── Read IDs from Script Properties (Project Settings → Script Properties) ──
// Required keys: SPREADSHEET_ID
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

// Sheet name constants
const SHEET_CANDIDATES = 'tbl_Candidates';
const SHEET_DOCUMENTS = 'tbl_Documents';
const SHEET_USERS = 'tbl_Users';
const SHEET_LOGS = 'tbl_SystemLogs';
const SHEET_EVENTS = 'tbl_Events';

/**
 * Internal: Returns the opened Spreadsheet, cached for this execution.
 * SpreadsheetApp.openById() is expensive — calling it once per execution
 * and reusing the object cuts API overhead for every multi-step operation.
 * The cache is a module-level variable: it resets to null on every fresh
 * GAS execution (i.e., every new api_* call from the frontend), so there
 * is no risk of serving stale data across requests.
 */
let _ssCache = null;
function getSpreadsheet_() {
  if (!_ssCache) _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ssCache;
}

/**
 * Helper: Returns a specific sheet by name.
 * Uses the cached Spreadsheet object — never opens by ID more than once
 * per execution regardless of how many sheets are accessed.
 */
function getSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  return sheet;
}

/**
 * Helper: Generates a unique UUID for primary keys.
 */
function generateUUID_() {
  return Utilities.getUuid();
}

// ─────────────────────────────────────────────
// AUTHORIZATION HELPERS
// ─────────────────────────────────────────────

/**
 * Looks up the current user's role from tbl_Users.
 * Returns the role string ('Admin', 'HR', 'Coordinator', 'Viewer')
 * or null if the user is not registered.
 */
function getCurrentUserRole_() {
  try {
    const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
    if (!email) return null;
    const sheet   = getSheet_(SHEET_USERS);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const roleCol  = headers.indexOf('Role');
    for (let i = 1; i < data.length; i++) {
      if ((data[i][emailCol] || '').toString().toLowerCase() === email) {
        return data[i][roleCol] || null;
      }
    }
    return null;
  } catch (e) {
    Logger.log('getCurrentUserRole_ error: ' + e.message);
    return null;
  }
}

/**
 * Checks that the current user holds one of the allowed roles.
 * @param {string[]} allowedRoles - e.g. ['Admin', 'HR']
 * @returns {{ authorized: boolean, role?: string, error?: string }}
 */
function requireRole_(allowedRoles) {
  const role = getCurrentUserRole_();
  if (!role) {
    return { authorized: false,
      error: 'Access denied: your account (' +
             (Session.getActiveUser().getEmail() || 'unknown') +
             ') is not registered in the HR system.' };
  }
  if (!allowedRoles.includes(role)) {
    return { authorized: false,
      error: 'Access denied: your role (' + role + ') does not have permission for this action.' };
  }
  return { authorized: true, role };
}

// ─────────────────────────────────────────────
// CANDIDATES
// ─────────────────────────────────────────────

/**
 * Returns all candidate rows as an array of objects.
 * Called from the frontend via: google.script.run.api_getAllCandidates()
 */
function api_getAllCandidates() {
  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const candidates = rows.map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return { success: true, data: candidates };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Creates a new candidate record in the Candidates sheet.
 * @param {object} candidateData - { fullName, position, department, email, phone, nationality, salary, coordinatorEmail }
 */
function api_createCandidate(candidateData) {
  const auth = requireRole_(['Admin', 'HR']);
  if (!auth.authorized) return { success: false, error: auth.error };
  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const id = generateUUID_();
    const now = new Date().toISOString();

    sheet.appendRow([
      id,                                   // CandidateID
      candidateData.fullName,               // FullName
      candidateData.position,               // Position
      candidateData.department,             // Department
      candidateData.email,                  // Email
      candidateData.phone,                  // Phone
      candidateData.nationality,            // Nationality
      candidateData.salary,                 // OfferSalary
      candidateData.coordinatorEmail,       // AssignedCoordinatorEmail
      'Documents Requested',                // CurrentStatus
      now,                                  // CreatedAt
      now,                                  // UpdatedAt
      ''                                    // DriveFolderID (filled by DriveManager)
    ]);

    api_writeLog_(id, 'SYSTEM', 'Candidate Created: ' + candidateData.fullName);
    // [CACHE POLICY] Write operation — invalidate dashboard cache immediately
    CacheService.getScriptCache().remove('dashboard_data');
    return { success: true, candidateId: id };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Updates a candidate's profile details.
 * @param {string} candidateId
 * @param {object} updates - { phone, notes }
 */
function api_updateCandidateDetails(candidateId, updates) {
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };
  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('CandidateID');
    const phoneCol = headers.indexOf('Phone');
    const notesCol = headers.indexOf('Notes'); // Column N
    const updatedCol = headers.indexOf('UpdatedAt');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === candidateId) {
        if (updates.phone !== undefined && phoneCol >= 0) {
          sheet.getRange(i + 1, phoneCol + 1).setValue(updates.phone);
        }
        if (updates.notes !== undefined) {
          if (notesCol === -1) {
            // Notes column missing — skip silently and log; do not mutate schema at runtime
            Logger.log('WARNING: Notes column not found in tbl_Candidates. Add it manually.');
          } else {
            sheet.getRange(i + 1, notesCol + 1).setValue(updates.notes);
          }
        }
        sheet.getRange(i + 1, updatedCol + 1).setValue(new Date().toISOString());
        api_writeLog_(candidateId, Session.getActiveUser().getEmail(), 'Profile Updated');
        // [CACHE POLICY] Write operation — invalidate dashboard cache immediately
        CacheService.getScriptCache().remove('dashboard_data');
        return { success: true };
      }
    }
    return { success: false, error: 'Candidate not found.' };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Updates a candidate's status.
 * @param {string} candidateId
 * @param {string} newStatus
 */
function api_updateCandidateStatus(candidateId, newStatus) {
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };
  const ALLOWED_STATUSES = new Set([
    'New Candidate',
    'Documents Requested',
    'Documents Under Preparing',
    'Pending Passport',
    'Pending Photo',
    'Pending Academic Certificate',
    'Pending Medical',
    'Booked a medical examination',
    'Documents Complete',
    'Visa Pending',
    'Visa Completed',
    'Mobilized',
    'Closed'
  ]);

  if (!ALLOWED_STATUSES.has(newStatus)) {
    return { success: false, error: 'Invalid status value: ' + newStatus };
  }

  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('CandidateID');
    const statusCol = headers.indexOf('CurrentStatus');
    const updatedCol = headers.indexOf('UpdatedAt');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === candidateId) {
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        sheet.getRange(i + 1, updatedCol + 1).setValue(new Date().toISOString());
        api_writeLog_(candidateId, Session.getActiveUser().getEmail(), 'Status Changed: ' + newStatus);
        // [CACHE POLICY] Write operation — invalidate dashboard cache immediately
        CacheService.getScriptCache().remove('dashboard_data');
        return { success: true };
      }
    }
    return { success: false, error: 'Candidate not found.' };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

// ─────────────────────────────────────────────
// DOCUMENTS
// ─────────────────────────────────────────────

/**
 * Gets all documents for a specific candidate.
 * @param {string} candidateId
 */
function api_getDocumentsByCandidate(candidateId) {
  try {
    const sheet = getSheet_(SHEET_DOCUMENTS);
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const documents = rows
      .filter(row => row[headers.indexOf('CandidateID')] === candidateId)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
    return { success: true, data: documents };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Returns ALL document records across all candidates.
 * Used by api_getDashboardData for doc-level KPI aggregation.
 */
function api_getAllDocuments() {
  try {
    const sheet = getSheet_(SHEET_DOCUMENTS);
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const documents = rows.map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return { success: true, data: documents };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

// ─────────────────────────────────────────────
// AUDIT LOG
// ─────────────────────────────────────────────

/**
 * Internal-only: Writes an immutable audit log entry.
 */
function api_writeLog_(candidateId, actor, event) {
  try {
    const sheet = getSheet_(SHEET_LOGS);
    sheet.appendRow([
      generateUUID_(),              // LogID
      new Date().toISOString(),     // Timestamp
      candidateId,                  // CandidateID
      actor,                        // Actor
      event                         // Event description
    ]);
  } catch (e) {
    Logger.log('AUDIT LOG ERROR: ' + e.message);
  }
}

/**
 * Public: Retrieves all audit log entries for a candidate.
 * @param {string} candidateId
 */
function api_getAuditLog(candidateId) {
  const auth = requireRole_(['Admin', 'HR']);
  if (!auth.authorized) return { success: false, error: auth.error };
  try {
    const sheet = getSheet_(SHEET_LOGS);
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const logs = rows
      .filter(row => row[headers.indexOf('CandidateID')] === candidateId)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
    return { success: true, data: logs };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}
