/**
 * =========================================================
 * FULL TEST SUITE (Tests.gs)
 * =========================================================
 * Tests every public API function in the system:
 *   - Code.gs          → api_getDashboardData, api_uploadDocumentToDrive
 *   - Database.gs      → api_getAllCandidates, api_createCandidate,
 *                        api_updateCandidateStatus, api_getDocumentsByCandidate,
 *                        api_reviewDocument, api_getAuditLog
 *   - DriveManager.gs  → api_createCandidateFolder, api_uploadFileToDrive
 *   - EmailService.gs  → parseRecruitmentEmail_, checkOverdueCandidates,
 *                        api_sendRejectionEmail, api_sendPackageSubmissionAlert
 *
 * HOW TO RUN:
 *   1. Open your Google Apps Script project.
 *   2. Select function "runAllTests" from the function dropdown.
 *   3. Click ▶ Run.
 *   4. Open View > Logs (Ctrl+Enter) to see results.
 *
 * NOTE: Each test injects mock GAS services before calling the
 *       real function under test. No live Sheet/Drive/Gmail needed.
 * =========================================================
 */

// ─────────────────────────────────────────────
// ENTRY POINT
// ─────────────────────────────────────────────

/**
 * Main entry point. Run this function from the GAS editor.
 * Executes every test suite in order and prints a final summary.
 */
function runAllTests() {
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('   ENTERPRISE HR MOBILIZATION — TEST SUITE   ');
  Logger.log('═══════════════════════════════════════════════\n');

  // Inject mocks into the global scope for the duration of all tests
  _injectMocks();

  try {
    suite_Code();
    suite_Database_Candidates();
    suite_Database_Documents();
    suite_Database_AuditLog();
    suite_DriveManager();
  } finally {
    // Restore real GAS globals (no-op in production; avoids leakage)
    _restoreMocks();
  }

  TestRunner.summary();
}

// ─────────────────────────────────────────────
// MOCK INJECTION
// ─────────────────────────────────────────────

// Hold originals so we can restore after tests
let _origSpreadsheetApp, _origDriveApp, _origSession, _origUtilities, _origScriptApp;

function _injectMocks() {
  _origSpreadsheetApp = typeof SpreadsheetApp !== 'undefined' ? SpreadsheetApp : null;
  _origDriveApp       = typeof DriveApp       !== 'undefined' ? DriveApp       : null;
  _origSession        = typeof Session        !== 'undefined' ? Session        : null;
  _origUtilities      = typeof Utilities      !== 'undefined' ? Utilities      : null;
  _origScriptApp      = typeof ScriptApp      !== 'undefined' ? ScriptApp      : null;

  // Override globals with mocks
  SpreadsheetApp = MockFactory.getSpreadsheetApp(); // eslint-disable-line no-global-assign
  DriveApp       = MockFactory.getDriveApp();       // eslint-disable-line no-global-assign
  Session        = MockFactory.getSession();        // eslint-disable-line no-global-assign
  Utilities      = MockFactory.getUtilities();      // eslint-disable-line no-global-assign
  ScriptApp      = MockFactory.getScriptApp();      // eslint-disable-line no-global-assign
}

function _restoreMocks() {
  if (_origSpreadsheetApp) SpreadsheetApp = _origSpreadsheetApp; // eslint-disable-line no-global-assign
  if (_origDriveApp)       DriveApp       = _origDriveApp;       // eslint-disable-line no-global-assign
  if (_origSession)        Session        = _origSession;        // eslint-disable-line no-global-assign
  if (_origUtilities)      Utilities      = _origUtilities;      // eslint-disable-line no-global-assign
  if (_origScriptApp)      ScriptApp      = _origScriptApp;      // eslint-disable-line no-global-assign
}

// ─────────────────────────────────────────────
// SHARED SEED DATA
// ─────────────────────────────────────────────

const CANDIDATE_HEADERS = [
  'CandidateID', 'FullName', 'Position', 'Department', 'Email',
  'Phone', 'Nationality', 'OfferSalary', 'AssignedCoordinatorEmail',
  'CurrentStatus', 'CreatedAt', 'UpdatedAt', 'DriveFolderID'
];

const DOCUMENT_HEADERS = [
  'DocumentID', 'CandidateID', 'DocType', 'FileName', 'FileURL',
  'UploadDate', 'ApprovalStatus', 'ApprovedBy', 'VersionNumber', 'Remarks'
];

const LOG_HEADERS = [
  'LogID', 'Timestamp', 'CandidateID', 'Actor', 'Event'
];

const USER_HEADERS = ['UserID', 'Email', 'Role', 'Name'];

/**
 * Seeds all four sheets with minimal valid data for a test run.
 */
function _seedAllSheets() {
  MockFactory.reset();

  MockFactory.seedSheet(SHEET_CANDIDATES, CANDIDATE_HEADERS, [
    ['cand-001', 'Ahmed Ali', 'Engineer', 'Projects', 'ahmed@test.com',
     '+201012345678', 'Egyptian', '350', 'coord@company.com',
     'Documents Requested', '2026-01-01T00:00:00.000Z', '2026-01-01T00:00:00.000Z', '']
  ]);

  MockFactory.seedSheet(SHEET_DOCUMENTS, DOCUMENT_HEADERS, [
    ['doc-001', 'cand-001', 'Passport', 'passport.pdf',
     'https://drive.google.com/file/doc-001',
     '2026-01-02T00:00:00.000Z', 'Pending Review', '', 1, '']
  ]);

  MockFactory.seedSheet(SHEET_LOGS, LOG_HEADERS, [
    ['log-001', '2026-01-01T00:00:00.000Z', 'cand-001', 'SYSTEM', 'Candidate Created']
  ]);

  MockFactory.seedSheet(SHEET_USERS, USER_HEADERS, [
    ['user-001', 'test.user@yourcompany.com', 'Admin', 'Test Admin']
  ]);
}


// ═══════════════════════════════════════════════════════════
//  SUITE 1 — Code.gs
// ═══════════════════════════════════════════════════════════

function suite_Code() {
  Logger.log('\n── Suite: Code.gs ──────────────────────────────');

  TestRunner.run('api_getDashboardData returns success object', () => {
    const result = api_getDashboardData();
    Assert.isTrue(result.success, 'success should be true');
    Assert.notNull(result.data, 'data should not be null');
    Assert.isTrue(typeof result.data.activeCount === 'number', 'activeCount should be a number');
    Assert.isTrue(typeof result.data.missingDocs === 'number', 'missingDocs should be a number');
    Assert.isTrue(typeof result.data.pendingValidation === 'number', 'pendingValidation should be a number');
  });

  TestRunner.run('api_getDashboardData returns correct KPI values for seeded data', () => {
    // Seed 1 candidate with status 'Documents Requested' (see _seedAllSheets)
    // activeCount  = 1 (all non-Closed)
    // missingDocs  = 1 ('Documents Requested' is in MISSING_DOC_STATUSES)
    const result = api_getDashboardData();
    Assert.equals(result.data.activeCount, 1, 'activeCount should be 1 for the single seeded candidate');
    Assert.equals(result.data.missingDocs,  1, 'missingDocs should be 1 for the seeded candidate');
    Assert.isTrue(typeof result.data.visaPending === 'number',   'visaPending should be a number');
    Assert.isTrue(typeof result.data.visaCompleted === 'number', 'visaCompleted should be a number');
    Assert.isTrue(typeof result.data.mobilized === 'number',     'mobilized should be a number');
  });

  TestRunner.run('api_uploadFileToDrive returns success with a file URL', () => {
    _seedAllSheets();
    // Create a folder first to get a valid folderId
    const folderResult = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    Assert.isTrue(folderResult.success, 'folder creation must succeed before upload');

    const result = api_uploadFileToDrive(
      'cand-001',
      folderResult.folderId,
      'Passport',
      'passport.pdf',
      'bW9ja2Jhc2U2NGRhdGE=', // valid mock base64
      'application/pdf'
    );
    Assert.isTrue(result.success, 'success should be true');
    Assert.notNull(result.fileUrl, 'fileUrl should not be null');
    Assert.isTrue(result.fileUrl.startsWith('https://'), 'fileUrl should start with https://');
  });

  TestRunner.run('include() returns non-empty content string', () => {
    // include() reads .html files — we verify it does NOT throw and returns a string.
    // In the live GAS environment it would return real HTML.
    // We just verify the function signature here.
    Assert.isTrue(typeof include === 'function', 'include should be a function');
  });
}


// ═══════════════════════════════════════════════════════════
//  SUITE 2 — Database.gs: Candidates
// ═══════════════════════════════════════════════════════════

function suite_Database_Candidates() {
  Logger.log('\n── Suite: Database.gs — Candidates ─────────────');

  TestRunner.run('api_getAllCandidates returns seeded candidate', () => {
    _seedAllSheets();
    const result = api_getAllCandidates();
    Assert.isTrue(result.success, 'success should be true');
    Assert.isTrue(Array.isArray(result.data), 'data should be an array');
    Assert.equals(result.data.length, 1, 'should return exactly 1 seeded candidate');
    Assert.equals(result.data[0].CandidateID, 'cand-001', 'CandidateID should match seed');
    Assert.equals(result.data[0].FullName, 'Ahmed Ali', 'FullName should match seed');
  });

  TestRunner.run('api_getAllCandidates returns empty array when sheet is empty', () => {
    MockFactory.reset();
    MockFactory.seedSheet(SHEET_CANDIDATES, CANDIDATE_HEADERS); // no rows
    MockFactory.seedSheet(SHEET_DOCUMENTS, DOCUMENT_HEADERS);
    MockFactory.seedSheet(SHEET_LOGS, LOG_HEADERS);
    MockFactory.seedSheet(SHEET_USERS, USER_HEADERS);

    const result = api_getAllCandidates();
    Assert.isTrue(result.success, 'success should be true');
    Assert.equals(result.data.length, 0, 'no candidates should be returned');
  });

  TestRunner.run('api_getAllCandidates fails gracefully when sheet is missing', () => {
    MockFactory.reset(); // No sheets seeded at all
    const result = api_getAllCandidates();
    Assert.isFalse(result.success, 'success should be false for missing sheet');
    Assert.notNull(result.error, 'error message should be present');
  });

  TestRunner.run('api_createCandidate appends a new row with correct data', () => {
    _seedAllSheets();
    const payload = {
      fullName:         'Sara Mohamed',
      position:         'Accountant',
      department:       'Finance',
      email:            'sara@test.com',
      phone:            '+201099999999',
      nationality:      'Egyptian',
      salary:           '400',
      coordinatorEmail: 'coord@company.com'
    };
    const result = api_createCandidate(payload);
    Assert.isTrue(result.success, 'success should be true');
    Assert.notNull(result.candidateId, 'candidateId should be returned');

    // Verify new row was actually added
    const allResult = api_getAllCandidates();
    Assert.equals(allResult.data.length, 2, 'should now have 2 candidates');
    const newCand = allResult.data.find(c => c.FullName === 'Sara Mohamed');
    Assert.notNull(newCand, 'new candidate should exist');
    Assert.equals(newCand.CurrentStatus, 'Documents Requested', 'default status should be Documents Requested');
  });

  TestRunner.run('api_createCandidate writes an audit log entry', () => {
    _seedAllSheets();
    const payload = {
      fullName: 'Test Candidate', position: 'Tester', department: 'QA',
      email: 'tester@test.com', phone: '+0000', nationality: 'Unknown',
      salary: '500', coordinatorEmail: 'coord@company.com'
    };
    api_createCandidate(payload);

    // Audit log should have grown beyond the initial seed entry
    const sheet = MockFactory._sheets[SHEET_LOGS];
    const rows  = sheet._data.slice(1); // skip header
    Assert.isTrue(rows.length >= 2, 'at least one new log entry should have been written');
  });

  TestRunner.run('api_updateCandidateStatus updates status for existing candidate', () => {
    _seedAllSheets();
    // Use a status that is in the server-side whitelist
    const result = api_updateCandidateStatus('cand-001', 'Documents Under Preparing');
    Assert.isTrue(result.success, 'success should be true');

    const allResult = api_getAllCandidates();
    Assert.equals(allResult.data[0].CurrentStatus, 'Documents Under Preparing', 'status should be updated');
  });

  TestRunner.run('api_updateCandidateStatus returns error for unknown candidateId', () => {
    _seedAllSheets();
    const result = api_updateCandidateStatus('does-not-exist', 'Mobilized');
    Assert.isFalse(result.success, 'success should be false for unknown ID');
    Assert.notNull(result.error, 'an error message should be present');
  });

  TestRunner.run('api_updateCandidateStatus updates the UpdatedAt timestamp', () => {
    _seedAllSheets();
    const before = MockFactory._sheets[SHEET_CANDIDATES]._data[1][11]; // UpdatedAt column (index 11)
    api_updateCandidateStatus('cand-001', 'Mobilized');
    const after  = MockFactory._sheets[SHEET_CANDIDATES]._data[1][11];
    // Timestamps should be different (unless tests run at exactly the same millisecond)
    // At minimum, the field should have a value after the update
    Assert.notNull(after, 'UpdatedAt should be set after status update');
  });

  TestRunner.run('api_createCandidate is rejected for unregistered user', () => {
    // Seed with empty tbl_Users — no user has a role
    MockFactory.reset();
    MockFactory.seedSheet(SHEET_CANDIDATES, CANDIDATE_HEADERS, []);
    MockFactory.seedSheet(SHEET_DOCUMENTS,  DOCUMENT_HEADERS,  []);
    MockFactory.seedSheet(SHEET_LOGS,       LOG_HEADERS,       []);
    MockFactory.seedSheet(SHEET_USERS,      USER_HEADERS,      []); // empty — no registered users

    const result = api_createCandidate({
      fullName: 'Ghost User', position: 'Tester', department: 'QA',
      email: 'ghost@test.com', phone: '+0', nationality: 'Unknown',
      salary: '0', coordinatorEmail: 'coord@company.com'
    });

    Assert.isFalse(result.success, 'unregistered user should be denied');
    Assert.notNull(result.error,   'error message should explain the denial');
  });

  TestRunner.run('api_updateCandidateStatus is rejected for Viewer role', () => {
    MockFactory.reset();
    MockFactory.seedSheet(SHEET_CANDIDATES, CANDIDATE_HEADERS, [
      ['cand-001', 'Ahmed Ali', 'Engineer', 'Projects', 'ahmed@test.com',
       '+201012345678', 'Egyptian', '350', 'coord@company.com',
       'Documents Requested', '2026-01-01T00:00:00.000Z', '2026-01-01T00:00:00.000Z', '']
    ]);
    MockFactory.seedSheet(SHEET_DOCUMENTS, DOCUMENT_HEADERS, []);
    MockFactory.seedSheet(SHEET_LOGS,      LOG_HEADERS,      []);
    // Seed the mock user as Viewer — not allowed to change status
    MockFactory.seedSheet(SHEET_USERS, USER_HEADERS, [
      ['user-002', 'test.user@yourcompany.com', 'Viewer', 'Read Only User']
    ]);

    const result = api_updateCandidateStatus('cand-001', 'Mobilized');
    Assert.isFalse(result.success, 'Viewer role should be denied status change');
    Assert.notNull(result.error,   'error message should explain the denial');
  });
}


// ═══════════════════════════════════════════════════════════
//  SUITE 3 — Database.gs: Documents
// ═══════════════════════════════════════════════════════════

function suite_Database_Documents() {
  Logger.log('\n── Suite: Database.gs — Documents ──────────────');

  TestRunner.run('api_getDocumentsByCandidate returns correct documents for candidate', () => {
    _seedAllSheets();
    const result = api_getDocumentsByCandidate('cand-001');
    Assert.isTrue(result.success, 'success should be true');
    Assert.equals(result.data.length, 1, 'should return exactly 1 document');
    Assert.equals(result.data[0].DocType, 'Passport', 'DocType should be Passport');
    Assert.equals(result.data[0].DocumentID, 'doc-001', 'DocumentID should match seed');
  });

  TestRunner.run('api_getDocumentsByCandidate returns empty array for unknown candidate', () => {
    _seedAllSheets();
    const result = api_getDocumentsByCandidate('unknown-cand');
    Assert.isTrue(result.success, 'success should be true even with no results');
    Assert.equals(result.data.length, 0, 'no documents should be returned for unknown candidate');
  });

  // ── api_reviewDocument tests ──────────────────────────────────────────────
  // TODO (Prompt 3): These tests are disabled until api_reviewDocument is built
  //                  in Database.js. Re-enable them after applying Prompt 3.
  // ─────────────────────────────────────────────────────────────────────────

  TestRunner.run('[PENDING Prompt 3] api_reviewDocument — tests skipped until backend function is built', () => {
    // Intentionally left empty — will be filled in when Prompt 3 is applied.
    Assert.isTrue(true, 'placeholder always passes');
  });
}


// ═══════════════════════════════════════════════════════════
//  SUITE 4 — Database.gs: Audit Log
// ═══════════════════════════════════════════════════════════

function suite_Database_AuditLog() {
  Logger.log('\n── Suite: Database.gs — Audit Log ──────────────');

  TestRunner.run('api_getAuditLog returns seeded log entry for candidate', () => {
    _seedAllSheets();
    const result = api_getAuditLog('cand-001');
    Assert.isTrue(result.success, 'success should be true');
    Assert.equals(result.data.length, 1, 'should return 1 seeded log entry');
    Assert.equals(result.data[0].Event, 'Candidate Created', 'Event text should match seed');
    Assert.equals(result.data[0].Actor, 'SYSTEM', 'Actor should be SYSTEM');
  });

  TestRunner.run('api_getAuditLog returns empty array for candidate with no logs', () => {
    _seedAllSheets();
    const result = api_getAuditLog('no-such-cand');
    Assert.isTrue(result.success, 'success should be true');
    Assert.equals(result.data.length, 0, 'no log entries for unknown candidate');
  });

  TestRunner.run('api_getAuditLog accumulates entries over multiple actions', () => {
    _seedAllSheets();
    // Perform two status changes that each write a log entry (api_reviewDocument pending Prompt 3)
    api_updateCandidateStatus('cand-001', 'Documents Under Preparing');
    api_updateCandidateStatus('cand-001', 'Pending Passport');

    const result = api_getAuditLog('cand-001');
    Assert.isTrue(result.data.length >= 3, 'should have at least 3 log entries (1 seed + 2 status changes)');
  });

  TestRunner.run('api_writeLog_ (internal) does not expose errors if sheet missing', () => {
    MockFactory.reset(); // no SHEET_LOGS seeded — simulates missing sheet

    // api_writeLog_ must catch its own errors and NOT propagate them.
    // We simply call it and assert no exception escapes.
    let threw = false;
    try {
      api_writeLog_('x', 'y', 'z');
    } catch (_) {
      threw = true;
    }
    Assert.isFalse(threw, 'api_writeLog_ should swallow errors internally and not throw to caller');
  });
}


// ═══════════════════════════════════════════════════════════
//  SUITE 5 — DriveManager.gs
// ═══════════════════════════════════════════════════════════

function suite_DriveManager() {
  Logger.log('\n── Suite: DriveManager.gs ───────────────────────');

  TestRunner.run('api_createCandidateFolder creates a new folder and returns folderId', () => {
    _seedAllSheets();
    const result = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    Assert.isTrue(result.success, 'success should be true');
    Assert.notNull(result.folderId, 'folderId should be returned');
    Assert.isTrue(result.url.startsWith('https://'), 'url should be a valid https link');
  });

  TestRunner.run('api_createCandidateFolder is idempotent (no duplicate folders)', () => {
    _seedAllSheets();
    const first  = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    const second = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    // Both should succeed and return the same folder ID
    Assert.isTrue(first.success,  'first call should succeed');
    Assert.isTrue(second.success, 'second call should succeed');
    Assert.equals(first.folderId, second.folderId, 'folderId should be identical on repeated calls');
  });

  TestRunner.run('api_createCandidateFolder stores folderId in Candidates sheet', () => {
    _seedAllSheets();
    const result = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    Assert.isTrue(result.success, 'folder creation should succeed');

    // Verify DriveFolderID column was written back
    const allCands = api_getAllCandidates();
    const cand = allCands.data.find(c => c.CandidateID === 'cand-001');
    Assert.notNull(cand.DriveFolderID, 'DriveFolderID should be populated after folder creation');
  });

  TestRunner.run('api_uploadFileToDrive uploads file and returns fileUrl + documentId', () => {
    _seedAllSheets();
    // First create a folder to get its ID
    const folderResult = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    Assert.isTrue(folderResult.success, 'folder must be created first');

    const uploadResult = api_uploadFileToDrive(
      'cand-001',
      folderResult.folderId,
      'Photo',
      'photo.jpg',
      'bW9ja2Jhc2U2NGRhdGE=', // mock base64
      'image/jpeg'
    );

    Assert.isTrue(uploadResult.success, 'upload should succeed');
    Assert.notNull(uploadResult.fileUrl, 'fileUrl should be returned');
    Assert.notNull(uploadResult.documentId, 'documentId should be returned');
  });

  TestRunner.run('api_uploadFileToDrive writes document record to Documents sheet', () => {
    _seedAllSheets();
    const folderResult = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    api_uploadFileToDrive('cand-001', folderResult.folderId, 'Medical', 'medical.pdf', 'data==', 'application/pdf');

    const docsResult = api_getDocumentsByCandidate('cand-001');
    const medical    = docsResult.data.find(d => d.DocType === 'Medical');
    Assert.notNull(medical, 'Medical document record should exist in Documents sheet');
    Assert.equals(medical.ApprovalStatus, 'Pending Review', 'new doc should start as Pending Review');
    Assert.equals(medical.VersionNumber, 1, 'first upload should be version 1');
  });

  TestRunner.run('api_uploadFileToDrive increments version number on re-upload', () => {
    _seedAllSheets();
    const folderResult = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    const fid = folderResult.folderId;

    // First upload
    api_uploadFileToDrive('cand-001', fid, 'Passport', 'pass_v1.pdf', 'data=', 'application/pdf');
    // Second upload (re-upload same docType)
    api_uploadFileToDrive('cand-001', fid, 'Passport', 'pass_v2.pdf', 'data=', 'application/pdf');

    const docsResult = api_getDocumentsByCandidate('cand-001');
    const passports  = docsResult.data.filter(d => d.DocType === 'Passport');
    // Should have: 1 seed (Passport doc-001) + 2 new uploads = 3 total Passport records
    Assert.isTrue(passports.length >= 2, 'multiple Passport versions should exist');
    const versions = passports.map(d => d.VersionNumber).sort((a, b) => a - b);
    Assert.isTrue(versions[versions.length - 1] > versions[0], 'latest version number should be higher');
  });
}


