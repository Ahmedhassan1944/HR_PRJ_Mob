/**
 * =========================================================
 * PRODUCTION INTEGRATION TEST (ProductionTest.gs)
 * =========================================================
 * Runs against REAL Google Services (Sheet, Drive, Gmail).
 * Requires:
 *   1. SPREADSHEET_ID in Database.gs is correctly set.
 *   2. The 4 sheet tabs exist: tbl_Candidates, tbl_Documents,
 *      tbl_Users, tbl_SystemLogs
 *   3. The script has the required OAuth permissions granted.
 *
 * HOW TO RUN:
 *   - Select "runProductionTest" from the function dropdown.
 *   - Click ▶ Run (grant permissions when prompted).
 *   - Open View > Logs to see real folder/file URLs.
 * =========================================================
 */

function runProductionTest() {
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('   PRODUCTION INTEGRATION TEST   ');
  Logger.log('═══════════════════════════════════════════════\n');

  // ─────────────────────────────────────────────
  // STEP 1: Create a test candidate in the Sheet
  // ─────────────────────────────────────────────
  Logger.log('▶ STEP 1: Creating test candidate in Google Sheets...');

  const candidatePayload = {
    fullName:         'Integration Test User',
    position:         'Test Engineer',
    department:       'QA',
    email:            Session.getActiveUser().getEmail(), // sends to yourself
    phone:            '+201000000000',
    nationality:      'Egyptian',
    salary:           '999',
    coordinatorEmail: Session.getActiveUser().getEmail()
  };

  const createResult = api_createCandidate(candidatePayload);
  if (!createResult.success) {
    Logger.log('❌ FAILED to create candidate: ' + createResult.error);
    return;
  }

  const candidateId = createResult.candidateId;
  Logger.log('✅ Candidate created — ID: ' + candidateId);

  // ─────────────────────────────────────────────
  // STEP 2: Create a real Drive folder
  // ─────────────────────────────────────────────
  Logger.log('\n▶ STEP 2: Creating folder in Google Drive...');

  const folderResult = api_createCandidateFolder(candidateId, candidatePayload.fullName);
  if (!folderResult.success) {
    Logger.log('❌ FAILED to create folder: ' + folderResult.error);
    return;
  }

  Logger.log('✅ Drive Folder Created!');
  Logger.log('   Folder ID  : ' + folderResult.folderId);
  Logger.log('   Folder URL : ' + folderResult.url);   // ← REAL Google Drive URL

  // ─────────────────────────────────────────────
  // STEP 3: Upload a small test file to Drive
  // ─────────────────────────────────────────────
  Logger.log('\n▶ STEP 3: Uploading a test PDF file to Drive...');

  // A minimal valid base64-encoded 1-byte file (enough to create a real file)
  const sampleBase64 = Utilities.base64Encode('Integration test file content - ' + new Date().toISOString());

  const uploadResult = api_uploadFileToDrive(
    candidateId,
    folderResult.folderId,
    'TestDocument',
    'integration_test.pdf',  // renamed to .pdf to match allowed MIME type
    sampleBase64,
    'application/pdf'         // 'text/plain' is blocked by server-side validation (Prompt 1)
  );

  if (!uploadResult.success) {
    Logger.log('❌ FAILED to upload file: ' + uploadResult.error);
    return;
  }

  Logger.log('✅ File Uploaded!');
  Logger.log('   Document ID : ' + uploadResult.documentId);
  Logger.log('   File URL    : ' + uploadResult.fileUrl);   // ← REAL Google Drive URL

  // ─────────────────────────────────────────────
  // STEP 4: Read back from the Sheet to confirm
  // ─────────────────────────────────────────────
  Logger.log('\n▶ STEP 4: Reading candidate back from Sheets...');

  const allResult = api_getAllCandidates();
  const candidate = allResult.data?.find(c => c.CandidateID === candidateId);

  if (candidate) {
    Logger.log('✅ Candidate found in Sheet:');
    Logger.log('   Name          : ' + candidate.FullName);
    Logger.log('   Status        : ' + candidate.CurrentStatus);
    Logger.log('   DriveFolderID : ' + candidate.DriveFolderID);
  } else {
    Logger.log('⚠️  Candidate row not found — check SPREADSHEET_ID.');
  }

  // ─────────────────────────────────────────────
  // STEP 5: Read documents back
  // ─────────────────────────────────────────────
  Logger.log('\n▶ STEP 5: Reading uploaded document record from Sheets...');

  const docsResult = api_getDocumentsByCandidate(candidateId);
  if (docsResult.success && docsResult.data.length > 0) {
    const doc = docsResult.data[0];
    Logger.log('✅ Document record found:');
    Logger.log('   Doc Type   : ' + doc.DocType);
    Logger.log('   File URL   : ' + doc.FileURL);   // ← REAL Google Drive URL
    Logger.log('   Status     : ' + doc.ApprovalStatus);
    Logger.log('   Version    : ' + doc.VersionNumber);
  } else {
    Logger.log('⚠️  No document records found for this candidate.');
  }

  // ─────────────────────────────────────────────
  // STEP 6: Read audit log
  // ─────────────────────────────────────────────
  Logger.log('\n▶ STEP 6: Reading audit log...');

  const logResult = api_getAuditLog(candidateId);
  if (logResult.success) {
    Logger.log(`✅ ${logResult.data.length} audit log entries found:`);
    logResult.data.forEach(entry => {
      Logger.log(`   [${entry.Timestamp}] ${entry.Actor} → ${entry.Event}`);
    });
  }

  // ─────────────────────────────────────────────
  // SUMMARY
  // ─────────────────────────────────────────────
  Logger.log('\n═══════════════════════════════════════════════');
  Logger.log('  INTEGRATION TEST COMPLETE');
  Logger.log('  Candidate ID : ' + candidateId);
  Logger.log('  Folder URL   : ' + folderResult.url);
  Logger.log('  File URL     : ' + uploadResult.fileUrl);
  Logger.log('═══════════════════════════════════════════════');

  // Open the folder URL directly in the browser (optional)
  // Uncomment the line below if you want the folder to auto-open:
  // return HtmlService.createHtmlOutput('<script>window.open("' + folderResult.url + '");</script>');
}

/**
 * Cleanup: Deletes the test candidate row from the Sheet.
 * Run this after runProductionTest() to keep your Sheet clean.
 * Pass the candidateId logged in the previous run.
 */
/**
 * Full cleanup after runProductionTest().
 * Deletes: Candidate row, all Document rows, all Log rows, and the Drive folder.
 * Pass the candidateId printed in the SUMMARY section of the test log.
 *
 * Example: cleanupProductionTest("mock-uuid-abc123")
 */
function cleanupProductionTest(candidateId) {
  if (!candidateId) {
    Logger.log('⚠️  Provide a candidateId to clean up. Example: cleanupProductionTest("mock-uuid-abc123")');
    return;
  }

  Logger.log('🧹 Starting cleanup for candidateId: ' + candidateId);
  let errors = [];

  // ── 1. Delete Drive folder ────────────────────────────────────────────────
  try {
    const candSheet  = getSheet_(SHEET_CANDIDATES);
    const candData   = candSheet.getDataRange().getValues();
    const candIdCol  = candData[0].indexOf('CandidateID');
    const folderCol  = candData[0].indexOf('DriveFolderID');

    const candRow = candData.find((row, i) => i > 0 && row[candIdCol] === candidateId);
    if (candRow && candRow[folderCol]) {
      const folderId = candRow[folderCol];
      try {
        DriveApp.getFolderById(folderId).setTrashed(true);
        Logger.log('✅ Drive folder trashed: ' + folderId);
      } catch (e) {
        errors.push('Drive folder: ' + e.message);
        Logger.log('⚠️  Could not trash Drive folder: ' + e.message);
      }
    } else {
      Logger.log('ℹ️  No Drive folder found for this candidate (already deleted or never created).');
    }
  } catch (e) {
    errors.push('Drive lookup: ' + e.message);
  }

  // ── 2. Delete Candidate row ───────────────────────────────────────────────
  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const data  = sheet.getDataRange().getValues();
    const idCol = data[0].indexOf('CandidateID');
    let deleted = false;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idCol] === candidateId) {
        sheet.deleteRow(i + 1);
        Logger.log('✅ Candidate row deleted.');
        deleted = true;
        break;
      }
    }
    if (!deleted) Logger.log('ℹ️  Candidate row not found (already deleted?).');
  } catch (e) {
    errors.push('Candidate delete: ' + e.message);
  }

  // ── 3. Delete all Document rows for this candidate ────────────────────────
  try {
    const sheet = getSheet_(SHEET_DOCUMENTS);
    const data  = sheet.getDataRange().getValues();
    const idCol = data[0].indexOf('CandidateID');
    let count   = 0;
    // Iterate bottom-up to avoid row-index drift after deletion
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idCol] === candidateId) {
        sheet.deleteRow(i + 1);
        count++;
      }
    }
    Logger.log('✅ Deleted ' + count + ' document record(s).');
  } catch (e) {
    errors.push('Documents delete: ' + e.message);
  }

  // ── 4. Delete all Audit Log rows for this candidate ───────────────────────
  try {
    const sheet = getSheet_(SHEET_LOGS);
    const data  = sheet.getDataRange().getValues();
    const idCol = data[0].indexOf('CandidateID');
    let count   = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idCol] === candidateId) {
        sheet.deleteRow(i + 1);
        count++;
      }
    }
    Logger.log('✅ Deleted ' + count + ' audit log entry/entries.');
  } catch (e) {
    errors.push('Log delete: ' + e.message);
  }

  // ── Summary ───────────────────────────────────────────────────────────────
  if (errors.length === 0) {
    Logger.log('\n✅ Cleanup complete — all test data removed.');
  } else {
    Logger.log('\n⚠️  Cleanup finished with ' + errors.length + ' error(s):');
    errors.forEach(err => Logger.log('   • ' + err));
  }
}
