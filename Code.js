/**
 * =========================================================
 * ENTERPRISE HR MOBILIZATION SYSTEM - BACKEND CONTROLLER
 * =========================================================
 */

/**
 * Standard GET handler. Evaluates and serves the Index.html template.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Enterprise HR Mobilization')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Required if embedding
}

/**
 * Helper function to inject Styles.html and Script.html into Index.html.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * =========================================================
 * API LAYER -> To be called via google.script.run
 * =========================================================
 */

/**
 * Fetches and aggregates ALL dashboard KPI data.
 *
 * Status-based KPIs (from tbl_Candidates):
 *   activeCount         → all candidates except Closed
 *   missingDocs         → New Candidate | Documents Requested | Documents Under Preparing |
 *                         Pending Passport | Pending Photo | Pending Academic Certificate |
 *                         Pending Medical | Booked a medical examination
 *   visaPending         → Visa Pending
 *   visaCompleted       → Visa Completed
 *   mobilized           → Mobilized
 *   pendingMedical      → Pending Medical
 *   bookedMedical       → Booked a medical examination
 *   docsUnderPreparing  → Documents Under Preparing
 *
 * Document-based KPIs (from tbl_Documents, per candidate):
 *   hasPassport, hasPhoto, hasAcademicCert,
 *   hasMedicalExam, hasMedicalAnalysis, hasVisa, hasCV
 *   → count of candidates who HAVE that doc (ApprovalStatus !== 'Rejected')
 *
 *   missingPassport, missingPhoto, missingAcademicCert,
 *   missingMedicalExam, missingMedicalAnalysis, missingVisa, missingCV
 *   → count of candidates who are MISSING that doc
 */

// ═══════════════════════════════════════════════════════════════════════
// DASHBOARD CACHE — INVALIDATION POLICY (project architecture rule)
// ═══════════════════════════════════════════════════════════════════════
// Cache key : 'dashboard_data'
// TTL       : 60 seconds (FALLBACK ONLY — not the primary mechanism)
// Strategy  : EVENT-BASED invalidation
//
// Any backend function that inserts, updates, deletes, approves, rejects,
// archives, restores, or otherwise modifies data displayed on the Dashboard
// MUST add this line immediately before its success return:
//
//   CacheService.getScriptCache().remove('dashboard_data');
//
// Read-only api_get* functions MUST NOT call remove().
//
// Cache lifecycle:
//   Request → cache HIT  → return cached JSON (0 Sheets calls)
//   Request → cache MISS → read Sheets → compute → cache.put() → return
//   Write op fires       → cache.remove() → next request regenerates
// ═══════════════════════════════════════════════════════════════════════

function api_getDashboardData() {
  try {
    // ── Cache read ─────────────────────────────────────────────────────
    // CacheService stores the computed result for 60 seconds (TTL = fallback).
    // Primary invalidation is EVENT-BASED: every write api_* function calls
    // CacheService.getScriptCache().remove('dashboard_data') before returning.
    // This means the cache is always fresh immediately after any data change.
    const _cache      = CacheService.getScriptCache();
    const _cachedJson = _cache.get('dashboard_data');
    if (_cachedJson) {
      return JSON.parse(_cachedJson);
    }
    // ───────────────────────────────────────────────────────────────────

    const candsRes = api_getAllCandidates();
    if (!candsRes.success) throw new Error(candsRes.error);
    const docsRes  = api_getAllDocuments();
    if (!docsRes.success) throw new Error(docsRes.error);
    const eventsRes = api_getAllUpcomingEvents();
    if (!eventsRes.success) throw new Error(eventsRes.error);
    const candidates = candsRes.data;
    const allDocs    = docsRes.data;
    const activeEvents = eventsRes.data;
    // ── Status-based counters ──────────────────────────────────────
    const MISSING_DOC_STATUSES = new Set([
      'New Candidate',
      'Documents Requested',
      'Documents Under Preparing',
      'Pending Passport',
      'Pending Photo',
      'Pending Academic Certificate',
      'Pending Medical',
      'Booked a medical examination'
    ]);
    let activeCount        = 0;
    let missingDocs        = 0;
    let visaPending        = 0;
    let visaCompleted      = 0;
    let mobilized          = 0;
    let pendingMedical     = 0;
    let bookedMedical      = 0;
    let docsUnderPreparing = 0;
    candidates.forEach(cand => {
      const status = (cand.CurrentStatus || '').trim();
      if (status !== 'Closed') activeCount++;
      if (MISSING_DOC_STATUSES.has(status))           missingDocs++;
      if (status === 'Visa Pending')                   visaPending++;
      if (status === 'Visa Completed')                 visaCompleted++;
      if (status === 'Mobilized')                      mobilized++;
      if (status === 'Pending Medical')                pendingMedical++;
      if (status === 'Booked a medical examination')   bookedMedical++;
      if (status === 'Documents Under Preparing')      docsUnderPreparing++;
    });
    // ── Document-based counters ────────────────────────────────────
    // Build a Set of {CandidateID}_{DocType} for approved/pending docs
    const REQUIRED_DOCS = [
      'Passport', 'Photo', 'Academic Certificate',
      'Medical Examination', 'Medical Analysis', 'Visa', 'CV'
    ];
    // Map: candidateId → Set of docTypes they HAVE (not Rejected)
    const candDocMap = {};
    candidates.forEach(c => { candDocMap[c.CandidateID] = new Set(); });
    allDocs.forEach(doc => {
      if ((doc.ApprovalStatus || '').trim() !== 'Rejected') {
        const key = (doc.CandidateID || '').trim();
        if (candDocMap[key]) {
          candDocMap[key].add((doc.DocType || '').trim());
        }
      }
    });
    // Count has/missing per docType across all candidates
    const hasCount     = {};
    const missingCount = {};
    REQUIRED_DOCS.forEach(dt => { hasCount[dt] = 0; missingCount[dt] = 0; });
    candidates.forEach(cand => {
      const myDocs = candDocMap[cand.CandidateID] || new Set();
      REQUIRED_DOCS.forEach(dt => {
        if (myDocs.has(dt)) hasCount[dt]++;
        else                missingCount[dt]++;
      });
    });
    
    // ── Calendar-based counters ────────────────────────────────────
    const todayDate = new Date();
    todayDate.setHours(0,0,0,0);
    const tomorrowDate = new Date(todayDate);
    tomorrowDate.setDate(tomorrowDate.getDate() + 1);
    const nextWeekDate = new Date(todayDate);
    nextWeekDate.setDate(nextWeekDate.getDate() + 7);

    let eventsToday = 0;
    let eventsTomorrow = 0;
    let eventsThisWeek = 0;
    let eventsOverdue = 0;
    let eventsHighPriority = 0;

    activeEvents.forEach(evt => {
      const dateStr = normalizeEventDateString_(evt.EventDate);
      if (!dateStr) return;
      const [y, m, d] = dateStr.split('-').map(Number);
      const evDate = new Date(y, m - 1, d);
      
      if (evDate.getTime() === todayDate.getTime()) eventsToday++;
      if (evDate.getTime() === tomorrowDate.getTime()) eventsTomorrow++;
      if (evDate >= todayDate && evDate <= nextWeekDate) eventsThisWeek++;
      if (evDate < todayDate) eventsOverdue++;
      if (evt.Priority === 'High' && evDate >= todayDate) eventsHighPriority++;
    });

    const _result = {
      success: true,
      data: {
        // Status-based
        activeCount,
        missingDocs,
        visaPending,
        visaCompleted,
        mobilized,
        pendingMedical,
        bookedMedical,
        docsUnderPreparing,
        // Document HAS counts
        hasPassport:      hasCount['Passport'],
        hasPhoto:         hasCount['Photo'],
        hasAcademicCert:  hasCount['Academic Certificate'],
        hasMedicalExam:   hasCount['Medical Examination'],
        hasMedicalAnalysis: hasCount['Medical Analysis'],
        hasVisa:          hasCount['Visa'],
        hasCV:            hasCount['CV'],
        // Document MISSING counts
        missingPassport:       missingCount['Passport'],
        missingPhoto:          missingCount['Photo'],
        missingAcademicCert:   missingCount['Academic Certificate'],
        missingMedicalExam:    missingCount['Medical Examination'],
        missingMedicalAnalysis:missingCount['Medical Analysis'],
        missingVisa:           missingCount['Visa'],
        missingCV:             missingCount['CV'],
        // Calendar-based
        eventsToday,
        eventsTomorrow,
        eventsThisWeek,
        eventsOverdue,
        eventsHighPriority,
        msg: 'Live data successfully retrieved from Google Sheets.'
      }
    };

    // ── Cache write ───────────────────────────────────────────────────
    // TTL = 60 seconds (fallback only). The cache is primarily invalidated
    // by event-based removal in every write api_* function — not by TTL.
    // If JSON.stringify exceeds 100KB, put() fails silently (try/catch).
    try { _cache.put('dashboard_data', JSON.stringify(_result), 60); } catch (_) {}
    // ───────────────────────────────────────────────────────────────────

    return _result;
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}
