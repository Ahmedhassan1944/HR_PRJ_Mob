You are fixing security and null-reference bugs in two backend files of a Google Apps Script HR web app.

─── FILE 1: Database.js ───

BUG: Three read API functions expose sensitive HR data with no authorization check.
Add requireRole_ at the top of each function (after the try/catch opening):

FUNCTION api_getAllCandidates (line ~118):
ADD at the very start of the function body:
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };

FUNCTION api_getDocumentsByCandidate (line ~277):
ADD at the very start:
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };

FUNCTION api_getAllDocuments (line ~299):
ADD at the very start:
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };

─── FILE 2: CalendarManager.js ───

BUG A: Two read functions expose calendar data with no authorization check.

FUNCTION api_getEventsByCandidate (line ~210):
ADD at the very start:
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };

FUNCTION api_getAllUpcomingEvents (line ~236):
ADD at the very start:
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };

BUG B: getAppCalendar_ does not null-check CalendarApp.getCalendarById() (line ~50).
CalendarApp.getCalendarById() can return null silently (no exception) when the calendar is deleted.

CURRENT:
  if (cachedId) {
    try {
      return CalendarApp.getCalendarById(cachedId);
    } catch (_) {
      props.deleteProperty('APP_CALENDAR_ID');
    }
  }

FIX:
  if (cachedId) {
    try {
      const cal = CalendarApp.getCalendarById(cachedId);
      if (cal) return cal;
      // null = calendar was deleted or unshared — clear stale ID and fall through
    } catch (_) {}
    props.deleteProperty('APP_CALENDAR_ID');
  }

After making all changes, commit with message:
"fix: add requireRole_ to unprotected read APIs + CalendarApp null-check"