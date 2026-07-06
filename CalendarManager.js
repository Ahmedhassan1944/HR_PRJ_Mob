/**
 * =========================================================
 * CALENDAR MANAGER MODULE (CalendarManager.gs)
 * =========================================================
 * Manages all Google Calendar operations for candidate follow-ups:
 * - Creating events in the dedicated HR calendar
 * - Setting reminders (email and popup)
 * - Keeping the tbl_Events sheet in sync with Google Calendar
 * =========================================================
 */

const APP_CALENDAR_NAME = 'HR Mobilization — Follow-ups';

/**
 * ONE-TIME AUTHORIZATION HELPER
 * Run this function ONCE from the Apps Script editor to grant Calendar permission.
 * After running it, redeploy the Web App as a "New version".
 * You can safely delete this function afterwards.
 */
function authorizeCalendarScope() {
  const calendar = getAppCalendar_();
  Logger.log('✅ Calendar authorized successfully: ' + calendar.getName());
}

/**
 * Normalizes a sheet EventDate value (which may be a real Date object
 * if Sheets auto-converted it, or a plain string) into a consistent
 * 'YYYY-MM-DD' string.
 */
function normalizeEventDateString_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(value);
}

/**
 * Gets or creates the dedicated Google Calendar, caching its ID in PropertiesService.
 */
function getAppCalendar_() {
  const props = PropertiesService.getScriptProperties();
  const cachedId = props.getProperty('APP_CALENDAR_ID');

  if (cachedId) {
    try {
      const cal = CalendarApp.getCalendarById(cachedId);
      if (cal) return cal;
      // null = calendar was deleted or unshared — clear stale ID and fall through
    } catch (_) {}
    props.deleteProperty('APP_CALENDAR_ID');
  }

  // First-time lookup: search by name
  const calendars = CalendarApp.getCalendarsByName(APP_CALENDAR_NAME);
  let calendar;
  if (calendars.length > 0) {
    calendar = calendars[0];
  } else {
    calendar = CalendarApp.createCalendar(APP_CALENDAR_NAME, {
      summary: 'Automated calendar for HR follow-up events.',
      timeZone: Session.getScriptTimeZone()
    });
  }

  props.setProperty('APP_CALENDAR_ID', calendar.getId());
  return calendar;
}

function getEventsSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_EVENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EVENTS);
    sheet.appendRow([
      'EventID', 'CandidateID', 'CandidateName', 'Title', 'Description',
      'EventDate', 'EventTime', 'Priority', 'ReminderMinutesBefore',
      'Status', 'GoogleCalendarEventId', 'GoogleCalendarLink',
      'CreatedBy', 'CreatedAt', 'UpdatedAt'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Returns a map of CandidateID -> Phone for enriching event objects.
 */
function getCandidatePhoneMap_() {
  const sheet = getSheet_(SHEET_CANDIDATES);
  const [headers, ...rows] = sheet.getDataRange().getValues();
  const idCol = headers.indexOf('CandidateID');
  const phoneCol = headers.indexOf('Phone');
  const map = {};
  rows.forEach(row => { map[row[idCol]] = row[phoneCol]; });
  return map;
}

/**
 * Creates a new calendar event for a candidate and logs it in tbl_Events.
 * @param {string} candidateId
 * @param {object} eventData - { title, description, eventDate, eventTime, priority, reminderMinutesBefore }
 */
function api_createCalendarEvent(candidateId, eventData) {
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };

  // Validate required fields
  if (!eventData.title) return { success: false, error: 'Title is required.' };
  if (!eventData.eventDate) return { success: false, error: 'Event date is required.' };
  const priority = eventData.priority || 'Medium';
  if (!['Low', 'Medium', 'High'].includes(priority)) {
    return { success: false, error: 'Invalid priority.' };
  }

  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('CandidateID');
    let candidateName = '';
    let currentStatus = '';

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === candidateId) {
        candidateName = data[i][headers.indexOf('FullName')];
        currentStatus = data[i][headers.indexOf('CurrentStatus')];
        break;
      }
    }

    if (!candidateName) return { success: false, error: 'Candidate not found.' };

    const calendar = getAppCalendar_();
    let calEvent;
    
    // Parse date (assumes YYYY-MM-DD from client)
    const [year, month, day] = eventData.eventDate.split('-').map(Number);
    const startDate = new Date(year, month - 1, day);

    const descriptionText = `Candidate: ${candidateName}\nStatus: ${currentStatus}\n\n${eventData.description || ''}`;

    if (eventData.eventTime) {
      // time provided (HH:mm)
      const [hours, minutes] = eventData.eventTime.split(':').map(Number);
      startDate.setHours(hours, minutes, 0, 0);
      
      const endDate = new Date(startDate.getTime());
      endDate.setMinutes(endDate.getMinutes() + 30); // 30 min duration

      calEvent = calendar.createEvent(eventData.title, startDate, endDate, {
        description: descriptionText
      });
    } else {
      calEvent = calendar.createAllDayEvent(eventData.title, startDate, {
        description: descriptionText
      });
    }

    if (eventData.reminderMinutesBefore) {
      const mins = parseInt(eventData.reminderMinutesBefore, 10);
      if (!isNaN(mins) && mins > 0) {
        calEvent.addPopupReminder(mins);
        calEvent.addEmailReminder(mins);
      }
    }

    let eventLink = '';
    try { eventLink = calEvent.getHtmlLink() || ''; } catch (e) {}

    const eventId = generateUUID_();
    const now = new Date().toISOString();
    const currentUser = Session.getActiveUser().getEmail();

    const eventsSheet = getEventsSheet_();
    eventsSheet.appendRow([
      eventId,
      candidateId,
      candidateName,
      eventData.title,
      eventData.description || '',
      eventData.eventDate,
      eventData.eventTime || '',
      priority,
      eventData.reminderMinutesBefore || '',
      'Active',
      calEvent.getId(),
      eventLink,
      currentUser,
      now,
      now
    ]);

    api_writeLog_(candidateId, currentUser, 'Calendar Event Created: ' + eventData.title);
    CacheService.getScriptCache().remove('dashboard_data');

    return { success: true, eventId, calendarEventId: calEvent.getId(), calendarLink: eventLink };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Gets all events for a specific candidate.
 */
function api_getEventsByCandidate(candidateId) {
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };
  try {
    const sheet = getEventsSheet_();
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const phoneMap = getCandidatePhoneMap_();
    const events = rows
      .filter(row => row[headers.indexOf('CandidateID')] === candidateId)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        obj.EventDate = normalizeEventDateString_(obj.EventDate);
        obj.CandidatePhone = phoneMap[obj.CandidateID] || '';
        return obj;
      })
      .sort((a, b) => new Date(a.EventDate) - new Date(b.EventDate));
      
    return { success: true, data: events };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Gets all active upcoming events for dashboard/calendar view.
 */
function api_getAllUpcomingEvents() {
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };
  try {
    const sheet = getEventsSheet_();
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const phoneMap = getCandidatePhoneMap_();
    const events = rows
      .filter(row => row[headers.indexOf('Status')] === 'Active')
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        obj.EventDate = normalizeEventDateString_(obj.EventDate);
        obj.CandidatePhone = phoneMap[obj.CandidateID] || '';
        return obj;
      });
      
    return { success: true, data: events };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Updates an existing calendar event.
 */
function api_updateCalendarEvent(eventId, updates) {
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };

  try {
    const sheet = getEventsSheet_();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('EventID');
    
    let rowIndex = -1;
    let rowData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === eventId) {
        rowIndex = i + 1;
        rowData = data[i];
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Event not found in database.' };

    const googleEventId = rowData[headers.indexOf('GoogleCalendarEventId')];
    let warning = null;

    try {
      const calendar = getAppCalendar_();
      const calEvent = calendar.getEventById(googleEventId);
      
      if (calEvent) {
        if (updates.title) calEvent.setTitle(updates.title);
        
        let newDate = normalizeEventDateString_(updates.eventDate || rowData[headers.indexOf('EventDate')]);
        let newTime = updates.eventTime !== undefined ? updates.eventTime : rowData[headers.indexOf('EventTime')];
        
        const [year, month, day] = newDate.split('-').map(Number);
        const startDate = new Date(year, month - 1, day);
        
        if (newTime) {
          const [hours, minutes] = newTime.split(':').map(Number);
          startDate.setHours(hours, minutes, 0, 0);
          const endDate = new Date(startDate.getTime());
          endDate.setMinutes(endDate.getMinutes() + 30);
          calEvent.setTime(startDate, endDate);
        } else {
          calEvent.setAllDayDate(startDate);
        }

        if (updates.description !== undefined) {
           const candName = rowData[headers.indexOf('CandidateName')];
           const candId = rowData[headers.indexOf('CandidateID')];
           let candStatus = 'Unknown';
           
           const candSheet = getSheet_(SHEET_CANDIDATES);
           const candData = candSheet.getDataRange().getValues();
           const candHeaders = candData[0];
           const candIdCol = candHeaders.indexOf('CandidateID');
           
           for (let i = 1; i < candData.length; i++) {
             if (candData[i][candIdCol] === candId) {
               candStatus = candData[i][candHeaders.indexOf('CurrentStatus')];
               break;
             }
           }
           
           const descriptionText = `Candidate: ${candName}\nStatus: ${candStatus}\n\n${updates.description || ''}`;
           calEvent.setDescription(descriptionText);
        }

        // Reminders update
        if (updates.reminderMinutesBefore !== undefined) {
           calEvent.removeAllReminders();
           const mins = parseInt(updates.reminderMinutesBefore, 10);
           if (!isNaN(mins) && mins > 0) {
             calEvent.addPopupReminder(mins);
             calEvent.addEmailReminder(mins);
           }
        }
      } else {
        warning = 'Google Calendar event was already deleted or is inaccessible.';
      }
    } catch (err) {
      Logger.log('Google Calendar update failed: ' + err.message);
      warning = 'Google Calendar update failed: ' + err.message;
    }

    // Update sheet
    if (updates.title) sheet.getRange(rowIndex, headers.indexOf('Title') + 1).setValue(updates.title);
    if (updates.description !== undefined) sheet.getRange(rowIndex, headers.indexOf('Description') + 1).setValue(updates.description);
    if (updates.eventDate) sheet.getRange(rowIndex, headers.indexOf('EventDate') + 1).setValue(updates.eventDate);
    if (updates.eventTime !== undefined) sheet.getRange(rowIndex, headers.indexOf('EventTime') + 1).setValue(updates.eventTime);
    if (updates.priority) sheet.getRange(rowIndex, headers.indexOf('Priority') + 1).setValue(updates.priority);
    if (updates.reminderMinutesBefore !== undefined) sheet.getRange(rowIndex, headers.indexOf('ReminderMinutesBefore') + 1).setValue(updates.reminderMinutesBefore);
    
    sheet.getRange(rowIndex, headers.indexOf('UpdatedAt') + 1).setValue(new Date().toISOString());

    const titleToLog = updates.title || rowData[headers.indexOf('Title')];
    api_writeLog_(rowData[headers.indexOf('CandidateID')], Session.getActiveUser().getEmail(), 'Calendar Event Updated: ' + titleToLog);
    CacheService.getScriptCache().remove('dashboard_data');

    return { success: true, warning: warning };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Soft deletes an event from the app and removes it from Google Calendar.
 */
function api_deleteCalendarEvent(eventId) {
  const auth = requireRole_(['Admin', 'HR']);
  if (!auth.authorized) return { success: false, error: auth.error };

  try {
    const sheet = getEventsSheet_();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('EventID');
    
    let rowIndex = -1;
    let rowData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === eventId) {
        rowIndex = i + 1;
        rowData = data[i];
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Event not found.' };

    const googleEventId = rowData[headers.indexOf('GoogleCalendarEventId')];
    try {
      const calendar = getAppCalendar_();
      const calEvent = calendar.getEventById(googleEventId);
      if (calEvent) {
        calEvent.deleteEvent();
      }
    } catch (err) {
      Logger.log('Error deleting from Google Calendar: ' + err.message);
    }

    // Soft delete
    sheet.getRange(rowIndex, headers.indexOf('Status') + 1).setValue('Cancelled');
    sheet.getRange(rowIndex, headers.indexOf('UpdatedAt') + 1).setValue(new Date().toISOString());

    api_writeLog_(rowData[headers.indexOf('CandidateID')], Session.getActiveUser().getEmail(), 'Calendar Event Cancelled: ' + rowData[headers.indexOf('Title')]);
    CacheService.getScriptCache().remove('dashboard_data');

    return { success: true };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Marks an event as completed.
 */
function api_markEventCompleted(eventId) {
  const auth = requireRole_(['Admin', 'HR', 'Coordinator']);
  if (!auth.authorized) return { success: false, error: auth.error };

  try {
    const sheet = getEventsSheet_();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('EventID');
    
    let rowIndex = -1;
    let rowData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === eventId) {
        rowIndex = i + 1;
        rowData = data[i];
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'Event not found.' };

    sheet.getRange(rowIndex, headers.indexOf('Status') + 1).setValue('Completed');
    sheet.getRange(rowIndex, headers.indexOf('UpdatedAt') + 1).setValue(new Date().toISOString());

    api_writeLog_(rowData[headers.indexOf('CandidateID')], Session.getActiveUser().getEmail(), 'Calendar Event Completed: ' + rowData[headers.indexOf('Title')]);
    CacheService.getScriptCache().remove('dashboard_data');

    return { success: true };
  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}
