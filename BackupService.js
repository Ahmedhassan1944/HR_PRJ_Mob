function createDailyBackup() {
  try {
    const props          = PropertiesService.getScriptProperties();
    const spreadsheetId  = props.getProperty('SPREADSHEET_ID');
    const backupFolderId = props.getProperty('BACKUP_FOLDER_ID');

    if (!spreadsheetId || !backupFolderId) {
      throw new Error('Missing Script Properties: SPREADSHEET_ID or BACKUP_FOLDER_ID not set.');
    }

    const ss        = SpreadsheetApp.openById(spreadsheetId);
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd_HH-mm'
    );
    const backupName = ss.getName() + '_Backup_' + timestamp;
    const file       = DriveApp.getFileById(ss.getId());

    file.makeCopy(backupName, DriveApp.getFolderById(backupFolderId));
    Logger.log('Backup Created: ' + backupName);

  } catch (e) {
    Logger.log('BACKUP FAILED: ' + e.message);
    MailApp.sendEmail(
      Session.getEffectiveUser().getEmail(),
      '⚠️ HR System — Daily Backup Failed',
      'The daily backup failed with the following error:\n\n' +
      e.message +
      '\n\nPlease check the Apps Script execution log immediately.'
    );
  }
}