// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 8_Backup.gs
// Daily Drive backups for restore / recovery
// ============================================================

function backupSpreadsheetNow() {
  const result = createDailyDriveBackup();
  SpreadsheetApp.getUi().alert(
    'Backup created successfully!\n\n'
    + 'File: ' + result.fileName + '\n'
    + 'Folder: ' + result.folderUrl + '\n\n'
    + 'Use this copy anytime to restore data.'
  );
}

function createDailyDriveBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet found for backup.');

  const tz        = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const now       = new Date();
  const dateLabel = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const timeLabel = Utilities.formatDate(now, tz, 'HHmmss');
  const folder    = getBackupFolder_();

  const sourceFile = DriveApp.getFileById(ss.getId());
  const fileName   = ss.getName() + ' - Backup - ' + dateLabel + ' ' + timeLabel;
  const copiedFile = sourceFile.makeCopy(fileName, folder);

  cleanupOldBackups_(folder, CONFIG.BACKUP_KEEP_DAYS || 60);

  Logger.log('Backup created: ' + copiedFile.getUrl());
  return {
    success  : true,
    fileId   : copiedFile.getId(),
    fileName : fileName,
    fileUrl  : copiedFile.getUrl(),
    folderUrl: folder.getUrl()
  };
}

function getBackupFolder_() {
  if (CONFIG.BACKUP_FOLDER_ID) {
    try {
      return DriveApp.getFolderById(CONFIG.BACKUP_FOLDER_ID);
    } catch (e) {
      Logger.log('Invalid BACKUP_FOLDER_ID, falling back to auto-created folder: ' + e);
    }
  }

  const rootFolderName = 'Kings Equestrian Backups';
  let rootFolders = DriveApp.getFoldersByName(rootFolderName);
  const rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(rootFolderName);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const childFolderName = ss.getName() + ' - Daily Backups';
  let childFolders = rootFolder.getFoldersByName(childFolderName);

  return childFolders.hasNext() ? childFolders.next() : rootFolder.createFolder(childFolderName);
}

function cleanupOldBackups_(folder, keepDays) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - Number(keepDays || 60));

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    try {
      if (file.getDateCreated() < cutoff) {
        file.setTrashed(true);
      }
    } catch (e) {
      Logger.log('cleanupOldBackups_ skipped one file: ' + e);
    }
  }
}

function showBackupRestoreHelp() {
  const folder = getBackupFolder_();
  const msg = ''
    + 'Daily backup is now supported.\n\n'
    + 'Restore steps:\n'
    + '1. Open the backup folder link below\n'
    + '2. Open the required dated backup copy\n'
    + '3. If needed, use that file as the restored master or copy sheets/data back\n\n'
    + 'Backup folder:\n' + folder.getUrl() + '\n\n'
    + 'Important: keep this backup folder private to admin-only access.';

  SpreadsheetApp.getUi().alert('Backup & Restore Help', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
