function getMasterSheet() {
  const props = PropertiesService.getScriptProperties();
  let masterId = props.getProperty('MASTER_SHEET_ID');
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  
  if (!masterId) {
    const files = folder.getFilesByName(CONFIG.MASTER_SHEET_NAME);
    if (files.hasNext()) {
      masterId = files.next().getId();
    } else {
      // अगर शीट नहीं है, तो बनाएगा
      const ss = SpreadsheetApp.create(CONFIG.MASTER_SHEET_NAME);
      masterId = ss.getId();
      const sheet = ss.getActiveSheet();
      // मास्टर हेडर्स
      sheet.appendRow(["User_ID", "Email", "Name", "Role", "Organization", "API_Key", "Status", "Key_Seen", "User_Sheet_ID"]);
      DriveApp.getFileById(masterId).moveTo(folder);
    }
    props.setProperty('MASTER_SHEET_ID', masterId);
  }
  return SpreadsheetApp.openById(masterId).getActiveSheet();
}

function createUserSheet(email) {
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const ss = SpreadsheetApp.create(`AI_STUDIO_USER_${email}`);
  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(folder);
  
  const sheet = ss.getActiveSheet();
  sheet.setName("Security_Audit_Logs");
  sheet.appendRow(["Timestamp", "Event_Type", "Device", "IP_Address", "Details"]);
  
  return ss.getId();
}

function findUserByEmail(email) {
  const data = getMasterSheet().getDataRange().getValues();
  return data.findIndex(row => row[1] === email); // Row index (0-based)
}
