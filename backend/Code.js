// Code.gs

// ==========================================
// ENVIRONMENT CONFIGURATION
// ==========================================
const APP_ENV = 'PROD'; // Change to 'DEV' for development mode

const ENV_CONFIG = {
  PROD: {
    sheetName: 'WERONE_Database',
    folderName: 'WERONE_Icons',
    sheetProp: 'SHEET_ID_PROD',
    folderProp: 'FOLDER_ID_PROD'
  },
  DEV: {
    sheetName: 'WERONE_Database_DEV',
    folderName: 'WERONE_Icons_DEV',
    sheetProp: 'SHEET_ID_DEV',
    folderProp: 'FOLDER_ID_DEV'
  }
};

// Initialize the database and folder automatically based on environment
function setupStorage() {
  const props = PropertiesService.getScriptProperties();
  const config = ENV_CONFIG[APP_ENV];
  
  let sheetId = props.getProperty(config.sheetProp);
  let folderId = props.getProperty(config.folderProp);
  
  // Setup Sheet for the current environment
  if (!sheetId) {
    let ss = SpreadsheetApp.create(config.sheetName);
    let sheet = ss.getActiveSheet();
    sheet.appendRow(['ID', 'Name', 'URL', 'IconURL', 'OrderIndex']);
    props.setProperty(config.sheetProp, ss.getId());
    sheetId = ss.getId();
  }
  
  // Setup Folder for the current environment
  if (!folderId) {
    let folders = DriveApp.getFoldersByName(config.folderName);
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(config.folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    props.setProperty(config.folderProp, folder.getId());
  }
  
  return SpreadsheetApp.openById(sheetId).getActiveSheet();
}

function doPost(e) {
  let request;
  try {
    request = JSON.parse(e.postData.contents);
  } catch(err) {
    return respond({ success: false, error: 'Invalid JSON request.' });
  }

  const action = request.action;
  const payload = request.payload;
  
  try {
    // Read operations (no password needed)
    if (action === 'getApps') {
      return respond({ success: true, data: fetchApps() });
    }
    
    // Auth operations
    if (action === 'verifyPassword') {
      const isValid = verifyPassword(payload.password);
      return respond({ success: isValid });
    }
    
    // Write operations (require password)
    if (!verifyPassword(payload.password)) {
      return respond({ success: false, error: 'Unauthorized: Incorrect Password' });
    }
    
    let result;
    switch(action) {
      case 'addApp':
        result = addApp(payload.app);
        break;
      case 'editApp':
        result = editApp(payload.app);
        break;
      case 'deleteApp':
        result = deleteApp(payload.id);
        break;
      case 'reorderApps':
        result = reorderApps(payload.orderedIds);
        break;
      case 'uploadIcon':
        result = uploadIcon(payload.filename, payload.base64, payload.mimeType);
        break;
      default:
        return respond({ success: false, error: 'Unknown action.' });
    }
    
    return respond({ success: true, data: result });
    
  } catch (error) {
    return respond({ success: false, error: error.toString() });
  }
}

function doGet(e) {
  return respond({ success: true, message: `WERONE Backend Active. Environment: ${APP_ENV}` });
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function verifyPassword(inputPass) {
  const actualPass = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return inputPass && inputPass === actualPass;
}

function fetchApps() {
  const sheet = setupStorage();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return[];
  
  const headers = data.shift();
  let apps = data.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  
  apps.sort((a, b) => Number(a.OrderIndex) - Number(b.OrderIndex));
  return apps;
}

function addApp(app) {
  const sheet = setupStorage();
  const id = Utilities.getUuid();
  const newRow =[id, app.name, app.url, app.iconUrl, fetchApps().length];
  sheet.appendRow(newRow);
  return { id: id };
}

function editApp(app) {
  const sheet = setupStorage();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === app.id) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[app.name, app.url, app.iconUrl]]);
      return true;
    }
  }
  throw new Error("App not found.");
}

function deleteApp(id) {
  const sheet = setupStorage();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error("App not found.");
}

function reorderApps(orderedIds) {
  const sheet = setupStorage();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    let id = data[i][0];
    let newIndex = orderedIds.indexOf(id);
    if (newIndex !== -1) {
      sheet.getRange(i + 1, 5).setValue(newIndex);
    }
  }
  return true;
}

function uploadIcon(filename, base64Data, mimeType) {
  const props = PropertiesService.getScriptProperties();
  const config = ENV_CONFIG[APP_ENV];
  const folderId = props.getProperty(config.folderProp);
  const folder = DriveApp.getFolderById(folderId);
  
  const decodedBytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decodedBytes, mimeType, filename);
  
  const file = folder.createFile(blob);
  return `https://drive.google.com/uc?export=view&id=${file.getId()}`;
}
