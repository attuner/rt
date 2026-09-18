/**
 * PULSE BROADCAST // LIVE MEDIA ENGINE (AUDIO & VIDEO)
 * Backing Google Apps Script
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet() ? SpreadsheetApp.getActiveSpreadsheet().getId() : "";
const DRIVE_FOLDER_NAME = "PULSE_MEDIA_STORAGE";
const MAX_BUFFER_CAPACITY = 50;

function doGet(e) {
  return handleRequest(e, "GET");
}

function doPost(e) {
  return handleRequest(e, "POST");
}

function handleRequest(e, method) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    initializeStorageInfrastructure();

    let params = {};
    if (method === "GET") {
      params = e.parameter || {};
    } else {
      if (e.postData && e.postData.contents) {
        params = JSON.parse(e.postData.contents);
      }
    }

    const action = params.action;
    let responseData = { success: false, error: "Action not recognized" };

    switch (action) {
      case "getSalt":
        responseData = handleGetSalt(params.username);
        break;
      case "registerUser":
        responseData = handleRegisterUser(params);
        break;
      case "authenticateUser":
        responseData = handleAuthenticateUser(params);
        break;
      case "uploadMedia":
      case "uploadAudioTrack":
        responseData = handleMediaUpload(params);
        break;
      case "getLiveManifest":
        responseData = handleGetLiveManifest(params.listenerId, params.etag);
        break;
      case "streamMedia":
      case "streamAudio":
        responseData = handleStreamMedia(params.fileId);
        break;
      default:
        responseData = { success: false, error: "Invalid action parameter" };
    }

    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function initializeStorageInfrastructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let userSheet = ss.getSheetByName("USERS");
  if (!userSheet) {
    userSheet = ss.insertSheet("USERS");
    userSheet.appendRow(["Username", "FullName", "Salt", "PasswordHash", "AuthToken", "TokenExpiry", "CreatedAt"]);
  }

  let mediaSheet = ss.getSheetByName("MEDIA_BUFFER");
  if (!mediaSheet) {
    mediaSheet = ss.insertSheet("MEDIA_BUFFER");
    mediaSheet.appendRow(["ContentId", "FileId", "Title", "FileName", "MediaType", "MimeType", "DurationSec", "SizeBytes", "UploaderUsername", "UploaderName", "Timestamp"]);
  }

  let listenerSheet = ss.getSheetByName("LISTENERS");
  if (!listenerSheet) {
    listenerSheet = ss.insertSheet("LISTENERS");
    listenerSheet.appendRow(["ListenerId", "LastPingTime"]);
  }

  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (!folders.hasNext()) {
    const folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
}

function getMediaFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  const folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder;
}

function handleGetSalt(username) {
  if (!username) return { success: false, error: "Missing username" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("USERS");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(username).toLowerCase()) {
      return { success: true, data: { salt: data[i][2] } };
    }
  }
  return { success: false, error: "User not found" };
}

function handleRegisterUser(params) {
  const { username, clientHash, fullName } = params;
  if (!username || !clientHash || !fullName) {
    return { success: false, error: "Missing registration payload" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("USERS");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(username).toLowerCase()) {
      return { success: false, error: "Username already taken" };
    }
  }

  const salt = Utilities.getUuid().substring(0, 16);
  const serverHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, clientHash + "::" + salt)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  const authToken = Utilities.getUuid();
  const expiry = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();

  sheet.appendRow([username.toLowerCase(), fullName, salt, serverHash, authToken, expiry, new Date().toISOString()]);

  return {
    success: true,
    data: { username: username.toLowerCase(), fullName: fullName, token: authToken }
  };
}

function handleAuthenticateUser(params) {
  const { username, clientHash } = params;
  if (!username || !clientHash) return { success: false, error: "Missing credentials" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("USERS");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(username).toLowerCase()) {
      const salt = data[i][2];
      const storedHash = data[i][3];
      const computed = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, clientHash + "::" + salt)
        .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

      if (computed === storedHash) {
        const token = Utilities.getUuid();
        const expiry = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();
        sheet.getRange(i + 1, 5, 1, 2).setValues([[token, expiry]]);

        return {
          success: true,
          data: { username: data[i][0], fullName: data[i][1], token: token }
        };
      } else {
        return { success: false, error: "Invalid credentials" };
      }
    }
  }
  return { success: false, error: "Contributor not found" };
}

function verifyAuth(username, token) {
  if (!username || !token) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("USERS");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(username).toLowerCase()) {
      if (data[i][4] === token) {
        const expiry = new Date(data[i][5]).getTime();
        if (Date.now() < expiry) {
          return { username: data[i][0], fullName: data[i][1] };
        }
      }
    }
  }
  return null;
}

function handleMediaUpload(params) {
  const { auth, metadata, fileData } = params;
  const user = verifyAuth(auth?.username, auth?.token);
  if (!user) return { success: false, error: "Unauthorized access token" };

  if (!fileData || !metadata) return { success: false, error: "Incomplete file payload" };

  const rawBytes = Utilities.base64Decode(fileData);
  const mime = metadata.mimeType || "application/octet-stream";
  const mediaType = mime.startsWith("video/") || (metadata.mediaType === "video") ? "video" : "audio";
  const blob = Utilities.newBlob(rawBytes, mime, metadata.fileName);

  const folder = getMediaFolder();
  const driveFile = folder.createFile(blob);
  driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = driveFile.getId();
  const contentId = "med_" + Utilities.getUuid().substring(0, 12);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MEDIA_BUFFER");

  sheet.appendRow([
    contentId,
    fileId,
    metadata.title || "Untitled Transmission",
    metadata.fileName,
    mediaType,
    mime,
    metadata.durationSeconds || 0,
    metadata.sizeBytes || rawBytes.length,
    user.username,
    user.fullName,
    new Date().toISOString()
  ]);

  enforceFifoCapacity(sheet);

  return { success: true, data: { contentId: contentId, fileId: fileId, mediaType: mediaType } };
}

function enforceFifoCapacity(sheet) {
  const totalRows = sheet.getLastRow();
  const headerCount = 1;
  const entryCount = totalRows - headerCount;

  if (entryCount > MAX_BUFFER_CAPACITY) {
    const excess = entryCount - MAX_BUFFER_CAPACITY;
    for (let i = 0; i < excess; i++) {
      const fileIdToDelete = sheet.getRange(2, 2).getValue();
      try {
        DriveApp.getFileById(fileIdToDelete).setTrashed(true);
      } catch (e) {}
      sheet.deleteRow(2);
    }
  }
}

/**
 * On-demand lazy reconciliation: syncs Google Drive additions/deletions
 * into the MEDIA_BUFFER sheet automatically without background triggers.
 */
function reconcileDriveFolderWithSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MEDIA_BUFFER");
  if (!sheet) return;

  const folder = getMediaFolder();

  // 1. Map all active (non-trashed) files currently in the folder
  const activeFilesInFolder = new Map();
  const folderFiles = folder.getFiles();
  while (folderFiles.hasNext()) {
    const f = folderFiles.next();
    if (!f.isTrashed()) {
      activeFilesInFolder.set(f.getId(), f);
    }
  }

  // 2. Scan sheet rows bottom-to-top to prune removed or trashed items
  const data = sheet.getDataRange().getValues();
  const existingSheetFileIds = new Set();

  for (let i = data.length - 1; i >= 1; i--) {
    const fileId = String(data[i][1]).trim();
    let fileExistsAndActive = false;

    if (activeFilesInFolder.has(fileId)) {
      fileExistsAndActive = true;
    } else {
      try {
        const driveFile = DriveApp.getFileById(fileId);
        if (!driveFile.isTrashed()) {
          fileExistsAndActive = true;
        }
      } catch (e) {
        fileExistsAndActive = false;
      }
    }

    if (!fileExistsAndActive) {
      sheet.deleteRow(i + 1);
    } else {
      existingSheetFileIds.add(fileId);
    }
  }

  // 3. Register any new files dropped into the folder
  activeFilesInFolder.forEach((file, fileId) => {
    if (!existingSheetFileIds.has(fileId)) {
      const mime = file.getMimeType();
      const isVideo = mime.startsWith("video/");
      const isAudio = mime.startsWith("audio/");

      if (isVideo || isAudio) {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        sheet.appendRow([
          "med_" + Utilities.getUuid().substring(0, 10),
          fileId,
          file.getName().replace(/\.[^/.]+$/, ""),
          file.getName(),
          isVideo ? "video" : "audio",
          mime,
          0,
          file.getSize(),
          "admin",
          "Station Manager",
          new Date().toISOString()
        ]);
      }
    }
  });

  enforceFifoCapacity(sheet);
}

function handleGetLiveManifest(listenerId, incomingEtag) {
  // Reconcile changes directly when the station manifest is queried
  reconcileDriveFolderWithSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Track active listeners
  if (listenerId) {
    const lSheet = ss.getSheetByName("LISTENERS");
    const lData = lSheet.getDataRange().getValues();
    const now = Date.now();
    let found = false;

    for (let i = 1; i < lData.length; i++) {
      if (lData[i][0] === listenerId) {
        lSheet.getRange(i + 1, 2).setValue(now);
        found = true;
        break;
      }
    }
    if (!found) lSheet.appendRow([listenerId, now]);
  }

  // Count listeners pinged within last 45 seconds
  const lSheet = ss.getSheetByName("LISTENERS");
  const lData = lSheet.getDataRange().getValues();
  const cutoff = Date.now() - 45000;
  let activeCount = 0;
  for (let i = 1; i < lData.length; i++) {
    if (Number(lData[i][1]) > cutoff) activeCount++;
  }

  const mSheet = ss.getSheetByName("MEDIA_BUFFER");
  const data = mSheet.getDataRange().getValues();
  const manifest = [];

  for (let i = 1; i < data.length; i++) {
    manifest.push({
      contentId: data[i][0],
      fileId: data[i][1],
      title: data[i][2],
      fileName: data[i][3],
      mediaType: data[i][4] || "audio",
      mimeType: data[i][5],
      durationSeconds: Number(data[i][6]) || 0,
      sizeBytes: Number(data[i][7]) || 0,
      uploaderUsername: data[i][8],
      uploaderName: data[i][9],
      timestamp: data[i][10]
    });
  }

  const serverEtag = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    JSON.stringify(manifest) + "_" + activeCount
  ).map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

  if (incomingEtag && incomingEtag === serverEtag) {
    return { success: true, data: { notModified: true, activeListeners: activeCount } };
  }

  return {
    success: true,
    data: {
      manifest: manifest,
      etag: serverEtag,
      activeListeners: activeCount
    }
  };
}

function handleStreamMedia(fileId) {
  if (!fileId) return { success: false, error: "Missing fileId" };
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return {
      success: true,
      mimeType: blob.getContentType(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}