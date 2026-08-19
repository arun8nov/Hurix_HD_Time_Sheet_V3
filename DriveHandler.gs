/**
 * HURIX DIGITAL - CENTRALIZED GOOGLE DRIVE ATTACHMENT HANDLER (DriveHandler.gs)
 */

// Shared Admin Folder ID for uploaded leave/permission proofs
var MASTER_FOLDER_ID = '1pvJQ9q6FuqRObkebq8GizkWOQ_QbX2RH';

/**
 * Uploads a base64 encoded proof document/attachment to Google Drive
 */
function uploadLeaveProofAttachment(base64Data, fileName, mimeType) {
  try {
    if (!base64Data) {
      return { success: false, message: 'No file data provided.' };
    }

    var folder;
    try {
      folder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    } catch (e) {
      // Fallback to root folder if specific folder ID is unreachable
      folder = DriveApp.getRootFolder();
    }

    // Process base64 data
    var rawData = base64Data.indexOf(',') !== -1 ? base64Data.split(',')[1] : base64Data;
    var data = Utilities.base64Decode(rawData);
    var blob = Utilities.newBlob(data, mimeType || 'application/octet-stream', fileName || 'Leave_Proof_' + new Date().getTime());
    
    var file = folder.createFile(blob);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      // Ignore domain permission restriction errors
    }

    return {
      success: true,
      fileUrl: file.getUrl(),
      fileId: file.getId()
    };
  } catch (error) {
    return { success: false, message: 'Drive Upload Error: ' + error.toString() };
  }
}
