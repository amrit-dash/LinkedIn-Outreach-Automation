function doPost(e) {
  const props = PropertiesService.getScriptProperties();
  const isEnabled = props.getProperty('WEBHOOK_MONITORING_ENABLED') || "DISABLED";
  
  if (String(isEnabled).toUpperCase() !== 'ENABLED') {
    return ContentService.createTextOutput(JSON.stringify({ status: "ignored", reason: "monitoring disabled in Apps Script" })).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", reason: "no payload" })).setMimeType(ContentService.MimeType.JSON);
  }
  
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", reason: "invalid json" })).setMimeType(ContentService.MimeType.JSON);
  }
  
  const event = data.event;
  if (event !== 'message_received' && event !== 'new_relation') {
    return ContentService.createTextOutput(JSON.stringify({ status: "ignored", reason: "unsupported event type" })).setMimeType(ContentService.MimeType.JSON);
  }
  
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", reason: "could not obtain lock" })).setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName("Database");
    if (!dbSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", reason: "Database sheet not found" })).setMimeType(ContentService.MimeType.JSON);
    
    const lastDbRow = dbSheet.getLastRow();
    if (lastDbRow < 2) return ContentService.createTextOutput(JSON.stringify({ status: "ignored", reason: "Database is empty" })).setMimeType(ContentService.MimeType.JSON);
    
    const dbData = dbSheet.getRange(2, 1, lastDbRow - 1, Math.max(26, dbSheet.getLastColumn())).getValues();
    const nowStr = new Date().toLocaleDateString('en-US', { timeZone: 'Asia/Kolkata' }) + " " + new Date().toLocaleTimeString('en-US', { timeZone: 'Asia/Kolkata' });
    
    let updated = false;
    
    if (event === 'message_received') {
      const accountId = data.account_id;
      const providerId = data.sender && data.sender.attendee_provider_id;
      const messageText = data.message || "";
      let timestamp = data.timestamp; // ISO string
      let formattedTime = "";
      
      if (timestamp) {
        try {
           const d = new Date(timestamp);
           formattedTime = d.toLocaleDateString('en-US', { timeZone: 'Asia/Kolkata' }) + " " + d.toLocaleTimeString('en-US', { timeZone: 'Asia/Kolkata' });
        } catch(ex) {
           formattedTime = timestamp;
        }
      } else {
        formattedTime = nowStr;
      }
      
      // Ensure we don't log messages sent by the account owner as replies
      const isSelf = data.account_info && data.account_info.user_id === providerId;
      
      if (!isSelf && providerId) {
        for (let i = dbData.length - 1; i >= 0; i--) {
          if (dbData[i][10] === accountId && dbData[i][11] === providerId) {
            // Found the matching prospect
            // Column 23 = reply_received, Column 24 = reply_text, Column 25 = reply_time
            dbSheet.getRange(i + 2, 23).setValue(true);
            dbSheet.getRange(i + 2, 24).setValue(messageText);
            dbSheet.getRange(i + 2, 25).setValue(formattedTime);
            updated = true;
            break;
          }
        }
      } else {
         return ContentService.createTextOutput(JSON.stringify({ status: "ignored", reason: "message sent by self or missing providerId" })).setMimeType(ContentService.MimeType.JSON);
      }
      
    } else if (event === 'new_relation') {
      const accountId = data.account_id;
      const providerId = data.user_provider_id;
      
      if (providerId) {
        for (let i = dbData.length - 1; i >= 0; i--) {
          if (dbData[i][10] === accountId && dbData[i][11] === providerId) {
            // Found the matching prospect
            // Column 13 = connection_request_status, Column 15 = connection_accepted, Column 16 = connection_accepted_time
            dbSheet.getRange(i + 2, 13).setValue("Accepted");
            dbSheet.getRange(i + 2, 15).setValue(true);
            dbSheet.getRange(i + 2, 16).setValue(nowStr);
            updated = true;
            break;
          }
        }
      }
    }
    
    if (updated) {
       SpreadsheetApp.flush();
       // Call the stats updater to roll up these new replies/connections into the Campaigns and Accounts sheets
       updateGlobalStats();
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", updated: updated })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: err.message })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}