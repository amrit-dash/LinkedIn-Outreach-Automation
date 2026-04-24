function processCampaignsWorker() {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("Could not obtain lock for processCampaignsWorker.");
    return;
  }
  
  // Track continuous execution
  let executionLoops = 0;
  const maxLoops = 2; // Run twice within a single 15-minute trigger
  const globalStartTime = Date.now();
  
  try {
    while (executionLoops < maxLoops) {
      if (Date.now() - globalStartTime > 240000) { // 4 minutes overall safety
        Logger.log("Overall worker execution nearing limit. Stopping early.");
        break;
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dbSheet = ss.getSheetByName("Database");
      const campaignsSheet = ss.getSheetByName("Campaigns");
      const invSheet = ss.getSheetByName("Invitations");
      
      if (!dbSheet || !campaignsSheet) break;
  
      let creds;
      try {
        creds = getCredentials();
      } catch (e) {
        Logger.log("Error loading credentials: " + e.message);
        break;
      }

      const lastDbRow = dbSheet.getLastRow();
      if (lastDbRow < 2) break;
  
      const lastDbCol = Math.max(26, dbSheet.getLastColumn());
      const dbRange = dbSheet.getRange(2, 1, lastDbRow - 1, lastDbCol);
      const dbData = dbRange.getValues();

      const lastCampRow = campaignsSheet.getLastRow();
      if (lastCampRow < 2) break;
      const campaignsData = campaignsSheet.getRange(2, 1, lastCampRow - 1, campaignsSheet.getLastColumn()).getValues();

      let invData = [];
      let invRange = null;
      if (invSheet) {
        const lastInvRow = invSheet.getLastRow();
        if (lastInvRow >= 2) {
          invRange = invSheet.getRange(2, 1, lastInvRow - 1, invSheet.getLastColumn());
          invData = invRange.getValues();
        }
      }
      
      const now = new Date();
      const nowTime = now.getTime();
      
      // O(1) Lookup Maps for Maximum Efficiency
      const campaignMap = {};
      for (let i = 0; i < campaignsData.length; i++) {
        campaignMap[campaignsData[i][0]] = campaignsData[i];
      }
      
      const invMap = {};
      for (let j = 0; j < invData.length; j++) {
        if (invData[j][3] === "Sent") {
          const key = invData[j][0] + "_" + invData[j][1]; // accountId_providerId
          invMap[key] = {
            invId: invData[j][2],
            rowIndex: j
          };
        }
      }

      const requests = [];
      const actions = [];
      let dbUpdated = false;
  
      for (let i = 0; i < dbData.length; i++) {
        const row = dbData[i];
        
        // Check if reply received, skip processing for messaging/uninviting
        let replyTextCheck = row[23];
        let replyTimeCheck = row[24];
        
        const hasReplyBoxChecked = (row[22] === true || String(row[22]).toUpperCase() === "TRUE");
        
        let hasReplyText = false;
        if (replyTextCheck !== null && replyTextCheck !== undefined && replyTextCheck !== "") {
          let strText = String(replyTextCheck).trim().toUpperCase();
          if (strText !== "" && strText !== "FALSE" && strText !== "NULL" && strText !== "UNDEFINED") {
            hasReplyText = true;
          }
        }

        let hasReplyTime = false;
        if (replyTimeCheck !== null && replyTimeCheck !== undefined && replyTimeCheck !== "") {
          let strTime = String(replyTimeCheck).trim().toUpperCase();
          if (strTime !== "" && strTime !== "FALSE" && strTime !== "NULL" && strTime !== "UNDEFINED") {
            hasReplyTime = true;
          }
        }
        
        if (hasReplyBoxChecked || hasReplyText || hasReplyTime) {
          continue; // Reply received
        }
        
        const campId = row[0];
        const accountId = row[10];
        const providerId = row[11];
        const connReqStatus = row[12];
        const connReqTime = row[13] ? new Date(row[13]) : null;
        const connAccepted = row[14];
        
        // Find campaign details directly using Hash Map
        const campaign = campaignMap[campId];
        if (!campaign) continue;
        
        // Case 1: Uninvite after 7 days
        if (connReqStatus === "Sent" && connAccepted !== true) {
          if (connReqTime) {
            const daysPassed = (nowTime - connReqTime.getTime()) / (1000 * 3600 * 24);
            if (daysPassed > 7) {
              const invKey = accountId + "_" + providerId;
              const invInfo = invMap[invKey];
              let invId = invInfo ? invInfo.invId : null;
              let invRowIndex = invInfo ? invInfo.rowIndex : -1;
              
              if (!invId) {
                 invId = findInvitationId(creds, accountId, providerId);
                 if (invId && invSheet) {
                    invSheet.appendRow([accountId, providerId, invId, "Sent", new Date()]);
                 }
              }
              
              if (invId) {
                requests.push({
                  url: `${creds.baseUrl}/users/invite/sent/${invId}?account_id=${accountId}`,
                  method: "DELETE",
                  headers: { "X-API-KEY": creds.apiKey, "Accept": "application/json" },
                  muteHttpExceptions: true
                });
                actions.push({ type: 'uninvite', dbIndex: i, invRowIndex: invRowIndex, providerId });
              } else {
                 // Invitation ID not found, just mark failed in DB so it doesn't keep checking
                 row[12] = "Failed";
                 row[25] = `[${now.toISOString()}] 7 days passed, could not uninvite (missing ID).`;
                 dbUpdated = true;
              }
            }
          }
        }
        
        // Messaging Logic
        if (connAccepted === true) {
           const msg1Status = row[16];
           const msg2Status = row[18];
           const msg3Status = row[20];
           
           const firstName = String(row[3] || "").trim();
           const acceptedTime = row[15] ? new Date(row[15]) : null;
           
           if (msg1Status === "Pending") {
              let msg1Text = campaign[5]; // Column F
              if (msg1Text && String(msg1Text).trim() !== "") {
                 msg1Text = String(msg1Text).replace(/\$name/g, firstName);
                 requests.push({
                   url: `${creds.baseUrl}/chats`,
                   method: "POST",
                   headers: { "X-API-KEY": creds.apiKey, "Accept": "application/json", "Content-Type": "application/json" },
                   payload: JSON.stringify({ account_id: accountId, text: msg1Text, attendees_ids: [providerId] }),
                   muteHttpExceptions: true
                 });
                 actions.push({ type: 'msg1', dbIndex: i });
              } else {
                 row[16] = "Skipped";
                 row[17] = new Date();
                 dbSheet.getRange(i + 2, 17, 1, 2).setValues([["Skipped", row[17]]]);
              }
           } else if ((msg1Status === "Sent" || msg1Status === "Skipped") && msg2Status === "Pending") {
              const msg1Time = row[17] ? new Date(row[17]) : acceptedTime;
              const delay2Hours = parseFloat(campaign[8]) || 0; // Column I
              
              if (msg1Time) {
                const hoursPassed = (nowTime - msg1Time.getTime()) / (1000 * 3600);
                if (hoursPassed >= delay2Hours) {
                   let msg2Text = campaign[6]; // Column G
                   if (msg2Text && String(msg2Text).trim() !== "") {
                      msg2Text = String(msg2Text).replace(/\$name/g, firstName);
                      requests.push({
                        url: `${creds.baseUrl}/chats`,
                        method: "POST",
                        headers: { "X-API-KEY": creds.apiKey, "Accept": "application/json", "Content-Type": "application/json" },
                        payload: JSON.stringify({ account_id: accountId, text: msg2Text, attendees_ids: [providerId] }),
                        muteHttpExceptions: true
                      });
                      actions.push({ type: 'msg2', dbIndex: i });
                   } else {
                      row[18] = "Skipped";
                      row[19] = new Date();
                      dbSheet.getRange(i + 2, 19, 1, 2).setValues([["Skipped", row[19]]]);
                   }
                }
              }
           } else if ((msg2Status === "Sent" || msg2Status === "Skipped") && msg3Status === "Pending") {
              const msg2Time = row[19] ? new Date(row[19]) : (row[17] ? new Date(row[17]) : acceptedTime);
              const delay3Hours = parseFloat(campaign[9]) || 0; // Column J
              
              if (msg2Time) {
                const hoursPassed = (nowTime - msg2Time.getTime()) / (1000 * 3600);
                if (hoursPassed >= delay3Hours) {
                   let msg3Text = campaign[7]; // Column H
                   if (msg3Text && String(msg3Text).trim() !== "") {
                      msg3Text = String(msg3Text).replace(/\$name/g, firstName);
                      requests.push({
                        url: `${creds.baseUrl}/chats`,
                        method: "POST",
                        headers: { "X-API-KEY": creds.apiKey, "Accept": "application/json", "Content-Type": "application/json" },
                        payload: JSON.stringify({ account_id: accountId, text: msg3Text, attendees_ids: [providerId] }),
                        muteHttpExceptions: true
                      });
                      actions.push({ type: 'msg3', dbIndex: i });
                   } else {
                      row[20] = "Skipped";
                      row[21] = new Date();
                      dbSheet.getRange(i + 2, 21, 1, 2).setValues([["Skipped", row[21]]]);
                   }
                }
              }
           }
        }
      }
  
    if (requests.length === 0) {
      if (dbUpdated) {
        dbSheet.getRange(2, 1, dbData.length, dbData[0].length).setValues(dbData);
      }
      executionLoops++;
      if (executionLoops < maxLoops) {
         Utilities.sleep(15000); // 15 second delay before restarting loop
         continue; 
      } else {
         break;
      }
    }
  
  // Process API requests sequentially to simulate human behavior and avoid flagging
  let invUpdated = false;
  const startTime = Date.now();
  
    for (let k = 0; k < requests.length; k++) {
       if (Date.now() - startTime > 180000) { // 3 minutes inner safety
         Logger.log("Nearing execution limit in worker loop. Stopping early.");
         break;
       }

     const req = requests[k];
     const action = actions[k];
     const row = dbData[action.dbIndex];
     const nowStr = new Date().toISOString();
     
     let response;
     try {
       response = fetchWithRetry(req.url, req);
     } catch(e) {
       Logger.log(`Worker request failed: ${e.message}`);
       continue;
     }
     
     let code = response.getResponseCode();
     let respText = "";
     try {
       respText = response.getContentText();
     } catch(e) {
       respText = "Parse error: " + e.message;
     }
     
     if (action.type === 'uninvite') {
        if (code === 200 || code === 204) {
            row[12] = "Failed";
            row[25] = `[${nowStr}] 7 days passed, so it is uninvited.`;
            dbSheet.getRange(action.dbIndex + 2, 13).setValue("Failed");
            dbSheet.getRange(action.dbIndex + 2, 26).setValue(`[${nowStr}] 7 days passed, so it is uninvited.`);
            if (action.invRowIndex >= 0) invData[action.invRowIndex][3] = "Uninvited";
            invUpdated = true;
        } else if (code === 429 || code >= 500) {
            Logger.log(`Worker uninvite: Rate limit or server error (${code}). Retrying next run.`);
        } else {
             if (respText.includes("invalid_invitation_id") || respText.includes("Resource not found")) {
                 row[12] = "Failed";
                 row[25] = `[${nowStr}] 7 days passed, uninvite failed (already gone).`;
                 dbSheet.getRange(action.dbIndex + 2, 13).setValue("Failed");
                 dbSheet.getRange(action.dbIndex + 2, 26).setValue(`[${nowStr}] 7 days passed, uninvite failed (already gone).`);
                 if (action.invRowIndex >= 0) invData[action.invRowIndex][3] = "Uninvited";
                 invUpdated = true;
             } else {
                 Logger.log(`Failed to uninvite ${action.providerId}: ${respText}`);
             }
        }
        Utilities.sleep(1000); // Small delay for uninvites
     } else if (action.type === 'msg1' || action.type === 'msg2' || action.type === 'msg3') {
        const statusCol = action.type === 'msg1' ? 16 : (action.type === 'msg2' ? 18 : 20);
        const timeCol = action.type === 'msg1' ? 17 : (action.type === 'msg2' ? 19 : 21);
        
        if (code === 201 || code === 200) {
             row[statusCol] = "Sent";
             row[timeCol] = new Date();
             dbSheet.getRange(action.dbIndex + 2, statusCol + 1, 1, 2).setValues([["Sent", row[timeCol]]]);
             
             // Random delay 3 to 7 seconds to prevent flagging and mimic human behavior
             const delayMs = Math.floor(Math.random() * (7000 - 3000 + 1) + 3000);
             Utilities.sleep(delayMs);
        } else if (code === 429 || code >= 500) {
             Logger.log(`Worker ${action.type}: Rate limit or server error (${code}). Retrying next run.`);
        } else {
             row[statusCol] = "Failed";
             row[25] = `[${nowStr}] ${action.type.toUpperCase()} Error: ${respText}`.substring(0, 500);
             dbSheet.getRange(action.dbIndex + 2, statusCol + 1).setValue("Failed");
             dbSheet.getRange(action.dbIndex + 2, 26).setValue(row[25]);
        }
     }
  }
  
  // Save state after batch to prevent data loss
  if (invUpdated && invRange && invData.length > 0) {
     invRange.setValues(invData);
     invUpdated = false;
  }
  SpreadsheetApp.flush();
  
    // Sync all global stats after processing requests in this loop
    updateGlobalStats();
    
    executionLoops++;
    if (executionLoops < maxLoops) {
      // Pause briefly before running the next cycle within the same execution
      Utilities.sleep(15000); 
    }
  } // End of while loop
  
  } finally {
    lock.releaseLock();
  }
}

function updateGlobalStats() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("Could not obtain lock for updateGlobalStats.");
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  const dbSheet = ss.getSheetByName("Database");
  const campaignsSheet = ss.getSheetByName("Campaigns");
  const accountsSheet = ss.getSheetByName("Accounts");
  const credSheet = ss.getSheetByName("Credentials");

  if (credSheet) {
    const props = PropertiesService.getScriptProperties();
    let propStatus = props.getProperty('WEBHOOK_MONITORING_ENABLED') || 'DISABLED';
    
    let statusValue = credSheet.getRange(4, 2).getValue();
    if (String(statusValue).toUpperCase() !== String(propStatus).toUpperCase()) {
      credSheet.getRange(4, 1).setValue("appsscript_webhook_status");
      credSheet.getRange(4, 2).setValue(propStatus);
    }
  }

  if (!dbSheet || !campaignsSheet || !accountsSheet) return;

  const dbData = dbSheet.getDataRange().getValues();
  const campData = campaignsSheet.getDataRange().getValues();
  const accData = accountsSheet.getDataRange().getValues();

  const campStats = {};
  const accStats = {};

  const todayStr = new Date().toLocaleDateString('en-US', { timeZone: 'Asia/Kolkata' });

  for (let i = 1; i < accData.length; i++) {
    const accId = accData[i][1];
    if (accId) {
      accStats[accId] = {
        connectionsToday: 0,
        dmsToday: 0
      };
    }
  }

  for (let i = 1; i < campData.length; i++) {
    const campId = campData[i][0];
    if (campId) {
      campStats[campId] = {
        connectionsSent: 0,
        connectionsAccepted: 0,
        messagesSent: 0,
        repliesReceived: 0,
        targetProspects: parseInt(campData[i][2]) || 0
      };
    }
  }

  for (let i = 1; i < dbData.length; i++) {
    const row = dbData[i];
    const campId = row[0];
    const accId = row[10];

    // Campaign stats
    if (campId && campStats[campId]) {
      let st = row[12]; // Connection Request Status
      if (st === "Sent" || st === "Accepted") campStats[campId].connectionsSent++;
      if (row[14] === true || String(row[14]).toUpperCase() === "TRUE") campStats[campId].connectionsAccepted++;
      
      if (row[16] === "Sent") campStats[campId].messagesSent++;
      if (row[18] === "Sent") campStats[campId].messagesSent++;
      if (row[20] === "Sent") campStats[campId].messagesSent++;
      
      let replyText = row[23];
      let replyTime = row[24];
      
      const hasReplyBoxChecked = (row[22] === true || String(row[22]).toUpperCase() === "TRUE");
      
      let hasReplyText = false;
      if (replyText !== null && replyText !== undefined && replyText !== "") {
        let strText = String(replyText).trim().toUpperCase();
        if (strText !== "" && strText !== "FALSE" && strText !== "NULL" && strText !== "UNDEFINED") {
          hasReplyText = true;
        }
      }

      let hasReplyTime = false;
      if (replyTime !== null && replyTime !== undefined && replyTime !== "") {
        let strTime = String(replyTime).trim().toUpperCase();
        if (strTime !== "" && strTime !== "FALSE" && strTime !== "NULL" && strTime !== "UNDEFINED") {
          hasReplyTime = true;
        }
      }
      
      if (hasReplyBoxChecked || hasReplyText || hasReplyTime) {
        campStats[campId].repliesReceived++;
      }
    }

    // Account stats (Today)
    if (accId && accStats[accId]) {
      if (row[13] && new Date(row[13]).toLocaleDateString('en-US', { timeZone: 'Asia/Kolkata' }) === todayStr) {
        accStats[accId].connectionsToday++;
      }
      
      if (row[17] && new Date(row[17]).toLocaleDateString('en-US', { timeZone: 'Asia/Kolkata' }) === todayStr) accStats[accId].dmsToday++;
      if (row[19] && new Date(row[19]).toLocaleDateString('en-US', { timeZone: 'Asia/Kolkata' }) === todayStr) accStats[accId].dmsToday++;
      if (row[21] && new Date(row[21]).toLocaleDateString('en-US', { timeZone: 'Asia/Kolkata' }) === todayStr) accStats[accId].dmsToday++;
    }
  }

  let campUpdated = false;
  for (let i = 1; i < campData.length; i++) {
    const campId = campData[i][0];
    const stats = campStats[campId];
    if (stats) {
      if (campData[i][10] !== stats.connectionsSent ||
          campData[i][11] !== stats.connectionsAccepted ||
          campData[i][12] !== stats.messagesSent ||
          campData[i][13] !== stats.repliesReceived) {
        
        campData[i][10] = stats.connectionsSent;
        campData[i][11] = stats.connectionsAccepted;
        campData[i][12] = stats.messagesSent;
        campData[i][13] = stats.repliesReceived;
        campUpdated = true;
      }
      
      if (stats.targetProspects > 0 && stats.repliesReceived >= stats.targetProspects && String(campData[i][3]).trim() !== "Completed") {
        campData[i][3] = "Completed";
        campUpdated = true;
      }
    }
  }

  if (campUpdated) {
    campaignsSheet.getRange(1, 1, campData.length, campData[0].length).setValues(campData);
  }

  let accUpdated = false;
  for (let i = 1; i < accData.length; i++) {
    const accId = accData[i][1];
    const stats = accStats[accId];
    if (stats) {
      if (accData[i][5] !== stats.connectionsToday || accData[i][7] !== stats.dmsToday) {
        accData[i][5] = stats.connectionsToday;
        accData[i][7] = stats.dmsToday;
        accUpdated = true;
      }
    }
  }

  if (accUpdated) {
    accountsSheet.getRange(1, 1, accData.length, accData[0].length).setValues(accData);
  }
  } finally {
    lock.releaseLock();
  }
}

function syncInvitationStatuses() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("Could not obtain lock for syncInvitationStatuses.");
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    
    const dbSheet = ss.getSheetByName("Database");
    const invSheet = ss.getSheetByName("Invitations");
    
    if (!dbSheet || !invSheet) return;
    
    const dbData = dbSheet.getDataRange().getValues();
    const invData = invSheet.getDataRange().getValues();
    
    if (dbData.length < 2 || invData.length < 2) return;
    
    // Map Database statuses
    const dbStatusMap = {};
    for (let i = 1; i < dbData.length; i++) {
      const accountId = dbData[i][10];
      const providerId = dbData[i][11];
      const status = dbData[i][12];
      const isAccepted = dbData[i][14];
      
      if (accountId && providerId) {
        if (isAccepted === true || String(isAccepted).toUpperCase() === "TRUE") {
          dbStatusMap[`${accountId}_${providerId}`] = "Accepted";
        } else {
          dbStatusMap[`${accountId}_${providerId}`] = status;
        }
      }
    }
    
    let invUpdated = false;
    for (let i = 1; i < invData.length; i++) {
      const accountId = invData[i][0];
      const providerId = invData[i][1];
      const currentStatus = invData[i][3];
      
      if (accountId && providerId) {
        const dbStatus = dbStatusMap[`${accountId}_${providerId}`];
        if (dbStatus && (dbStatus === "Sent" || dbStatus === "Accepted")) {
          if (currentStatus !== dbStatus) {
            invData[i][3] = dbStatus;
            invUpdated = true;
          }
        }
      }
    }
    
    if (invUpdated) {
      invSheet.getRange(1, 1, invData.length, invData[0].length).setValues(invData);
    }
  } finally {
    lock.releaseLock();
  }
}

function processStatsWorker() {
  updateGlobalStats();
  syncInvitationStatuses();
}

function startStatsWorker() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processStatsWorker') {
      SpreadsheetApp.getUi().alert("Stats Worker is already running.");
      return;
    }
  }
  
  ScriptApp.newTrigger('processStatsWorker')
    .timeBased()
    .everyMinutes(10)
    .create();
    
  SpreadsheetApp.getUi().alert("Stats Worker started! It will run every 10 minutes to sync campaign & account stats in the background.");
}

function stopStatsWorker() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processStatsWorker') {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }
  SpreadsheetApp.getUi().alert(`Stats Worker stopped. (Deleted ${count} triggers)`);
}

function startBackgroundWorker() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processCampaignsWorker') {
      SpreadsheetApp.getUi().alert("Worker is already running.");
      return;
    }
  }
  
  ScriptApp.newTrigger('processCampaignsWorker')
    .timeBased()
    .everyMinutes(12)
    .create();
    
  SpreadsheetApp.getUi().alert("Background Worker started! It will run every 15 minutes to process messages and 7-day uninvites.");
}

function stopBackgroundWorker() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processCampaignsWorker') {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }
  SpreadsheetApp.getUi().alert(`Background Worker stopped. (Deleted ${count} triggers)`);
}
