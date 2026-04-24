
function findInvitationId(creds, accountId, providerId) {
  let cursor = null;
  do {
    let url = `${creds.baseUrl}/users/invite/sent?account_id=${accountId}&limit=100`;
    if (cursor) {
      url += `&cursor=${cursor}`;
    }
    const options = {
      "method": "GET",
      "headers": {
        "X-API-KEY": creds.apiKey,
        "Accept": "application/json"
      },
      "muteHttpExceptions": true
    };
    try {
      const response = fetchWithRetry(url, options);
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        const items = data.items || data || [];
        cursor = data.cursor || null;
        
        for (let i = 0; i < items.length; i++) {
          if (items[i].provider_id === providerId || items[i].attendee_provider_id === providerId) {
            return items[i].id || items[i].invitation_id;
          }
        }
        
        if (items.length === 0) break;
      } else {
        break;
      }
    } catch(e) {
      break;
    }
  } while (cursor);
  return null;
}

function createDatabaseEntries() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const campaignsSheet = ss.getSheetByName("Campaigns");
  const prospectsSheet = ss.getSheetByName("Prospects");
  const dbSheet = ss.getSheetByName("Database");
  
  const lastCampRow = campaignsSheet.getLastRow();
  if (lastCampRow < 2) {
    ui.alert("No campaigns found. Create a campaign first.");
    return;
  }
  const campaignsData = campaignsSheet.getRange(2, 1, lastCampRow - 1, campaignsSheet.getLastColumn()).getValues();
  
  const notStartedCampaigns = campaignsData.map((row, i) => ({ row, rowIndex: i + 2 })).filter(item => item.row[3] === "Not Started");
  
  if (notStartedCampaigns.length === 0) {
    ui.alert("No campaigns in 'Not Started' status found.");
    return;
  }

  let selectedListIndex;
  
  if (notStartedCampaigns.length === 1) {
    selectedListIndex = 0;
  } else {
    const campaignListStr = notStartedCampaigns.map((item, i) => `${i+1}. ${item.row[1]} (ID: ${item.row[0]}) - Target Prospects: ${item.row[2]}`).join('\n');
    const campResp = ui.prompt("Select Campaign to Start", `Enter the number (1-${notStartedCampaigns.length}) of the campaign to start:\n\n${campaignListStr}`, ui.ButtonSet.OK_CANCEL);
    if (campResp.getSelectedButton() !== ui.Button.OK) return;
    
    selectedListIndex = parseInt(campResp.getResponseText()) - 1;
    if (isNaN(selectedListIndex) || selectedListIndex < 0 || selectedListIndex >= notStartedCampaigns.length) {
      ui.alert("Invalid selection.");
      return;
    }
  }
  
  const selectedItem = notStartedCampaigns[selectedListIndex];
  const selectedCampaignId = selectedItem.row[0];
  const selectedCampaignName = selectedItem.row[1];
  const targetProspects = selectedItem.row[2]; 
  const campaignRowIndex = selectedItem.rowIndex;
  
  const accountsSheet = ss.getSheetByName("Accounts");
  const lastAccRow = accountsSheet.getLastRow();
  if (lastAccRow < 2) {
    ui.alert("No accounts found. Sync accounts first.");
    return;
  }
  
  const allAccountsData = accountsSheet.getRange(2, 1, lastAccRow - 1, accountsSheet.getLastColumn()).getValues();
  
  const activeAccounts = allAccountsData.map((row, i) => ({ row, rowIndex: i + 2 }))
    .filter(item => {
      const st = String(item.row[4]).toLowerCase().trim();
      return st === "active" || st === "ok";
    });
    
  if (activeAccounts.length === 0) {
    ui.alert("No active accounts found. Please ensure your accounts are active and try syncing again.");
    return;
  }
  
  let sendingAccountId;
  
  if (activeAccounts.length === 1) {
    sendingAccountId = activeAccounts[0].row[1];
  } else {
    const accListStr = activeAccounts.map((item, i) => `${i+1}. ${item.row[0]} (ID: ${item.row[1]})`).join('\n');
    const accResp = ui.prompt("Select Sending Account", `Enter the number (1-${activeAccounts.length}) of the active account to send from:\n\n${accListStr}`, ui.ButtonSet.OK_CANCEL);
    if (accResp.getSelectedButton() !== ui.Button.OK) return;
    
    const selectedAccIndex = parseInt(accResp.getResponseText()) - 1;
    if (isNaN(selectedAccIndex) || selectedAccIndex < 0 || selectedAccIndex >= activeAccounts.length) {
      ui.alert("Invalid selection.");
      return;
    }
    sendingAccountId = activeAccounts[selectedAccIndex].row[1];
  }
  
  const lastProspRow = prospectsSheet.getLastRow();
  if (lastProspRow < 2) {
    ui.alert("No prospects found in Prospects sheet.");
    return;
  }
  
  const lastProspCol = Math.max(12, prospectsSheet.getLastColumn());
  const prospectsDataRange = prospectsSheet.getRange(2, 1, lastProspRow - 1, lastProspCol);
  const prospectsData = prospectsDataRange.getValues();
  
  let takeCount = prospectsData.length;
  if (String(targetProspects).toLowerCase() !== "all") {
    const n = parseInt(targetProspects);
    if (!isNaN(n) && n > 0 && n < takeCount) {
      takeCount = n;
    }
  }
  
  let creds;
  try {
    creds = getCredentials();
  } catch (e) {
    ui.alert(`Error reading credentials: ${e.message}`);
    return;
  }
  
  const startTime = Date.now();
  const takenProspects = prospectsData.slice(0, takeCount);
  
  // Bulk enrichment to avoid execution time limit
  const result = enrichProspectsBatch(creds, sendingAccountId, takenProspects, startTime);
  
  if (result.enrichedCount > 0 || result.errorCount > 0) {
    // Write back the updated taken prospects to the main prospects array
    for (let i = 0; i < takenProspects.length; i++) {
      prospectsData[i] = takenProspects[i];
    }
    prospectsDataRange.setValues(prospectsData);
    SpreadsheetApp.flush();
  }
  
  // Create db rows only for successfully enriched or already enriched prospects
  const newDbRows = [];
  let skippedDueToError = 0;
  
  for (let i = 0; i < takenProspects.length; i++) {
    const p = takenProspects[i];
    const enrichedFlag = String(p[11] || "").trim().toLowerCase();
    
    // Only import if we have a provider ID (successfully enriched)
    if ((enrichedFlag === 'yes' || enrichedFlag === 'true') && p[10]) {
      newDbRows.push([
        selectedCampaignId,
        selectedCampaignName,
        p[2], // linkedin_url
        p[0], // first_name
        p[1], // last_name
        p[8], // company_name
        p[7], // title
        p[5], // city
        p[6], // country
        p[9], // company_website
        sendingAccountId, // sending_account
        p[10] || "", // provider_id from enriched prospect
        "Pending", // connection_request_status
        "", // connection_request_time
        false, // connection_accepted
        "", // connection_accepted_time
        "Pending", // message_1_status
        "", // message_1_sent_time
        "Pending", // message_2_status
        "", // message_2_sent_time
        "Pending", // message_3_status
        "", // message_3_sent_time
        false, // reply_received
        "", // reply_text
        "", // reply_time
        "", // failed_reason
        new Date() // creation_date
      ]);
    } else {
      skippedDueToError++;
    }
  }
  
  if (newDbRows.length > 0) {
    dbSheet.getRange(dbSheet.getLastRow() + 1, 1, newDbRows.length, newDbRows[0].length).setValues(newDbRows);
    campaignsSheet.getRange(campaignRowIndex, 3).setValue(newDbRows.length);
    campaignsSheet.getRange(campaignRowIndex, 4).setValue("Active");
  } else {
    ui.alert("No new prospects were successfully enriched to add to the database. Check Prospects sheet for errors.");
    return;
  }
  
  const postActionResp = ui.alert(
    "Success",
    `Successfully set up campaign database entries: ${selectedCampaignName}\nStatus changed to Active.\nImported ${newDbRows.length} prospects into the Database.\nEnriched ${result.enrichedCount} prospects inline.\nSkipped ${skippedDueToError} due to enrichment errors/invalid URLs.\n\nDo you want to send connection requests now?`,
    ui.ButtonSet.YES_NO
  );

  if (postActionResp === ui.Button.YES) {
    sendConnectionRequests(selectedCampaignId);
  }
}

function sendConnectionRequests(campaignIdToUse) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const campaignsSheet = ss.getSheetByName("Campaigns");
  const dbSheet = ss.getSheetByName("Database");
  
  let selectedCampaignId = campaignIdToUse;
  let connectionNote = "";
  
  if (!selectedCampaignId || typeof selectedCampaignId !== 'string') {
    const lastCampRow = campaignsSheet.getLastRow();
    if (lastCampRow < 2) {
      ui.alert("No campaigns found.");
      return;
    }
    const campaignsData = campaignsSheet.getRange(2, 1, lastCampRow - 1, campaignsSheet.getLastColumn()).getValues();
    const activeCampaigns = campaignsData.filter(row => row[3] === "Active");
    
    if (activeCampaigns.length === 0) {
      ui.alert("No active campaigns found. Please 'Create database entries' to activate a campaign first.");
      return;
    }
    
    if (activeCampaigns.length === 1) {
      selectedCampaignId = activeCampaigns[0][0];
      connectionNote = activeCampaigns[0][4]; 
    } else {
      const campaignListStr = activeCampaigns.map((item, i) => `${i+1}. ${item[1]} (ID: ${item[0]})`).join('\n');
      const campResp = ui.prompt("Select Campaign", `Enter the number (1-${activeCampaigns.length}) of the active campaign to process:\n\n${campaignListStr}`, ui.ButtonSet.OK_CANCEL);
      if (campResp.getSelectedButton() !== ui.Button.OK) return;
      
      const selectedListIndex = parseInt(campResp.getResponseText()) - 1;
      if (isNaN(selectedListIndex) || selectedListIndex < 0 || selectedListIndex >= activeCampaigns.length) {
        ui.alert("Invalid selection.");
        return;
      }
      selectedCampaignId = activeCampaigns[selectedListIndex][0];
      connectionNote = activeCampaigns[selectedListIndex][4];
    }
  } else {
    const campaignsData = campaignsSheet.getRange(2, 1, campaignsSheet.getLastRow() - 1, campaignsSheet.getLastColumn()).getValues();
    const campRow = campaignsData.find(row => row[0] === selectedCampaignId);
    if (campRow) {
      connectionNote = campRow[4];
    }
  }

  const accountsSheet = ss.getSheetByName("Accounts");
  const lastAccRow = accountsSheet.getLastRow();
  if (lastAccRow < 2) {
    ui.alert("No accounts found. Sync accounts first.");
    return;
  }
  
  const allAccountsDataRange = accountsSheet.getRange(2, 1, lastAccRow - 1, accountsSheet.getLastColumn());
  const allAccountsData = allAccountsDataRange.getValues();
  
  let accountsMap = {};
  allAccountsData.forEach((row, i) => {
    const st = String(row[4]).toLowerCase().trim();
    if (st === "active" || st === "ok") {
      accountsMap[row[1]] = {
        arrayIndex: i, // index in allAccountsData
        sentToday: parseInt(row[5]) || 0,
        dailyLimit: parseInt(row[6]) || 100,
        currentErrorCount: parseInt(row[10]) || 0,
        updated: false
      };
    }
  });

  if (Object.keys(accountsMap).length === 0) {
    ui.alert("No active accounts found. Please ensure your accounts are active and try syncing again.");
    return;
  }

  let creds;
  try {
    creds = getCredentials();
  } catch (e) {
    ui.alert(`Error reading credentials: ${e.message}`);
    return;
  }
  
  const apiUrl = `${creds.baseUrl}/users/invite`;
  
  const invSheet = ss.getSheetByName("Invitations");
  const existingInvitations = new Set();
  if (invSheet) {
    const lastInvRow = invSheet.getLastRow();
    if (lastInvRow >= 2) {
      const invData = invSheet.getRange(2, 1, lastInvRow - 1, 3).getValues();
      invData.forEach(row => {
        if (row[2]) existingInvitations.add(String(row[2]));
        if (row[0] && row[1]) existingInvitations.add(row[0] + "_" + row[1]);
      });
    }
  }
  
  const lastDbRow = dbSheet.getLastRow();
  if (lastDbRow < 2) {
    ui.alert("Database is empty.");
    return;
  }
  
  const lastDbCol = Math.max(26, dbSheet.getLastColumn());
  const dbRange = dbSheet.getRange(2, 1, lastDbRow - 1, lastDbCol);
  const dbData = dbRange.getValues();
  
  let sentCount = 0;
  let skippedCount = 0;
  let skippedInactiveAccountCount = 0;
  let errorCount = 0;
  let autoCorrectedCount = 0;
  let processedInBatch = 0;
  const BATCH_SIZE = 10;
  
  for (let i = 0; i < dbData.length; i++) {
    const row = dbData[i];
    const campId = row[0];
    const sendingAccountId = row[10];
    const providerId = row[11]; 
    const status = row[12]; 
    
    if (campId === selectedCampaignId && status === "Pending") {
      const acc = accountsMap[sendingAccountId];
      
      if (!acc) {
        skippedInactiveAccountCount++;
        continue;
      }
      
      if (acc.sentToday >= acc.dailyLimit) {
        row[12] = "Pending"; // Leave as pending so it can be picked up the next day
        row[25] = `[${new Date().toISOString()}] Daily limit reached (${acc.dailyLimit}). Will retry tomorrow.`;
        errorCount++;
        processedInBatch++;
        
        if (processedInBatch > 0 && processedInBatch % BATCH_SIZE === 0) {
          // Instead of dbRange.setValues(dbData), update specific cells
          dbSheet.getRange(i + 2, 13).setValue("Pending");
          dbSheet.getRange(i + 2, 26).setValue(row[25]);
          Object.keys(accountsMap).forEach(id => {
            let act = accountsMap[id];
            if (act.updated) {
              allAccountsData[act.arrayIndex][5] = act.sentToday;
              allAccountsData[act.arrayIndex][10] = act.currentErrorCount;
              act.updated = false;
            }
          });
          allAccountsDataRange.setValues(allAccountsData);
          SpreadsheetApp.flush();
        }
        continue;
      }
      
      if (!providerId) {
        skippedCount++;
        continue;
      }
      
      const payload = {
        provider_id: providerId,
        account_id: sendingAccountId
      };
      
      if (connectionNote && connectionNote.trim() !== "") {
        const firstName = String(row[3] || "").trim();
        payload.message = String(connectionNote).replace(/\$name/g, firstName);
      }
      
      const options = {
        "method": "POST",
        "headers": {
          "X-API-KEY": creds.apiKey,
          "Accept": "application/json",
          "Content-Type": "application/json"
        },
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      };
      
      try {
        const response = fetchWithRetry(apiUrl, options);
        if (response.getResponseCode() === 201 || response.getResponseCode() === 200) {
          const respData = JSON.parse(response.getContentText());
          const invitationId = respData.id || respData.invitation_id || "";

          row[12] = "Sent";
          row[13] = new Date(); 
          row[25] = ""; // clear any previous error
          
          if (invitationId) {
             if (invSheet && !existingInvitations.has(String(invitationId)) && !existingInvitations.has(sendingAccountId + "_" + providerId)) {
               invSheet.appendRow([sendingAccountId, providerId, invitationId, "Sent", new Date()]);
               existingInvitations.add(String(invitationId));
               existingInvitations.add(sendingAccountId + "_" + providerId);
             }
          }

          acc.sentToday++;
          acc.updated = true;
          sentCount++;
          processedInBatch++;
          
          Utilities.sleep(500); 
        } else {
          const respText = response.getContentText();
          let isAlreadyConnected = false;
          let isInvitationAlreadySent = false;
          let extractedError = "";
          
          try {
            const errJson = JSON.parse(respText);
            const errType = String(errJson.type || "").toLowerCase();
            const errDetail = String(errJson.detail || "").toLowerCase();
            
            extractedError = errJson.detail || errJson.message || errJson.error || respText.substring(0, 200);
            
            if (errType.includes("already_connected") || errType.includes("is_connection") || errDetail.includes("already connected") || errDetail.includes("already a connection")) {
              isAlreadyConnected = true;
            } else if (errType.includes("invitation_already_sent") || errDetail.includes("invitation already been sent") || errDetail.includes("invitation has already been sent")) {
              isInvitationAlreadySent = true;
            }
          } catch(e) {
            extractedError = respText.substring(0, 200);
          }
          
          if (isAlreadyConnected) {
             row[12] = "Accepted"; 
             row[14] = true; 
             row[15] = new Date(); 
             row[25] = `[${new Date().toISOString()}] Auto-corrected: Already a connection. Moved to Accepted status.`;
             autoCorrectedCount++;
             processedInBatch++;
          } else if (isInvitationAlreadySent) {
             const profileUrl = `${creds.baseUrl}/users/${providerId}?account_id=${sendingAccountId}`;
             const profileOptions = {
               "method": "GET",
               "headers": {
                 "X-API-KEY": creds.apiKey,
                 "Accept": "application/json"
               },
               "muteHttpExceptions": true
             };
             
             let connectedAt = null;
             try {
               const profileResp = fetchWithRetry(profileUrl, profileOptions);
               if (profileResp.getResponseCode() === 200) {
                 const profileData = JSON.parse(profileResp.getContentText());
                 if (profileData.connected_at) {
                   connectedAt = profileData.connected_at;
                 }
               }
             } catch(e) {}
             
             row[12] = "Sent";
             row[13] = new Date(); // Treat as sent right now to move it out of pending
             
             if (connectedAt) {
               row[14] = true;
               row[15] = new Date(connectedAt); 
               row[25] = `[${new Date().toISOString()}] Auto-corrected: Invitation already sent and is now connected.`;
             } else {
               // We need to find the missing invitation_id so we can uninvite later if needed
               const foundInvId = findInvitationId(creds, sendingAccountId, providerId);
               if (foundInvId) {
                 if (invSheet && !existingInvitations.has(String(foundInvId)) && !existingInvitations.has(sendingAccountId + "_" + providerId)) {
                   invSheet.appendRow([sendingAccountId, providerId, foundInvId, "Sent", new Date()]);
                   existingInvitations.add(String(foundInvId));
                   existingInvitations.add(sendingAccountId + "_" + providerId);
                 }
                 row[25] = `[${new Date().toISOString()}] Auto-corrected: Invitation already sent. Found ID ${foundInvId}`;
               } else {
                 row[25] = `[${new Date().toISOString()}] Auto-corrected: Invitation already sent. Could not find ID.`;
               }
             }
             
             autoCorrectedCount++;
             processedInBatch++;
          } else {
             errorCount++;
             acc.currentErrorCount++;
             acc.updated = true;
             
             row[12] = "Failed";
             row[25] = `[${new Date().toISOString()}] ${extractedError}`; 
             processedInBatch++;
          }
        }
      } catch (e) {
        errorCount++;
        acc.currentErrorCount++;
        acc.updated = true;
        
        row[12] = "Failed";
        row[25] = `[${new Date().toISOString()}] Error: ${e.message}`.substring(0, 500);
        processedInBatch++;
      }
      
      // Update specific row cells to prevent overwriting entire sheet
      dbSheet.getRange(i + 2, 13, 1, 4).setValues([[row[12], row[13] || "", row[14] || false, row[15] || ""]]);
      dbSheet.getRange(i + 2, 26).setValue(row[25] || "");
      
      if (processedInBatch > 0 && processedInBatch % BATCH_SIZE === 0) {
        // update accounts array then sheet
        Object.keys(accountsMap).forEach(id => {
          let act = accountsMap[id];
          if (act.updated) {
            allAccountsData[act.arrayIndex][5] = act.sentToday;
            allAccountsData[act.arrayIndex][10] = act.currentErrorCount;
            act.updated = false; // reset flag
          }
        });
        allAccountsDataRange.setValues(allAccountsData);
        SpreadsheetApp.flush();
      }
    }
  }
  
  // Final flush and recalculate Campaign Stats
  if (processedInBatch > 0 || sentCount > 0) {
    Object.keys(accountsMap).forEach(id => {
      let act = accountsMap[id];
      if (act.updated) {
        allAccountsData[act.arrayIndex][5] = act.sentToday;
        allAccountsData[act.arrayIndex][10] = act.currentErrorCount;
        act.updated = false;
      }
    });
    allAccountsDataRange.setValues(allAccountsData);
    SpreadsheetApp.flush();
  }

  // Calculate & Update Campaign Stats
  let campConnectionsSent = 0;
  let campConnectionsAccepted = 0;
  let campMessagesSent = 0;
  let campRepliesReceived = 0;

  for (let i = 0; i < dbData.length; i++) {
    if (dbData[i][0] === selectedCampaignId) {
      let st = dbData[i][12]; // Connection Request Status
      if (st === "Sent" || st === "Accepted") campConnectionsSent++;
      
      if (dbData[i][14] === true) campConnectionsAccepted++; // Connection Accepted (Boolean)
      
      // Messages 1, 2, 3
      if (dbData[i][16] === "Sent") campMessagesSent++;
      if (dbData[i][18] === "Sent") campMessagesSent++;
      if (dbData[i][20] === "Sent") campMessagesSent++;
      
      let replyText = dbData[i][23];
      let replyTime = dbData[i][24];
      
      const hasReplyBoxChecked = (dbData[i][22] === true || String(dbData[i][22]).toUpperCase() === "TRUE");
      
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
        campRepliesReceived++;
      }
    }
  }

  const campaignsDataRange = campaignsSheet.getRange(2, 1, campaignsSheet.getLastRow() - 1, campaignsSheet.getLastColumn());
  const cData = campaignsDataRange.getValues();
  for (let i = 0; i < cData.length; i++) {
    if (cData[i][0] === selectedCampaignId) {
      cData[i][10] = campConnectionsSent;     // Column K
      cData[i][11] = campConnectionsAccepted; // Column L
      cData[i][12] = campMessagesSent;        // Column M
      cData[i][13] = campRepliesReceived;     // Column N
      break;
    }
  }
  campaignsDataRange.setValues(cData);
  SpreadsheetApp.flush();
  
  let alertMsg = `Connection Requests Processed!\n\nSuccessfully Sent: ${sentCount}\nSkipped (No Provider ID): ${skippedCount}\nAuto-Corrected: ${autoCorrectedCount}\nErrors (incl. Limits): ${errorCount}`;
  
  if (skippedInactiveAccountCount > 0) {
    alertMsg += `\nSkipped ${skippedInactiveAccountCount} prospects because their assigned account is inactive.`;
  }
  if (errorCount > 0) {
    alertMsg += `\n\nCheck the "failed_reason" column in your Database sheet to see the exact errors or limit notifications.`;
  }
  
  ui.alert(alertMsg);
}

function forceCheckRequests() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const campaignsSheet = ss.getSheetByName("Campaigns");
  const dbSheet = ss.getSheetByName("Database");
  
  const lastCampRow = campaignsSheet.getLastRow();
  if (lastCampRow < 2) {
    ui.alert("No campaigns found.");
    return;
  }
  
  const campaignsData = campaignsSheet.getRange(2, 1, lastCampRow - 1, campaignsSheet.getLastColumn()).getValues();
  const activeCampaigns = campaignsData.filter(row => row[3] === "Active" || row[3] === "Completed");
  
  if (activeCampaigns.length === 0) {
    ui.alert("No active or completed campaigns found to check.");
    return;
  }
  
  let selectedCampaignId;
  let selectedCampaignName;
  
  if (activeCampaigns.length === 1) {
    selectedCampaignId = activeCampaigns[0][0];
    selectedCampaignName = activeCampaigns[0][1];
  } else {
    const campaignListStr = activeCampaigns.map((item, i) => `${i+1}. ${item[1]} (ID: ${item[0]})`).join('\n');
    const campResp = ui.prompt("Select Campaign", `Enter the number (1-${activeCampaigns.length}) of the campaign to check:\n\n${campaignListStr}`, ui.ButtonSet.OK_CANCEL);
    if (campResp.getSelectedButton() !== ui.Button.OK) return;
    
    const selectedListIndex = parseInt(campResp.getResponseText()) - 1;
    if (isNaN(selectedListIndex) || selectedListIndex < 0 || selectedListIndex >= activeCampaigns.length) {
      ui.alert("Invalid selection.");
      return;
    }
    selectedCampaignId = activeCampaigns[selectedListIndex][0];
    selectedCampaignName = activeCampaigns[selectedListIndex][1];
  }
  
  let creds;
  try {
    creds = getCredentials();
  } catch (e) {
    ui.alert(`Error reading credentials: ${e.message}`);
    return;
  }
  
  const invSheet = ss.getSheetByName("Invitations");
  const existingInvitations = new Set();
  if (invSheet) {
    const lastInvRow = invSheet.getLastRow();
    if (lastInvRow >= 2) {
      const invData = invSheet.getRange(2, 1, lastInvRow - 1, 3).getValues();
      invData.forEach(row => {
        if (row[2]) existingInvitations.add(String(row[2]));
        if (row[0] && row[1]) existingInvitations.add(row[0] + "_" + row[1]);
      });
    }
  }

  const lastDbRow = dbSheet.getLastRow();
  if (lastDbRow < 2) {
    ui.alert("Database is empty.");
    return;
  }
  
  const lastDbCol = Math.max(26, dbSheet.getLastColumn());
  const dbRange = dbSheet.getRange(2, 1, lastDbRow - 1, lastDbCol);
  const dbData = dbRange.getValues();
  
  const indicesToCheck = [];
  for (let i = 0; i < dbData.length; i++) {
    const row = dbData[i];
    // Check pending or sent, if provider id exists, and not accepted
    if (row[0] === selectedCampaignId && row[11] && row[14] !== true && (row[12] === "Sent" || row[12] === "Pending")) {
      indicesToCheck.push(i);
    }
  }
  
  if (indicesToCheck.length === 0) {
    ui.alert(`No pending/sent prospects found to check for '${selectedCampaignName}'.`);
    return;
  }
  
  let checkedCount = 0;
  let updatedConnectedCount = 0;
  let updatedInvitationCount = 0;
  const BATCH_SIZE = 40;
  const startTime = Date.now();
  
  for (let b = 0; b < indicesToCheck.length; b += BATCH_SIZE) {
    if (Date.now() - startTime > 280000) { // Safety limit: approx 4.5 minutes
       ui.alert(`Nearing 6-minute execution limit. Stopping early.\nProcessed ${checkedCount} out of ${indicesToCheck.length} prospects.`);
       break;
    }
    
    const batchIndices = indicesToCheck.slice(b, b + BATCH_SIZE);
    const reqs = [];
    
    for (let k = 0; k < batchIndices.length; k++) {
      const idx = batchIndices[k];
      const row = dbData[idx];
      const providerId = row[11];
      const accountId = row[10];
      
      reqs.push({
        url: `${creds.baseUrl}/users/${providerId}?account_id=${accountId}`,
        method: "GET",
        headers: { "X-API-KEY": creds.apiKey, "Accept": "application/json" },
        muteHttpExceptions: true
      });
    }
    
    let responses;
    try {
      responses = UrlFetchApp.fetchAll(reqs);
    } catch(e) {
      Logger.log("Batch fetch failed: " + e.message);
      continue;
    }
    
    for (let k = 0; k < responses.length; k++) {
      const idx = batchIndices[k];
      const row = dbData[idx];
      const response = responses[k];
      const code = response.getResponseCode();
      
      let connectedAt = null;
      if (code === 200) {
         try {
           const profileData = JSON.parse(response.getContentText());
           if (profileData.connected_at) connectedAt = profileData.connected_at;
         } catch(e) {}
      }
      
      checkedCount++;
      
      if (connectedAt) {
        row[12] = "Accepted";
        row[14] = true;
        row[15] = new Date(connectedAt); 
        row[25] = `[${new Date().toISOString()}] Force Check: Confirmed connected.`;
        updatedConnectedCount++;
      } else {
        const connReqStatus = row[12];
        const sendingAccountId = row[10];
        const providerId = row[11];
        
        // Find invitation id fallback
        const foundInvId = findInvitationId(creds, sendingAccountId, providerId);
        if (foundInvId) {
          if (connReqStatus === "Pending") {
            row[12] = "Sent";
            row[13] = new Date();
          }
          if (invSheet && !existingInvitations.has(String(foundInvId)) && !existingInvitations.has(sendingAccountId + "_" + providerId)) {
            invSheet.appendRow([sendingAccountId, providerId, foundInvId, "Sent", new Date()]);
            existingInvitations.add(String(foundInvId));
            existingInvitations.add(sendingAccountId + "_" + providerId);
            updatedInvitationCount++;
          }
          row[25] = `[${new Date().toISOString()}] Force Check: Found missing invite ID ${foundInvId}`;
        }
      }
      dbSheet.getRange(idx + 2, 13, 1, 4).setValues([[row[12], row[13] || "", row[14] || false, row[15] || ""]]);
      dbSheet.getRange(idx + 2, 26).setValue(row[25] || "");
    }
    
    SpreadsheetApp.flush();
  }
  
  // Call updateGlobalStats to sync changes to campaign tab
  updateGlobalStats();
  
  ui.alert(`Force Check Complete for '${selectedCampaignName}'!\n\nProspects Checked: ${checkedCount}\nNewly marked as Connected: ${updatedConnectedCount}\nMissing Invitations Re-linked: ${updatedInvitationCount}`);
}
