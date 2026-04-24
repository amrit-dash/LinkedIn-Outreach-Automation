
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
