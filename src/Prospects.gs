function extractLinkedInId(url) {
  if (!url) return null;
  const match = url.match(/linkedin\.com\/in\/([^\/\?]+)/i);
  return match ? match[1] : null;
}

// Reusable function to enrich a single prospect and return the updated row array
function enrichProspectInline(creds, accountId, url, row) {
  const identifier = extractLinkedInId(url);
  if (!identifier) return { row: row, success: false };
  
  const apiUrl = `${creds.baseUrl}/users/${identifier}?account_id=${accountId}`;
  const options = {
    "method": "GET",
    "headers": {
      "X-API-KEY": creds.apiKey,
      "Accept": "application/json"
    },
    "muteHttpExceptions": true
  };
  
  try {
    const response = fetchWithRetry(apiUrl, options);
    if (response.getResponseCode() === 200) {
      const profile = JSON.parse(response.getContentText());
      
      if (!row[0] && profile.first_name) row[0] = profile.first_name;
      if (!row[1] && profile.last_name) row[1] = profile.last_name;
      
      if (profile.primary_locale && profile.primary_locale.country) {
        row[3] = profile.primary_locale.country;
      }
      
      if (profile.location) {
        const parts = profile.location.split(',').map(s => s.trim());
        row[5] = parts[0] || ""; 
        row[6] = parts.length > 1 ? parts[parts.length - 1] : ""; 
      }
      
      if (profile.headline) {
        row[7] = profile.headline;
      }
      
      if (profile.provider_id) {
        row[10] = profile.provider_id;
      }
      
      if (profile.experiences && profile.experiences.length > 0) {
        const exp = profile.experiences[0];
        if (exp.company_name) row[8] = exp.company_name;
        if (exp.company_website) row[9] = exp.company_website;
      }
      
      row[11] = "Yes"; // Enriched flag
      return { row: row, success: true };
    }
  } catch (e) {
    Logger.log(`Error enriching prospect: ${e.message}`);
  }
  return { row: row, success: false };
}

// Bulk enrichment with robust error handling and execution time protection
function enrichProspectsBatch(creds, accountId, prospectsData, startTime) {
  let enrichedCount = 0;
  let errorCount = 0;
  const BATCH_SIZE = 40; // Use UrlFetchApp.fetchAll batches
  
  for (let b = 0; b < prospectsData.length; b += BATCH_SIZE) {
    // If we've exceeded 5 minutes, stop processing to prevent timeout crash
    if (Date.now() - startTime > 300000) {
      Logger.log("Approaching 6-minute limit. Stopping enrichment early.");
      break; 
    }
    
    const batchReqs = [];
    const validIndices = [];
    const end = Math.min(b + BATCH_SIZE, prospectsData.length);
    
    for (let i = b; i < end; i++) {
      const row = prospectsData[i];
      let enrichedFlag = String(row[11] || "").trim().toLowerCase();
      
      if (enrichedFlag !== 'true' && enrichedFlag !== 'yes') {
        const identifier = extractLinkedInId(row[2]);
        if (identifier) {
          batchReqs.push({
            url: `${creds.baseUrl}/users/${identifier}?account_id=${accountId}`,
            method: "GET",
            headers: { "X-API-KEY": creds.apiKey, "Accept": "application/json" },
            muteHttpExceptions: true
          });
          validIndices.push(i);
        } else {
          // Can't extract ID, mark as failed
          row[11] = "Failed (Invalid URL)";
          errorCount++;
        }
      }
    }
    
    if (batchReqs.length === 0) continue;
    
    let responses;
    try {
      responses = UrlFetchApp.fetchAll(batchReqs);
    } catch(e) {
      Logger.log("Batch fetch failed: " + e.message);
      errorCount += batchReqs.length;
      continue;
    }
    
    // Check for rate limits and prepare retries
    const retryReqs = [];
    const retryIndices = [];
    
    for (let k = 0; k < responses.length; k++) {
      const code = responses[k].getResponseCode();
      if (code === 429 || code >= 500) {
        retryReqs.push(batchReqs[k]);
        retryIndices.push(validIndices[k]);
      } else {
        processEnrichmentResponse(responses[k], prospectsData[validIndices[k]]);
      }
    }
    
    // One level of retry with exponential backoff for 429s in the batch
    if (retryReqs.length > 0) {
      Utilities.sleep(2000); // 2 second backoff
      try {
        const retryResponses = UrlFetchApp.fetchAll(retryReqs);
        for (let k = 0; k < retryResponses.length; k++) {
          processEnrichmentResponse(retryResponses[k], prospectsData[retryIndices[k]]);
        }
      } catch(e) {
        Logger.log("Retry batch fetch failed: " + e.message);
      }
    }
  }
  
  // Count final successes/errors
  for (let i = 0; i < prospectsData.length; i++) {
    const row = prospectsData[i];
    let flag = String(row[11] || "").trim().toLowerCase();
    if (flag === 'yes' || flag === 'true') enrichedCount++;
    else if (flag.includes("failed")) errorCount++;
  }
  
  return { enrichedCount, errorCount };
}

function processEnrichmentResponse(response, row) {
  const code = response.getResponseCode();
  if (code === 200) {
    try {
      const profile = JSON.parse(response.getContentText());
      if (!row[0] && profile.first_name) row[0] = profile.first_name;
      if (!row[1] && profile.last_name) row[1] = profile.last_name;
      if (profile.primary_locale && profile.primary_locale.country) row[3] = profile.primary_locale.country;
      
      if (profile.location) {
        const parts = profile.location.split(',').map(s => s.trim());
        row[5] = parts[0] || ""; 
        row[6] = parts.length > 1 ? parts[parts.length - 1] : ""; 
      }
      
      if (profile.headline) row[7] = profile.headline;
      if (profile.provider_id) row[10] = profile.provider_id;
      if (profile.experiences && profile.experiences.length > 0) {
        const exp = profile.experiences[0];
        if (exp.company_name) row[8] = exp.company_name;
        if (exp.company_website) row[9] = exp.company_website;
      }
      row[11] = "Yes"; 
    } catch(e) {
      row[11] = "Failed (Parse Error)";
    }
  } else if (code === 404) {
    row[11] = "Failed (Not Found)";
  } else {
    row[11] = `Failed (${code})`;
  }
}

function enrichProspects() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prospectsSheet = ss.getSheetByName("Prospects");
  const accountsSheet = ss.getSheetByName("Accounts");
  
  if (!prospectsSheet) {
    ui.alert("Prospects sheet not found.");
    return;
  }
  
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
  
  let accountId;
  if (activeAccounts.length === 1) {
    accountId = activeAccounts[0].row[1];
  } else {
    const accListStr = activeAccounts.map((item, i) => `${i+1}. ${item.row[0]} (ID: ${item.row[1]})`).join('\n');
    const accResp = ui.prompt("Select Account for Enrichment", `Enter the number (1-${activeAccounts.length}) of the active account to use for searching:\n\n${accListStr}`, ui.ButtonSet.OK_CANCEL);
    if (accResp.getSelectedButton() !== ui.Button.OK) return;
    
    const selectedAccIndex = parseInt(accResp.getResponseText()) - 1;
    if (isNaN(selectedAccIndex) || selectedAccIndex < 0 || selectedAccIndex >= activeAccounts.length) {
      ui.alert("Invalid selection.");
      return;
    }
    accountId = activeAccounts[selectedAccIndex].row[1];
  }
  
  let creds;
  try {
    creds = getCredentials();
  } catch (e) {
    ui.alert(`Error reading credentials: ${e.message}`);
    return;
  }
  
  const lastRow = prospectsSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No prospects found.");
    return;
  }
  
  const lastCol = Math.max(12, prospectsSheet.getLastColumn());
  const range = prospectsSheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = range.getValues();
  
  const startTime = Date.now();
  
  // To handle batching appropriately when processing the entire sheet sequentially:
  const result = enrichProspectsBatch(creds, accountId, data, startTime);
  
  // Final flush for all data
  range.setValues(data);
  SpreadsheetApp.flush();
  
  ui.alert(`Enrichment complete!\nSuccessfully enriched total: ${result.enrichedCount}\nErrors/Failed: ${result.errorCount}`);
}
// Added batch enrichment
