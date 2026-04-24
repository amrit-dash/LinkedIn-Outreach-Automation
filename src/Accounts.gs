function syncAccounts() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Accounts");
  if (!sheet) {
    ui.alert("Accounts sheet not found.");
    return;
  }
  
  let creds;
  try {
    creds = getCredentials();
  } catch (e) {
    ui.alert(`Error reading credentials: ${e.message}`);
    return;
  }
  
  const url = `${creds.baseUrl}/accounts`;
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
      const fetchedAccounts = data.items || data; 
      
      const lastRow = sheet.getLastRow();
      
      let existingData = [];
      let existingIds = {};
      
      if (lastRow > 1) {
        existingData = sheet.getRange(2, 1, lastRow - 1, Math.max(11, sheet.getLastColumn())).getValues();
        existingData.forEach((row, index) => {
          const id = row[1]; // unipile_id
          if (id) {
            existingIds[id] = {
              rowNum: index + 2,
              data: row
            };
          }
        });
      }
      
      let updatedCount = 0;
      let newCount = 0;
      let newRows = [];
      let fetchedIds = {};
      
      fetchedAccounts.forEach(acc => {
        const id = acc.id;
        fetchedIds[id] = true;
        const name = acc.name || "";
        
        let rawStatus = "";
        if (acc.sources && acc.sources.length > 0 && acc.sources[0].status) {
          rawStatus = acc.sources[0].status;
        } else {
          rawStatus = acc.status || acc.state || "";
        }
        
        let status = "Inactive";
        if (String(rawStatus).toUpperCase() === "OK") {
          status = "Active";
        }
        
        const now = new Date();
        
        let invitesToday = 0;
        let messagesToday = 0;
        if (existingIds[id]) {
          invitesToday = existingIds[id].data[5] || 0;
          messagesToday = existingIds[id].data[7] || 0;
        }
        
        if (existingIds[id]) {
          const rowNum = existingIds[id].rowNum;
          
          sheet.getRange(rowNum, 1).setValue(name);
          sheet.getRange(rowNum, 3, 1, 3).setValues([[creds.apiKey, creds.dsn, status]]);
          sheet.getRange(rowNum, 10).setValue(now);
          
          updatedCount++;
        } else {
          newRows.push([
            name,
            id,
            creds.apiKey,
            creds.dsn,
            status,
            invitesToday, // connections_sent_today
            100, // connections_daily_limit (default)
            messagesToday, // dms_sent_today
            100, // dms_daily_limit (default)
            now, // last_checked
            0 // error_count
          ]);
          newCount++;
        }
      });
      
      // Update accounts that were not returned by API to Inactive
      let inactiveCount = 0;
      Object.keys(existingIds).forEach(id => {
        if (!fetchedIds[id]) {
          const rowNum = existingIds[id].rowNum;
          const currentStatus = existingIds[id].data[4]; // column E is index 4
          if (String(currentStatus).toLowerCase() !== "inactive") {
            sheet.getRange(rowNum, 5).setValue("Inactive");
            inactiveCount++;
          }
        }
      });
      
      if (newRows.length > 0) {
        sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      }
      
      ui.alert(`Sync Complete.\nUpdated existing accounts: ${updatedCount}\nAdded new accounts: ${newCount}\nMarked disconnected/missing as inactive: ${inactiveCount}`);
      
    } else {
      ui.alert(`Failed to sync accounts. Status code: ${response.getResponseCode()}\nResponse: ${response.getContentText()}`);
    }
  } catch (e) {
    ui.alert(`Error fetching from Unipile API: ${e.message}`);
  }
}
