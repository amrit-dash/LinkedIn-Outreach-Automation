function getCredentials() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const credSheet = ss.getSheetByName("Credentials");
  if (!credSheet) {
    throw new Error("Credentials sheet not found.");
  }
  
  const data = credSheet.getRange(1, 1, 3, 2).getValues();
  
  const apiKey = data[0][1];
  let dsn = data[1][1];
  
  if (!apiKey || !dsn) {
    throw new Error("Missing API Key or DSN in Credentials sheet (Row 1 and 2, Column B).");
  }
  
  dsn = dsn.replace(/^https?:\/\//, '');
  
  const baseUrl = `https://${dsn}/api/v1`;
  
  if (data[2][1] !== baseUrl) {
    credSheet.getRange(3, 2).setValue(baseUrl);
  }
  
  return { apiKey, dsn, baseUrl };
}

// Helper to fetch with exponential backoff for rate limits and server errors
function fetchWithRetry(url, options, maxRetries = 3) {
  let retries = 0;
  let delay = 1000; // start with 1 second delay
  while (true) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      if ((code === 429 || code >= 500) && retries < maxRetries) {
        retries++;
        Logger.log(`Rate limit or server error (${code}). Retrying in ${delay}ms...`);
        Utilities.sleep(delay);
        delay *= 2; // exponential backoff
        continue;
      }
      return response;
    } catch (e) {
      if (retries < maxRetries) {
        retries++;
        Logger.log(`Fetch error: ${e.message}. Retrying in ${delay}ms...`);
        Utilities.sleep(delay);
        delay *= 2;
        continue;
      }
      throw e;
    }
  }
}

// getTodayCount has been removed as it is no longer used. Daily counts are tracked internally in the spreadsheet.
