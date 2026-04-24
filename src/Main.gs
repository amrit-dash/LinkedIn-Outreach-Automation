function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  const manualTriggersMenu = ui.createMenu('⚙️ . Manual Triggers')
    .addItem('1 - Create database entries', 'createDatabaseEntries')
    .addItem('2 - Send connection requests', 'sendConnectionRequests')
    .addItem('3 - Check connection requests', 'forceCheckRequests')
    .addItem('4 - Send First Message', 'sendFirstMessageManual')
    .addItem('5 - Send Second Message', 'sendSecondMessageManual')
    .addItem('6 - Send Third Message', 'sendThirdMessageManual');

  const startCampaignMenu = ui.createMenu('🎯 . Campaign Settings')
    .addItem('📢 . Create New Campaign', 'showCreateCampaignDialog')
    .addSeparator()
    .addItem('🚀 . Start Campaign', 'createDatabaseEntries')
    .addSeparator()
    .addSubMenu(manualTriggersMenu);

  const automationMenu = ui.createMenu('🤖 . Automate & Monitor')
    .addItem('▶️ . Start Campaign Background Worker', 'startBackgroundWorker')
    .addItem('⏹️ . Stop Campaign Background Worker', 'stopBackgroundWorker')
    .addSeparator()
    .addItem('▶️ . Start Stats Worker', 'startStatsWorker')
    .addItem('⏹️ . Stop Stats Worker', 'stopStatsWorker')
    .addSeparator()
    .addItem('▶️ . Start Webhook Monitoring Process', 'startMonitoringProcess')
    .addItem('⏹️ . Stop Webhook Monitoring Process', 'stopMonitoringProcess');

  ui.createMenu('LinkedIn Outreach System')
    .addItem('🌀 . Sync Unipile Accounts', 'syncAccounts')
    .addSeparator()
    .addItem('🔍 . Enrich LinkedIn Prospects', 'enrichProspects')
    .addSeparator()
    .addItem('📊 . Update Global Stats', 'updateGlobalStats')
    .addSeparator()
    .addSubMenu(startCampaignMenu)
    .addSeparator()
    .addSubMenu(automationMenu)
    .addToUi();

  // Update stats in the background whenever the sheet is opened
  updateGlobalStats();
}

function onEdit(e) {
  // Update global stats whenever the sheet is edited by a user.
  // This will pick up manual ticks (like reply_received) and automatically roll them up.
  if (e && e.user) {
    updateGlobalStats();
  }
}

function startMonitoringProcess() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const credSheet = ss.getSheetByName("Credentials");

  if (credSheet) {
    const n8nStatus = credSheet.getRange(5, 2).getValue();
    if (String(n8nStatus).trim().toUpperCase() !== 'DISABLED') {
      const cell = credSheet.getRange(5, 2);
      credSheet.setActiveRange(cell);
      cell.setBackground('#FBE6A3');
      ui.alert('Action Cancelled', 'You must disable the n8n_webhook_status before starting Apps Script monitoring.', ui.ButtonSet.OK);
      return;
    }
    credSheet.getRange(5, 2).setBackground(null);
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('WEBHOOK_MONITORING_ENABLED', 'ENABLED');
  
  if (credSheet) {
    credSheet.getRange(4, 1).setValue("appsscript_webhook_status");
    credSheet.getRange(4, 2).setValue("ENABLED");
  }
  
  ui.alert(
    'Monitoring Enabled',
    'Apps Script Webhook Monitoring is now ENABLED.\n\nIMPORTANT: Please ensure you disable or pause your n8n webhook scenario, as this Apps Script will now handle incoming replies and new relations directly.',
    ui.ButtonSet.OK
  );
}

function stopMonitoringProcess() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const credSheet = ss.getSheetByName("Credentials");

  if (credSheet) {
    const n8nStatus = credSheet.getRange(5, 2).getValue();
    if (String(n8nStatus).trim().toUpperCase() !== 'ENABLED') {
      const cell = credSheet.getRange(5, 2);
      credSheet.setActiveRange(cell);
      cell.setBackground('#FBE6A3');
      ui.alert('Action Cancelled', 'You must manually set the n8n_webhook_status to ENABLED before stopping Apps Script monitoring.', ui.ButtonSet.OK);
      return;
    }
    credSheet.getRange(5, 2).setBackground(null);
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('WEBHOOK_MONITORING_ENABLED', 'DISABLED');
  
  if (credSheet) {
    credSheet.getRange(4, 1).setValue("appsscript_webhook_status");
    credSheet.getRange(4, 2).setValue("DISABLED");
  }
  
  ui.alert(
    'Monitoring Disabled',
    'Apps Script Webhook Monitoring is now DISABLED.\n\nIMPORTANT: Please ensure your n8n webhook scenario is active if you want n8n to continue handling replies and new relations.',
    ui.ButtonSet.OK
  );
}
