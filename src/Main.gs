function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  const startCampaignMenu = ui.createMenu('Start Campaign Now')
    .addItem('1 - Create database entries', 'createDatabaseEntries')
    .addItem('2 - Send connection requests', 'sendConnectionRequests')
    .addItem('3 - Check connection requests', 'forceCheckRequests');

  const automationMenu = ui.createMenu('Automate & Monitor')
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
    .addItem('🚀 . Create New Campaign', 'showCreateCampaignDialog')
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
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('WEBHOOK_MONITORING_ENABLED', 'TRUE');
  
  ui.alert(
    'Monitoring Enabled',
    'Apps Script Webhook Monitoring is now ENABLED.\n\nIMPORTANT: Please ensure you disable or pause your n8n webhook scenario, as this Apps Script will now handle incoming replies and new relations directly.',
    ui.ButtonSet.OK
  );
}

function stopMonitoringProcess() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('WEBHOOK_MONITORING_ENABLED', 'FALSE');
  
  ui.alert(
    'Monitoring Disabled',
    'Apps Script Webhook Monitoring is now DISABLED.\n\nIMPORTANT: Please ensure you re-enable your n8n webhook scenario if you want n8n to continue handling replies and new relations.',
    ui.ButtonSet.OK
  );
}
// UI Helpers
// Ready for V1 deployment
