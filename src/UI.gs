function showCreateCampaignDialog() {
  const html = HtmlService.createHtmlOutputFromFile('CampaignForm')
      .setWidth(500)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create New Campaign');
}

function processCampaignForm(formObject) {
  const campaignName = formObject.campaignName;
  const numProspects = formObject.numProspects;
  const connectionNote = formObject.connectionNote;
  const msg1 = formObject.msg1;
  const delay2 = formObject.delay2;
  const msg2 = formObject.msg2;
  const delay3 = formObject.delay3;
  const msg3 = formObject.msg3;
  
  const campaignId = Utilities.getUuid();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Campaigns");
  
  sheet.appendRow([
    campaignId,
    campaignName,
    numProspects, 
    "Not Started", 
    connectionNote,
    msg1,
    msg2,
    msg3, 
    delay2 || 0,
    delay3 || 0, 
    0, 
    0, 
    0, 
    0, 
    new Date() 
  ]);
  
  SpreadsheetApp.getUi().alert(`Campaign '${campaignName}' created successfully!\nStatus: Not Started\nID: ${campaignId}`);
  return { success: true, name: campaignName, id: campaignId };
}
