function extractDriveFolderId_(url) {
  const match = url.match(/[-\w]{25,}/);
  if (!match) {
    throw new Error("Could not extract a Google Drive folder ID from the link provided.");
  }
  return match[0];
}

function clearStatusCell() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const overviewSheet = ss.getSheetByName("Overview");
  if (overviewSheet) {
    overviewSheet.getRange("P4").clearContent().setBackground("#ffffff").setFontColor("#000000");
  }
}

function createAutoTrigger() { 
  deleteAutoTrigger(); 
  ScriptApp.newTrigger('exportSheetsToPDF').timeBased().after(60000).create(); 
}

function deleteAutoTrigger() { 
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'exportSheetsToPDF') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}