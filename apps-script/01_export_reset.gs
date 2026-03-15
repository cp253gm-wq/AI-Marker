function runUpdateAllSheets() {
  // SAFETY: Deletes any active PDF timers as soon as the update starts.
  deleteAutoTrigger(); 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const excludedSheets = ["Marking", "Overview"];

  for (let s = 0; s < allSheets.length; s++) {
    const sheet = allSheets[s];
    let sheetName = sheet.getName();

    if (excludedSheets.includes(sheetName)) continue;

    try {
      const newName = sheet.getRange("R120").getValue().toString().trim();
      if (newName !== "" && newName !== sheetName) {
        sheet.setName(newName);
        sheetName = newName; 
      }
    } catch (err) {
      console.log("Could not rename sheet '" + sheetName + "': " + err.message);
    }

    const startRow = 20;
    const endRow = 119;
    const numRows = endRow - startRow + 1;
    const values = sheet.getRange(startRow, 18, numRows).getValues(); 

    for (let i = 0; i < values.length; i++) {
      const currentRow = startRow + i;
      const val = values[i][0].toString().toUpperCase();

      if (val === "Y") {
        sheet.hideRows(currentRow);
      } 
      else if (val === "N") {
        sheet.showRows(currentRow);
        sheet.autoResizeRows(currentRow, 1);
        if (sheet.getRowHeight(currentRow) < 40) {
          sheet.setRowHeight(currentRow, 40);
        }
      }
    }
  }
  linkStudentNamesToReports();
  SpreadsheetApp.getUi().alert("All sheets have been updated!");
}

function exportSheetsToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const excludedSheets = ["Overview", "Marking"];
  const startTime = new Date().getTime();
  
  const overviewSheet = ss.getSheetByName("Overview");
  const marksSheet = ss.getSheetByName("Marking");
  const statusCell = overviewSheet.getRange("P4");
  
const exportFolderLink = overviewSheet.getRange("Q8").getValue().toString().trim();

if (exportFolderLink === "") {
  statusCell.setValue("Error: Q8 export folder link is empty").setBackground("#ea4335").setFontColor("white");
  return;
}

  const studentList = marksSheet.getRange("C14:C43").getValues();
  const totalTarget = studentList.filter(row => row[0].toString().trim() !== "").length;

  if (totalTarget === 0) {
    statusCell.setValue("Error: No students in C14:C43").setBackground("#ea4335").setFontColor("white");
    return;
  }

const exportFolderId = extractDriveFolderId_(exportFolderLink);
const exportFolder = DriveApp.getFolderById(exportFolderId);

let existingFiles = [];
let files = exportFolder.getFiles();
while (files.hasNext()) { existingFiles.push(files.next().getName()); }

  statusCell.setBackground("#fbbc04").setFontColor("black");

  const url_base = ss.getUrl().replace(/\/edit$/, '/export?');
  const exportOptions = '&exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true' +
                        '&top_margin=0.2&bottom_margin=0.2&left_margin=0&right_margin=0' +
                        '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';
  
  const token = ScriptApp.getOAuthToken();

  for (let i = 0; i < allSheets.length; i++) {
    if (new Date().getTime() - startTime > 240000) { 
      createAutoTrigger();
      statusCell.setValue("Resuming (" + existingFiles.length + "/" + totalTarget + ")...");
      return; 
    }

    const sheet = allSheets[i];
    const sheetName = sheet.getName();

    if (excludedSheets.includes(sheetName)) continue;
    if (sheet.getRange("H4").getValue().toString().trim() === "") continue;
    if (existingFiles.includes(sheetName + ".pdf")) continue;

    try {
      statusCell.setValue("Now Generating " + (existingFiles.length + 1) + "/" + totalTarget + " PDFs...");
      SpreadsheetApp.flush();

      const url = url_base + exportOptions + '&gid=' + sheet.getSheetId();
      const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
      
      if (response.getResponseCode() === 200) {
        exportFolder.createFile(response.getBlob().setName(sheetName + ".pdf"));
        existingFiles.push(sheetName + ".pdf");
      }
    } catch (err) { console.log(err.message); }
  }

  if (existingFiles.length >= totalTarget) {
    statusCell.setValue("All " + existingFiles.length + " PDFs Generated").setBackground("#71a35e").setFontColor("white");
    deleteAutoTrigger();
  } else {
    createAutoTrigger();
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Admin Tools')
    .addItem('Clear All Data', 'showCustomConfirm') 
    .addToUi();
}

function showCustomConfirm() {
  const html = HtmlService.createHtmlOutput(
    '<html><head><style>' +
    'body { font-family: "Google Sans", Roboto, sans-serif; text-align: center; margin: 0; padding: 10px; overflow: hidden; background-color: #fff; }' +
    'h1 { color: #202124; margin: 5px 0 12px 0; font-size: 24px; }' + 
    'p { color: #5f6368; font-size: 18px; margin: 12px 0; }' + 
    '.warning { color: #d93025; font-weight: bold; font-size: 18px; margin: 12px 0 20px 0; }' +
    '.btn-container { display: flex; justify-content: center; gap: 10px; margin-top: 10px; }' +
    'button { padding: 8px 18px; border-radius: 4px; cursor: pointer; font-size: 18px; font-weight: 500; border: none; }' +
    '.btn-no { background-color: #fff; color: #3c4043; border: 1px solid #dadce0; }' +
    '.btn-yes { background-color: #c53929; color: #fff; }' +
    '*:focus { outline: none !important; }' +
    '</style></head><body>' +
    '<h1>⚠️ Delete Data? ⚠️</h1>' +
    '<p>Are you sure you want to delete all data? This action cannot be undone.</p>' +
    '<div class="btn-container">' +
      '<button class="btn-no" onclick="google.script.host.close()">No, Cancel</button>' +
      '<button class="btn-yes" onclick="runDelete()">Yes, I\'m sure</button>' +
    '</div>' +
    '<script>function runDelete(){google.script.run.withSuccessHandler(function(){google.script.host.close();}).resetWorkbookData();}</script>' +
    '</body></html>'
  ).setWidth(350).setHeight(210); 
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function resetWorkbookData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const markingSheet = ss.getSheetByName("Marking");
  const overviewSheet = ss.getSheetByName("Overview");

  if (markingSheet) {
    // Clear student and marking data
    markingSheet.getRange("C14:C43").clearContent();
    markingSheet.getRange("H14:H43").clearContent();
    markingSheet.getRange("P14:P43").clearContent();
    markingSheet.getRange("X14:FR43").clearContent();

    // Clear setup / status cells on Marking
    markingSheet.getRange("J2").clearContent(); // Student Papers Folder
    markingSheet.getRange("K4").clearContent(); // Answer Key File
    markingSheet.getRange("G6").clearContent(); // Last Gemini Error
    markingSheet.getRange("O10").clearContent(); // System Status

    // Restore defaults on Marking
    markingSheet.getRange("C14").setValue("Add Student Names Here...");
    markingSheet.getRange("E10").setValue("Standard");
  }

  if (overviewSheet) {
    // Clear and restore key setup fields
    overviewSheet.getRange("H2").clearContent().setValue("G?U? Knowledge and Skills Assessment (KSA) OR Cornerstone ? Assessment - NAME (Standards)");
    overviewSheet.getRange("C6").clearContent().setValue("Class Name");
    overviewSheet.getRange("I8").clearContent();
    overviewSheet.getRange("F10").clearContent();
    overviewSheet.getRange("D12").clearContent().setValue("T. Name");
    overviewSheet.getRange("E14").clearContent().setValue("G? U? KSA OR C#");
    overviewSheet.getRange("Q8").clearContent();

    // Clear question setup table
    overviewSheet.getRange("B17:J68").clearContent();

    // Restore default question numbers
    const defaultQuestions = [
      ["1"],
      ["2"],
      ["3a"],
      ["3b"],
      ["4"],
      ["5"],
      ["6"]
    ];
    overviewSheet.getRange("B17:B23").setValues(defaultQuestions);

    // Restore default marks available
    const defaultMarks = [
      [1],
      [1],
      [2],
      [1],
      [3],
      [1],
      [1]
    ];
    overviewSheet.getRange("F17:F23").setValues(defaultMarks);
  }

  clearStatusCell();
  deleteAutoTrigger();
  ss.toast("Workbook data successfully cleared.", "Success");