function linkStudentNamesToReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const markingSheet = ss.getSheetByName("Marking");
  const allSheets = ss.getSheets();
  const excludedSheets = ["Overview", "Marking"];

  const startRow = 14;
  const endRow = 43;

  for (let row = startRow; row <= endRow; row++) {
    const studentNumber = markingSheet.getRange(row, 2).getValue().toString().trim(); // Column B
    const nameCell = markingSheet.getRange(row, 3); // Column C
    const studentName = nameCell.getValue().toString().trim();

    if (!studentNumber || !studentName) {
      nameCell.setRichTextValue(
        SpreadsheetApp.newRichTextValue().setText(studentName).build()
      );
      continue;
    }

    let targetSheet = null;

    for (let i = 0; i < allSheets.length; i++) {
      const sheet = allSheets[i];
      if (excludedSheets.includes(sheet.getName())) continue;

      const a1Value = sheet.getRange("A1").getValue().toString().trim();
      if (a1Value === studentNumber) {
        targetSheet = sheet;
        break;
      }
    }

    if (!targetSheet) {
      nameCell.setRichTextValue(
        SpreadsheetApp.newRichTextValue().setText(studentName).build()
      );
      continue;
    }

    const gid = targetSheet.getSheetId();
    const link = ss.getUrl() + `#gid=${gid}&range=R1`;

    const richText = SpreadsheetApp.newRichTextValue()
      .setText(studentName)
      .setLinkUrl(link)
      .build();

    nameCell.setRichTextValue(richText);
  }
}

function markStudentAtRow(row, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Marking");

  const studentName = sheet.getRange(row, 3).getValue().toString().trim(); // Column C

  if (!studentName) {
    SpreadsheetApp.getUi().alert("No student name in this row.");
    return;
  }

  const existingData = sheet.getRange(`X${row}:FR${row}`).getValues()[0];
  const hasExistingMarking = existingData.some(value => value !== "");

  if (hasExistingMarking) {
    const overwrite = SpreadsheetApp.getUi().alert(
      "Overwrite existing marks?",
      "Are you sure you want to overwrite existing question data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (overwrite !== SpreadsheetApp.getUi().Button.YES) return;
  }

  SpreadsheetApp.getActive().toast(
    "Marking started for " + studentName,
    "Marking System"
  );

  markStudentPass1(row, mode);
  markStudentPass2(row);
  resizeSingleRowAndCentreButtons(row);
}

function markRow14() { markStudentAtRow(14, "Standard"); }
function remarkRow14() { markStudentAtRow(14, "Compassionate"); }

function markRow15() { markStudentAtRow(15, "Standard"); }
function remarkRow15() { markStudentAtRow(15, "Compassionate"); }

function markRow16() { markStudentAtRow(16, "Standard"); }
function remarkRow16() { markStudentAtRow(16, "Compassionate"); }

function markRow17() { markStudentAtRow(17, "Standard"); }
function remarkRow17() { markStudentAtRow(17, "Compassionate"); }

function markRow18() { markStudentAtRow(18, "Standard"); }
function remarkRow18() { markStudentAtRow(18, "Compassionate"); }

function markRow19() { markStudentAtRow(19, "Standard"); }
function remarkRow19() { markStudentAtRow(19, "Compassionate"); }

function markRow20() { markStudentAtRow(20, "Standard"); }
function remarkRow20() { markStudentAtRow(20, "Compassionate"); }

function markRow21() { markStudentAtRow(21, "Standard"); }
function remarkRow21() { markStudentAtRow(21, "Compassionate"); }

function markRow22() { markStudentAtRow(22, "Standard"); }
function remarkRow22() { markStudentAtRow(22, "Compassionate"); }

function markRow23() { markStudentAtRow(23, "Standard"); }
function remarkRow23() { markStudentAtRow(23, "Compassionate"); }

function markRow24() { markStudentAtRow(24, "Standard"); }
function remarkRow24() { markStudentAtRow(24, "Compassionate"); }

function markRow25() { markStudentAtRow(25, "Standard"); }
function remarkRow25() { markStudentAtRow(25, "Compassionate"); }

function markRow26() { markStudentAtRow(26, "Standard"); }
function remarkRow26() { markStudentAtRow(26, "Compassionate"); }

function markRow27() { markStudentAtRow(27, "Standard"); }
function remarkRow27() { markStudentAtRow(27, "Compassionate"); }

function markRow28() { markStudentAtRow(28, "Standard"); }
function remarkRow28() { markStudentAtRow(28, "Compassionate"); }

function markRow29() { markStudentAtRow(29, "Standard"); }
function remarkRow29() { markStudentAtRow(29, "Compassionate"); }

function markRow30() { markStudentAtRow(30, "Standard"); }
function remarkRow30() { markStudentAtRow(30, "Compassionate"); }

function markRow31() { markStudentAtRow(31, "Standard"); }
function remarkRow31() { markStudentAtRow(31, "Compassionate"); }

function markRow32() { markStudentAtRow(32, "Standard"); }
function remarkRow32() { markStudentAtRow(32, "Compassionate"); }

function markRow33() { markStudentAtRow(33, "Standard"); }
function remarkRow33() { markStudentAtRow(33, "Compassionate"); }

function markRow34() { markStudentAtRow(34, "Standard"); }
function remarkRow34() { markStudentAtRow(34, "Compassionate"); }

function markRow35() { markStudentAtRow(35, "Standard"); }
function remarkRow35() { markStudentAtRow(35, "Compassionate"); }

function markRow36() { markStudentAtRow(36, "Standard"); }
function remarkRow36() { markStudentAtRow(36, "Compassionate"); }

function markRow37() { markStudentAtRow(37, "Standard"); }
function remarkRow37() { markStudentAtRow(37, "Compassionate"); }

function markRow38() { markStudentAtRow(38, "Standard"); }
function remarkRow38() { markStudentAtRow(38, "Compassionate"); }

function markRow39() { markStudentAtRow(39, "Standard"); }
function remarkRow39() { markStudentAtRow(39, "Compassionate"); }

function markRow40() { markStudentAtRow(40, "Standard"); }
function remarkRow40() { markStudentAtRow(40, "Compassionate"); }

function markRow41() { markStudentAtRow(41, "Standard"); }
function remarkRow41() { markStudentAtRow(41, "Compassionate"); }

function markRow42() { markStudentAtRow(42, "Standard"); }
function remarkRow42() { markStudentAtRow(42, "Compassionate"); }

function markRow43() { markStudentAtRow(43, "Standard"); }
function remarkRow43() { markStudentAtRow(43, "Compassionate"); }
