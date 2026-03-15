/*********************************
 * 04_gemini_pass1.gs
 * Gemini API setup and PDF test
 *********************************/

const GEMINI_API_KEY_PROPERTY = "GEMINI_API_KEY";

function setGeminiApiKeyOnce() {
  const apiKey = "PASTE_YOUR_API_KEY_HERE";
  PropertiesService.getScriptProperties().setProperty(GEMINI_API_KEY_PROPERTY, apiKey);
  SpreadsheetApp.getUi().alert("Gemini API key saved to Script Properties.");
}

function getGeminiApiKey_() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(GEMINI_API_KEY_PROPERTY);
  if (!apiKey) {
    throw new Error("Gemini API key not found in Script Properties.");
  }
  return apiKey;
}

function setGeminiStatus_(message, minSeconds = 3) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Marking");
  sheet.getRange("O10").setValue(message);
  SpreadsheetApp.flush();

  if (minSeconds > 0) {
    Utilities.sleep(minSeconds * 1000);
  }
}

function clearGeminiError_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Marking");
  sheet.getRange("G6").clearContent();
}

function setGeminiError_(message) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Marking");
  sheet.getRange("G6").setValue(message);
}

function formatBytes_(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function uploadFileToGemini_(blob, displayName, statusMessage) {
  const apiKey = getGeminiApiKey_();
  const numBytes = blob.getBytes().length;
  const sizeText = formatBytes_(numBytes);

  if (statusMessage) {
    setGeminiStatus_(`${statusMessage} (${sizeText})`);
  }

  const startResponse = UrlFetchApp.fetch(
    `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${encodeURIComponent(apiKey)}`,
    {
      method: "post",
      headers: {
        "X-Goog-Upload-Protocol": "resumable",
        "X-Goog-Upload-Command": "start",
        "X-Goog-Upload-Header-Content-Length": String(numBytes),
        "X-Goog-Upload-Header-Content-Type": blob.getContentType()
      },
      contentType: "application/json",
      payload: JSON.stringify({
        file: {
          display_name: displayName
        }
      }),
      muteHttpExceptions: true
    }
  );

  const uploadUrl =
    startResponse.getHeaders()["X-Goog-Upload-URL"] ||
    startResponse.getHeaders()["x-goog-upload-url"];

  if (!uploadUrl) {
    throw new Error("Gemini upload start failed: " + startResponse.getContentText());
  }

  const uploadResponse = UrlFetchApp.fetch(uploadUrl, {
    method: "post",
    headers: {
      "X-Goog-Upload-Offset": "0",
      "X-Goog-Upload-Command": "upload, finalize"
    },
    contentType: blob.getContentType(),
    payload: blob.getBytes(),
    muteHttpExceptions: true
  });

  const json = JSON.parse(uploadResponse.getContentText());

  if (!json.file || !json.file.uri) {
    throw new Error("Gemini upload finalize failed: " + uploadResponse.getContentText());
  }

  return json.file;
}

function testGeminiConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const markingSheet = ss.getSheetByName("Marking");

  try {
    clearGeminiError_();
    setGeminiStatus_("Connecting to Gemini...");

    const modelId = markingSheet.getRange("F8").getValue().toString().trim();
    if (!modelId) {
      throw new Error("Gemini model is blank in Marking!F8.");
    }

    const apiKey = getGeminiApiKey_();
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(modelId)}:generateContent?key=${encodeURIComponent(apiKey)}`;

    const payload = {
      contents: [
        {
          role: "user",
          parts: [
            { text: "Reply with exactly this text and nothing else: GEMINI CONNECTION OK" }
          ]
        }
      ],
      generationConfig: {
        temperature: 0
      }
    };

    setGeminiStatus_("Waiting for Gemini response...");

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseText = response.getContentText();

    if (response.getResponseCode() !== 200) {
      setGeminiError_(`Gemini connection error: ${responseText}`);
      setGeminiStatus_("Gemini connection failed.");
      throw new Error(`Gemini connection failed: ${responseText}`);
    }

    const json = JSON.parse(responseText);

    setGeminiStatus_("Gemini connection confirmed.", 0);
    clearGeminiError_();

  } catch (error) {
    setGeminiError_(error.message || String(error));
    if (markingSheet.getRange("O10").getValue().toString().trim() === "") {
      setGeminiStatus_("Gemini connection failed.");
    }
    throw error;
  }
}

function listGeminiModels() {
  clearGeminiError_();
  setGeminiStatus_("Listing Gemini models...");

  const apiKey = getGeminiApiKey_();
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${encodeURIComponent(apiKey)}`;

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true
  });

  const text = response.getContentText();

  if (response.getResponseCode() !== 200) {
    setGeminiError_("Failed to list models: " + text);
    setGeminiStatus_("Model listing failed.");
    throw new Error("Failed to list models: " + text);
  }

  const json = JSON.parse(text);
  const models = json.models.map(m => m.name).join("\n");

  setGeminiStatus_("Model listing complete.");
  clearGeminiError_();

  Logger.log(models);
  SpreadsheetApp.getUi().alert("Available Gemini models logged. Check Apps Script Logs.");
}

function testGeminiPDFRead() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Marking");

  try {
    clearGeminiError_();
    setGeminiStatus_("Preparing PDF read test...");

    const modelId = sheet.getRange("F8").getValue().toString().trim();
    const studentFolderLink = sheet.getRange("J2").getValue().toString().trim();
    const answerKeyLink = sheet.getRange("K4").getValue().toString().trim();

    if (!modelId) throw new Error("Gemini model missing in F8");
    if (!studentFolderLink) throw new Error("Student folder link missing in J2");
    if (!answerKeyLink) throw new Error("Answer key file link missing in K4");

    setGeminiStatus_("Finding student PDF...");
    const folderId = extractDriveFolderId_(studentFolderLink);
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();

    if (!files.hasNext()) {
      throw new Error("No student PDFs found in the folder.");
    }

    const studentPdf = files.next();

    setGeminiStatus_("Opening answer key...");
    const answerKeyFileId = extractDriveFileId_(answerKeyLink);
    const answerKeyFile = DriveApp.getFileById(answerKeyFileId);

    const uploadedStudent = uploadFileToGemini_(
      studentPdf.getBlob(),
      studentPdf.getName(),
      "Uploading student PDF..."
    );

    const uploadedAnswerKey = uploadFileToGemini_(
      answerKeyFile.getBlob(),
      answerKeyFile.getName(),
      "Uploading answer key..."
    );

    const apiKey = getGeminiApiKey_();
    const url =
      `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(modelId)}:generateContent?key=${encodeURIComponent(apiKey)}`;

    const payload = {
      contents: [
        {
          role: "user",
          parts: [
            { text: "Confirm you can read both PDFs. Reply only with: PDF READ SUCCESSFUL." },
            {
              file_data: {
                mime_type: uploadedStudent.mimeType || "application/pdf",
                file_uri: uploadedStudent.uri
              }
            },
            {
              file_data: {
                mime_type: uploadedAnswerKey.mimeType || "application/pdf",
                file_uri: uploadedAnswerKey.uri
              }
            }
          ]
        }
      ],
      generationConfig: {
        temperature: 0
      }
    };

    setGeminiStatus_("Waiting for Gemini response...");

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseText = response.getContentText();

    if (response.getResponseCode() !== 200) {
      setGeminiError_(responseText);
      setGeminiStatus_("Gemini PDF read failed.");
      throw new Error("Gemini PDF read failed: " + responseText);
    }

    const json = JSON.parse(responseText);

    setGeminiStatus_("PDF check complete.", 0);
    clearGeminiError_();

  } catch (error) {
    setGeminiError_(error.message || String(error));
    if (sheet.getRange("O10").getValue().toString().trim() === "") {
      setGeminiStatus_("Gemini PDF read failed.");
    }
    throw error;
  }
}