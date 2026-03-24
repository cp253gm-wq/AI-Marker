/*********************************
 * 04_gemini_pass1.gs
 * Gemini API setup and Pass 1 marking
 *********************************/

const MARKING_SHEET_NAME = "Marking";
const OVERVIEW_SHEET_NAME = "Overview";
const GEMINI_API_KEY_CELL = "U69";
const MARKING_STATUS_CELL = "O10";
const GEMINI_ERROR_CELL = "G6";
const LAST_STUDENT_MARKED_CELL = "Q8";
const STUDENT_FOLDER_LINK_CELL = "J2";
const ANSWER_KEY_LINK_CELL = "K4";
const MODEL_ID_CELL = "F8";
const FIRST_STUDENT_ROW = 14;
const LAST_STUDENT_ROW = 43;
const STUDENT_NUMBER_COLUMN = 2;
const STUDENT_NAME_COLUMN = 3;
const MODE_COLUMN = 8;
const TIMESTAMP_COLUMN = 16;
const GENERAL_FEEDBACK_COLUMN = 24;
const FIRST_QUESTION_MARK_COLUMN = 25;
const OVERVIEW_MARKS_RANGE = "F17:F68";
const OVERVIEW_LABELS_RANGE = "K17:K68";

function getGeminiApiKey_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const overviewSheet = ss.getSheetByName(OVERVIEW_SHEET_NAME);

  if (!overviewSheet) {
    throw new Error(`Sheet "${OVERVIEW_SHEET_NAME}" not found.`);
  }

  const sheetKey = overviewSheet.getRange(GEMINI_API_KEY_CELL).getValue().toString().trim();

  if (sheetKey !== "") {
    return sheetKey;
  }

  throw new Error(`Gemini API key not found. Please add it to Overview!${GEMINI_API_KEY_CELL}.`);
}

function setGeminiStatus_(message, minSeconds) {
  const delaySeconds = minSeconds === undefined ? 3 : minSeconds;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MARKING_SHEET_NAME);
  sheet.getRange(MARKING_STATUS_CELL).setValue(message);
  SpreadsheetApp.flush();

  if (delaySeconds > 0) {
    Utilities.sleep(delaySeconds * 1000);
  }
}

function clearGeminiError_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MARKING_SHEET_NAME);
  sheet.getRange(GEMINI_ERROR_CELL).clearContent();
}

function setGeminiError_(message) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MARKING_SHEET_NAME);
  sheet.getRange(GEMINI_ERROR_CELL).setValue(message);
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

function normalizeQuestionLabel_(label) {
  return String(label || "")
    .trim()
    .replace(/\)+$/g, "");
}

function getQuestionStructure_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const overviewSheet = ss.getSheetByName(OVERVIEW_SHEET_NAME);

  if (!overviewSheet) {
    throw new Error(`Sheet "${OVERVIEW_SHEET_NAME}" not found.`);
  }

  const labels = overviewSheet.getRange(OVERVIEW_LABELS_RANGE).getDisplayValues();
  const maxMarks = overviewSheet.getRange(OVERVIEW_MARKS_RANGE).getValues();
  const questions = [];

  for (let i = 0; i < labels.length; i++) {
    const normalizedLabel = normalizeQuestionLabel_(labels[i][0]);
    const maxMark = Number(maxMarks[i][0]);

    if (!normalizedLabel) continue;
    if (isNaN(maxMark)) {
      throw new Error(`Overview mark value is not numeric for ${normalizedLabel}.`);
    }

    questions.push({
      question: normalizedLabel,
      max_mark: maxMark
    });
  }

  if (questions.length === 0) {
    throw new Error("No question structure found on the Overview sheet.");
  }

  return questions;
}

function findStudentPdfFile_(folderLink, studentName) {
  const folderId = extractDriveFolderId_(folderLink);
  const folder = DriveApp.getFolderById(folderId);
  const expectedFileName = `${studentName}.pdf`;
  const files = folder.getFilesByName(expectedFileName);

  if (!files.hasNext()) {
    throw new Error(`Student PDF not found for ${studentName}. Expected file name: ${expectedFileName}`);
  }

  const studentFile = files.next();

  if (files.hasNext()) {
    throw new Error(`Multiple PDFs found for ${studentName}. Please keep only one file named ${expectedFileName}`);
  }

  return studentFile;
}

function getAnswerKeyPdfFile_(answerKeyLink) {
  const answerKeyFileId = extractDriveFileId_(answerKeyLink);
  return DriveApp.getFileById(answerKeyFileId);
}

function buildPass1Prompt_(studentName, studentNumber, mode, questions) {
  const questionLines = questions.map(function(question) {
    return `${question.question} (max ${question.max_mark})`;
  }).join("\n");

  const markingModeInstruction = mode === "Compassionate"
    ? "Compassionate mode must still be evidence-based, but be slightly more generous on borderline cases where visible mathematical understanding deserves credit."
    : "Standard mode must still award fair partial credit whenever visible evidence justifies it.";

  return [
    "You are marking a student's paper against the supplied answer key.",
    `Student name: ${studentName}`,
    `Student number: ${studentNumber}`,
    `Marking mode: ${mode}`,
    "",
    "Return Pass 1 results only.",
    "For each question, decide a numeric mark and one short evidence note.",
    "Do not include polished general feedback.",
    "Do not include worked solutions.",
    "Use the answer key as the marking reference.",
    markingModeInstruction,
    "",
    "Partial credit rules:",
    "1. Half marks such as 0.5 are allowed when justified by visible evidence.",
    "2. Half marks are not reserved only for Compassionate mode.",
    "3. In Standard mode, award fair partial credit for partially correct method, correct mathematical setup with an arithmetic slip, valid intermediate reasoning that deserves some credit, and partially completed but mathematically meaningful progress.",
    "4. In Compassionate mode, stay evidence-based but be slightly more generous on borderline cases.",
    "5. Avoid binary all-or-nothing marking when visible partial credit is deserved.",
    "6. Award the most defensible numeric mark supported by the student's actual work.",
    "",
    "Questions and maximum marks:",
    questionLines,
    "",
    "Output rules:",
    "1. Return valid JSON only.",
    "2. Include every question exactly once.",
    "3. question must match one of the provided question labels exactly.",
    "4. mark must be a JSON number, not a string.",
    "5. evidence_note must be short and evidence-based.",
    "6. Marks such as 0.5 are valid when justified by evidence.",
    "7. Do not use formats such as 1/2 or 2 out of 3.",
    "8. Keep marks within the maximum available for each question."
  ].join("\n");
}

function buildPass1ResponseSchema_(questions) {
  const allowedQuestions = questions.map(function(question) {
    return question.question;
  });

  return {
    type: "OBJECT",
    required: ["results"],
    propertyOrdering: ["results"],
    properties: {
      results: {
        type: "ARRAY",
        minItems: questions.length,
        maxItems: questions.length,
        items: {
          type: "OBJECT",
          required: ["question", "mark", "evidence_note"],
          propertyOrdering: ["question", "mark", "evidence_note"],
          properties: {
            question: {
              type: "STRING",
              enum: allowedQuestions
            },
            mark: {
              type: "NUMBER"
            },
            evidence_note: {
              type: "STRING"
            }
          }
        }
      }
    }
  };
}

function extractGeminiText_(responseJson) {
  if (
    responseJson &&
    responseJson.candidates &&
    responseJson.candidates.length > 0 &&
    responseJson.candidates[0].content &&
    responseJson.candidates[0].content.parts &&
    responseJson.candidates[0].content.parts.length > 0
  ) {
    for (let i = 0; i < responseJson.candidates[0].content.parts.length; i++) {
      const part = responseJson.candidates[0].content.parts[i];
      if (part.text) {
        return part.text;
      }
    }
  }

  throw new Error("Gemini returned no text content.");
}

function parseGeminiPass1Response_(responseText, questions) {
  let parsed;

  try {
    parsed = JSON.parse(responseText);
  } catch (error) {
    throw new Error("Gemini returned invalid JSON for Pass 1.");
  }

  if (!parsed || !parsed.results || !Array.isArray(parsed.results)) {
    throw new Error("Gemini Pass 1 JSON is missing a results array.");
  }

  const allowedMap = {};
  for (let i = 0; i < questions.length; i++) {
    allowedMap[questions[i].question] = questions[i];
  }

  const resultMap = {};

  for (let j = 0; j < parsed.results.length; j++) {
    const item = parsed.results[j];
    if (!item || typeof item.question !== "string") {
      throw new Error("Gemini Pass 1 JSON contains a result with no valid question label.");
    }

    const questionKey = normalizeQuestionLabel_(item.question);
    const questionConfig = allowedMap[questionKey];

    if (!questionConfig) {
      throw new Error(`Gemini returned an unexpected question label: ${item.question}`);
    }

    if (typeof item.mark !== "number" || isNaN(item.mark)) {
      throw new Error(`Gemini returned a non-numeric mark for ${questionKey}.`);
    }

    if (item.mark < 0 || item.mark > questionConfig.max_mark) {
      throw new Error(`Gemini returned an out-of-range mark for ${questionKey}.`);
    }

    resultMap[questionKey] = {
      mark: item.mark,
      evidence_note: item.evidence_note ? String(item.evidence_note).trim() : ""
    };
  }

  for (let k = 0; k < questions.length; k++) {
    const questionLabel = questions[k].question;
    if (!resultMap[questionLabel]) {
      throw new Error(`Gemini did not return a result for ${questionLabel}.`);
    }
  }

  return resultMap;
}

function callGeminiPass1_(modelId, prompt, studentPdfFile, answerKeyFile, questions) {
  const uploadedStudent = uploadFileToGemini_(
    studentPdfFile.getBlob(),
    studentPdfFile.getName(),
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
          { text: prompt },
          {
            file_data: {
              mime_type: uploadedAnswerKey.mimeType || "application/pdf",
              file_uri: uploadedAnswerKey.uri
            }
          },
          {
            file_data: {
              mime_type: uploadedStudent.mimeType || "application/pdf",
              file_uri: uploadedStudent.uri
            }
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0,
      responseMimeType: "application/json",
      responseSchema: buildPass1ResponseSchema_(questions)
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
    throw new Error("Gemini Pass 1 request failed: " + responseText);
  }

  const responseJson = JSON.parse(responseText);
  const jsonText = extractGeminiText_(responseJson);
  return parseGeminiPass1Response_(jsonText, questions);
}

function applyCompassionateMarkFloor_(sheet, row, questions, resultMap) {
  const existingValues = sheet
    .getRange(row, FIRST_QUESTION_MARK_COLUMN, 1, questions.length * 3)
    .getValues()[0];
  const mergedResultMap = {};

  for (let i = 0; i < questions.length; i++) {
    const questionLabel = questions[i].question;
    const newResult = resultMap[questionLabel];
    const existingMark = existingValues[i * 3];
    const finalMark =
      typeof existingMark === "number" && !isNaN(existingMark)
        ? Math.max(existingMark, newResult.mark)
        : newResult.mark;

    mergedResultMap[questionLabel] = {
      mark: finalMark,
      evidence_note: newResult.evidence_note
    };
  }

  return mergedResultMap;
}

function writePass1Results_(sheet, row, mode, studentNumber, studentName, questions, resultMap) {
  const blockWidth = questions.length * 3;
  const rowValues = [];

  for (let i = 0; i < questions.length; i++) {
    const questionLabel = questions[i].question;
    const result = resultMap[questionLabel];
    rowValues.push(result.mark, result.evidence_note, "");
  }

  sheet.getRange(row, GENERAL_FEEDBACK_COLUMN).setValue("PENDING");
  sheet.getRange(row, FIRST_QUESTION_MARK_COLUMN, 1, blockWidth).setValues([rowValues]);
  sheet.getRange(row, MODE_COLUMN).setValue(mode);
  sheet.getRange(row, TIMESTAMP_COLUMN).setValue(new Date());
  sheet.getRange(LAST_STUDENT_MARKED_CELL).setValue(`${studentNumber} - ${studentName}`);
}

function markStudentPass1(row, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MARKING_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet "${MARKING_SHEET_NAME}" not found.`);
  }

  try {
    clearGeminiError_();
    setGeminiStatus_("Connecting to Gemini...");

    if (row < FIRST_STUDENT_ROW || row > LAST_STUDENT_ROW) {
      throw new Error(`Student row must be between ${FIRST_STUDENT_ROW} and ${LAST_STUDENT_ROW}.`);
    }

    const studentNumber = sheet.getRange(row, STUDENT_NUMBER_COLUMN).getDisplayValue().toString().trim();
    const studentName = sheet.getRange(row, STUDENT_NAME_COLUMN).getDisplayValue().toString().trim();
    const studentFolderLink = sheet.getRange(STUDENT_FOLDER_LINK_CELL).getDisplayValue().toString().trim();
    const answerKeyLink = sheet.getRange(ANSWER_KEY_LINK_CELL).getDisplayValue().toString().trim();
    const modelId = sheet.getRange(MODEL_ID_CELL).getDisplayValue().toString().trim();

    if (!studentName) throw new Error("Student name is blank in the selected row.");
    if (!studentNumber) throw new Error("Student number is blank in the selected row.");
    if (!studentFolderLink) throw new Error("Student folder link is blank in Marking!J2.");
    if (!answerKeyLink) throw new Error("Answer key file link is blank in Marking!K4.");
    if (!modelId) throw new Error("Gemini model is blank in Marking!F8.");

    const questions = getQuestionStructure_();
    const studentPdfFile = findStudentPdfFile_(studentFolderLink, studentName);
    const answerKeyFile = getAnswerKeyPdfFile_(answerKeyLink);
    const prompt = buildPass1Prompt_(studentName, studentNumber, mode, questions);
    const rawResultMap = callGeminiPass1_(modelId, prompt, studentPdfFile, answerKeyFile, questions);
    const resultMap = mode === "Compassionate"
      ? applyCompassionateMarkFloor_(sheet, row, questions, rawResultMap)
      : rawResultMap;

    writePass1Results_(sheet, row, mode, studentNumber, studentName, questions, resultMap);
    clearGeminiError_();
    setGeminiStatus_("Pass 1 complete.", 0);
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    setGeminiError_(message);
    setGeminiStatus_("Pass 1 failed.", 0);
    throw error;
  }
}

function testGeminiConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const markingSheet = ss.getSheetByName(MARKING_SHEET_NAME);

  try {
    clearGeminiError_();
    setGeminiStatus_("Connecting to Gemini...");

    const modelId = markingSheet.getRange(MODEL_ID_CELL).getValue().toString().trim();
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

    JSON.parse(responseText);

    setGeminiStatus_("Gemini connection confirmed.", 0);
    clearGeminiError_();
  } catch (error) {
    setGeminiError_(error.message || String(error));
    if (markingSheet.getRange(MARKING_STATUS_CELL).getValue().toString().trim() === "") {
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
  const models = json.models.map(function(model) {
    return model.name;
  }).join("\n");

  setGeminiStatus_("Model listing complete.");
  clearGeminiError_();

  Logger.log(models);
  SpreadsheetApp.getUi().alert("Available Gemini models logged. Check Apps Script Logs.");
}

function testGeminiPDFRead() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MARKING_SHEET_NAME);

  try {
    clearGeminiError_();
    setGeminiStatus_("Preparing PDF read test...");

    const modelId = sheet.getRange(MODEL_ID_CELL).getValue().toString().trim();
    const studentFolderLink = sheet.getRange(STUDENT_FOLDER_LINK_CELL).getValue().toString().trim();
    const answerKeyLink = sheet.getRange(ANSWER_KEY_LINK_CELL).getValue().toString().trim();

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
    const answerKeyFile = getAnswerKeyPdfFile_(answerKeyLink);

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

    JSON.parse(responseText);

    setGeminiStatus_("PDF check complete.", 0);
    clearGeminiError_();
  } catch (error) {
    setGeminiError_(error.message || String(error));
    if (sheet.getRange(MARKING_STATUS_CELL).getValue().toString().trim() === "") {
      setGeminiStatus_("Gemini PDF read failed.");
    }
    throw error;
  }
}