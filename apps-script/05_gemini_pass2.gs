/*********************************
 * 05_gemini_pass2.gs
 * Gemini Pass 2 feedback and worked solutions
 *********************************/

function getPass1ResultsForRow_(sheet, row, questions) {
  const blockWidth = questions.length * 3;
  const rowValues = sheet.getRange(row, FIRST_QUESTION_MARK_COLUMN, 1, blockWidth).getValues()[0];
  const results = [];

  for (let i = 0; i < questions.length; i++) {
    const question = questions[i];
    const offset = i * 3;
    const rawMark = rowValues[offset];
    const feedback = String(rowValues[offset + 1] || "").trim();

    if (rawMark === "" || rawMark === null) {
      throw new Error(`Pass 1 mark is missing for ${question.question}.`);
    }

    if (typeof rawMark !== "number" || isNaN(rawMark)) {
      throw new Error(`Pass 1 mark is not numeric for ${question.question}.`);
    }

    if (rawMark < 0 || rawMark > question.max_mark) {
      throw new Error(`Pass 1 mark is out of range for ${question.question}.`);
    }

    results.push({
      question: question.question,
      max_mark: question.max_mark,
      mark: rawMark,
      pass1_feedback: feedback,
      full_marks_awarded: rawMark === question.max_mark
    });
  }

  return results;
}

function buildPass2Prompt_(studentName, studentNumber, mode, pass1Results) {
  const questionLines = pass1Results.map(function(result) {
    return [
      `Question: ${result.question}`,
      `Max mark: ${result.max_mark}`,
      `Awarded mark: ${result.mark}`,
      `Pass 1 evidence note: ${result.pass1_feedback || "(blank)"}`,
      `Full marks awarded: ${result.full_marks_awarded ? "yes" : "no"}`
    ].join("\n");
  }).join("\n\n");

  const workedSolutionExample = [
    "Worked solution example:",
    "2x - 3 = -(2/3)x + 5",
    "",
    "2x + (2/3)x = 5 + 3",
    "",
    "(8/3)x = 8",
    "",
    "x = 3",
    "",
    "y = 2(3) - 3",
    "",
    "y = 3",
    "",
    "Solution: (3, 3)"
  ].join("\n");

  return [
    "You are performing Pass 2 of a maths marking workflow using the supplied student paper and answer key PDFs.",
    `Student name: ${studentName}`,
    `Student number: ${studentNumber}`,
    `Marking mode: ${mode}`,
    "",
    "Use the existing Pass 1 marks and evidence notes as context.",
    "Do not change or reinterpret the awarded numeric marks.",
    "Your task is to produce final teacher-friendly and student-friendly written feedback only.",
    "",
    "Return:",
    "1. general_feedback for the whole paper",
    "2. one short feedback sentence for each question",
    "3. one worked solution for each question that needs correction",
    "",
    "General feedback rules:",
    "1. Supportive and encouraging.",
    "2. Maximum 2 to 3 sentences.",
    "3. No markdown.",
    "",
    "Question feedback rules:",
    "1. One short sentence where possible.",
    "2. No markdown.",
    "3. Refer to the student's next step clearly and kindly.",
    "",
    "Worked solution rules:",
    "1. Return only mathematical working, with no explanatory sentences.",
    "2. Put each algebraic step on a new line.",
    "3. Put one blank line between steps.",
    "4. Each line must contain exactly one mathematical step only.",
    "5. No two steps may appear on the same line.",
    "6. Do not compress multiple equations into one long line.",
    "7. Do not write prose such as First, Then, Therefore, So, or Because.",
    "8. Use clean mathematical notation such as √, ×, ÷, powers, and bracketed fractions where appropriate.",
    "9. The final answer must be on its own final line.",
    "10. Stop when the final answer is reached.",
    "11. If full marks were awarded for a question, working must be an empty string.",
    "12. If no worked solution is needed, working may be an empty string.",
    "13. Do not use bullets, markdown, or LaTeX delimiters.",
    "",
    workedSolutionExample,
    "",
    "Questions with Pass 1 context:",
    questionLines,
    "",
    "Output rules:",
    "1. Return valid JSON only.",
    "2. Include every question exactly once.",
    "3. question must match the provided question label exactly.",
    "4. general_feedback must be a string.",
    "5. feedback must be a string.",
    "6. working must be a string.",
    "7. Do not include any extra keys."
  ].join("\n");
}

function buildPass2ResponseSchema_(questions) {
  const allowedQuestions = questions.map(function(question) {
    return question.question;
  });

  return {
    type: "OBJECT",
    required: ["general_feedback", "questions"],
    propertyOrdering: ["general_feedback", "questions"],
    properties: {
      general_feedback: {
        type: "STRING"
      },
      questions: {
        type: "ARRAY",
        minItems: questions.length,
        maxItems: questions.length,
        items: {
          type: "OBJECT",
          required: ["question", "feedback", "working"],
          propertyOrdering: ["question", "feedback", "working"],
          properties: {
            question: {
              type: "STRING",
              enum: allowedQuestions
            },
            feedback: {
              type: "STRING"
            },
            working: {
              type: "STRING"
            }
          }
        }
      }
    }
  };
}

function normalizePass2Working_(workingText) {
  let normalized = String(workingText || "").replace(/\r\n?/g, "\n").trim();

  if (!normalized) {
    return "";
  }

  normalized = normalized.replace(/([^\n])\s*(Solution:)/g, "$1\n\n$2");
  normalized = normalized.replace(/([0-9A-Za-z\)])([A-Za-z]\s*=)/g, "$1\n\n$2");
  normalized = normalized.replace(
    /([0-9A-Za-z\)])((?:-?\d+(?:\.\d+)?|[A-Za-z(√][^\n=]{0,40}[+\-−×÷\/][^\n=]{0,40}=))/g,
    "$1\n\n$2"
  );

  return normalized;
}

function parseGeminiPass2Response_(responseText, pass1Results) {
  let parsed;

  try {
    parsed = JSON.parse(responseText);
  } catch (error) {
    throw new Error("Gemini returned invalid JSON for Pass 2.");
  }

  if (!parsed || typeof parsed.general_feedback !== "string") {
    throw new Error("Gemini Pass 2 JSON is missing general_feedback.");
  }

  if (!parsed.general_feedback.trim()) {
    throw new Error("Gemini Pass 2 returned blank general_feedback.");
  }

  if (!Array.isArray(parsed.questions)) {
    throw new Error("Gemini Pass 2 JSON is missing a questions array.");
  }

  const pass1Map = {};
  for (let i = 0; i < pass1Results.length; i++) {
    pass1Map[pass1Results[i].question] = pass1Results[i];
  }

  const resultMap = {};

  for (let j = 0; j < parsed.questions.length; j++) {
    const item = parsed.questions[j];

    if (!item || typeof item.question !== "string") {
      throw new Error("Gemini Pass 2 JSON contains a question entry with no valid question label.");
    }

    const questionKey = normalizeQuestionLabel_(item.question);
    const pass1Result = pass1Map[questionKey];

    if (!pass1Result) {
      throw new Error(`Gemini returned an unexpected question label in Pass 2: ${item.question}`);
    }

    if (typeof item.feedback !== "string") {
      throw new Error(`Gemini returned non-text feedback for ${questionKey}.`);
    }

    if (typeof item.working !== "string") {
      throw new Error(`Gemini returned non-text working for ${questionKey}.`);
    }

    const feedback = item.feedback.trim();
    const working = String(item.working || "");
    const normalizedWorking = normalizePass2Working_(working);

    if (!feedback) {
      throw new Error(`Gemini returned blank feedback for ${questionKey}.`);
    }

    if (pass1Result.full_marks_awarded && normalizedWorking !== "") {
      throw new Error(`Gemini returned worked solution text for full-mark question ${questionKey}.`);
    }

    resultMap[questionKey] = {
      feedback: feedback,
      working: normalizedWorking
    };
  }

  for (let k = 0; k < pass1Results.length; k++) {
    const questionLabel = pass1Results[k].question;
    if (!resultMap[questionLabel]) {
      throw new Error(`Gemini did not return Pass 2 data for ${questionLabel}.`);
    }
  }

  return {
    general_feedback: parsed.general_feedback.trim(),
    questions: resultMap
  };
}

function callGeminiPass2_(modelId, prompt, studentPdfFile, answerKeyFile, questions) {
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
      responseSchema: buildPass2ResponseSchema_(questions)
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
    throw new Error("Gemini Pass 2 request failed: " + responseText);
  }

  const responseJson = JSON.parse(responseText);
  return extractGeminiText_(responseJson);
}

function writePass2Results_(sheet, row, questions, parsedResponse) {
  setGeminiStatus_("Writing Pass 2 results...");

  sheet.getRange(row, GENERAL_FEEDBACK_COLUMN).setValue(parsedResponse.general_feedback);

  for (let i = 0; i < questions.length; i++) {
    const questionLabel = questions[i].question;
    const pass2Result = parsedResponse.questions[questionLabel];
    const feedbackColumn = FIRST_QUESTION_MARK_COLUMN + (i * 3) + 1;
    const workingColumn = FIRST_QUESTION_MARK_COLUMN + (i * 3) + 2;

    sheet.getRange(row, feedbackColumn).setValue(pass2Result.feedback);
    sheet.getRange(row, workingColumn).setValue(pass2Result.working);
  }
}

function markStudentPass2(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MARKING_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet "${MARKING_SHEET_NAME}" not found.`);
  }

  try {
    clearGeminiError_();
    setGeminiStatus_("Preparing Pass 2...");

    if (row < FIRST_STUDENT_ROW || row > LAST_STUDENT_ROW) {
      throw new Error(`Student row must be between ${FIRST_STUDENT_ROW} and ${LAST_STUDENT_ROW}.`);
    }

    const studentNumber = sheet.getRange(row, STUDENT_NUMBER_COLUMN).getDisplayValue().toString().trim();
    const studentName = sheet.getRange(row, STUDENT_NAME_COLUMN).getDisplayValue().toString().trim();
    const mode = sheet.getRange(row, MODE_COLUMN).getDisplayValue().toString().trim();
    const studentFolderLink = sheet.getRange(STUDENT_FOLDER_LINK_CELL).getDisplayValue().toString().trim();
    const answerKeyLink = sheet.getRange(ANSWER_KEY_LINK_CELL).getDisplayValue().toString().trim();
    const modelId = sheet.getRange(MODEL_ID_CELL).getDisplayValue().toString().trim();

    if (!studentName) throw new Error("Student name is blank in the selected row.");
    if (!studentNumber) throw new Error("Student number is blank in the selected row.");
    if (!mode) throw new Error("Marking mode is blank in the selected row. Pass 1 must run first.");
    if (!studentFolderLink) throw new Error("Student folder link is blank in Marking!J2.");
    if (!answerKeyLink) throw new Error("Answer key file link is blank in Marking!K4.");
    if (!modelId) throw new Error("Gemini model is blank in Marking!F8.");

    const questions = getQuestionStructure_();

    setGeminiStatus_("Reading Pass 1 results...");
    const pass1Results = getPass1ResultsForRow_(sheet, row, questions);

    setGeminiStatus_("Finding student PDF...");
    const studentPdfFile = findStudentPdfFile_(studentFolderLink, studentName);

    setGeminiStatus_("Opening answer key...");
    const answerKeyFile = getAnswerKeyPdfFile_(answerKeyLink);

    setGeminiStatus_("Building Pass 2 prompt...");
    const prompt = buildPass2Prompt_(studentName, studentNumber, mode, pass1Results);
    const responseText = callGeminiPass2_(modelId, prompt, studentPdfFile, answerKeyFile, questions);
    const parsedResponse = parseGeminiPass2Response_(responseText, pass1Results);

    writePass2Results_(sheet, row, questions, parsedResponse);
    clearGeminiError_();
    setGeminiStatus_("Pass 2 complete.", 0);
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    setGeminiError_(message);
    setGeminiStatus_("Pass 2 failed.", 0);
    throw error;
  }
}
