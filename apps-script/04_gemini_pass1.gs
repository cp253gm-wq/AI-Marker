/*********************************
 * 04_gemini_pass1.gs
 * Gemini Pass 1: setup only
 *********************************/

const GEMINI_API_KEY_PROPERTY = "GEMINI_API_KEY";

/**
 * Run once to save your Gemini API key into Script Properties.
 * After it works, remove your real key from this file.
 */
function setGeminiApiKeyOnce_() {
  const apiKey = "PASTE_YOUR_GEMINI_API_KEY_HERE";
  PropertiesService.getScriptProperties().setProperty(GEMINI_API_KEY_PROPERTY, apiKey);
  SpreadsheetApp.getUi().alert("Gemini API key saved to Script Properties.");
}

function getGeminiApiKey_() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(GEMINI_API_KEY_PROPERTY);
  if (!apiKey) {
    throw new Error("Gemini API key not found in Script Properties. Run setGeminiApiKeyOnce_() first.");
  }
  return apiKey;
}