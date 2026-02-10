/***********************
 * AccuGrader Universal - Free Grading Tool
 * 
 * Features:
 * - Multiple Choice: Auto-grading (FREE, no API needed)
 * - Essay: AI-powered grading (requires OpenAI API key)
 * 
 * Setup:
 * 1. For Multiple Choice: Just use! No setup needed.
 * 2. For Essay Grading: Add your OpenAI API key in Script Properties
 *    - Extensions > Apps Script > Project Settings > Script Properties
 *    - Add: OPENAI_API_KEY = your_key_here
 * 
 * v1.0 - Universal Template
 ***********************/

/***************
 * Configuration
 ***************/
const CONFIG = {
  // Sheet names
  MC_SUBMISSIONS: "MC_Submissions",
  MC_ANSWER_KEY: "MC_AnswerKey",
  ESSAY_SUBMISSIONS: "Essay_Submissions",
  ESSAY_RUBRIC: "Essay_Rubric",
  
  // API settings
  MODEL: "gpt-4o-mini",
  TEMPERATURE: 0.2,
  MAX_TOKENS: 1500,
};

/***************
 * Menu Creation
 ***************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📝 AccuGrader")
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Multiple Choice")
      .addItem("Grade selected student", "gradeMCSelected")
      .addItem("Grade ALL students", "gradeMCAll")
      .addItem("Reset selected student", "resetMCSelected"))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Essay")
      .addItem("Grade selected student", "gradeEssaySelected")
      .addItem("Grade ALL students", "gradeEssayAll")
      .addItem("Reset selected student", "resetEssaySelected"))
    .addSeparator()
    .addItem("⚙️ Test API Connection", "testAPIConnection")
    .addItem("ℹ️ Help / Instructions", "showHelp")
    .addToUi();
}

function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <h2>AccuGrader Universal</h2>
    <h3>Multiple Choice (FREE)</h3>
    <ol>
      <li>Enter answers in MC_AnswerKey sheet</li>
      <li>Enter student responses in MC_Submissions sheet</li>
      <li>Run: AccuGrader > Multiple Choice > Grade ALL students</li>
    </ol>
    <h3>Essay Grading (requires OpenAI API)</h3>
    <ol>
      <li>Get API key from <a href="https://platform.openai.com/api-keys" target="_blank">platform.openai.com</a></li>
      <li>Go to Extensions > Apps Script > Project Settings > Script Properties</li>
      <li>Add property: OPENAI_API_KEY = your_key</li>
      <li>Customize rubric in Essay_Rubric sheet</li>
      <li>Enter essays in Essay_Submissions sheet</li>
      <li>Run: AccuGrader > Essay > Grade ALL students</li>
    </ol>
    <p><strong>Cost:</strong> ~$0.01-0.03 per essay (GPT-4o-mini)</p>
  `)
    .setWidth(450)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "AccuGrader Help");
}

function testAPIConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const key = getOpenAIKey_();
    if (!key) {
      ui.alert("❌ No API Key", "OpenAI API key not found.\n\nFor Essay grading, add your key:\nExtensions > Apps Script > Project Settings > Script Properties\nAdd: OPENAI_API_KEY = your_key\n\nMultiple Choice grading works without an API key!", ui.ButtonSet.OK);
      return;
    }
    // Quick test call
    const response = callOpenAI_("Say 'API connection successful' in exactly those words.");
    if (response.toLowerCase().includes("successful")) {
      ui.alert("✅ Success", "OpenAI API connection is working!\n\nYou can now use Essay grading.", ui.ButtonSet.OK);
    } else {
      ui.alert("✅ Connected", "API responded: " + response.substring(0, 100), ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert("❌ Error", "API connection failed:\n" + e.message, ui.ButtonSet.OK);
  }
}

/***********************************************
 * MULTIPLE CHOICE GRADING (FREE - No API)
 ***********************************************/

function gradeMCSelected() {
  const sheet = getSheet_(CONFIG.MC_SUBMISSIONS);
  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("Please select a student row (not the header).");
    return;
  }
  
  try {
    gradeMCRow_(sheet, row);
    SpreadsheetApp.getUi().alert("✅ Grading complete!");
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error: " + e.message);
  }
}

function gradeMCAll() {
  const sheet = getSheet_(CONFIG.MC_SUBMISSIONS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert("No students to grade.");
    return;
  }
  
  const answerKey = getMCAnswerKey_();
  if (Object.keys(answerKey).length === 0) {
    SpreadsheetApp.getUi().alert("❌ Error: No answers found in MC_AnswerKey sheet.");
    return;
  }
  
  let graded = 0;
  let skipped = 0;
  
  for (let row = 2; row <= lastRow; row++) {
    const status = String(sheet.getRange(row, getColIndex_(sheet, "Status")).getValue()).trim().toUpperCase();
    if (status === "DONE") {
      skipped++;
      continue;
    }
    
    const name = String(sheet.getRange(row, getColIndex_(sheet, "Student Name")).getValue()).trim();
    if (!name) {
      skipped++;
      continue;
    }
    
    try {
      gradeMCRow_(sheet, row, answerKey);
      graded++;
    } catch (e) {
      sheet.getRange(row, getColIndex_(sheet, "Status")).setValue("ERROR");
    }
  }
  
  SpreadsheetApp.getUi().alert(`✅ Grading complete!\n\nGraded: ${graded}\nSkipped: ${skipped}`);
}

function gradeMCRow_(sheet, row, answerKey) {
  if (!answerKey) {
    answerKey = getMCAnswerKey_();
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const questionCols = [];
  
  // Find question columns (Q1, Q2, Q3, etc.)
  headers.forEach((h, idx) => {
    if (/^Q\d+$/i.test(String(h).trim())) {
      questionCols.push({ col: idx + 1, question: String(h).trim().toUpperCase() });
    }
  });
  
  let correct = 0;
  let total = 0;
  let totalPoints = 0;
  let earnedPoints = 0;
  const details = [];
  
  questionCols.forEach(qc => {
    const studentAnswer = String(sheet.getRange(row, qc.col).getValue()).trim().toUpperCase();
    const keyData = answerKey[qc.question];
    
    if (keyData) {
      total++;
      const correctAnswer = keyData.answer;
      const points = keyData.points || 1;
      totalPoints += points;
      
      if (studentAnswer === correctAnswer) {
        correct++;
        earnedPoints += points;
        details.push(`${qc.question}: ✓`);
      } else {
        details.push(`${qc.question}: ✗ (Your: ${studentAnswer || '-'}, Correct: ${correctAnswer})`);
      }
    }
  });
  
  // Write results
  const scoreCol = getColIndex_(sheet, "Score");
  const percentCol = getColIndex_(sheet, "Percentage");
  const correctCol = getColIndex_(sheet, "Correct");
  const totalCol = getColIndex_(sheet, "Total Questions");
  const feedbackCol = getColIndex_(sheet, "Feedback");
  const statusCol = getColIndex_(sheet, "Status");
  
  if (scoreCol) sheet.getRange(row, scoreCol).setValue(earnedPoints + "/" + totalPoints);
  if (percentCol) sheet.getRange(row, percentCol).setValue(total > 0 ? Math.round((earnedPoints / totalPoints) * 100) + "%" : "0%");
  if (correctCol) sheet.getRange(row, correctCol).setValue(correct);
  if (totalCol) sheet.getRange(row, totalCol).setValue(total);
  if (feedbackCol) sheet.getRange(row, feedbackCol).setValue(details.join("\n"));
  if (statusCol) sheet.getRange(row, statusCol).setValue("DONE");
}

function getMCAnswerKey_() {
  const sheet = getSheet_(CONFIG.MC_ANSWER_KEY);
  const data = sheet.getDataRange().getValues();
  const answerKey = {};
  
  // Expected format: Question | Correct Answer | Points
  for (let i = 1; i < data.length; i++) {
    const question = String(data[i][0] || "").trim().toUpperCase();
    const answer = String(data[i][1] || "").trim().toUpperCase();
    const points = Number(data[i][2]) || 1;
    
    if (question && answer) {
      answerKey[question] = { answer, points };
    }
  }
  
  return answerKey;
}

function resetMCSelected() {
  const sheet = getSheet_(CONFIG.MC_SUBMISSIONS);
  const row = sheet.getActiveCell().getRow();
  if (row <= 1) return;
  
  const colsToClear = ["Score", "Percentage", "Correct", "Total Questions", "Feedback", "Status"];
  colsToClear.forEach(colName => {
    const col = getColIndex_(sheet, colName);
    if (col) sheet.getRange(row, col).clearContent();
  });
  
  SpreadsheetApp.getUi().alert("✅ Student reset complete.");
}

/***********************************************
 * ESSAY GRADING (Requires OpenAI API Key)
 ***********************************************/

function gradeEssaySelected() {
  const sheet = getSheet_(CONFIG.ESSAY_SUBMISSIONS);
  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("Please select a student row (not the header).");
    return;
  }
  
  // Check for API key
  const key = getOpenAIKey_();
  if (!key) {
    SpreadsheetApp.getUi().alert("❌ OpenAI API key not found.\n\nTo use Essay grading:\n1. Go to Extensions > Apps Script\n2. Project Settings > Script Properties\n3. Add: OPENAI_API_KEY = your_key");
    return;
  }
  
  try {
    gradeEssayRow_(sheet, row);
    SpreadsheetApp.getUi().alert("✅ Grading complete!");
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error: " + e.message);
  }
}

function gradeEssayAll() {
  const sheet = getSheet_(CONFIG.ESSAY_SUBMISSIONS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert("No students to grade.");
    return;
  }
  
  // Check for API key
  const key = getOpenAIKey_();
  if (!key) {
    SpreadsheetApp.getUi().alert("❌ OpenAI API key not found.\n\nTo use Essay grading:\n1. Go to Extensions > Apps Script\n2. Project Settings > Script Properties\n3. Add: OPENAI_API_KEY = your_key");
    return;
  }
  
  const confirm = SpreadsheetApp.getUi().alert(
    "Grade All Essays",
    "This will grade all students with Status = TODO or blank.\nAPI costs apply (~$0.01-0.03 per essay).\n\nContinue?",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;
  
  let graded = 0;
  let skipped = 0;
  let errors = 0;
  
  for (let row = 2; row <= lastRow; row++) {
    const status = String(sheet.getRange(row, getColIndex_(sheet, "Status")).getValue()).trim().toUpperCase();
    if (status === "DONE") {
      skipped++;
      continue;
    }
    
    const name = String(sheet.getRange(row, getColIndex_(sheet, "Student Name")).getValue()).trim();
    const essay = String(sheet.getRange(row, getColIndex_(sheet, "Essay Text")).getValue()).trim();
    
    if (!name || !essay) {
      skipped++;
      continue;
    }
    
    try {
      gradeEssayRow_(sheet, row);
      graded++;
      Utilities.sleep(1500); // Rate limiting
    } catch (e) {
      errors++;
      sheet.getRange(row, getColIndex_(sheet, "Status")).setValue("ERROR");
    }
  }
  
  SpreadsheetApp.getUi().alert(`✅ Grading complete!\n\nGraded: ${graded}\nSkipped: ${skipped}\nErrors: ${errors}`);
}

function gradeEssayRow_(sheet, row) {
  const studentName = String(sheet.getRange(row, getColIndex_(sheet, "Student Name")).getValue()).trim();
  const essay = String(sheet.getRange(row, getColIndex_(sheet, "Essay Text")).getValue()).trim();
  const assignment = String(sheet.getRange(row, getColIndex_(sheet, "Assignment")).getValue()).trim() || "Essay Assignment";
  
  if (!studentName) throw new Error("Student Name is empty.");
  if (!essay) throw new Error("Essay Text is empty.");
  
  // Get rubric
  const rubric = getEssayRubric_();
  
  // Build prompt
  const prompt = buildEssayPrompt_(studentName, essay, assignment, rubric);
  
  // Call API
  const response = callOpenAI_(prompt);
  const parsed = safeParseJson_(response);
  
  if (!parsed || !parsed.sectionScores) {
    throw new Error("Failed to parse AI response");
  }
  
  // Write scores
  let totalScore = 0;
  let totalMax = 0;
  
  rubric.sections.forEach(section => {
    const col = getColIndex_(sheet, section.label);
    const score = Math.min(parsed.sectionScores[section.key] || 0, section.max);
    if (col) sheet.getRange(row, col).setValue(score);
    totalScore += score;
    totalMax += section.max;
  });
  
  // Write total and feedback
  const totalCol = getColIndex_(sheet, "Total Score");
  const feedbackCol = getColIndex_(sheet, "Feedback");
  const statusCol = getColIndex_(sheet, "Status");
  
  if (totalCol) sheet.getRange(row, totalCol).setValue(totalScore + "/" + totalMax);
  if (feedbackCol) sheet.getRange(row, feedbackCol).setValue(parsed.feedback || "");
  if (statusCol) sheet.getRange(row, statusCol).setValue("DONE");
}

function getEssayRubric_() {
  const sheet = getSheet_(CONFIG.ESSAY_RUBRIC);
  const data = sheet.getDataRange().getValues();
  
  const sections = [];
  
  // Expected format: Category | Max Points | Description
  for (let i = 1; i < data.length; i++) {
    const label = String(data[i][0] || "").trim();
    const max = Number(data[i][1]) || 0;
    const description = String(data[i][2] || "").trim();
    
    if (label && max > 0) {
      const key = label.toLowerCase().replace(/[^a-z0-9]/g, "_");
      sections.push({ key, label, max, description });
    }
  }
  
  if (sections.length === 0) {
    // Default rubric if none defined
    return {
      sections: [
        { key: "content", label: "Content & Ideas", max: 40, description: "Quality and depth of ideas" },
        { key: "organization", label: "Organization", max: 20, description: "Structure and flow" },
        { key: "evidence", label: "Evidence & Support", max: 20, description: "Use of examples and citations" },
        { key: "grammar", label: "Grammar & Mechanics", max: 20, description: "Writing quality and correctness" },
      ]
    };
  }
  
  return { sections };
}

function buildEssayPrompt_(studentName, essay, assignment, rubric) {
  const sectionLines = rubric.sections
    .map(s => `- ${s.key}: ${s.label} (0-${s.max}) - ${s.description}`)
    .join("\n");
  
  const scoreKeys = rubric.sections.map(s => `"${s.key}": <integer 0-${s.max}>`).join(",\n    ");
  
  return `
You are a teaching assistant grading student essays.

ASSIGNMENT: ${assignment}

RUBRIC (use INTEGER scores only):
${sectionLines}

OUTPUT FORMAT (strict JSON only, no markdown):
{
  "sectionScores": {
    ${scoreKeys}
  },
  "feedback": "2-4 sentences of constructive feedback highlighting strengths and areas for improvement."
}

GRADING GUIDELINES:
- Be fair but critical
- Use the full range of scores
- Provide specific, actionable feedback
- No perfect scores unless truly exceptional

STUDENT ESSAY:
${essay}
`.trim();
}

function resetEssaySelected() {
  const sheet = getSheet_(CONFIG.ESSAY_SUBMISSIONS);
  const row = sheet.getActiveCell().getRow();
  if (row <= 1) return;
  
  const rubric = getEssayRubric_();
  const colsToClear = ["Total Score", "Feedback", "Status"];
  rubric.sections.forEach(s => colsToClear.push(s.label));
  
  colsToClear.forEach(colName => {
    const col = getColIndex_(sheet, colName);
    if (col) sheet.getRange(row, col).clearContent();
  });
  
  SpreadsheetApp.getUi().alert("✅ Student reset complete.");
}

/***********************************************
 * UTILITY FUNCTIONS
 ***********************************************/

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet "${name}" not found. Please check sheet names.`);
  return sheet;
}

function getColIndex_(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === headerName.toLowerCase()) {
      return i + 1;
    }
  }
  return null;
}

function getOpenAIKey_() {
  return PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
}

function callOpenAI_(prompt) {
  const apiKey = getOpenAIKey_();
  if (!apiKey) throw new Error("OpenAI API key not configured.");
  
  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: CONFIG.MODEL,
    temperature: CONFIG.TEMPERATURE,
    max_tokens: CONFIG.MAX_TOKENS,
    messages: [
      { role: "system", content: "Return strict JSON only. No markdown, no extra text." },
      { role: "user", content: prompt },
    ],
  };
  
  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  
  const code = response.getResponseCode();
  const body = response.getContentText();
  
  if (code < 200 || code >= 300) {
    throw new Error("API error: HTTP " + code);
  }
  
  const data = JSON.parse(body);
  return data?.choices?.[0]?.message?.content || "";
}

function safeParseJson_(text) {
  const s = String(text || "").trim();
  const a = s.indexOf("{");
  const b = s.lastIndexOf("}");
  if (a >= 0 && b > a) {
    try { return JSON.parse(s.slice(a, b + 1)); } catch (e) {}
  }
  try { return JSON.parse(s); } catch (e) { return null; }
}
