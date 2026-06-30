function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Quiz')
    .addItem('Set up sheets', 'setupSheet')
    .addItem('Add Question and Rubric', 'showAddQuestionDialog')
    .addItem('Open Logs', 'openLogsSheet')
    .addItem('Authorize API access (run once)', 'authorizeExternalRequests')
    .addItem('Run AI Evaluation', 'runEvaluationWithAlert')
    .addItem('Seed Histogram Pilot Quiz', 'seedHistogramPilotQuizWithAlert')
    .addItem('Seed Extended Pilot Quiz', 'seedExtendedHistogramPilotQuizWithAlert')
    .addSeparator()
    .addItem('Backfill evaluation points...', 'backfillEvaluationPointsWithConfirm')
    .addItem('Clear all quiz data...', 'clearAllQuizDataWithConfirm')
    .addToUi();
}

function openLogsSheet() {
  var ss = getQuizSpreadsheet_();
  ensureLogsSheet_(ss);
  ss.setActiveSheet(ss.getSheetByName('Logs'));
}

function clearAllQuizDataWithConfirm() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Clear all quiz data?',
    'This deletes every data row in Questions, Responses, Evaluations, and EvaluationHistory. Column headers are kept.\n\n' +
      'Logs are not cleared. Use File > Version history on the spreadsheet to restore if needed.',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) {
    return;
  }

  try {
    var result = clearAllQuizData_();
    ui.alert(
      'Quiz data cleared.\n\n' +
      'Questions removed: ' + result.questionsRemoved + '\n' +
      'Responses removed: ' + result.responsesRemoved + '\n' +
      'Evaluations removed: ' + result.evaluationsRemoved + '\n' +
      'Evaluation history removed: ' + result.evaluationHistoryRemoved
    );
  } catch (error) {
    ui.alert('Could not clear quiz data:\n\n' + error.message);
  }
}

function clearAllQuizData_() {
  var ss = getQuizSpreadsheet_();
  var questionsRemoved = clearQuizSheetData_(ss, 'Questions', ensureQuestionsHeaders_);
  var responsesRemoved = clearQuizSheetData_(ss, 'Responses', ensureResponsesHeaders_);
  var evaluationsRemoved = clearQuizSheetData_(ss, 'Evaluations', ensureEvaluationsHeaders_);
  var evaluationHistoryRemoved = clearQuizSheetData_(ss, 'EvaluationHistory', ensureEvaluationHistoryHeaders_);
  SpreadsheetApp.flush();

  logQuizEvent_('info', 'clearAllQuizData', 'quiz data cleared', {
    questionsRemoved: questionsRemoved,
    responsesRemoved: responsesRemoved,
    evaluationsRemoved: evaluationsRemoved,
    evaluationHistoryRemoved: evaluationHistoryRemoved
  });

  return {
    questionsRemoved: questionsRemoved,
    responsesRemoved: responsesRemoved,
    evaluationsRemoved: evaluationsRemoved,
    evaluationHistoryRemoved: evaluationHistoryRemoved
  };
}

function clearQuizSheetData_(ss, sheetName, ensureHeadersFn) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return 0;
  }

  ensureHeadersFn(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }

  var count = lastRow - 1;
  sheet.deleteRows(2, count);
  return count;
}

function showAddQuestionDialog() {
  var html = HtmlService.createHtmlOutputFromFile('add-question')
    .setWidth(520)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Question and Rubric');
}

function authorizeExternalRequests() {
  UrlFetchApp.fetch('https://example.com', { muteHttpExceptions: true });
  SpreadsheetApp.getUi().alert(
    'Done. If Google asked for permission, make sure you clicked Allow for external network access. ' +
    'You can now use Run AI Evaluation.'
  );
}

function runEvaluationWithAlert() {
  try {
    var result = triggerEvaluation();
    SpreadsheetApp.getUi().alert(formatEvaluationAlertText_(result));
  } catch (error) {
    SpreadsheetApp.getUi().alert('Evaluation failed:\n\n' + error.message);
  }
}

var EVAL_BATCH_SIZE_ = 15;
var EVAL_BATCH_MAX_TOKENS_ = 8192;
var EVAL_MODEL_ = 'claude-haiku-4-5-20251001';
var EVAL_SCORE_MAX_ = 4;
// Calibrated from pilot runs: 2 API calls ~13s; 16 API calls ~80s.
var EVAL_ESTIMATE_SEC_PER_API_CALL_ = 5;
var EVAL_ESTIMATE_SEC_OVERHEAD_ = 3;
var EVAL_GRADING_PHILOSOPHY_ = [
  'Grading philosophy (Grading for Equity — 4-point rubric):',
  'Score each answer from 0 to 4 based on how many requested facts from the question-specific rubric the student addressed correctly.',
  '- 4: Correctly addressed all requested facts from the rubric.',
  '- 3: Correctly addressed most requested facts.',
  '- 2: Correctly addressed some requested facts.',
  '- 1: Addressed at least one requested fact correctly.',
  '- 0: None of the requested facts from the rubric were addressed correctly.',
  'Use the question-specific rubric below to identify the requested facts. Apply this scale consistently across all students in the batch.',
  'Feedback has two parts in this order: (1) substantive, question-specific comments citing the student answer and rubric facts; (2) a closing sentence that justifies the score using the scale (all / most / some / one / none of requested facts). Do not replace detailed feedback with only generic scale language.'
].join('\n');

function ensureResponsesHeaders_(sheet) {
  if (sheet.getRange('A1').getValue() === 'Timestamp') {
    if (sheet.getRange('E1').getValue() !== 'Question ID') {
      sheet.getRange('E1').setValue('Question ID');
    }
    if (sheet.getRange('F1').getValue() !== 'Quiz ID') {
      sheet.getRange('F1').setValue('Quiz ID');
    }
    formatHeaderRow_(sheet, 1, 6);
    return;
  }
  if (sheet.getLastRow() > 0) {
    sheet.insertRowBefore(1);
  }
  setupResponsesSheet_(sheet);
}

function getQuizSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    return ss;
  }

  var parents = DriveApp.getFileById(ScriptApp.getScriptId()).getParents();
  if (parents.hasNext()) {
    var parent = parents.next();
    if (parent.getMimeType() === MimeType.GOOGLE_SHEETS) {
      return SpreadsheetApp.openById(parent.getId());
    }
  }

  throw new Error('Could not find the quiz spreadsheet.');
}

function getSpreadsheet_(spreadsheetId) {
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    return ss;
  }
  throw new Error('No active spreadsheet. Run setupSheet from the script editor or pass a spreadsheet ID.');
}

function setupSheet(spreadsheetId) {
  var ss = getSpreadsheet_(spreadsheetId);
  var questionsSheet = ensureSheet_(ss, 'Questions');
  var responsesSheet = ensureSheet_(ss, 'Responses');
  var evaluationsSheet = ensureSheet_(ss, 'Evaluations');

  ensureQuestionsHeaders_(questionsSheet);
  ensureResponsesHeaders_(responsesSheet);
  ensureEvaluationsHeaders_(evaluationsSheet);
  ensureEvaluationHistorySheet_(ss);
  setupInstructionsSheet_(ss);
  ensureLogsSheet_(ss);
  ensureQuestionImagesSheet_(ss);

  removeDefaultSheet_(ss);

  return 'Sheet setup complete. Existing quiz data was preserved; headers were added or updated as needed.';
}

function ensureQuestionsHeaders_(sheet) {
  if (sheet.getRange('B1').getValue() === 'Question') {
    if (sheet.getRange('D1').getValue() !== 'Quiz Name') {
      sheet.getRange('D1').setValue('Quiz Name');
    }
    if (sheet.getRange('E1').getValue() !== 'Quiz ID') {
      sheet.getRange('E1').setValue('Quiz ID');
    }
    if (sheet.getRange('F1').getValue() !== 'Question ID') {
      sheet.getRange('F1').setValue('Question ID');
    }
    if (sheet.getRange('G1').getValue() !== 'Question Image') {
      sheet.getRange('G1').setValue('Question Image');
    }
    formatHeaderRow_(sheet, 1, 7);
    return;
  }
  if (sheet.getLastRow() > 0) {
    sheet.insertRowBefore(1);
  }
  setupQuestionsSheet_(sheet);
}

function ensureSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function setColumnWidths_(sheet, widths) {
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}

function setupQuestionsSheet_(sheet) {
  sheet.clear();
  sheet.getRange('A1:G1').setValues([['Row', 'Question', 'Rubric', 'Quiz Name', 'Quiz ID', 'Question ID', 'Question Image']]);
  sheet.getRange('A2').setValue(2);
  sheet.getRange('B2').setValue('What is the capital of France?');
  sheet.getRange('C2').setValue(
    'Award full credit for "Paris". Partial credit for mentioning France. ' +
    'Deduct points for incorrect capitals.'
  );
  sheet.getRange('D2').setValue('Sample Quiz');
  sheet.getRange('E2').setValue(generateShortId_());
  sheet.getRange('F2').setValue(generateShortId_());
  setColumnWidths_(sheet, [60, 360, 360, 160, 100, 100, 80]);
  sheet.setFrozenRows(1);
  formatHeaderRow_(sheet, 1, 7);
}

function setupResponsesSheet_(sheet) {
  sheet.getRange('A1:F1').setValues([['Timestamp', 'Student Name', 'Answer', 'Status', 'Question ID', 'Quiz ID']]);
  setColumnWidths_(sheet, [160, 160, 360, 100, 100, 100]);
  sheet.setFrozenRows(1);
  formatHeaderRow_(sheet, 1, 6);
}

function setupEvaluationsSheet_(sheet) {
  sheet.getRange('A1:I1').setValues([
    ['Timestamp', 'Student Name', 'Answer', 'Rubric', 'Evaluation', 'Question ID', 'Quiz ID', 'Student Review', 'Points']
  ]);
  setColumnWidths_(sheet, [160, 160, 300, 300, 400, 100, 100, 320, 60]);
  sheet.setFrozenRows(1);
  formatHeaderRow_(sheet, 1, 9);
}

function ensureEvaluationsHeaders_(sheet) {
  if (sheet.getRange('A1').getValue() === 'Timestamp') {
    if (sheet.getRange('F1').getValue() !== 'Question ID') {
      sheet.getRange('F1').setValue('Question ID');
    }
    if (sheet.getRange('G1').getValue() !== 'Quiz ID') {
      sheet.getRange('G1').setValue('Quiz ID');
    }
    if (sheet.getRange('H1').getValue() !== 'Student Review') {
      sheet.getRange('H1').setValue('Student Review');
    }
    if (sheet.getRange('I1').getValue() !== 'Points') {
      sheet.getRange('I1').setValue('Points');
    }
    if (sheet.getRange('E1').getValue() === 'AI Evaluation') {
      sheet.getRange('E1').setValue('Evaluation');
    }
    formatHeaderRow_(sheet, 1, 9);
    return;
  }
  setupEvaluationsSheet_(sheet);
}

function setupEvaluationHistorySheet_(sheet) {
  sheet.getRange('A1:H1').setValues([[
    'Timestamp',
    'Source',
    'Student Name',
    'Question ID',
    'Quiz ID',
    'Evaluations Row',
    'Evaluation',
    'Points'
  ]]);
  setColumnWidths_(sheet, [160, 80, 160, 100, 100, 110, 400, 60]);
  sheet.setFrozenRows(1);
  formatHeaderRow_(sheet, 1, 8);
}

function ensureEvaluationHistoryHeaders_(sheet) {
  if (sheet.getRange('A1').getValue() === 'Timestamp') {
    if (sheet.getRange('H1').getValue() !== 'Points') {
      sheet.getRange('H1').setValue('Points');
    }
    formatHeaderRow_(sheet, 1, 8);
    return;
  }
  setupEvaluationHistorySheet_(sheet);
}

function ensureEvaluationHistorySheet_(ss) {
  var sheet = ensureSheet_(ss, 'EvaluationHistory');
  ensureEvaluationHistoryHeaders_(sheet);
  return sheet;
}

function appendEvaluationHistoryRows_(historyRows) {
  if (!historyRows || !historyRows.length) {
    return;
  }

  var ss = getQuizSpreadsheet_();
  var sheet = ensureEvaluationHistorySheet_(ss);
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange('A' + startRow + ':H' + (startRow + historyRows.length - 1)).setValues(historyRows);
}

function backfillEvaluationPointsWithConfirm() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Backfill evaluation points?',
    'This reads legacy Score: X/4 prefixes from the Evaluation column, writes Points (column I), ' +
      'and removes the score line from feedback text. Rows that already have points are left unchanged ' +
      '(except legacy score lines may still be stripped from feedback). Safe to run more than once.',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) {
    return;
  }

  try {
    var result = backfillEvaluationPoints_();
    ui.alert(
      'Backfill complete.\n\n' +
        'Points backfilled: ' + result.backfilled + '\n' +
        'Already had points: ' + result.alreadyHadPoints + '\n' +
        'Feedback prefixes stripped: ' + result.stripped + '\n' +
        'Rows without parseable score: ' + result.noParseableScore
    );
  } catch (error) {
    ui.alert('Backfill failed:\n\n' + error.message);
  }
}

function backfillEvaluationPoints_() {
  var ss = getQuizSpreadsheet_();
  var evalSheet = ss.getSheetByName('Evaluations');
  if (!evalSheet || evalSheet.getLastRow() < 2) {
    return {
      backfilled: 0,
      alreadyHadPoints: 0,
      stripped: 0,
      noParseableScore: 0
    };
  }

  ensureEvaluationsHeaders_(evalSheet);
  var lastRow = evalSheet.getLastRow();
  var rows = evalSheet.getRange('A2:I' + lastRow).getValues();
  var backfilled = 0;
  var alreadyHadPoints = 0;
  var stripped = 0;
  var noParseableScore = 0;
  var i;

  for (i = 0; i < rows.length; i++) {
    var rowNum = i + 2;
    var evaluation = cellText_(rows[i][4]).trim();
    var existingPoints = normalizeEvaluationPoints_(rows[i][8]);

    if (existingPoints !== null) {
      alreadyHadPoints++;
      if (stripLegacyScorePrefixFromEvaluation_(evaluation) !== evaluation) {
        evalSheet.getRange('E' + rowNum).setValue(stripLegacyScorePrefixFromEvaluation_(evaluation));
        stripped++;
      }
      continue;
    }

    if (!evaluation) {
      continue;
    }

    var parsed = parseLegacyScoreFromEvaluation_(evaluation);
    if (parsed === null) {
      noParseableScore++;
      continue;
    }

    evalSheet.getRange('I' + rowNum).setValue(parsed);
    evalSheet.getRange('E' + rowNum).setValue(stripLegacyScorePrefixFromEvaluation_(evaluation));
    backfilled++;
    stripped++;
  }

  SpreadsheetApp.flush();
  logQuizEvent_('info', 'backfillEvaluationPoints', 'evaluation points backfilled', {
    backfilled: backfilled,
    alreadyHadPoints: alreadyHadPoints,
    stripped: stripped,
    noParseableScore: noParseableScore
  });

  return {
    backfilled: backfilled,
    alreadyHadPoints: alreadyHadPoints,
    stripped: stripped,
    noParseableScore: noParseableScore
  };
}

function normalizeEvaluationPoints_(value) {
  if (value === '' || value === null || value === undefined) {
    return null;
  }
  var n = Math.round(Number(value));
  if (isNaN(n)) {
    return null;
  }
  if (n < 0) {
    n = 0;
  }
  if (n > EVAL_SCORE_MAX_) {
    n = EVAL_SCORE_MAX_;
  }
  return n;
}

function parseLegacyScoreFromEvaluation_(text) {
  var raw = String(text || '').trim();
  var match = raw.match(/^Score:\s*(\d+)\s*\/\s*4\b\s*(?:\r?\n\r?\n|\r?\n|$)/i);
  if (!match) {
    return null;
  }
  return normalizeEvaluationPoints_(match[1]);
}

function stripLegacyScorePrefixFromEvaluation_(text) {
  return String(text || '').replace(/^Score:\s*\d+\s*\/\s*4\b\s*(?:\r?\n\r?\n|\r?\n)?/i, '').trim();
}

function resolveEvaluationPointsForRow_(pointsCell, evaluationText) {
  var points = normalizeEvaluationPoints_(pointsCell);
  if (points !== null) {
    return points;
  }
  return parseLegacyScoreFromEvaluation_(evaluationText);
}

function resolveEvaluationFeedbackForClient_(pointsCell, evaluationText) {
  var feedback = String(evaluationText || '').trim();
  if (normalizeEvaluationPoints_(pointsCell) !== null) {
    return feedback;
  }
  var stripped = stripLegacyScorePrefixFromEvaluation_(feedback);
  return stripped || feedback;
}

function setupInstructionsSheet_(ss) {
  var sheet = ss.getSheetByName('Setup');
  if (!sheet) {
    sheet = ss.insertSheet('Setup', 0);
  }

  sheet.clear();
  sheet.getRange('A1').setValue('Quiz Prototype Setup').setFontSize(14).setFontWeight('bold').setFontColor('#1a73e8');
  sheet.getRange('A3:A9').setValues([
    ['1. Add questions via the teacher page (?teach) or Quiz > Add Question and Rubric.'],
    ['2. Each response stores Question ID (E) and Quiz ID (F) to link back to Questions.'],
    ['3. Questions in the same quiz share a Quiz Name and Quiz ID (column E).'],
    ['4. Deploy the script as a web app (Deploy > New deployment > Web app).'],
    ['5. Share the student URL: <web app url> (the deployment URL with no parameters).'],
    ['6. Open the teacher URL (<web app url>?teach) to add questions and run evaluations.'],
    ['7. Student answers appear on Responses with Status = Pending; evaluation sets Complete.']
  ]);
  sheet.setColumnWidth(1, 640);
  sheet.getRange('A3:A9').setWrap(true);
}

function formatHeaderRow_(sheet, row, numColumns) {
  var header = sheet.getRange(row, 1, 1, numColumns);
  header
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
}

function removeDefaultSheet_(ss) {
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
}

function doGet(e) {
  var params = e && e.parameter;
  if (params && params.teach !== undefined) {
    return HtmlService.createHtmlOutputFromFile('teacher')
      .setTitle('Teacher Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('student')
    .setTitle('Student Quiz')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getNextResponseRow_(responsesSheet) {
  ensureResponsesHeaders_(responsesSheet);
  var lastRow = responsesSheet.getLastRow();
  return lastRow < 2 ? 2 : lastRow + 1;
}

function getNextQuestionsRow_(questionsSheet) {
  ensureQuestionsHeaders_(questionsSheet);
  var lastRow = questionsSheet.getLastRow();
  return lastRow < 2 ? 2 : lastRow + 1;
}

function generateShortId_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  var id = '';
  for (var i = 0; i < 8; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}

function generateQuizId_() {
  return generateShortId_();
}

function collectExistingIds_(questionsSheet, column) {
  var existing = {};
  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) {
    return existing;
  }

  var ids = questionsSheet.getRange(column + '2:' + column + lastRow).getValues();
  for (var i = 0; i < ids.length; i++) {
    var value = String(ids[i][0]).trim();
    if (value) {
      existing[value] = true;
    }
  }

  return existing;
}

function generateUniqueId_(existing) {
  var id;
  do {
    id = generateShortId_();
  } while (existing[id]);
  existing[id] = true;
  return id;
}

function generateUniqueQuizId_(questionsSheet) {
  return generateUniqueId_(collectExistingIds_(questionsSheet, 'E'));
}

function generateUniqueQuestionId_(questionsSheet) {
  return generateUniqueId_(collectExistingIds_(questionsSheet, 'F'));
}

function ensureQuestionIds_(questionsSheet) {
  ensureQuestionsHeaders_(questionsSheet);
  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  var existing = collectExistingIds_(questionsSheet, 'F');
  var rows = questionsSheet.getRange('B2:F' + lastRow).getValues();
  var idValues = questionsSheet.getRange('F2:F' + lastRow).getValues();
  var changed = false;

  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][0]) {
      continue;
    }
    if (String(idValues[i][0]).trim()) {
      continue;
    }
    idValues[i][0] = generateUniqueId_(existing);
    changed = true;
  }

  if (changed) {
    questionsSheet.getRange('F2:F' + lastRow).setValues(idValues);
  }
}

function findQuizIdByName_(questionsSheet, quizName) {
  var normalized = String(quizName).trim().toLowerCase();
  if (!normalized) {
    return '';
  }

  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) {
    return '';
  }

  var rows = questionsSheet.getRange('D2:E' + lastRow).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === normalized && rows[i][1]) {
      return String(rows[i][1]).trim();
    }
  }

  return '';
}

function getQuestionAndRubricForRow_(questionsSheet, row, skipImages) {
  if (!questionsSheet) {
    throw new Error('Questions sheet not found.');
  }
  var questionId = questionsSheet.getRange('F' + row).getValue();
  var quizId = questionsSheet.getRange('E' + row).getValue();
  var legacyImage = questionsSheet.getRange('G' + row).getValue();
  var imageData = '';
  if (!skipImages) {
    var ss = questionsSheet.getParent();
    imageData = readQuestionImageDataForQuestion_(ss, questionId, quizId, legacyImage);
  }
  return {
    question: questionsSheet.getRange('B' + row).getValue(),
    rubric: questionsSheet.getRange('C' + row).getValue(),
    quizName: questionsSheet.getRange('D' + row).getValue(),
    quizId: quizId,
    questionId: questionId,
    imageData: imageData
  };
}

function getQuestionAndRubricByQuestionId_(questionsSheet, questionId, quizId, skipImages) {
  ensureQuestionsHeaders_(questionsSheet);
  questionId = normalizeSheetId_(questionId);
  quizId = normalizeSheetId_(quizId);
  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2 || !questionId) {
    return {
      questionText: '',
      rubricText: '',
      quizName: '',
      quizId: '',
      questionId: questionId
    };
  }

  var rows = questionsSheet.getRange('A2:G' + lastRow).getValues();
  var ss = questionsSheet.getParent();
  var fallback = null;

  for (var i = 0; i < rows.length; i++) {
    if (normalizeSheetId_(rows[i][5]) !== questionId) {
      continue;
    }

    var entry = {
      questionText: cellText_(rows[i][1]),
      rubricText: cellText_(rows[i][2]),
      imageData: skipImages
        ? ''
        : readQuestionImageDataForQuestion_(ss, questionId, normalizeSheetId_(rows[i][4]), rows[i][6]),
      quizName: cellText_(rows[i][3]),
      quizId: normalizeSheetId_(rows[i][4]),
      questionId: questionId
    };

    if (!quizId || entry.quizId === quizId) {
      return entry;
    }

    if (!fallback) {
      fallback = entry;
    }
  }

  if (fallback) {
    return fallback;
  }

  return {
    questionText: '',
    rubricText: '',
    quizName: '',
    quizId: quizId,
    questionId: questionId
  };
}

function normalizeSheetId_(value) {
  return String(value == null ? '' : value).trim().replace(/^'/, '');
}

function cellText_(value) {
  if (value == null) {
    return '';
  }
  return String(value);
}

function ensureLogsSheet_(ss) {
  var sheet = ensureSheet_(ss, 'Logs');
  if (sheet.getRange('A1').getValue() !== 'Timestamp') {
    sheet.getRange('A1:E1').setValues([['Timestamp', 'Level', 'Source', 'Message', 'Details']]);
    setColumnWidths_(sheet, [160, 80, 140, 280, 420]);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, 1, 5);
  }
  return sheet;
}

function logQuizEvent_(level, source, message, details) {
  try {
    var ss = getQuizSpreadsheet_();
    var sheet = ensureLogsSheet_(ss);
    var detailsText = '';
    if (details != null) {
      try {
        detailsText = JSON.stringify(details);
      } catch (jsonError) {
        detailsText = String(details);
      }
    }
    if (detailsText.length > 45000) {
      detailsText = detailsText.substring(0, 45000) + '... (truncated)';
    }
    sheet.appendRow([
      new Date(),
      String(level || 'info'),
      String(source || ''),
      String(message || ''),
      detailsText
    ]);
  } catch (logError) {
    console.error(logError);
  }
}

function logFromClient(level, source, message, detailsJson) {
  var details = {};
  if (detailsJson) {
    try {
      details = JSON.parse(detailsJson);
    } catch (parseError) {
      details = { raw: String(detailsJson) };
    }
  }
  logQuizEvent_(level || 'info', source || 'client', message || '', details);
  return { ok: true };
}

function parseLogDetails_(detailsText) {
  if (!detailsText) {
    return null;
  }
  try {
    return JSON.parse(String(detailsText));
  } catch (parseError) {
    return null;
  }
}

function normalizeLogQuizId_(quizId) {
  var id = String(quizId == null ? '' : quizId).trim();
  if (!id || id === '(all)') {
    return '';
  }
  return id;
}

function getLastEvaluationTimingByQuiz_() {
  var ss = getQuizSpreadsheet_();
  var sheet = ensureLogsSheet_(ss);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {};
  }

  var rows = sheet.getRange('A2:E' + lastRow).getValues();
  var byQuiz = {};
  var byRunId = {};
  var i;
  var j;
  var source;
  var message;
  var details;
  var quizId;
  var runId;

  for (i = rows.length - 1; i >= 0; i--) {
    source = String(rows[i][2] || '').trim();
    message = String(rows[i][3] || '').trim();
    details = parseLogDetails_(rows[i][4]);
    if (!details) {
      continue;
    }

    if (source === 'evaluation' && details.event === 'evaluationTiming' && details.phase === 'completed') {
      quizId = normalizeLogQuizId_(details.quizId);
      runId = String(details.evalRunId || '').trim();
      if (!quizId || byQuiz[quizId]) {
        continue;
      }
      byQuiz[quizId] = {
        evalRunId: runId,
        estimatedSec: details.estimatedSec,
        actualSec: details.actualSec,
        deltaSec: details.deltaSec,
        deltaPct: details.deltaPct,
        pendingCount: details.pendingCount,
        apiCallCount: details.apiCallCount,
        completedAt: formatLogTimestampForClient_(rows[i][0])
      };
      if (runId) {
        byRunId[runId] = byQuiz[quizId];
      }
    }
  }

  for (j = rows.length - 1; j >= 0; j--) {
    source = String(rows[j][2] || '').trim();
    message = String(rows[j][3] || '').trim();
    details = parseLogDetails_(rows[j][4]);
    if (!details) {
      continue;
    }

    if (source === 'teacherEvaluate' && message === 'client timing reported') {
      runId = String(details.evalRunId || '').trim();
      quizId = normalizeLogQuizId_(details.quizId);
      if (runId && byRunId[runId]) {
        byRunId[runId].clientWaitSec = details.clientWaitSec;
      } else if (quizId && byQuiz[quizId] && byQuiz[quizId].clientWaitSec == null) {
        byQuiz[quizId].clientWaitSec = details.clientWaitSec;
      }
    }
  }

  return byQuiz;
}

function formatLogTimestampForClient_(value) {
  if (!value) {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value.toISOString();
  }
  return String(value);
}

function serializeEvaluationTimingForClient_(timing) {
  if (!timing) {
    return null;
  }

  return {
    evalRunId: String(timing.evalRunId || ''),
    estimatedSec: timing.estimatedSec == null ? null : Number(timing.estimatedSec),
    actualSec: timing.actualSec == null ? null : Number(timing.actualSec),
    deltaSec: timing.deltaSec == null ? null : Number(timing.deltaSec),
    deltaPct: timing.deltaPct == null ? null : Number(timing.deltaPct),
    pendingCount: timing.pendingCount == null ? null : Number(timing.pendingCount),
    apiCallCount: timing.apiCallCount == null ? null : Number(timing.apiCallCount),
    clientWaitSec: timing.clientWaitSec == null ? null : Number(timing.clientWaitSec),
    completedAt: formatLogTimestampForClient_(timing.completedAt)
  };
}

function serializeQuizEvaluationListForClient_(list) {
  var out = [];
  var i;

  for (i = 0; i < list.length; i++) {
    var quiz = list[i];
    out.push({
      quizId: String(quiz.quizId || ''),
      quizName: String(quiz.quizName || 'Unnamed Quiz'),
      questionCount: Number(quiz.questionCount) || 0,
      pendingCount: Number(quiz.pendingCount) || 0,
      lastEvaluation: serializeEvaluationTimingForClient_(quiz.lastEvaluation)
    });
  }

  return out;
}

function logEvaluationUiEvent_(phase, details) {
  var payload = {
    event: 'evaluationUi',
    phase: phase
  };
  var key;
  for (key in details) {
    if (details.hasOwnProperty(key)) {
      payload[key] = details[key];
    }
  }
  logQuizEvent_('info', 'evaluationUi', 'evaluation ui ' + phase, payload);
}

var MAX_QUESTION_IMAGE_BYTES_ = 1048576;
var QUESTION_IMAGE_UPLOAD_CHUNK_SIZE_ = 40000;
var QUESTION_IMAGE_UPLOAD_MAX_PARTS_ = 40;
var QUESTION_IMAGE_STORAGE_CHUNK_SIZE_ = 49000;

function ensureQuestionImagesSheet_(ss) {
  var sheet = ensureSheet_(ss, 'QuestionImages');
  if (sheet.getRange('A1').getValue() !== 'Question ID') {
    sheet.getRange('A1:D1').setValues([['Question ID', 'Quiz ID', 'Part', 'Image Data']]);
    setColumnWidths_(sheet, [100, 100, 60, 420]);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, 1, 4);
  }
  return sheet;
}

function questionImageStorageKey_(quizId, questionId) {
  return normalizeSheetId_(quizId) + '|' + normalizeSheetId_(questionId);
}

function splitQuestionImageForStorage_(imageData) {
  var chunks = [];
  for (var i = 0; i < imageData.length; i += QUESTION_IMAGE_STORAGE_CHUNK_SIZE_) {
    chunks.push(imageData.substring(i, i + QUESTION_IMAGE_STORAGE_CHUNK_SIZE_));
  }
  return chunks;
}

function deleteQuestionImageRows_(imageSheet, questionId, quizId) {
  questionId = normalizeSheetId_(questionId);
  quizId = normalizeSheetId_(quizId);
  var lastRow = imageSheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  var rows = imageSheet.getRange('A2:B' + lastRow).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (normalizeSheetId_(rows[i][0]) === questionId && normalizeSheetId_(rows[i][1]) === quizId) {
      imageSheet.deleteRow(i + 2);
    }
  }
}

function buildQuestionImageMap_(ss) {
  var map = {};
  var sheet = ss.getSheetByName('QuestionImages');
  if (!sheet) {
    return map;
  }

  ensureQuestionImagesSheet_(ss);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return map;
  }

  var rows = sheet.getRange('A2:D' + lastRow).getValues();
  var grouped = {};

  for (var i = 0; i < rows.length; i++) {
    var questionId = normalizeSheetId_(rows[i][0]);
    var quizId = normalizeSheetId_(rows[i][1]);
    if (!questionId) {
      continue;
    }

    var key = questionImageStorageKey_(quizId, questionId);
    if (!grouped[key]) {
      grouped[key] = [];
    }
    grouped[key].push({
      part: Number(rows[i][2]) || 0,
      data: cellText_(rows[i][3])
    });
  }

  for (var storageKey in grouped) {
    if (!grouped.hasOwnProperty(storageKey)) {
      continue;
    }
    grouped[storageKey].sort(function(a, b) {
      return a.part - b.part;
    });
    var joined = '';
    for (var j = 0; j < grouped[storageKey].length; j++) {
      joined += grouped[storageKey][j].data;
    }
    map[storageKey] = joined;
  }

  return map;
}

function readQuestionImageData_(imageMap, legacyCellValue, questionId, quizId) {
  var key = questionImageStorageKey_(quizId, questionId);
  if (imageMap && imageMap[key]) {
    return imageMap[key];
  }

  var legacy = cellText_(legacyCellValue).trim();
  if (legacy && /^data:image\//i.test(legacy)) {
    return legacy;
  }

  return '';
}

function questionHasImageMarker_(legacyCellValue) {
  var legacy = cellText_(legacyCellValue).trim();
  if (/^PARTS:\d+$/i.test(legacy)) {
    return true;
  }
  return /^data:image\//i.test(legacy);
}

function readQuestionImageDataForQuestion_(ss, questionId, quizId, legacyCellValue) {
  var legacy = cellText_(legacyCellValue).trim();
  if (legacy && /^data:image\//i.test(legacy)) {
    return legacy;
  }

  questionId = normalizeSheetId_(questionId);
  quizId = normalizeSheetId_(quizId);
  if (!questionId) {
    return '';
  }

  var sheet = ss.getSheetByName('QuestionImages');
  if (!sheet) {
    return '';
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return '';
  }

  var rows = sheet.getRange('A2:D' + lastRow).getValues();
  var parts = [];

  for (var i = 0; i < rows.length; i++) {
    if (normalizeSheetId_(rows[i][0]) !== questionId) {
      continue;
    }
    if (quizId && normalizeSheetId_(rows[i][1]) !== quizId) {
      continue;
    }
    parts.push({
      part: Number(rows[i][2]) || 0,
      data: cellText_(rows[i][3])
    });
  }

  if (!parts.length) {
    return '';
  }

  parts.sort(function(a, b) {
    return a.part - b.part;
  });

  var joined = '';
  for (var j = 0; j < parts.length; j++) {
    joined += parts[j].data;
  }

  return joined;
}

function questionHasStoredImage_(imageMap, legacyCellValue, questionId, quizId) {
  if (readQuestionImageData_(imageMap, legacyCellValue, questionId, quizId)) {
    return true;
  }
  return /^PARTS:\d+$/i.test(cellText_(legacyCellValue).trim());
}

function lookupQuestionImageData_(quizId, questionId) {
  quizId = normalizeSheetId_(quizId);
  questionId = normalizeSheetId_(questionId);
  if (!questionId) {
    return '';
  }

  var ss = getQuizSpreadsheet_();
  var questionsSheet = ss.getSheetByName('Questions');
  var legacy = '';

  if (questionsSheet) {
    ensureQuestionsHeaders_(questionsSheet);
    var lastRow = questionsSheet.getLastRow();
    if (lastRow >= 2) {
      var rows = questionsSheet.getRange('A2:G' + lastRow).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (normalizeSheetId_(rows[i][5]) === questionId && normalizeSheetId_(rows[i][4]) === quizId) {
          legacy = rows[i][6];
          break;
        }
      }
    }
  }

  return readQuestionImageDataForQuestion_(ss, questionId, quizId, legacy);
}

function getQuestionImage(quizId, questionId) {
  quizId = normalizeSheetId_(quizId);
  questionId = normalizeSheetId_(questionId);
  if (!quizId || !questionId) {
    throw new Error('Quiz ID and Question ID are required.');
  }

  var imageData = lookupQuestionImageData_(quizId, questionId);
  return {
    quizId: quizId,
    questionId: questionId,
    hasImage: !!imageData,
    imageData: imageData || ''
  };
}

function writeQuestionImageToSheet_(questionsSheet, row, questionId, quizId, imageData) {
  questionId = normalizeSheetId_(questionId);
  quizId = normalizeSheetId_(quizId);
  if (!questionId) {
    throw new Error('Question ID is required to store an image.');
  }

  var ss = questionsSheet.getParent();
  var imageSheet = ensureQuestionImagesSheet_(ss);
  deleteQuestionImageRows_(imageSheet, questionId, quizId);

  var chunks = splitQuestionImageForStorage_(imageData);
  if (chunks.length) {
    var writeRows = [];
    for (var i = 0; i < chunks.length; i++) {
      writeRows.push([questionId, quizId, i, chunks[i]]);
    }
    var startRow = imageSheet.getLastRow() + 1;
    var endRow = startRow + writeRows.length - 1;
    imageSheet.getRange('A' + startRow + ':D' + endRow).setValues(writeRows);
  }

  questionsSheet.getRange(row, 7).setValue(chunks.length ? ('PARTS:' + chunks.length) : '');
  SpreadsheetApp.flush();

  return { partCount: chunks.length, imageBytes: imageData.length };
}

function questionImageUploadKey_(row, suffix) {
  return 'qi_' + row + '_' + suffix;
}

function clearQuestionImageUpload_(row) {
  var cache = CacheService.getScriptCache();
  var metaJson = cache.get(questionImageUploadKey_(row, 'meta'));
  var partCount = 0;
  if (metaJson) {
    try {
      partCount = JSON.parse(metaJson).partCount || 0;
    } catch (e) {
      partCount = 0;
    }
  }
  cache.remove(questionImageUploadKey_(row, 'meta'));
  for (var i = 0; i < QUESTION_IMAGE_UPLOAD_MAX_PARTS_; i++) {
    cache.remove(questionImageUploadKey_(row, 'part_' + i));
  }
}

function prepareQuestionImageData_(base64Image) {
  if (!base64Image || typeof base64Image !== 'string') {
    logQuizEvent_('warn', 'prepareQuestionImageData_', 'missing or non-string image payload', {
      type: typeof base64Image,
      length: base64Image ? String(base64Image).length : 0
    });
    return '';
  }

  var photoData = base64Image.trim().replace(/\s+/g, '');
  if (!photoData) {
    logQuizEvent_('warn', 'prepareQuestionImageData_', 'empty image payload after trim');
    return '';
  }
  if (!photoData.startsWith('data:image')) {
    logQuizEvent_('error', 'prepareQuestionImageData_', 'invalid image prefix', {
      prefix: photoData.substring(0, 32)
    });
    throw new Error('Question image must be a PNG or GIF file.');
  }
  if (!/^data:image\/(png|gif);base64,/i.test(photoData)) {
    logQuizEvent_('error', 'prepareQuestionImageData_', 'unsupported image type', {
      prefix: photoData.substring(0, 32)
    });
    throw new Error('Question image must be a PNG or GIF file.');
  }

  var commaIndex = photoData.indexOf(',');
  if (commaIndex < 1) {
    throw new Error('Question image must be a PNG or GIF file.');
  }

  var base64 = photoData.substring(commaIndex + 1);
  if (!base64) {
    throw new Error('Question image must be a PNG or GIF file.');
  }

  var approxBytes = Math.floor(base64.length * 3 / 4);
  if (approxBytes > MAX_QUESTION_IMAGE_BYTES_) {
    throw new Error('Question image must be 1 MB or smaller.');
  }

  return photoData;
}

function beginQuestionImageUpload(row, partCount, totalLength, questionId, quizId) {
  row = Number(row);
  partCount = Number(partCount);
  totalLength = Number(totalLength);
  questionId = normalizeSheetId_(questionId);
  quizId = normalizeSheetId_(quizId);

  if (!row || row < 2) {
    throw new Error('Invalid question row.');
  }
  if (!partCount || partCount < 1) {
    throw new Error('Image upload is missing data parts.');
  }
  if (partCount > QUESTION_IMAGE_UPLOAD_MAX_PARTS_) {
    throw new Error('Image is too large to upload.');
  }

  clearQuestionImageUpload_(row);
  CacheService.getScriptCache().put(
    questionImageUploadKey_(row, 'meta'),
    JSON.stringify({
      partCount: partCount,
      totalLength: totalLength,
      questionId: questionId,
      quizId: quizId
    }),
    600
  );

  return { ok: true, row: row, partCount: partCount };
}

function uploadQuestionImagePart(row, partIndex, partData) {
  row = Number(row);
  partIndex = Number(partIndex);

  if (!row || row < 2) {
    throw new Error('Invalid question row.');
  }
  if (!partData || typeof partData !== 'string' || !partData.length) {
    logQuizEvent_('error', 'uploadQuestionImagePart', 'missing part payload', {
      row: row,
      partIndex: partIndex
    });
    throw new Error('Image part ' + partIndex + ' did not reach the server.');
  }

  var cache = CacheService.getScriptCache();
  var metaJson = cache.get(questionImageUploadKey_(row, 'meta'));
  if (!metaJson) {
    throw new Error('Image upload session expired. Save the question again.');
  }

  cache.put(questionImageUploadKey_(row, 'part_' + partIndex), partData, 600);
  return { ok: true, row: row, partIndex: partIndex, partLength: partData.length };
}

function uploadQuestionImageParts(row, startIndex, partsJson) {
  row = Number(row);
  startIndex = Number(startIndex);
  var parts;
  try {
    parts = JSON.parse(partsJson);
  } catch (parseError) {
    throw new Error('Image upload payload is invalid.');
  }
  if (!parts || !parts.length) {
    throw new Error('Image upload payload is empty.');
  }

  for (var i = 0; i < parts.length; i++) {
    uploadQuestionImagePart(row, startIndex + i, parts[i]);
  }

  return { ok: true, row: row, startIndex: startIndex, count: parts.length };
}

function finalizeQuestionImageUpload(row, partCount, questionId, quizId) {
  row = Number(row);
  partCount = Number(partCount);

  if (!row || row < 2) {
    throw new Error('Invalid question row.');
  }

  var cache = CacheService.getScriptCache();
  var metaJson = cache.get(questionImageUploadKey_(row, 'meta'));
  if (!metaJson) {
    throw new Error('Image upload session expired. Save the question again.');
  }

  var meta;
  try {
    meta = JSON.parse(metaJson);
  } catch (parseError) {
    throw new Error('Image upload metadata is invalid.');
  }

  if (partCount !== meta.partCount) {
    logQuizEvent_('error', 'finalizeQuestionImageUpload', 'part count mismatch', {
      row: row,
      expected: meta.partCount,
      received: partCount
    });
    throw new Error('Image upload part count mismatch.');
  }

  questionId = normalizeSheetId_(questionId || meta.questionId);
  quizId = normalizeSheetId_(quizId || meta.quizId);
  if (!questionId) {
    throw new Error('Question ID is missing for image upload.');
  }

  var parts = [];
  for (var i = 0; i < partCount; i++) {
    var chunk = cache.get(questionImageUploadKey_(row, 'part_' + i));
    if (!chunk) {
      logQuizEvent_('error', 'finalizeQuestionImageUpload', 'missing cached part', {
        row: row,
        partIndex: i
      });
      throw new Error('Image part ' + i + ' was missing during upload.');
    }
    parts.push(chunk);
  }

  var assembled = parts.join('');

  if (meta.totalLength && assembled.length !== meta.totalLength) {
    logQuizEvent_('error', 'finalizeQuestionImageUpload', 'assembled length mismatch', {
      row: row,
      assembledLength: assembled.length,
      expectedLength: meta.totalLength
    });
    throw new Error('Image upload was incomplete. Check the Logs sheet for details.');
  }

  var imageData = prepareQuestionImageData_(assembled);
  if (!imageData) {
    throw new Error('Image data was empty after upload.');
  }

  var ss = getQuizSpreadsheet_();
  var questionsSheet = ensureSheet_(ss, 'Questions');
  ensureQuestionsHeaders_(questionsSheet);

  if (row > questionsSheet.getLastRow()) {
    throw new Error('Question row ' + row + ' was not found.');
  }

  var stored = writeQuestionImageToSheet_(questionsSheet, row, questionId, quizId, imageData);
  clearQuestionImageUpload_(row);
  return {
    row: row,
    imageSaved: true,
    imageBytes: imageData.length,
    partCount: stored.partCount
  };
}

function setQuestionImage(row, base64Image, questionId, quizId) {
  beginQuestionImageUpload(row, 1, base64Image ? base64Image.length : 0, questionId, quizId);
  uploadQuestionImagePart(row, 0, base64Image);
  return finalizeQuestionImageUpload(row, 1, questionId, quizId);
}

function getIdsFromResponseRow_(responsesSheet, row) {
  if (!responsesSheet || row < 2 || row > responsesSheet.getLastRow()) {
    return { questionId: '', quizId: '' };
  }

  ensureResponsesHeaders_(responsesSheet);
  var data = responsesSheet.getRange('A' + row + ':F' + row).getValues()[0];
  return {
    questionId: normalizeSheetId_(data[4]),
    quizId: normalizeSheetId_(data[5])
  };
}

function getQuestionBankForQuiz_(questionsSheet, quizId, skipImages) {
  ensureQuestionsHeaders_(questionsSheet);
  ensureQuestionIds_(questionsSheet);

  quizId = normalizeSheetId_(quizId);
  var bank = {};
  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) {
    return bank;
  }

  var rows = questionsSheet.getRange('A2:G' + lastRow).getValues();
  var ss = questionsSheet.getParent();
  var number = 0;

  for (var i = 0; i < rows.length; i++) {
    var rowQuizId = normalizeSheetId_(rows[i][4]);
    var questionId = normalizeSheetId_(rows[i][5]);
    if (rowQuizId !== quizId || !questionId || !cellText_(rows[i][1])) {
      continue;
    }

    number++;
    bank[questionId] = {
      questionText: cellText_(rows[i][1]),
      rubricText: cellText_(rows[i][2]),
      imageData: skipImages
        ? ''
        : readQuestionImageDataForQuestion_(ss, questionId, rowQuizId, rows[i][6]),
      hasImage: questionHasImageMarker_(rows[i][6]),
      questionNumber: number
    };
  }

  return bank;
}

function lookupQuestionInBank_(bank, questionId) {
  questionId = normalizeSheetId_(questionId);
  if (!questionId) {
    return null;
  }
  if (bank[questionId]) {
    return bank[questionId];
  }

  for (var key in bank) {
    if (bank.hasOwnProperty(key) && normalizeSheetId_(key) === questionId) {
      return bank[key];
    }
  }

  return null;
}

function resolveQuestionIdForEval_(evalRowData, responseIds, selectedQuizId) {
  selectedQuizId = normalizeSheetId_(selectedQuizId);
  var evalColF = normalizeSheetId_(evalRowData[5]);
  var evalColG = normalizeSheetId_(evalRowData[6]);

  if (evalColF && evalColF !== selectedQuizId) {
    return evalColF;
  }
  if (evalColG && evalColG !== selectedQuizId) {
    return evalColG;
  }
  if (responseIds.questionId && responseIds.questionId !== selectedQuizId) {
    return responseIds.questionId;
  }
  return evalColF || evalColG || responseIds.questionId;
}

function findQuestionById_(questionsSheet, questionId, quizId, questionBank) {
  questionId = normalizeSheetId_(questionId);
  quizId = normalizeSheetId_(quizId);
  if (!questionId) {
    return null;
  }

  var bankEntry = questionBank ? lookupQuestionInBank_(questionBank, questionId) : null;
  if (bankEntry) {
    return bankEntry;
  }

  ensureQuestionsHeaders_(questionsSheet);
  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  var values = questionsSheet.getRange('A2:G' + lastRow).getValues();
  var questionDisplay = questionsSheet.getRange('B2:B' + lastRow).getDisplayValues();
  var fallback = null;

  for (var i = 0; i < values.length; i++) {
    if (normalizeSheetId_(values[i][5]) !== questionId) {
      continue;
    }

    var entry = {
      questionText: cellText_(questionDisplay[i][0] || values[i][1]),
      rubricText: cellText_(values[i][2]),
      imageData: '',
      hasImage: questionHasImageMarker_(values[i][6]),
      quizName: cellText_(values[i][3]),
      quizId: normalizeSheetId_(values[i][4]),
      questionId: questionId,
      questionNumber: 999
    };

    if (!quizId || entry.quizId === quizId) {
      return entry;
    }

    if (!fallback) {
      fallback = entry;
    }
  }

  return fallback;
}

function resolveEvalRowIds_(evalRowData, responseIds, selectedQuizId) {
  var questionId = resolveQuestionIdForEval_(evalRowData, responseIds, selectedQuizId);
  var evalQuizId = normalizeSheetId_(evalRowData[6]);
  var quizId = evalQuizId || responseIds.quizId;

  if (normalizeSheetId_(evalRowData[5]) === selectedQuizId && evalQuizId) {
    quizId = evalQuizId;
  } else if (responseIds.quizId) {
    quizId = responseIds.quizId;
  }

  return {
    questionId: questionId,
    quizId: quizId
  };
}

function isPendingStatus_(status) {
  return String(status).trim().toLowerCase() === 'pending';
}

function getQuestion() {
  return getQuestionInfo().question;
}

function getQuestionInfo() {
  var ss = getQuizSpreadsheet_();
  var questionsSheet = ss.getSheetByName('Questions');
  if (!questionsSheet) {
    throw new Error('Questions sheet not found.');
  }
  ensureQuestionsHeaders_(questionsSheet);
  var responsesSheet = ss.getSheetByName('Responses');
  var row = responsesSheet ? getNextResponseRow_(responsesSheet) : 2;
  var qr = getQuestionAndRubricForRow_(questionsSheet, row);
  return {
    row: row,
    question: qr.question,
    rubric: qr.rubric,
    quizName: qr.quizName,
    quizId: qr.quizId
  };
}

function addQuestionAndRubric(quizNameOrPayload, question, rubric) {
  var imageData = '';
  var quizName;
  if (quizNameOrPayload && typeof quizNameOrPayload === 'object') {
    quizName = quizNameOrPayload.quizName;
    question = quizNameOrPayload.question;
    rubric = quizNameOrPayload.rubric;
    imageData = quizNameOrPayload.imageData || '';
  } else {
    quizName = quizNameOrPayload;
  }

  quizName = String(quizName).trim();
  question = String(question).trim();
  rubric = String(rubric).trim();

  if (!quizName) {
    throw new Error('Quiz name is required.');
  }
  if (!question) {
    throw new Error('Question is required.');
  }
  if (!rubric) {
    throw new Error('Rubric is required.');
  }

  var photoData = '';
  if (imageData) {
    photoData = prepareQuestionImageData_(imageData);
  }

  var ss = getQuizSpreadsheet_();
  var questionsSheet = ensureSheet_(ss, 'Questions');
  ensureQuestionsHeaders_(questionsSheet);

  var quizId = findQuizIdByName_(questionsSheet, quizName);
  var isNewQuiz = !quizId;
  if (!quizId) {
    quizId = generateUniqueQuizId_(questionsSheet);
  }

  var row = getNextQuestionsRow_(questionsSheet);
  var questionId = generateUniqueQuestionId_(questionsSheet);
  questionsSheet.appendRow([row, question, rubric, quizName, quizId, questionId, photoData]);

  SpreadsheetApp.flush();

  var message = 'Question saved in row ' + row + '.';
  if (isNewQuiz) {
    message += ' New quiz ID: ' + quizId + '.';
  } else {
    message += ' Using existing quiz ID: ' + quizId + '.';
  }
  message += ' Question ID: ' + questionId + '.';
  if (photoData) {
    message += ' Image saved (' + photoData.length + ' characters).';
  }

  return {
    row: row,
    quizName: quizName,
    quizId: quizId,
    questionId: questionId,
    isNewQuiz: isNewQuiz,
    imageSaved: !!photoData,
    imageBytes: photoData ? photoData.length : 0,
    message: message
  };
}

function getQuizList() {
  var ss = getQuizSpreadsheet_();
  var questionsSheet = ss.getSheetByName('Questions');
  if (!questionsSheet) {
    return [];
  }

  ensureQuestionsHeaders_(questionsSheet);
  ensureQuestionIds_(questionsSheet);

  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var rows = questionsSheet.getRange('B2:F' + lastRow).getValues();
  var quizzes = {};
  var list = [];

  for (var i = 0; i < rows.length; i++) {
    var question = rows[i][0];
    var quizName = String(rows[i][2]).trim();
    var quizId = normalizeSheetId_(rows[i][3]);
    if (!question || !quizId) {
      continue;
    }

    if (!quizzes[quizId]) {
      quizzes[quizId] = {
        quizId: quizId,
        quizName: quizName || 'Unnamed Quiz',
        questionCount: 0
      };
      list.push(quizzes[quizId]);
    }
    quizzes[quizId].questionCount++;
  }

  return list;
}

function getQuizEvaluationList() {
  var list = getQuizList();
  var ss = getQuizSpreadsheet_();
  var responsesSheet = ss.getSheetByName('Responses');
  var pendingByQuiz = {};

  if (responsesSheet && responsesSheet.getLastRow() >= 2) {
    ensureResponsesHeaders_(responsesSheet);
    var lastRow = responsesSheet.getLastRow();
    var rows = responsesSheet.getRange('A2:F' + lastRow).getValues();

    for (var i = 0; i < rows.length; i++) {
      if (!isPendingStatus_(rows[i][3])) {
        continue;
      }
      var rowQuizId = normalizeSheetId_(rows[i][5]);
      if (!rowQuizId) {
        continue;
      }
      pendingByQuiz[rowQuizId] = (pendingByQuiz[rowQuizId] || 0) + 1;
    }
  }

  for (var j = 0; j < list.length; j++) {
    list[j].pendingCount = pendingByQuiz[list[j].quizId] || 0;
  }

  var timingByQuiz = getLastEvaluationTimingByQuiz_();
  var quizzesWithTiming = 0;
  for (j = 0; j < list.length; j++) {
    list[j].lastEvaluation = timingByQuiz[list[j].quizId] || null;
    if (list[j].lastEvaluation) {
      quizzesWithTiming++;
    }
  }

  logEvaluationUiEvent_('evaluateListLoaded', {
    quizCount: list.length,
    quizzesWithTiming: quizzesWithTiming
  });

  return serializeQuizEvaluationListForClient_(list);
}

function getQuizReviewList() {
  var list = getQuizList();
  var ss = getQuizSpreadsheet_();
  var evalSheet = ss.getSheetByName('Evaluations');
  var evaluatedByQuiz = {};
  var commentedByQuiz = {};

  if (evalSheet && evalSheet.getLastRow() >= 2) {
    ensureEvaluationsHeaders_(evalSheet);
    var lastRow = evalSheet.getLastRow();
    var rows = evalSheet.getRange('A2:H' + lastRow).getValues();

    for (var i = 0; i < rows.length; i++) {
      var evaluation = String(rows[i][4] || '').trim();
      var rowQuizId = String(rows[i][6] || '').trim();
      if (!rowQuizId || !evaluation) {
        continue;
      }
      evaluatedByQuiz[rowQuizId] = (evaluatedByQuiz[rowQuizId] || 0) + 1;
      if (cellText_(rows[i][7]).trim()) {
        commentedByQuiz[rowQuizId] = (commentedByQuiz[rowQuizId] || 0) + 1;
      }
    }
  }

  for (var j = 0; j < list.length; j++) {
    list[j].evaluatedCount = evaluatedByQuiz[list[j].quizId] || 0;
    list[j].commentedCount = commentedByQuiz[list[j].quizId] || 0;
  }

  return list;
}

function getStudentReviewQuizList(studentName) {
  studentName = String(studentName).trim();
  if (!studentName) {
    throw new Error('Student name is required.');
  }

  var list = getQuizList();
  var ss = getQuizSpreadsheet_();
  var evalSheet = ss.getSheetByName('Evaluations');
  var reviewedByQuiz = {};

  if (evalSheet && evalSheet.getLastRow() >= 2) {
    ensureEvaluationsHeaders_(evalSheet);
    var lastRow = evalSheet.getLastRow();
    var rows = evalSheet.getRange('A2:H' + lastRow).getValues();

    for (var i = 0; i < rows.length; i++) {
      var rowStudent = cellText_(rows[i][1]).trim();
      var evaluation = cellText_(rows[i][4]).trim();
      var rowQuizId = normalizeSheetId_(rows[i][6]);
      if (!rowQuizId || !evaluation || rowStudent !== studentName) {
        continue;
      }
      reviewedByQuiz[rowQuizId] = (reviewedByQuiz[rowQuizId] || 0) + 1;
    }
  }

  var result = [];
  for (var j = 0; j < list.length; j++) {
    var count = reviewedByQuiz[list[j].quizId] || 0;
    if (!count) {
      continue;
    }
    result.push({
      quizId: list[j].quizId,
      quizName: list[j].quizName,
      questionCount: list[j].questionCount,
      evaluatedCount: count
    });
  }

  return result;
}

function getStudentEvaluationsForQuiz(studentName, quizId) {
  studentName = String(studentName).trim();
  quizId = normalizeSheetId_(quizId);
  if (!studentName) {
    throw new Error('Student name is required.');
  }
  if (!quizId) {
    throw new Error('Quiz ID is required.');
  }

  var items = getEvaluationsForQuiz(quizId);
  var filtered = [];
  for (var i = 0; i < items.length; i++) {
    if (String(items[i].studentName).trim() === studentName) {
      filtered.push(items[i]);
    }
  }

  if (!filtered.length) {
    throw new Error('No evaluated results found for you on this quiz yet.');
  }

  filtered.sort(function(a, b) {
    return a.questionNumber - b.questionNumber;
  });

  return filtered;
}

function updateStudentReview(row, studentName, reviewText) {
  row = Number(row);
  studentName = String(studentName).trim();
  reviewText = String(reviewText || '').trim();

  if (!row || row < 2) {
    throw new Error('Invalid evaluation row.');
  }
  if (!studentName) {
    throw new Error('Student name is required.');
  }

  var ss = getQuizSpreadsheet_();
  var evalSheet = ss.getSheetByName('Evaluations');
  if (!evalSheet) {
    throw new Error('Evaluations sheet not found.');
  }

  ensureEvaluationsHeaders_(evalSheet);
  if (row > evalSheet.getLastRow()) {
    throw new Error('Evaluation row not found.');
  }

  var rowStudent = cellText_(evalSheet.getRange('B' + row).getValue()).trim();
  if (rowStudent !== studentName) {
    throw new Error('You can only update your own review comments.');
  }

  evalSheet.getRange('H' + row).setValue(reviewText);
  SpreadsheetApp.flush();
  return 'Your comment was saved.';
}

function getEvaluationsForQuiz(quizId) {
  quizId = normalizeSheetId_(quizId);
  if (!quizId) {
    throw new Error('Quiz ID is required.');
  }

  var ss = getQuizSpreadsheet_();
  var questionsSheet = ss.getSheetByName('Questions');
  var evalSheet = ss.getSheetByName('Evaluations');
  var responsesSheet = ss.getSheetByName('Responses');

  if (!questionsSheet) {
    throw new Error('Questions sheet not found.');
  }
  if (!evalSheet || evalSheet.getLastRow() < 2) {
    throw new Error('No evaluations found for this quiz yet.');
  }

  ensureQuestionsHeaders_(questionsSheet);
  ensureEvaluationsHeaders_(evalSheet);

  var questionBank = getQuestionBankForQuiz_(questionsSheet, quizId, true);
  if (!Object.keys(questionBank).length) {
    throw new Error('No questions found for that quiz.');
  }

  var lastRow = evalSheet.getLastRow();
  var rows = evalSheet.getRange('A2:I' + lastRow).getValues();
  var responseRows = [];
  if (responsesSheet && responsesSheet.getLastRow() >= 2) {
    ensureResponsesHeaders_(responsesSheet);
    responseRows = responsesSheet.getRange('A2:F' + responsesSheet.getLastRow()).getValues();
  }
  var items = [];
  var idBackfillRows = [];
  var idBackfillValues = [];

  for (var i = 0; i < rows.length; i++) {
    var evalRow = i + 2;
    var responseIds = { questionId: '', quizId: '' };
    if (evalRow - 2 < responseRows.length) {
      responseIds = {
        questionId: normalizeSheetId_(responseRows[evalRow - 2][4]),
        quizId: normalizeSheetId_(responseRows[evalRow - 2][5])
      };
    }
    var ids = resolveEvalRowIds_(rows[i], responseIds, quizId);
    var rowQuizId = ids.quizId;
    var questionId = ids.questionId;

    if (rowQuizId !== quizId && responseIds.quizId !== quizId) {
      continue;
    }
    if (rowQuizId !== quizId) {
      rowQuizId = responseIds.quizId;
      questionId = questionId || responseIds.questionId;
    }

    var evaluation = cellText_(rows[i][4]).trim();
    if (!evaluation) {
      continue;
    }

    var points = resolveEvaluationPointsForRow_(rows[i][8], evaluation);
    var evaluationText = resolveEvaluationFeedbackForClient_(rows[i][8], evaluation);

    if (!normalizeSheetId_(rows[i][5]) && questionId) {
      idBackfillRows.push(evalRow);
      idBackfillValues.push([questionId]);
    }
    if (!normalizeSheetId_(rows[i][6]) && rowQuizId) {
      evalSheet.getRange('G' + evalRow).setValue(rowQuizId);
    }

    var details = findQuestionById_(questionsSheet, questionId, quizId, questionBank);
    if (!details || !details.questionText) {
      var bankEntry = lookupQuestionInBank_(questionBank, questionId);
      if (bankEntry) {
        details = bankEntry;
      }
    }

    items.push({
      row: evalRow,
      timestamp: rows[i][0] instanceof Date ? rows[i][0].toISOString() : rows[i][0],
      studentName: cellText_(rows[i][1]),
      answerText: cellText_(rows[i][2]),
      rubricText: cellText_(rows[i][3]) || (details ? details.rubricText : ''),
      evaluationText: evaluationText,
      points: points,
      studentReviewText: cellText_(rows[i][7]),
      questionId: questionId,
      quizId: rowQuizId,
      qPrompt: details ? details.questionText : '',
      questionText: details ? details.questionText : '',
      hasImage: details ? !!details.hasImage : false,
      imageData: '',
      questionNumber: details ? details.questionNumber : 999
    });
  }

  if (idBackfillRows.length) {
    for (var b = 0; b < idBackfillRows.length; b++) {
      evalSheet.getRange('F' + idBackfillRows[b]).setValue(idBackfillValues[b][0]);
    }
  }

  SpreadsheetApp.flush();

  if (items.length === 0) {
    throw new Error('No evaluations found for this quiz yet.');
  }

  items.sort(function(a, b) {
    var nameCompare = String(a.studentName).localeCompare(String(b.studentName));
    if (nameCompare !== 0) {
      return nameCompare;
    }
    return a.questionNumber - b.questionNumber;
  });

  return items;
}

function updateEvaluation(row, evaluationText, points) {
  row = Number(row);
  if (!row || row < 2) {
    throw new Error('Invalid evaluation row.');
  }

  evaluationText = String(evaluationText).trim();
  if (!evaluationText) {
    throw new Error('Evaluation cannot be empty.');
  }

  points = normalizeEvaluationPoints_(points);
  if (points === null) {
    throw new Error('Points must be a whole number from 0 to ' + EVAL_SCORE_MAX_ + '.');
  }

  var ss = getQuizSpreadsheet_();
  var evalSheet = ss.getSheetByName('Evaluations');
  if (!evalSheet) {
    throw new Error('Evaluations sheet not found.');
  }

  ensureEvaluationsHeaders_(evalSheet);
  if (row > evalSheet.getLastRow()) {
    throw new Error('Evaluation row not found.');
  }

  var rowValues = evalSheet.getRange('B' + row + ':I' + row).getValues()[0];
  var currentEvaluation = cellText_(rowValues[3]).trim();
  var currentPoints = normalizeEvaluationPoints_(rowValues[7]);
  if (evaluationText !== currentEvaluation || points !== currentPoints) {
    appendEvaluationHistoryRows_([[
      new Date(),
      'Teacher',
      cellText_(rowValues[0]),
      normalizeSheetId_(rowValues[4]),
      normalizeSheetId_(rowValues[5]),
      row,
      evaluationText,
      points
    ]]);
  }

  evalSheet.getRange('E' + row).setValue(evaluationText);
  evalSheet.getRange('I' + row).setValue(points);
  SpreadsheetApp.flush();
  return 'Evaluation updated successfully.';
}

function getQuestionsForQuiz(quizId) {
  quizId = String(quizId).trim();
  if (!quizId) {
    throw new Error('Quiz ID is required.');
  }

  var ss = getQuizSpreadsheet_();
  var questionsSheet = ss.getSheetByName('Questions');
  if (!questionsSheet) {
    throw new Error('Questions sheet not found.');
  }

  ensureQuestionsHeaders_(questionsSheet);
  ensureQuestionIds_(questionsSheet);

  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var rows = questionsSheet.getRange('A2:G' + lastRow).getValues();
  var questionDisplay = questionsSheet.getRange('B2:B' + lastRow).getDisplayValues();
  var questions = [];

  for (var i = 0; i < rows.length; i++) {
    var rowQuizId = String(rows[i][4]).trim();
    var questionId = String(rows[i][5]).trim();
    var questionText = cellText_(questionDisplay[i][0] || rows[i][1]).trim();

    if (rowQuizId !== quizId || !questionText || !questionId) {
      continue;
    }

    questions.push({
      questionId: questionId,
      questionText: questionText,
      hasImage: questionHasImageMarker_(rows[i][6]),
      number: questions.length + 1
    });
  }

  if (questions.length === 0) {
    throw new Error('No questions found for that quiz.');
  }

  return questions;
}

function getQuestionIdsForQuiz_(quizId) {
  var questions = getQuestionsForQuiz(quizId);
  var ids = {};
  for (var i = 0; i < questions.length; i++) {
    ids[questions[i].questionId] = true;
  }
  return ids;
}

function submitQuiz(studentName, quizId, answers) {
  studentName = String(studentName).trim();
  quizId = String(quizId).trim();

  if (!studentName) {
    throw new Error('Student name is required.');
  }
  if (!quizId) {
    throw new Error('Quiz ID is required.');
  }
  if (typeof answers === 'string') {
    answers = JSON.parse(answers);
  }
  if (!answers || !answers.length) {
    throw new Error('No answers to submit.');
  }

  var expectedIds = getQuestionIdsForQuiz_(quizId);

  var submittedIds = {};
  for (var j = 0; j < answers.length; j++) {
    var questionId = String(answers[j].questionId || '').trim();
    var answer = String(answers[j].answer || '').trim();

    if (!questionId || !expectedIds[questionId]) {
      throw new Error('Invalid question in submission.');
    }
    if (!answer) {
      throw new Error('Please answer all questions before submitting.');
    }
    if (submittedIds[questionId]) {
      throw new Error('Duplicate answer for question ' + questionId + '.');
    }
    submittedIds[questionId] = true;
  }

  for (var expectedId in expectedIds) {
    if (expectedIds.hasOwnProperty(expectedId) && !submittedIds[expectedId]) {
      throw new Error('Please answer all questions before submitting.');
    }
  }

  var ss = getQuizSpreadsheet_();
  var sheet = ss.getSheetByName('Responses');
  if (!sheet) {
    sheet = ss.insertSheet('Responses');
  }
  ensureResponsesHeaders_(sheet);

  var timestamp = new Date();
  var newRows = [];
  for (var k = 0; k < answers.length; k++) {
    newRows.push([
      timestamp,
      studentName,
      String(answers[k].answer).trim(),
      'Pending',
      String(answers[k].questionId).trim(),
      quizId
    ]);
  }

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange('A' + startRow + ':F' + (startRow + newRows.length - 1)).setValues(newRows);

  SpreadsheetApp.flush();
  return 'Quiz submitted successfully! ' + answers.length + ' answer(s) saved with Pending status.';
}

function getAnthropicApiKey_() {
  var anthropicApiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!anthropicApiKey) {
    throw new Error(
      'ANTHROPIC_API_KEY is not set. Add it in Apps Script Project Settings under Script Properties.'
    );
  }
  return anthropicApiKey;
}

function chunkArray_(items, chunkSize) {
  var chunks = [];
  chunkSize = Number(chunkSize) || 1;
  for (var i = 0; i < items.length; i += chunkSize) {
    chunks.push(items.slice(i, i + chunkSize));
  }
  return chunks;
}

function countPendingInGroups_(questionOrder, groups) {
  var total = 0;
  var g;
  for (g = 0; g < questionOrder.length; g++) {
    total += groups[questionOrder[g]].entries.length;
  }
  return total;
}

function countEvaluationApiCalls_(questionOrder, groups) {
  var total = 0;
  var g;
  for (g = 0; g < questionOrder.length; g++) {
    var entryCount = groups[questionOrder[g]].entries.length;
    total += Math.ceil(entryCount / EVAL_BATCH_SIZE_);
  }
  return total;
}

function estimateEvaluationDurationMs_(apiCallCount) {
  var seconds = EVAL_ESTIMATE_SEC_OVERHEAD_ + (apiCallCount * EVAL_ESTIMATE_SEC_PER_API_CALL_);
  return Math.round(seconds * 1000);
}

function roundEvaluationSeconds_(ms) {
  return Math.round(ms / 100) / 10;
}

function formatEvaluationDuration_(seconds) {
  seconds = Math.round(Number(seconds) * 10) / 10;
  if (!isFinite(seconds) || seconds < 0) {
    return '--';
  }
  if (seconds < 60) {
    return seconds + ' sec';
  }
  var mins = Math.floor(seconds / 60);
  var secs = Math.round(seconds - (mins * 60));
  if (secs === 60) {
    mins += 1;
    secs = 0;
  }
  if (secs === 0) {
    return mins + ' min';
  }
  return mins + ' min ' + secs + ' sec';
}

function getQuizDisplayNameForId_(questionsSheet, quizId) {
  if (!quizId) {
    return 'All quizzes';
  }

  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) {
    return 'Unnamed Quiz';
  }

  var rows = questionsSheet.getRange('D2:E' + lastRow).getValues();
  var i;
  for (i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === quizId) {
      return String(rows[i][0]).trim() || 'Unnamed Quiz';
    }
  }

  return 'Unnamed Quiz';
}

function formatEvaluationAlertText_(result) {
  if (typeof result === 'string') {
    return result;
  }

  var lines = [result.message || 'AI evaluation completed.'];
  lines.push(
    result.questionCount + ' question' + (result.questionCount === 1 ? '' : 's') +
    ' · ' + result.pendingCount + ' pending response' + (result.pendingCount === 1 ? '' : 's')
  );
  lines.push(
    'Estimated time: ' + formatEvaluationDuration_(result.estimatedSec) +
    ' · Actual time: ' + formatEvaluationDuration_(result.actualSec)
  );
  return lines.join('\n');
}

function logEvaluationTiming_(phase, details) {
  var payload = {
    event: 'evaluationTiming',
    phase: phase
  };
  var key;
  for (key in details) {
    if (details.hasOwnProperty(key)) {
      payload[key] = details[key];
    }
  }

  if (payload.estimatedMs != null) {
    payload.estimatedSec = roundEvaluationSeconds_(payload.estimatedMs);
  }
  if (payload.actualMs != null) {
    payload.actualSec = roundEvaluationSeconds_(payload.actualMs);
  }
  if (payload.estimatedMs != null && payload.actualMs != null) {
    payload.deltaMs = payload.actualMs - payload.estimatedMs;
    payload.deltaSec = roundEvaluationSeconds_(payload.deltaMs);
    if (payload.estimatedMs > 0) {
      payload.deltaPct = Math.round((payload.deltaMs / payload.estimatedMs) * 1000) / 10;
    }
  }

  var message = 'evaluation ' + phase;
  var level = phase === 'failed' ? 'error' : 'info';
  logQuizEvent_(level, 'evaluation', message, payload);
}

function extractJsonFromModelText_(text) {
  var raw = String(text || '').trim();
  if (!raw) {
    throw new Error('Model returned empty text.');
  }

  if (raw.indexOf('```') !== -1) {
    var fenceMatch = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fenceMatch && fenceMatch[1]) {
      raw = fenceMatch[1].trim();
    }
  }

  var start = raw.indexOf('[');
  var end = raw.lastIndexOf(']');
  if (start === -1 || end === -1 || end <= start) {
    throw new Error('Model response did not include a JSON array.');
  }

  return JSON.parse(raw.substring(start, end + 1));
}

function formatBatchEvaluationText_(score, feedback) {
  feedback = String(feedback || '').trim();
  return feedback || '(no feedback returned)';
}

function collectPendingResponsesByQuestion_(allResponses, quizId) {
  quizId = quizId ? String(quizId).trim() : '';
  var groups = {};
  var questionOrder = [];

  for (var i = 0; i < allResponses.length; i++) {
    var rowData = allResponses[i];
    var status = rowData[3];
    var questionId = normalizeSheetId_(rowData[4]);
    var rowQuizId = normalizeSheetId_(rowData[5]);
    var studentName = cellText_(rowData[1]).trim();
    var answer = cellText_(rowData[2]).trim();

    if (quizId && rowQuizId !== quizId) {
      continue;
    }
    if (!isPendingStatus_(status)) {
      continue;
    }
    if (!studentName && !answer) {
      continue;
    }

    var groupKey = rowQuizId + '|' + questionId;
    if (!groups[groupKey]) {
      groups[groupKey] = {
        quizId: rowQuizId,
        questionId: questionId,
        entries: []
      };
      questionOrder.push(groupKey);
    }

    groups[groupKey].entries.push({
      responseRow: i + 2,
      timestamp: rowData[0],
      studentName: studentName || 'Unknown student',
      answer: answer || '(no answer provided)',
      questionId: questionId,
      quizId: rowQuizId
    });
  }

  return {
    groups: groups,
    questionOrder: questionOrder
  };
}

function sortQuestionGroupKeys_(questionOrder, groups, questionBank) {
  function questionNumberForKey(groupKey) {
    var group = groups[groupKey];
    if (!group || !group.questionId) {
      return 9999;
    }
    var entry = lookupQuestionInBank_(questionBank, group.questionId);
    return entry ? entry.questionNumber : 9999;
  }

  questionOrder.sort(function(a, b) {
    return questionNumberForKey(a) - questionNumberForKey(b);
  });

  return questionOrder;
}

function callAnthropicBatchGrade_(apiKey, question, rubric, entries) {
  var studentPayload = [];
  for (var i = 0; i < entries.length; i++) {
    studentPayload.push({
      id: String(entries[i].responseRow),
      studentName: entries[i].studentName,
      answer: entries[i].answer
    });
  }

  var userMessage = [
    'You are grading student quiz answers for ONE question.',
    '',
    EVAL_GRADING_PHILOSOPHY_,
    '',
    'Question:',
    question,
    '',
    'Question-specific rubric (lists requested facts and common misconceptions):',
    rubric,
    '',
    'Students to grade (JSON):',
    JSON.stringify(studentPayload),
    '',
    'Return ONLY valid JSON (no markdown fences) as an array with one object per student:',
    '[{"id":"<response row id>","score":3,"feedback":"Brief feedback here."}, ...]',
    '',
    'Rules:',
    '- Include exactly one object for every student in the input.',
    '- Use the same id values from the input.',
    '- score must be a whole number from 0 to ' + EVAL_SCORE_MAX_ + ' (per the grading philosophy above).',
    '- feedback is stored separately from the score; do not write "Score: X/4" in feedback.',
    '- feedback has two parts in this order:',
    '  1) Context-specific feedback: cite what the student got right or wrong on this question, referencing their answer and the rubric requested facts; note misconceptions when relevant.',
    '  2) Score justification (closing sentence): state why the score fits the Grading for Equity level (all, most, some, one, or none of the requested facts).',
    '- Example structure: "You correctly identified … but did not … which the rubric asked for. This earns a 3 because most requested facts are correct, but not all."',
    '- Be specific and constructive; use as many sentences as needed for substantive feedback (typically 3-6).'
  ].join('\n');

  var payload = {
    model: EVAL_MODEL_,
    max_tokens: EVAL_BATCH_MAX_TOKENS_,
    messages: [
      {
        role: 'user',
        content: userMessage
      }
    ]
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var responseBody = response.getContentText();
  var responseData;
  try {
    responseData = JSON.parse(responseBody);
  } catch (parseError) {
    return {
      ok: false,
      error: 'API returned non-JSON: ' + responseBody.substring(0, 500)
    };
  }

  if (response.getResponseCode() !== 200) {
    return {
      ok: false,
      error: 'API error: ' + (responseData.error && responseData.error.message
        ? responseData.error.message
        : responseBody)
    };
  }

  try {
    var parsed = extractJsonFromModelText_(responseData.content[0].text);
    return {
      ok: true,
      results: parsed
    };
  } catch (jsonError) {
    return {
      ok: false,
      error: 'Could not parse model JSON: ' + jsonError.message,
      raw: responseData.content && responseData.content[0] ? responseData.content[0].text : ''
    };
  }
}

function writeBatchEvaluationResults_(
  evalSheet,
  responsesSheet,
  entries,
  rubric,
  questionId,
  quizId,
  batchResult
) {
  var resultMap = {};
  var i;
  var written = 0;

  if (batchResult.ok) {
    for (i = 0; i < batchResult.results.length; i++) {
      var item = batchResult.results[i];
      resultMap[String(item.id)] = item;
    }
  }

  var historyRows = [];

  for (i = 0; i < entries.length; i++) {
    var entry = entries[i];
    var evaluationText;
    var newStatus;

    if (!batchResult.ok) {
      evaluationText = batchResult.error;
      newStatus = 'Error';
    } else if (!resultMap[String(entry.responseRow)]) {
      evaluationText = 'Missing from batch evaluation response.';
      newStatus = 'Error';
    } else {
      var graded = resultMap[String(entry.responseRow)];
      evaluationText = formatBatchEvaluationText_(graded.score, graded.feedback);
      newStatus = 'Complete';
    }

    var points = null;
    if (batchResult.ok && resultMap[String(entry.responseRow)]) {
      points = normalizeEvaluationPoints_(resultMap[String(entry.responseRow)].score);
    }

    evalSheet.getRange('A' + entry.responseRow + ':G' + entry.responseRow).setValues([[
      entry.timestamp,
      entry.studentName,
      entry.answer,
      rubric,
      evaluationText,
      questionId,
      quizId
    ]]);
    if (points !== null) {
      evalSheet.getRange('I' + entry.responseRow).setValue(points);
    }
    responsesSheet.getRange('D' + entry.responseRow).setValue(newStatus);
    historyRows.push([
      new Date(),
      'AI',
      entry.studentName,
      questionId,
      quizId,
      entry.responseRow,
      evaluationText,
      points !== null ? points : ''
    ]);
    written++;
  }

  appendEvaluationHistoryRows_(historyRows);
  return written;
}

function triggerEvaluation(quizId) {
  return triggerEvaluationByQuestion_(quizId);
}

function buildEvaluationSteps_(questionsSheet, questionOrder, groups) {
  var steps = [];
  var g;
  var b;

  for (g = 0; g < questionOrder.length; g++) {
    var groupKey = questionOrder[g];
    var group = groups[groupKey];
    var questionId = group.questionId;
    var rowQuizId = group.quizId;
    var qr = questionId
      ? getQuestionAndRubricByQuestionId_(questionsSheet, questionId, rowQuizId, true)
      : { questionText: '', rubricText: '' };
    var question = qr.questionText || qr.question || '';
    var rubric = qr.rubricText || qr.rubric || '';
    var batches = chunkArray_(group.entries, EVAL_BATCH_SIZE_);

    for (b = 0; b < batches.length; b++) {
      steps.push({
        questionId: questionId,
        rowQuizId: rowQuizId,
        question: question,
        rubric: rubric,
        entries: batches[b]
      });
    }
  }

  return steps;
}

function buildEvaluationRunPlan_(quizId) {
  quizId = quizId ? String(quizId).trim() : '';

  var ss = getQuizSpreadsheet_();
  var questionsSheet = ss.getSheetByName('Questions');
  var responsesSheet = ss.getSheetByName('Responses');

  if (!questionsSheet) {
    throw new Error('Questions sheet not found.');
  }
  if (!responsesSheet || responsesSheet.getLastRow() < 2) {
    throw new Error('No responses found to evaluate. Submit an answer from the student page first.');
  }

  ensureResponsesHeaders_(responsesSheet);
  ensureSheet_(ss, 'Evaluations');
  ensureEvaluationsHeaders_(ss.getSheetByName('Evaluations'));

  var lastRow = responsesSheet.getLastRow();
  var allResponses = responsesSheet.getRange('A2:F' + lastRow).getValues();
  var collected = collectPendingResponsesByQuestion_(allResponses, quizId);
  var questionOrder = collected.questionOrder;
  var groups = collected.groups;

  if (!questionOrder.length) {
    if (quizId) {
      throw new Error('No Pending responses found to evaluate for this quiz.');
    }
    throw new Error('No Pending responses found to evaluate.');
  }

  var questionBank = quizId ? getQuestionBankForQuiz_(questionsSheet, quizId, true) : {};
  questionOrder = sortQuestionGroupKeys_(questionOrder, groups, questionBank);
  var steps = buildEvaluationSteps_(questionsSheet, questionOrder, groups);
  var estimatedApiCalls = steps.length;

  return {
    evalRunId: generateShortId_(),
    startedAt: Date.now(),
    quizId: quizId,
    pendingCount: countPendingInGroups_(questionOrder, groups),
    questionGroupCount: questionOrder.length,
    estimatedMs: estimateEvaluationDurationMs_(estimatedApiCalls),
    apiCallCount: estimatedApiCalls,
    steps: steps,
    batchCallsSoFar: 0,
    responsesWrittenSoFar: 0
  };
}

function executeEvaluationPlanStep_(plan, stepIndex, evalSheet, responsesSheet, anthropicApiKey) {
  var step = plan.steps[stepIndex];
  var batchResult;
  var apiCallMade = false;

  if (!step.question && !step.rubric) {
    batchResult = {
      ok: false,
      error: 'Skipped: no question or rubric found for question ID ' +
        (step.questionId || '(missing)') + '.'
    };
  } else {
    batchResult = callAnthropicBatchGrade_(anthropicApiKey, step.question, step.rubric, step.entries);
    apiCallMade = true;
  }

  var written = writeBatchEvaluationResults_(
    evalSheet,
    responsesSheet,
    step.entries,
    step.rubric,
    step.questionId,
    step.rowQuizId,
    batchResult
  );

  return {
    written: written,
    apiCallMade: apiCallMade
  };
}

function buildEvaluationResult_(plan, actualMs) {
  var summary = 'AI evaluation completed for ' + plan.responsesWrittenSoFar +
    ' Pending response(s) using question-first batching (' +
    plan.batchCallsSoFar + ' API call' + (plan.batchCallsSoFar === 1 ? '' : 's') +
    '). Check the Evaluations tab.';
  var ss = getQuizSpreadsheet_();

  var result = {
    evalRunId: plan.evalRunId,
    message: summary,
    quizId: plan.quizId,
    quizName: getQuizDisplayNameForId_(ss.getSheetByName('Questions'), plan.quizId),
    questionCount: plan.questionGroupCount,
    pendingCount: plan.pendingCount,
    responsesWritten: plan.responsesWrittenSoFar,
    apiCallCount: plan.batchCallsSoFar,
    estimatedSec: roundEvaluationSeconds_(plan.estimatedMs),
    actualSec: roundEvaluationSeconds_(actualMs),
    deltaSec: roundEvaluationSeconds_(actualMs - plan.estimatedMs)
  };

  logEvaluationUiEvent_('serverResultReady', {
    evalRunId: result.evalRunId,
    quizId: result.quizId,
    estimatedSec: result.estimatedSec,
    actualSec: result.actualSec,
    deltaSec: result.deltaSec,
    pendingCount: result.pendingCount,
    apiCallCount: result.apiCallCount
  });

  return result;
}

function triggerEvaluationByQuestion_(quizId) {
  // Set ANTHROPIC_API_KEY in Apps Script: Project Settings > Script Properties.
  var plan = buildEvaluationRunPlan_(quizId);

  logEvaluationTiming_('started', {
    evalRunId: plan.evalRunId,
    quizId: plan.quizId || '(all)',
    pendingCount: plan.pendingCount,
    questionGroupCount: plan.questionGroupCount,
    apiCallCount: plan.apiCallCount,
    estimatedMs: plan.estimatedMs
  });

  var ss = getQuizSpreadsheet_();
  var evalSheet = ensureSheet_(ss, 'Evaluations');
  var responsesSheet = ss.getSheetByName('Responses');
  ensureEvaluationsHeaders_(evalSheet);
  ensureEvaluationHistorySheet_(ss);
  ensureResponsesHeaders_(responsesSheet);
  var anthropicApiKey = getAnthropicApiKey_();
  var stepIndex;

  try {
    for (stepIndex = 0; stepIndex < plan.steps.length; stepIndex++) {
      var stepResult = executeEvaluationPlanStep_(
        plan,
        stepIndex,
        evalSheet,
        responsesSheet,
        anthropicApiKey
      );
      if (stepResult.apiCallMade) {
        plan.batchCallsSoFar += 1;
      }
      plan.responsesWrittenSoFar += stepResult.written;
    }

    SpreadsheetApp.flush();
    var actualMs = Date.now() - plan.startedAt;

    logEvaluationTiming_('completed', {
      evalRunId: plan.evalRunId,
      quizId: plan.quizId || '(all)',
      pendingCount: plan.pendingCount,
      responsesWritten: plan.responsesWrittenSoFar,
      questionGroupCount: plan.questionGroupCount,
      apiCallCount: plan.batchCallsSoFar,
      estimatedMs: plan.estimatedMs,
      actualMs: actualMs
    });

    return buildEvaluationResult_(plan, actualMs);
  } catch (error) {
    logEvaluationTiming_('failed', {
      evalRunId: plan.evalRunId,
      quizId: plan.quizId || '(all)',
      pendingCount: plan.pendingCount,
      responsesWritten: plan.responsesWrittenSoFar,
      questionGroupCount: plan.questionGroupCount,
      apiCallCount: plan.batchCallsSoFar,
      estimatedMs: plan.estimatedMs,
      actualMs: Date.now() - plan.startedAt,
      error: error.message
    });
    throw error;
  }
}

function debugQuizState() {
  var ss = getQuizSpreadsheet_();
  var responsesSheet = ss.getSheetByName('Responses');
  var evalSheet = ss.getSheetByName('Evaluations');

  return {
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    responseCount: responsesSheet ? Math.max(responsesSheet.getLastRow() - 1, 0) : 0,
    responses: responsesSheet ? responsesSheet.getRange('A1:F' + responsesSheet.getLastRow()).getValues() : [],
    evaluationsHasHeaders: evalSheet ? evalSheet.getRange('A1').getValue() : null,
    evaluationCount: evalSheet ? Math.max(evalSheet.getLastRow() - 1, 0) : 0
  };
}

function escapeForFormula(text) {
  return String(text)
    .replace(/"/g, '""')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function seedHistogramPilotQuizWithAlert() {
  try {
    var result = seedHistogramPilotQuiz();
    SpreadsheetApp.getUi().alert(
      'Histogram pilot quiz created.\n\n' +
      'Quiz: ' + result.quizName + '\n' +
      'Quiz ID: ' + result.quizId + '\n' +
      'Questions: ' + result.questionIds.join(', ') + '\n' +
      'Pending responses: ' + result.responseCount
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Could not seed pilot quiz:\n\n' + error.message);
  }
}

function seedHistogramPilotQuiz() {
  var quizName = 'Histograms Pilot Quiz';
  var ss = getQuizSpreadsheet_();
  var questionsSheet = ensureSheet_(ss, 'Questions');
  ensureQuestionsHeaders_(questionsSheet);
  var responsesSheet = ensureSheet_(ss, 'Responses');
  ensureResponsesHeaders_(responsesSheet);

  if (findQuizIdByName_(questionsSheet, quizName)) {
    throw new Error(
      'A quiz named "' + quizName + '" already exists. Rename or delete it before seeding again.'
    );
  }

  var quizId = generateUniqueQuizId_(questionsSheet);
  var question1Id = generateUniqueQuestionId_(questionsSheet);
  var question2Id = generateUniqueQuestionId_(questionsSheet);

  var question1Text =
    'The histogram summarizes pages read per week for 30 students. Bin counts: 0-10 pages: 8 students; ' +
    '11-20 pages: 6 students; 21-30 pages: 4 students; 31-40 pages: 2 students; 41-50 pages: 1 student; ' +
    '51-60 pages: 1 student. Describe the shape of this distribution. State whether it is symmetric, ' +
    'skewed left, or skewed right. Identify the most common bin and explain what the tail shows.';

  var question1Rubric =
    'Requested facts: distribution is unimodal; correctly identifies right skew and explains the tail extends ' +
    'toward higher page counts; identifies 0-10 pages as the most common bin with 8 students; notes fewer students ' +
    'in higher bins forming the tail; clear answer tied to the data. Common misconceptions: symmetric or left skew ' +
    'without evidence; naming the peak only without shape or tail reasoning.';

  var question2Text =
    'The histogram shows hours of sleep per night for 30 students. Bin counts: 4-5 hours: 2 students; ' +
    '5-6 hours: 4 students; 6-7 hours: 9 students; 7-8 hours: 10 students; 8-9 hours: 4 students; ' +
    '9-10 hours: 1 student. Describe the center and spread of this distribution. Give a reasonable typical ' +
    'value and explain how much variability you see, using evidence from the bins.';

  var question2Rubric =
    'Requested facts: center estimate in the 6-8 hour range with justification from the highest bins ' +
    '(6-7: 9 students, 7-8: 10 students); describes spread/variability using the range across bins ' +
    '(roughly 4-5 to 9-10 hours) or similar reasoning; notes the distribution is approximately symmetric; ' +
    'cites specific bin counts as evidence. Partial understanding: center OR spread alone without the other.';

  var row1 = getNextQuestionsRow_(questionsSheet);
  var row2 = row1 + 1;
  questionsSheet.getRange('A' + row1 + ':G' + row2).setValues([
    [row1, question1Text, question1Rubric, quizName, quizId, question1Id, ''],
    [row2, question2Text, question2Rubric, quizName, quizId, question2Id, '']
  ]);

  var students = [
    {
      name: 'Alex Chen',
      q1: 'The distribution is unimodal and right-skewed. The peak is the 0-10 pages bin with 8 students, ' +
        'so the most common amount is relatively low reading. There is a long tail toward higher page counts, ' +
        'with fewer students in each bin from 21-30 upward, including 1 student in 51-60. It is not symmetric ' +
        'because the tail extends to the right while most students cluster on the left.',
      q2: 'The center is around 7 hours of sleep because the 7-8 bin has 10 students and the 6-7 bin has 9. ' +
        'The distribution is fairly symmetric with most values between 6 and 8 hours. Spread is moderate: ' +
        'students range from about 4-5 hours up to 9-10 hours, so there is roughly a 5-hour spread across bins.'
    },
    {
      name: 'Jordan Lee',
      q1: 'The histogram is unimodal with the highest bar at 0-10 pages (8 students). I think it is roughly ' +
        'symmetric because the counts look somewhat balanced around the middle bins near 11-20 pages.',
      q2: 'Typical sleep is about 7 hours since the 7-8 bin is largest with 10 students. Students vary from ' +
        'around 4-5 hours to 9-10 hours, so there is noticeable variability.'
    },
    {
      name: 'Sam Rivera',
      q1: 'Most students read between 0 and 10 pages per week because that bin has 8 people. A few students read ' +
        'more pages in the higher bins, but I am not sure about the exact shape name.',
      q2: 'The middle is around 7-8 hours because those bins have the most students. Some sleep less and some ' +
        'sleep more than that.'
    },
    {
      name: 'Taylor Kim',
      q1: 'The shape is skewed left because more students are on the left side of the histogram. The most common ' +
        'bin is 0-10 pages.',
      q2: 'Center is about 7 hours. Spread is big because the bin counts are different across the histogram.'
    },
    {
      name: 'Morgan Davis',
      q1: 'It goes up and then down.',
      q2: 'Most people sleep around 7 hours.'
    }
  ];

  var timestamp = new Date();
  var responseRows = [];
  var i;
  for (i = 0; i < students.length; i++) {
    responseRows.push([
      timestamp,
      students[i].name,
      students[i].q1,
      'Pending',
      question1Id,
      quizId
    ]);
    responseRows.push([
      timestamp,
      students[i].name,
      students[i].q2,
      'Pending',
      question2Id,
      quizId
    ]);
  }

  var startRow = responsesSheet.getLastRow() + 1;
  responsesSheet.getRange('A' + startRow + ':F' + (startRow + responseRows.length - 1)).setValues(responseRows);
  SpreadsheetApp.flush();

  logQuizEvent_('info', 'seedHistogramPilotQuiz', 'pilot quiz seeded', {
    quizName: quizName,
    quizId: quizId,
    questionIds: [question1Id, question2Id],
    responseCount: responseRows.length
  });

  return {
    quizName: quizName,
    quizId: quizId,
    questionIds: [question1Id, question2Id],
    responseCount: responseRows.length,
    studentCount: students.length
  };
}

function seedExtendedHistogramPilotQuizWithAlert() {
  try {
    var result = seedExtendedHistogramPilotQuiz();
    SpreadsheetApp.getUi().alert(
      'Extended histogram pilot quiz created.\n\n' +
      'Quiz: ' + result.quizName + '\n' +
      'Quiz ID: ' + result.quizId + '\n' +
      'Questions in sheet: ' + result.questionIds.length + '\n' +
      'Students: ' + result.studentCount + '\n' +
      'Response rows (students x questions): ' + result.responseCount + '\n' +
      'Expected API calls at evaluate: ' + result.expectedApiCalls
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Could not seed extended pilot quiz:\n\n' + error.message);
  }
}

function seedExtendedHistogramPilotQuiz() {
  var quizName = 'Histograms Extended Pilot Quiz';
  var ss = getQuizSpreadsheet_();
  var questionsSheet = ensureSheet_(ss, 'Questions');
  ensureQuestionsHeaders_(questionsSheet);
  var responsesSheet = ensureSheet_(ss, 'Responses');
  ensureResponsesHeaders_(responsesSheet);

  if (findQuizIdByName_(questionsSheet, quizName)) {
    throw new Error(
      'A quiz named "' + quizName + '" already exists. Delete those question rows (and related responses) before seeding again.'
    );
  }

  var questions = getExtendedPilotQuestionBank_();
  var quizId = generateUniqueQuizId_(questionsSheet);
  var questionIds = [];
  var q;
  for (q = 0; q < questions.length; q++) {
    questionIds.push(generateUniqueQuestionId_(questionsSheet));
  }

  var startRow = getNextQuestionsRow_(questionsSheet);
  var questionRows = [];
  for (q = 0; q < questions.length; q++) {
    questionRows.push([
      startRow + q,
      questions[q].text,
      questions[q].rubric,
      quizName,
      quizId,
      questionIds[q],
      ''
    ]);
  }
  questionsSheet
    .getRange('A' + startRow + ':G' + (startRow + questions.length - 1))
    .setValues(questionRows);

  var students = buildExtendedPilotStudents_(questions.length);
  var timestamp = new Date();
  var responseRows = [];
  var i;
  for (i = 0; i < students.length; i++) {
    if (students[i].answers.length !== questions.length) {
      throw new Error(
        'Seed data mismatch for ' + students[i].name + ': expected ' +
        questions.length + ' answers, got ' + students[i].answers.length + '.'
      );
    }
    var s;
    for (s = 0; s < questions.length; s++) {
      responseRows.push([
        timestamp,
        students[i].name,
        students[i].answers[s],
        'Pending',
        questionIds[s],
        quizId
      ]);
    }
  }

  var responseStartRow = getNextResponseRow_(responsesSheet);
  responsesSheet
    .getRange('A' + responseStartRow + ':F' + (responseStartRow + responseRows.length - 1))
    .setValues(responseRows);
  SpreadsheetApp.flush();

  var batchesPerQuestion = Math.ceil(students.length / EVAL_BATCH_SIZE_);
  var expectedApiCalls = questions.length * batchesPerQuestion;

  logQuizEvent_('info', 'seedExtendedHistogramPilotQuiz', 'extended pilot quiz seeded', {
    quizName: quizName,
    quizId: quizId,
    questionIds: questionIds,
    studentCount: students.length,
    responseCount: responseRows.length,
    expectedApiCalls: expectedApiCalls
  });

  return {
    quizName: quizName,
    quizId: quizId,
    questionIds: questionIds,
    responseCount: responseRows.length,
    studentCount: students.length,
    expectedApiCalls: expectedApiCalls
  };
}

function getExtendedPilotQuestionBank_() {
  return [
    {
      text:
        'The histogram summarizes pages read per week for 30 students. Bin counts: 0-10 pages: 8 students; ' +
        '11-20 pages: 6 students; 21-30 pages: 4 students; 31-40 pages: 2 students; 41-50 pages: 1 student; ' +
        '51-60 pages: 1 student. Describe the shape of this distribution. State whether it is symmetric, ' +
        'skewed left, or skewed right. Identify the most common bin and explain what the tail shows.',
      rubric:
        'Requested facts: distribution is unimodal; correctly identifies right skew and explains the tail extends ' +
        'toward higher page counts; identifies 0-10 pages as the most common bin with 8 students; notes fewer students ' +
        'in higher bins forming the tail; clear answer tied to the data. Common misconceptions: symmetric or left skew ' +
        'without evidence; naming the peak only without shape or tail reasoning.'
    },
    {
      text:
        'The histogram shows hours of sleep per night for 30 students. Bin counts: 4-5 hours: 2 students; ' +
        '5-6 hours: 4 students; 6-7 hours: 9 students; 7-8 hours: 10 students; 8-9 hours: 4 students; ' +
        '9-10 hours: 1 student. Describe the center and spread of this distribution. Give a reasonable typical ' +
        'value and explain how much variability you see, using evidence from the bins.',
      rubric:
        'Requested facts: center estimate in the 6-8 hour range with justification from the highest bins ' +
        '(6-7: 9 students, 7-8: 10 students); describes spread/variability using the range across bins ' +
        '(roughly 4-5 to 9-10 hours) or similar reasoning; notes the distribution is approximately symmetric; ' +
        'cites specific bin counts as evidence. Partial understanding: center OR spread alone without the other.'
    },
    {
      text:
        'Histogram A shows minutes to walk to school for 25 students. Bin counts: 0-5 min: 2; 6-10 min: 4; ' +
        '11-15 min: 8; 16-20 min: 7; 21-25 min: 3; 26-30 min: 1. Histogram B shows daily minutes of homework ' +
        'for the same 25 students. Bin counts: 0-15 min: 1; 16-30 min: 3; 31-45 min: 5; 46-60 min: 8; ' +
        '61-75 min: 5; 76-90 min: 2; 91-105 min: 1. Which distribution has greater spread? Explain using ' +
        'evidence from the bin ranges and counts.',
      rubric:
        'Requested facts: Histogram B has greater spread; B spans roughly 0-105 minutes versus A spanning roughly ' +
        '0-30 minutes; cites specific bin ranges as evidence; explains spread as wider range or more variability ' +
        'across bins; may note A is more clustered. Common misconceptions: picks A because it has more bins in the ' +
        'middle; confuses spread with skew; compares tallest bars without discussing range.'
    },
    {
      text:
        'A histogram of quiz scores for 28 students has bins: 50-59: 1 student; 60-69: 4 students; 70-79: 7 students; ' +
        '80-89: 11 students; 90-99: 5 students. Describe the shape of the distribution and identify the most common ' +
        'score range. Is the distribution skewed? Explain using evidence from the bins.',
      rubric:
        'Requested facts: unimodal distribution; peak or most common bin is 80-89 with 11 students; shape is ' +
        'approximately symmetric or slightly left-skewed with justification from fewer students in low versus high ' +
        'extreme bins; uses bin counts as evidence. Common misconceptions: wrong peak bin; claims heavy skew without ' +
        'evidence; describes shape without naming the modal bin.'
    },
    {
      text:
        'A histogram of insects per plant has bins: 0-2 insects: 6 plants; 3-5 insects: 9 plants; 6-8 insects: 5 plants; ' +
        '9-11 insects: 0 plants; 12-14 insects: 3 plants. A student claims the empty 9-11 bin proves no plant ever has ' +
        '9 to 11 insects. Evaluate that claim. What might the gap between 6-8 and 12-14 represent?',
      rubric:
        'Requested facts: empty bin does not prove no values exist in that range; notes 12-14 bin has 3 plants so ' +
        'higher counts do occur; explains gap may reflect bin width, clustering, or limited sample rather than ' +
        'impossibility; rejects the absolute claim. Common misconceptions: treats empty bin as proof of zero frequency ' +
        'in the population; ignores nearby nonzero bins.'
    },
    {
      text:
        'Histogram A shows daily screen time (minutes) for 20 students: 0-30: 1; 31-60: 3; 61-90: 6; 91-120: 7; ' +
        '121-150: 3. Histogram B shows daily exercise time (minutes) for the same students: 0-15: 4; 16-30: 8; ' +
        '31-45: 5; 46-60: 2; 61-75: 1. Which distribution has a higher typical value? Estimate the center of each ' +
        'and compare using evidence from the bins.',
      rubric:
        'Requested facts: screen time (A) has higher typical value than exercise (B); A center near 91-120 with ' +
        'justification from the 7 students in that bin and neighboring bins; B center near 16-30 with justification ' +
        'from the 8 students there; compares both centers explicitly. Common misconceptions: picks taller total bins ' +
        'without comparing scales; confuses spread with center.'
    },
    {
      text:
        'A histogram of customer wait times (minutes) has bins: 0-2: 12 customers; 3-5: 8 customers; 6-8: 3 customers; ' +
        '9-11: 1 customer; 12-14: 0 customers; 15-17: 4 customers. Describe the shape. Is it skewed left, skewed right, ' +
        'or roughly symmetric? Use the bin counts to justify your answer.',
      rubric:
        'Requested facts: identifies right skew with most customers in low wait bins and a tail toward higher waits; ' +
        'notes concentration in 0-5 minutes (20 customers) versus fewer in high bins including 15-17 (4 customers); ' +
        'may mention gap at 12-14 without claiming impossibility. Common misconceptions: symmetric because bins exist ' +
        'on both sides; left skew because low bins are on the left side of the axis.'
    },
    {
      text:
        'Two classes took the same quiz. Class 1 scores (20 students): 50-59: 0; 60-69: 2; 70-79: 5; 80-89: 8; 90-99: 5. ' +
        'Class 2 scores (20 students): 50-59: 3; 60-69: 6; 70-79: 7; 80-89: 3; 90-99: 1. Which class has greater spread ' +
        'in scores? Which class has a higher typical score? Answer both parts with evidence from the bins.',
      rubric:
        'Requested facts: Class 2 has greater spread with counts across a wider range including more low scores ' +
        '(50-59: 3, 60-69: 6) and fewer high scores; Class 1 has higher typical score with peak at 80-89 (8 students) ' +
        'versus Class 2 peak at 70-79 (7 students) and more low-end counts; cites bin evidence for both spread and center. ' +
        'Common misconceptions: confuses taller middle bar with spread; compares only one class for both parts.'
    }
  ];
}

function buildExtendedPilotStudents_(questionCount) {
  var names = [
    'Alex Chen', 'Jordan Lee', 'Sam Rivera', 'Taylor Kim', 'Morgan Davis',
    'Riley Patel', 'Casey Nguyen', 'Jamie Okafor', 'Avery Brooks', 'Quinn Martinez',
    'Drew Sullivan', 'Skyler Tan', 'Cameron Walsh', 'Reese Johnson', 'Parker Gomez',
    'Hayden Wright'
  ];
  var tiers = [
    'strong', 'mixed', 'strong', 'mixed', 'weak',
    'strong', 'mixed', 'weak', 'mixed', 'strong',
    'mixed', 'weak', 'mixed', 'strong', 'mixed', 'weak'
  ];
  var students = [];
  var i;
  for (i = 0; i < names.length; i++) {
    students.push({
      name: names[i],
      answers: buildExtendedPilotAnswersForTier_(tiers[i], i, questionCount)
    });
  }
  return students;
}

function buildExtendedPilotAnswersForTier_(tier, index, questionCount) {
  var pools = getExtendedPilotAnswerPools_();
  var answers = [];
  var q;
  for (q = 0; q < questionCount; q++) {
    var pool = pools[q];
    if (!pool) {
      throw new Error('Missing answer pool for extended pilot question ' + (q + 1) + '.');
    }
    var tierPool = pool[tier] || pool.weak;
    answers.push(tierPool[index % tierPool.length]);
  }
  return answers;
}

function getExtendedPilotAnswerPools_() {
  return [
    {
      strong: [
        'The distribution is unimodal and right-skewed. The peak is the 0-10 pages bin with 8 students, so most students ' +
          'read relatively few pages. The tail extends toward higher page counts with fewer students in each bin from ' +
          '21-30 upward, including 1 student in 51-60. It is not symmetric because the tail is on the right.',
        'Unimodal with a right skew: the highest bar is 0-10 pages (8 students) and counts drop as pages increase, ' +
          'forming a long tail to the right toward 51-60 pages.'
      ],
      mixed: [
        'The histogram is unimodal with the highest bar at 0-10 pages (8 students). I think it is roughly symmetric ' +
          'because the counts look somewhat balanced around the middle bins near 11-20 pages.',
        'Most students read between 0 and 10 pages per week because that bin has 8 people. A few students read more ' +
          'in higher bins, but I am not sure about the exact shape name.',
        'The shape is skewed left because more students are on the left side of the histogram. The most common bin is 0-10 pages.'
      ],
      weak: [
        'It goes up and then down.',
        'A lot of students read books.',
        'The histogram has bars.'
      ]
    },
    {
      strong: [
        'The center is around 7 hours of sleep because the 7-8 bin has 10 students and the 6-7 bin has 9. The ' +
          'distribution is fairly symmetric with most values between 6 and 8 hours. Spread is moderate: students range ' +
          'from about 4-5 hours up to 9-10 hours, roughly a 5-hour spread across bins.',
        'Typical sleep is near 7 hours since the 6-7 and 7-8 bins dominate (9 and 10 students). Variability spans ' +
          'from 4-5 to 9-10 hours, so the distribution is symmetric with moderate spread.'
      ],
      mixed: [
        'Typical sleep is about 7 hours since the 7-8 bin is largest with 10 students. Students vary from around ' +
          '4-5 hours to 9-10 hours, so there is noticeable variability.',
        'The middle is around 7-8 hours because those bins have the most students. Some sleep less and some sleep more.',
        'Center is about 7 hours. Spread is big because the bin counts are different across the histogram.'
      ],
      weak: [
        'Most people sleep around 7 hours.',
        'Sleep varies.',
        'The 7-8 bin is tallest.'
      ]
    },
    {
      strong: [
        'Histogram B has greater spread. Homework minutes range from about 0-15 up to 91-105, while walk times only ' +
          'range from 0-5 to 26-30 minutes. B is spread across many wider bins with counts in both low and high ranges, ' +
          'whereas A is clustered mostly between 6 and 20 minutes.',
        'B has more spread because its values cover roughly 105 minutes compared to about 30 minutes for A. The homework ' +
          'bins extend much farther on both ends, showing more variability than commute times.'
      ],
      mixed: [
        'Histogram A has more spread because it has more bins with students in them across the middle range from 6-30 minutes.',
        'They look about the same spread to me because both have several bins with counts.',
        'B has greater spread, but mainly because the tallest bar in B is higher.'
      ],
      weak: [
        'Histogram A because it is longer.',
        'I am not sure.',
        'The one with more students.'
      ]
    },
    {
      strong: [
        'The distribution is unimodal and approximately symmetric. The most common score range is 80-89 with 11 students. ' +
          'Counts decrease toward both tails, with only 1 student in 50-59 and 5 in 90-99, so there is no strong skew.',
        'Unimodal with peak at 80-89 (11 students). Shape is roughly symmetric since the lower and upper tails have similar ' +
          'declining counts (1 in 50-59 versus 5 in 90-99, with gradual steps in between).'
      ],
      mixed: [
        'The most common range is 80-89. I think it is skewed right because the high scores go up to 90-99.',
        'Peak is 70-79 with 7 students. The shape is skewed left.',
        'Most students scored in the 80s. The distribution has a tail.'
      ],
      weak: [
        'Most students did well.',
        'The middle bin is biggest.',
        'Scores go from 50 to 99.'
      ]
    },
    {
      strong: [
        'The empty 9-11 bin does not prove no plant ever has 9 to 11 insects. Three plants fall in 12-14, so higher counts ' +
          'occur just above the gap. The gap may reflect how bins are cut or natural clustering, not that those values are impossible.',
        'A zero count in one bin only means none of the sampled plants landed there. Because 12-14 has 3 plants, the claim is ' +
          'too strong; the gap likely separates two clusters rather than proving none exist between them.'
      ],
      mixed: [
        'The student might be wrong because bugs could still be between 9 and 11 even if this sample has none there.',
        'The gap means no plants had exactly 9-11 insects in this data set, but that might change with more plants.',
        'I think the claim is true because the bin is empty.'
      ],
      weak: [
        'The bin is zero so there are no insects in that range.',
        'Gaps mean nothing.',
        'There are insects.'
      ]
    },
    {
      strong: [
        'Screen time (A) has a higher typical value. Its center is near 91-120 minutes with 7 students there and strong counts ' +
          'in neighboring bins, while exercise (B) centers near 16-30 minutes with 8 students. A is clearly higher on average.',
        'Histogram A is centered higher: most screen times cluster around 61-150 minutes, especially 91-120, whereas exercise ' +
          'clusters around 16-30 with 8 students.'
      ],
      mixed: [
        'Screen time is higher because the tallest bar in A is 91-120. Exercise might be similar because both have a peak bin.',
        'Exercise is higher because 16-30 has 8 students.',
        'A is higher but I am not sure about the center of B.'
      ],
      weak: [
        'Screen time.',
        'They are the same.',
        'Exercise is higher.'
      ]
    },
    {
      strong: [
        'The distribution is right-skewed. Most customers wait 0-5 minutes (20 total in 0-2 and 3-5 bins), with a tail toward ' +
          'longer waits including 15-17 minutes (4 customers). It is not symmetric because the tail extends to the right.',
        'Right-skewed: counts are highest at short waits and taper toward longer waits, with only 1 customer in 9-11 and 4 in 15-17.'
      ],
      mixed: [
        'Roughly symmetric because there are customers in low and high bins.',
        'Left-skewed because the low bins are on the left.',
        'Most people wait a short time but I cannot name the skew.'
      ],
      weak: [
        'A lot of short waits.',
        'Symmetric.',
        'The histogram goes down.'
      ]
    },
    {
      strong: [
        'Class 2 has greater spread with more scores in low bins (50-59: 3, 60-69: 6) and only 1 in 90-99, while Class 1 is ' +
          'more concentrated high with peak 80-89 (8 students). Class 1 has the higher typical score because its center is ' +
          'around 80-89 versus Class 2 peaking at 70-79 (7 students).',
        'Spread is larger in Class 2 because scores range more widely across low and middle bins. Class 1 has higher typical ' +
          'scores with 8 students in 80-89 and 5 in 90-99 compared with Class 2\'s peak at 70-79.'
      ],
      mixed: [
        'Class 1 has greater spread because it has more students in 80-89. Class 2 has the higher typical score.',
        'Class 2 spreads more. Class 1 is higher because of the 90-99 bin only.',
        'They are about the same on both measures.'
      ],
      weak: [
        'Class 1 on both.',
        'Class 2 on both.',
        'The classes are different.'
      ]
    }
  ];
}
