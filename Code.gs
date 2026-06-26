function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Quiz')
    .addItem('Set up sheets', 'setupSheet')
    .addItem('Add Question and Rubric', 'showAddQuestionDialog')
    .addItem('Open Logs', 'openLogsSheet')
    .addItem('Authorize API access (run once)', 'authorizeExternalRequests')
    .addItem('Run AI Evaluation', 'runEvaluationWithAlert')
    .addToUi();
}

function openLogsSheet() {
  var ss = getQuizSpreadsheet_();
  ensureLogsSheet_(ss);
  ss.setActiveSheet(ss.getSheetByName('Logs'));
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
    SpreadsheetApp.getUi().alert(result);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Evaluation failed:\n\n' + error.message);
  }
}

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

  setupQuestionsSheet_(questionsSheet);
  setupResponsesSheet_(responsesSheet);
  setupEvaluationsSheet_(evaluationsSheet);
  setupInstructionsSheet_(ss);
  ensureLogsSheet_(ss);
  ensureQuestionImagesSheet_(ss);

  removeDefaultSheet_(ss);

  return 'Sheet setup complete. Add questions via the teacher page (?teach) or Quiz > Add Question and Rubric.';
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
  sheet.getRange('A1:G1').setValues([
    ['Timestamp', 'Student Name', 'Answer', 'Rubric', 'AI Evaluation', 'Question ID', 'Quiz ID']
  ]);
  setColumnWidths_(sheet, [160, 160, 300, 300, 400, 100, 100]);
  sheet.setFrozenRows(1);
  formatHeaderRow_(sheet, 1, 7);
}

function ensureEvaluationsHeaders_(sheet) {
  if (sheet.getRange('A1').getValue() === 'Timestamp') {
    if (sheet.getRange('F1').getValue() !== 'Question ID') {
      sheet.getRange('F1').setValue('Question ID');
    }
    if (sheet.getRange('G1').getValue() !== 'Quiz ID') {
      sheet.getRange('G1').setValue('Quiz ID');
    }
    formatHeaderRow_(sheet, 1, 7);
    return;
  }
  setupEvaluationsSheet_(sheet);
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

  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][0]) {
      continue;
    }
    if (String(rows[i][4]).trim()) {
      continue;
    }
    var questionId = generateUniqueId_(existing);
    questionsSheet.getRange('F' + (i + 2)).setValue(questionId);
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

function getQuestionAndRubricForRow_(questionsSheet, row) {
  if (!questionsSheet) {
    throw new Error('Questions sheet not found.');
  }
  var ss = questionsSheet.getParent();
  var imageMap = buildQuestionImageMap_(ss);
  var questionId = questionsSheet.getRange('F' + row).getValue();
  var quizId = questionsSheet.getRange('E' + row).getValue();
  var legacyImage = questionsSheet.getRange('G' + row).getValue();
  return {
    question: questionsSheet.getRange('B' + row).getValue(),
    rubric: questionsSheet.getRange('C' + row).getValue(),
    quizName: questionsSheet.getRange('D' + row).getValue(),
    quizId: quizId,
    questionId: questionId,
    imageData: readQuestionImageData_(imageMap, legacyImage, questionId, quizId)
  };
}

function getQuestionAndRubricByQuestionId_(questionsSheet, questionId, quizId) {
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
  var imageMap = buildQuestionImageMap_(questionsSheet.getParent());
  var fallback = null;

  for (var i = 0; i < rows.length; i++) {
    if (normalizeSheetId_(rows[i][5]) !== questionId) {
      continue;
    }

    var entry = {
      questionText: cellText_(rows[i][1]),
      rubricText: cellText_(rows[i][2]),
      imageData: readQuestionImageData_(
        imageMap,
        rows[i][6],
        questionId,
        normalizeSheetId_(rows[i][4])
      ),
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
  var imageMap = buildQuestionImageMap_(ss);
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

  return readQuestionImageData_(imageMap, legacy, questionId, quizId);
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

  var saved = readQuestionImageData_(buildQuestionImageMap_(ss), '', questionId, quizId);
  logQuizEvent_('info', 'writeQuestionImageToSheet_', 'stored image', {
    row: row,
    questionId: questionId,
    quizId: quizId,
    partCount: chunks.length,
    savedLength: saved.length,
    expectedLength: imageData.length
  });

  if (saved.length !== imageData.length) {
    throw new Error('Image could not be stored in the sheet. See the Logs tab for details.');
  }

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

  logQuizEvent_('info', 'beginQuestionImageUpload', 'starting upload', {
    row: row,
    partCount: partCount,
    totalLength: totalLength,
    questionId: questionId,
    quizId: quizId
  });

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

  logQuizEvent_('info', 'uploadQuestionImagePart', 'received part', {
    row: row,
    partIndex: partIndex,
    partLength: partData ? String(partData).length : 0
  });

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

function finalizeQuestionImageUpload(row, partCount, questionId, quizId) {
  row = Number(row);
  partCount = Number(partCount);

  logQuizEvent_('info', 'finalizeQuestionImageUpload', 'assembling image', {
    row: row,
    partCount: partCount,
    questionId: questionId,
    quizId: quizId
  });

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
  logQuizEvent_('info', 'finalizeQuestionImageUpload', 'assembled image', {
    row: row,
    assembledLength: assembled.length,
    expectedLength: meta.totalLength
  });

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

function getQuestionBankForQuiz_(questionsSheet, quizId) {
  ensureQuestionsHeaders_(questionsSheet);
  ensureQuestionIds_(questionsSheet);

  quizId = normalizeSheetId_(quizId);
  var bank = {};
  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) {
    return bank;
  }

  var rows = questionsSheet.getRange('A2:G' + lastRow).getValues();
  var imageMap = buildQuestionImageMap_(questionsSheet.getParent());
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
      imageData: readQuestionImageData_(imageMap, rows[i][6], questionId, rowQuizId),
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
  var imageMap = buildQuestionImageMap_(questionsSheet.getParent());
  var fallback = null;

  for (var i = 0; i < values.length; i++) {
    if (normalizeSheetId_(values[i][5]) !== questionId) {
      continue;
    }

    var entry = {
      questionText: cellText_(questionDisplay[i][0] || values[i][1]),
      rubricText: cellText_(values[i][2]),
      imageData: readQuestionImageData_(
        imageMap,
        values[i][6],
        questionId,
        normalizeSheetId_(values[i][4])
      ),
      questionNumber: 999
    };

    if (!quizId || normalizeSheetId_(values[i][4]) === quizId) {
      return entry;
    }

    if (!fallback) {
      fallback = entry;
    }
  }

  return fallback;
}

function resolveEvalRowIds_(evalRowData, responsesSheet, evalRow, selectedQuizId) {
  var responseIds = getIdsFromResponseRow_(responsesSheet, evalRow);
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
  var quizName;
  if (quizNameOrPayload && typeof quizNameOrPayload === 'object') {
    quizName = quizNameOrPayload.quizName;
    question = quizNameOrPayload.question;
    rubric = quizNameOrPayload.rubric;
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

  var ss = getQuizSpreadsheet_();
  var questionsSheet = ensureSheet_(ss, 'Questions');
  ensureQuestionsHeaders_(questionsSheet);
  ensureQuestionImagesSheet_(ss);

  var quizId = findQuizIdByName_(questionsSheet, quizName);
  var isNewQuiz = !quizId;
  if (!quizId) {
    quizId = generateUniqueQuizId_(questionsSheet);
  }

  var row = getNextQuestionsRow_(questionsSheet);
  var questionId = generateUniqueQuestionId_(questionsSheet);
  questionsSheet.appendRow([row, question, rubric, quizName, quizId, questionId, '']);

  SpreadsheetApp.flush();
  logQuizEvent_('info', 'addQuestionAndRubric', 'question saved', {
    row: row,
    quizId: quizId,
    questionId: questionId,
    quizName: quizName
  });

  var message = 'Question saved in row ' + row + '.';
  if (isNewQuiz) {
    message += ' New quiz ID: ' + quizId + '.';
  } else {
    message += ' Using existing quiz ID: ' + quizId + '.';
  }
  message += ' Question ID: ' + questionId + '.';

  return {
    row: row,
    quizName: quizName,
    quizId: quizId,
    questionId: questionId,
    isNewQuiz: isNewQuiz,
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
    var quizId = String(rows[i][3]).trim();
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
      var rowQuizId = String(rows[i][5] || '').trim();
      if (!rowQuizId) {
        continue;
      }
      pendingByQuiz[rowQuizId] = (pendingByQuiz[rowQuizId] || 0) + 1;
    }
  }

  for (var j = 0; j < list.length; j++) {
    list[j].pendingCount = pendingByQuiz[list[j].quizId] || 0;
  }

  return list;
}

function getQuizReviewList() {
  var list = getQuizList();
  var ss = getQuizSpreadsheet_();
  var evalSheet = ss.getSheetByName('Evaluations');
  var evaluatedByQuiz = {};

  if (evalSheet && evalSheet.getLastRow() >= 2) {
    ensureEvaluationsHeaders_(evalSheet);
    var lastRow = evalSheet.getLastRow();
    var rows = evalSheet.getRange('A2:G' + lastRow).getValues();

    for (var i = 0; i < rows.length; i++) {
      var evaluation = String(rows[i][4] || '').trim();
      var rowQuizId = String(rows[i][6] || '').trim();
      if (!rowQuizId || !evaluation) {
        continue;
      }
      evaluatedByQuiz[rowQuizId] = (evaluatedByQuiz[rowQuizId] || 0) + 1;
    }
  }

  for (var j = 0; j < list.length; j++) {
    list[j].evaluatedCount = evaluatedByQuiz[list[j].quizId] || 0;
  }

  return list;
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

  var questionBank = getQuestionBankForQuiz_(questionsSheet, quizId);
  if (!Object.keys(questionBank).length) {
    throw new Error('No questions found for that quiz.');
  }

  var lastRow = evalSheet.getLastRow();
  var rows = evalSheet.getRange('A2:G' + lastRow).getValues();
  var items = [];

  for (var i = 0; i < rows.length; i++) {
    var evalRow = i + 2;
    var responseIds = getIdsFromResponseRow_(responsesSheet, evalRow);
    var ids = resolveEvalRowIds_(rows[i], responsesSheet, evalRow, quizId);
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

    if (!normalizeSheetId_(rows[i][5]) && questionId) {
      evalSheet.getRange('F' + evalRow).setValue(questionId);
    }
    if (!normalizeSheetId_(rows[i][6]) && rowQuizId) {
      evalSheet.getRange('G' + evalRow).setValue(rowQuizId);
    }

    var details = findQuestionById_(questionsSheet, questionId, quizId, questionBank);
    if (!details || !details.questionText) {
      var rowAligned = getQuestionAndRubricForRow_(questionsSheet, evalRow);
      if (cellText_(rowAligned.question)) {
        details = {
          questionText: cellText_(rowAligned.question),
          rubricText: cellText_(rowAligned.rubric),
          imageData: cellText_(rowAligned.imageData),
          questionNumber: lookupQuestionInBank_(questionBank, questionId)
            ? lookupQuestionInBank_(questionBank, questionId).questionNumber
            : 999
        };
      }
    }

    items.push({
      row: evalRow,
      timestamp: rows[i][0] instanceof Date ? rows[i][0].toISOString() : rows[i][0],
      studentName: cellText_(rows[i][1]),
      answerText: cellText_(rows[i][2]),
      rubricText: cellText_(rows[i][3]) || (details ? details.rubricText : ''),
      evaluationText: evaluation,
      questionId: questionId,
      quizId: rowQuizId,
      qPrompt: details ? details.questionText : '',
      questionText: details ? details.questionText : '',
      imageData: details ? details.imageData : '',
      questionNumber: details ? details.questionNumber : 999
    });
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

function updateEvaluation(row, evaluationText) {
  row = Number(row);
  if (!row || row < 2) {
    throw new Error('Invalid evaluation row.');
  }

  evaluationText = String(evaluationText).trim();
  if (!evaluationText) {
    throw new Error('Evaluation cannot be empty.');
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

  evalSheet.getRange('E' + row).setValue(evaluationText);
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
  ensureQuestionImagesSheet_(ss);

  var lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var rows = questionsSheet.getRange('A2:G' + lastRow).getValues();
  var questionDisplay = questionsSheet.getRange('B2:B' + lastRow).getDisplayValues();
  var imageMap = buildQuestionImageMap_(ss);
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
      hasImage: questionHasStoredImage_(imageMap, rows[i][6], questionId, rowQuizId),
      number: questions.length + 1
    });
  }

  if (questions.length === 0) {
    throw new Error('No questions found for that quiz.');
  }

  return questions;
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

  var quizQuestions = getQuestionsForQuiz(quizId);
  var expectedIds = {};
  for (var i = 0; i < quizQuestions.length; i++) {
    expectedIds[quizQuestions[i].questionId] = true;
  }

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
  for (var k = 0; k < answers.length; k++) {
    sheet.appendRow([
      timestamp,
      studentName,
      String(answers[k].answer).trim(),
      'Pending',
      String(answers[k].questionId).trim(),
      quizId
    ]);
  }

  SpreadsheetApp.flush();
  return 'Quiz submitted successfully! ' + answers.length + ' answer(s) saved with Pending status.';
}

function triggerEvaluation(quizId) {
  // Set ANTHROPIC_API_KEY in Apps Script: Project Settings > Script Properties.
  quizId = quizId ? String(quizId).trim() : '';
  var anthropicApiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!anthropicApiKey) {
    throw new Error(
      'ANTHROPIC_API_KEY is not set. Add it in Apps Script Project Settings under Script Properties.'
    );
  }

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

  var evalSheet = ensureSheet_(ss, 'Evaluations');
  ensureEvaluationsHeaders_(evalSheet);

  var lastRow = responsesSheet.getLastRow();
  var count = 0;

  for (var responseRow = 2; responseRow <= lastRow; responseRow++) {
    var rowData = responsesSheet.getRange('A' + responseRow + ':F' + responseRow).getValues()[0];
    var timestamp = rowData[0];
    var studentName = rowData[1];
    var answer = rowData[2];
    var status = rowData[3];
    var questionId = String(rowData[4] || '').trim();
    var rowQuizId = String(rowData[5] || '').trim();

    if (quizId && rowQuizId !== quizId) {
      continue;
    }

    if (!isPendingStatus_(status)) {
      continue;
    }

    if (!studentName && !answer) {
      continue;
    }

    studentName = studentName || 'Unknown student';
    answer = answer || '(no answer provided)';

    var qr;
    if (questionId) {
      qr = getQuestionAndRubricByQuestionId_(questionsSheet, questionId, rowQuizId);
    } else {
      qr = getQuestionAndRubricForRow_(questionsSheet, responseRow);
    }

    var question = qr.question || qr.questionText;
    var rubric = qr.rubric || qr.rubricText;
    var evaluationText;
    var newStatus;

    if (!question && !rubric) {
      var lookupHint = questionId
        ? 'question ID ' + questionId + (rowQuizId ? ' in quiz ' + rowQuizId : '')
        : 'Questions row ' + responseRow + ' (columns B and C)';
      evaluationText = 'Skipped: no question or rubric found for ' + lookupHint + '.';
      newStatus = 'Error';
    } else {
      var userMessage = [
        'You are grading a student quiz answer.',
        'Question: ' + question,
        'Rubric: ' + rubric,
        'Student answer: ' + answer,
        'Score the answer out of 10 and provide brief feedback.'
      ].join('\n');

      var payload = {
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 300,
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
          'x-api-key': anthropicApiKey,
          'anthropic-version': '2023-06-01',
          'content-type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      var responseBody = response.getContentText();

      try {
        var responseData = JSON.parse(responseBody);
        if (response.getResponseCode() !== 200) {
          evaluationText = 'API error: ' + (responseData.error && responseData.error.message
            ? responseData.error.message
            : responseBody);
          newStatus = 'Error';
        } else {
          evaluationText = responseData.content[0].text;
          newStatus = 'Complete';
        }
      } catch (parseError) {
        evaluationText = 'API error: ' + responseBody;
        newStatus = 'Error';
      }
    }

    evalSheet.getRange('A' + responseRow + ':G' + responseRow).setValues([[
      timestamp,
      studentName,
      answer,
      rubric,
      evaluationText,
      questionId,
      rowQuizId
    ]]);
    responsesSheet.getRange('D' + responseRow).setValue(newStatus);
    count++;
  }

  SpreadsheetApp.flush();

  if (count === 0) {
    if (quizId) {
      throw new Error('No Pending responses found to evaluate for this quiz.');
    }
    throw new Error('No Pending responses found to evaluate.');
  }

  if (quizId) {
    return 'AI evaluation completed for ' + count + ' Pending response(s) in this quiz. Check the Evaluations tab.';
  }

  return 'AI evaluation completed for ' + count + ' Pending response(s). Check the Evaluations tab.';
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
