function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Quiz')
    .addItem('Set up sheets', 'setupSheet')
    .addItem('Authorize API access (run once)', 'authorizeExternalRequests')
    .addItem('Run AI Evaluation', 'runEvaluationWithAlert')
    .addToUi();
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

  removeDefaultSheet_(ss);

  return 'Sheet setup complete. Add one question and rubric per row in Questions (columns B and C); row numbers match Responses.';
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
  sheet.getRange('A1:C1').setValues([['Row', 'Question', 'Rubric']]);
  sheet.getRange('A2').setValue(2);
  sheet.getRange('B2').setValue('What is the capital of France?');
  sheet.getRange('C2').setValue(
    'Award full credit for "Paris". Partial credit for mentioning France. ' +
    'Deduct points for incorrect capitals.'
  );
  setColumnWidths_(sheet, [60, 360, 360]);
  sheet.setFrozenRows(1);
  formatHeaderRow_(sheet, 1, 3);
}

function setupResponsesSheet_(sheet) {
  sheet.getRange('A1:D1').setValues([['Timestamp', 'Student Name', 'Answer', 'Status']]);
  setColumnWidths_(sheet, [160, 160, 360, 100]);
  sheet.setFrozenRows(1);
  formatHeaderRow_(sheet, 1, 4);
}

function setupEvaluationsSheet_(sheet) {
  sheet.getRange('A1:E1').setValues([['Timestamp', 'Student Name', 'Answer', 'Rubric', 'AI Evaluation']]);
  setColumnWidths_(sheet, [160, 160, 300, 300, 400]);
  sheet.setFrozenRows(1);
  formatHeaderRow_(sheet, 1, 5);
}

function setupInstructionsSheet_(ss) {
  var sheet = ss.getSheetByName('Setup');
  if (!sheet) {
    sheet = ss.insertSheet('Setup', 0);
  }

  sheet.clear();
  sheet.getRange('A1').setValue('Quiz Prototype Setup').setFontSize(14).setFontWeight('bold').setFontColor('#1a73e8');
  sheet.getRange('A3:A9').setValues([
    ['1. Add each question in Questions column B and its rubric in column C.'],
    ['2. Use the same row number as the matching response (row 2 with row 2, etc.).'],
    ['3. Deploy the script as a web app (Deploy > New deployment > Web app).'],
    ['4. Share the student URL: <web app url>?page=student'],
    ['5. Open the teacher URL (web app url without parameters) to run evaluations.'],
    ['6. Student answers appear on the Responses tab with Status = Pending.'],
    ['7. Run AI Evaluation to grade Pending rows; Status becomes Complete when done.']
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
  var page = e && e.parameter && e.parameter.page;
  if (page === 'student') {
    return HtmlService.createHtmlOutputFromFile('student')
      .setTitle('Student Quiz')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('teacher')
    .setTitle('Teacher Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getNextResponseRow_(responsesSheet) {
  ensureResponsesHeaders_(responsesSheet);
  var lastRow = responsesSheet.getLastRow();
  return lastRow < 2 ? 2 : lastRow + 1;
}

function getQuestionAndRubricForRow_(questionsSheet, row) {
  if (!questionsSheet) {
    throw new Error('Questions sheet not found.');
  }
  return {
    question: questionsSheet.getRange('B' + row).getValue(),
    rubric: questionsSheet.getRange('C' + row).getValue()
  };
}

function isPendingStatus_(status) {
  return String(status).trim().toLowerCase() === 'pending';
}

function getQuestion() {
  var ss = getQuizSpreadsheet_();
  var questionsSheet = ss.getSheetByName('Questions');
  if (!questionsSheet) {
    throw new Error('Questions sheet not found.');
  }
  var responsesSheet = ss.getSheetByName('Responses');
  var row = responsesSheet ? getNextResponseRow_(responsesSheet) : 2;
  return getQuestionAndRubricForRow_(questionsSheet, row).question;
}

function submitAnswer(studentName, answer) {
  var ss = getQuizSpreadsheet_();
  var sheet = ss.getSheetByName('Responses');
  if (!sheet) {
    sheet = ss.insertSheet('Responses');
  }
  ensureResponsesHeaders_(sheet);
  sheet.appendRow([new Date(), studentName, answer, 'Pending']);
  return 'Your answer has been submitted successfully!';
}

function triggerEvaluation() {
  // Set ANTHROPIC_API_KEY in Apps Script: Project Settings > Script Properties.
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
  setupEvaluationsSheet_(evalSheet);

  var lastRow = responsesSheet.getLastRow();
  var count = 0;

  for (var responseRow = 2; responseRow <= lastRow; responseRow++) {
    var rowData = responsesSheet.getRange('A' + responseRow + ':D' + responseRow).getValues()[0];
    var timestamp = rowData[0];
    var studentName = rowData[1];
    var answer = rowData[2];
    var status = rowData[3];

    if (!isPendingStatus_(status)) {
      continue;
    }

    if (!studentName && !answer) {
      continue;
    }

    studentName = studentName || 'Unknown student';
    answer = answer || '(no answer provided)';

    var qr = getQuestionAndRubricForRow_(questionsSheet, responseRow);
    var question = qr.question;
    var rubric = qr.rubric;
    var evaluationText;
    var newStatus;

    if (!question && !rubric) {
      evaluationText = 'Skipped: no question or rubric found in Questions row ' + responseRow + ' (columns B and C).';
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

    evalSheet.getRange('A' + responseRow + ':D' + responseRow).setValues([[timestamp, studentName, answer, rubric]]);
    evalSheet.getRange('E' + responseRow).setValue(evaluationText);
    responsesSheet.getRange('D' + responseRow).setValue(newStatus);
    count++;
  }

  SpreadsheetApp.flush();

  if (count === 0) {
    throw new Error('No Pending responses found to evaluate.');
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
    responses: responsesSheet ? responsesSheet.getRange('A1:D' + responsesSheet.getLastRow()).getValues() : [],
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
