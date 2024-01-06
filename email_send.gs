function loadTemplate() {
  var archivoHTML = HtmlService.createHtmlOutputFromFile('email_template.html').getContent();
  return archivoHTML;
}

function sendEmail(adress, subject, contentHTML) {
  GmailApp.sendEmail(adress, subject, contentHTML,{ htmlBody: contentHTML });
}

function obtainColumn(text,matrix) {
  for (var i=0;i<matrix.length;i++) {
    if (matrix[0][i]==text) return i
  }
  return -1
}

function emailEveryStudent() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "You will send an email to EVERY STUDENT with the DONE flag to yes. Are you sure?",
    ui.ButtonSet.OK_CANCEL
  );
  // User clicked "OK" on first prompt
  if (result == ui.Button.OK) {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = ss.getDataRange().getValues();

    var columnDone = obtainColumn(constants().DONE_COLUMN_TITLE,data);
    var columnFullName = obtainColumn(constants().STUDENT_DATA_STUDENT_FIRST_NAME, data);
    var columnEmail = obtainColumn(constants().STUDENT_DATA_STUDENT_EMAIL,data);
    var columnMark = obtainColumn(constants().MARK_COLUMN_TITLE,data);
    var columnComment = obtainColumn(constants().COMMENT_COLUMN_TITLE,data);
    var activityName = ss.getSheetName();
    var subject = constants().EMAIL_SUBJECT + " " + activityName;

    for (var i = 1; i < data.length-1; i++) {
      var done = data[i][columnDone];
      if (done == "Yes") {
        var recipient = data[i][columnEmail];
        var mark = data[i][columnMark];
        var comment = data[i][columnComment];
        var fullName = data[i][columnFullName];
        
        if (recipient && mark && comment) {
          var contentHTML = loadTemplate();
          contentHTML = contentHTML.replace("[Full name]", fullName);
          contentHTML = contentHTML.replace("[Exercise_name]", activityName);
          contentHTML = contentHTML.replace("[Mark]", mark);
          contentHTML = contentHTML.replace("[Comment]", comment);

          sendEmail(recipient, subject, contentHTML);
        }
      }
    }
  }
}
