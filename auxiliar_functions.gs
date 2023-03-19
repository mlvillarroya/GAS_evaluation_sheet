function onOpen() {
  const ctes = constants();
  SpreadsheetApp.getUi().createMenu('Evaluation')
    .addSubMenu(
        SpreadsheetApp.getUi().createMenu('Students')
        .addItem('Create student data sheet','student_data_sheet')) 
    .addSubMenu(
        SpreadsheetApp.getUi().createMenu('Avaluation sheets')
        .addItem('Create blank evaluation sheet','evaluation_sheet')
        .addItem('Create evaluation columns','compute_evaluation_sheet')
        .addItem('Fill undone rows with \"Correct\"','fill_undone_rows'))
    .addToUi();
}

function student_data_sheet(){
  const ctes = constants();
  var ss = ctes.SS;
  if (spreadsheet_exists(ss,ctes.STUDENT_DATA_SHEET_NAME)) {
    SpreadsheetApp.getUi().alert('Sheet already exists (' + ctes.STUDENT_DATA_SHEET_NAME + '). Please fill in this sheet');
    return;
    }
  ss.insertSheet();
  const sheet = ss.getActiveSheet();
  sheet.setName(ctes.STUDENT_DATA_SHEET_NAME);
  create_student_sheet_content(ctes,sheet);
  ss.moveActiveSheet(2);  
  }

function evaluation_sheet(){
  const ctes = constants();
  var ss = ctes.SS;
  if (!spreadsheet_exists(ss,ctes.STUDENT_DATA_SHEET_NAME))  {
    SpreadsheetApp.getUi().alert('Before creating evaluation sheet, student sheet is needed. Please, create and fill it first');
    return;
    }
  if (spreadsheet_exists(ss,ctes.BASE_EVALUATION_SHEET_NAME))  {
    SpreadsheetApp.getUi().alert('Sheet already exists (' + ctes.BASE_EVALUATION_SHEET_NAME + '). Please fill in this sheet');
    return;
    }
  ss.insertSheet();
  const sheet = ss.getActiveSheet();
  create_evaluation_blank_sheet_content(ctes,sheet);
  ss.moveActiveSheet(3);
}

function compute_evaluation_sheet(){
  const ctes = constants();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet_exists(ss,ctes.STUDENT_DATA_SHEET_NAME))  {
    SpreadsheetApp.getUi().alert('Before creating evaluation sheet, student sheet is needed. Please, create and fill it first');
    return;
    }
  var students_sheet = ss.getSheetByName(ctes.STUDENT_DATA_SHEET_NAME);
  var students_data = students_sheet.getDataRange().getValues();
  var students = [];
  for (var i = 1; i < students_sheet.getLastRow(); i++) {
    var student = students_data[i];
    students.push(student);
  }
  var sheet = ss.getActiveSheet();
  var exercise_name = sheet.getRange(1,2).getValue();
  if (exercise_name == null || exercise_name == '') {
    SpreadsheetApp.getUi().alert('Exercise name has to be filled');
    return;
  }
  var exercise_short_name = sheet.getRange(2,2).getValue();
  if (exercise_short_name == null || exercise_short_name == '') {
    SpreadsheetApp.getUi().alert('Exercise short name has to be filled');
    return;
  }
  var items_row = sheet.getRange(2,4,1,sheet.getLastColumn()-3).getValues()[0];
  var items=[];
  items_row.forEach((element)=>{
    if (!element.toString().startsWith("Item") && element != '') items.push(element);
  });
  var weights_row = sheet.getRange(4,4,1,items.length).getValues()[0];
  var weights = [];
  weights_row.forEach((element)=>{
    if (!isNaN(Number(element))) weights.push(Number(element));
    else weights.push(0);
  });
  var max_mark_cell_content = sheet.getRange(ctes.EVALUATION_MAX_MARK_CELL).getValue();
  var max_mark = !isNaN(Number(max_mark_cell_content)) ? max_mark_cell_content : 10;

  //ss.deleteSheet(sheet);
  sheet = ss.insertSheet();
  sheet.setName(exercise_short_name);
  ss.moveActiveSheet(ss.getNumSheets());
  create_student_sheet_content(ctes,sheet);
  sheet.getRange(ctes.EVALUATION_FIRST_ROW_NUMBER,1).activate();
  students.forEach((st)=>{
    var i = ss.getActiveRange().getRow();
    var j = ss.getActiveRange().getColumn();
    sheet.getRange(i,j).setValue(st[0]);
    sheet.getRange(i,j+1).setValue(st[1]);
    sheet.getRange(i,j+2).setValue(st[2]);
    sheet.getRange(i+1,j).activate();
  });
  sheet.autoResizeColumns(3, 1);
  sheet.getRange(1,ctes.EVALUATION_FIRST_COLUMN_NUMBER).activate();
  items.forEach((it)=>{
    var j = ss.getActiveRange().getColumn();
    sheet.getRange(ctes.EVALUATION_ITEMS_ROW,j).setValue(it);
    sheet.getRange(ctes.EVALUATION_ITEMS_ROW,j+2).activate();
  });
  var j = ss.getActiveRange().getColumn();
  sheet.getRange(1,j).setValue(ctes.PENALTY_COLUMN_TITLE).activate();
  var penalty_cell_column = get_active_cell_column_letter(sheet);
  sheet.getRange(1,j+1).setValue(ctes.EXTRA_COMMENT_COLUMN_TITLE).activate();
  var extra_comment_cell_column = get_active_cell_column_letter(sheet);
  sheet.getRange(1,j+2).setValue(ctes.MARK_COLUMN_TITLE).activate();
  var mark_cell_column = get_active_cell_column_letter(sheet);
  sheet.getRange(1,j+3).setValue(ctes.COMMENT_COLUMN_TITLE).activate();
  var comment_cell_column = get_active_cell_column_letter(sheet);
  sheet.getRange(1,j+4).setValue(ctes.DONE_COLUMN_TITLE).activate();
  var done_cell_column = get_active_cell_column_letter(sheet);
  var last_row = sheet.getLastRow() + 1;
  sheet.getRange(last_row, ctes.EVALUATION_FIRST_COLUMN_NUMBER).activate();
  weights.forEach((we)=>{
    var j = ss.getActiveRange().getColumn();
    sheet.getRange(last_row,j).setValue(we);
    sheet.getRange(last_row,j+2).activate();
  });
  sheet.getRange(1,ctes.EVALUATION_FIRST_COLUMN_NUMBER).activate();
  var evaluation_first_column_letter = get_active_cell_column_letter(sheet);
  sheet.getRange(1,ctes.EVALUATION_FIRST_COLUMN_NUMBER + 2 * (items.length - 1)).activate();
  var evaluation_last_column_letter = get_active_cell_column_letter(sheet);
  var weights_interval = "$" + evaluation_first_column_letter + '$' + last_row +":" + '$' + evaluation_last_column_letter + '$' + last_row;
  sheet.getRange(mark_cell_column + "1").activate();
  go_down_one_cell(sheet);
  var row_number = get_active_cell_row_number(sheet);
  var current_marks_interval = evaluation_first_column_letter + get_active_cell_row_number(sheet) +":" + evaluation_last_column_letter + get_active_cell_row_number(sheet);  
  sheet.getActiveCell().setValue("=IF(" + evaluation_first_column_letter + get_active_cell_row_number(sheet) + "=\"\";\"\";(SUMPRODUCT("+ current_marks_interval + ";" + weights_interval + ")/SUM(" + weights_interval + ")*"+ max_mark +"/10) + " + penalty_cell_column + get_active_cell_row_number(sheet) + ")");
  go_right_one_cell(sheet);
  var comment_phrase = "=IF(" + evaluation_first_column_letter + get_active_cell_row_number(sheet) + "<>\"\";\"<div>\"&  ";
  items.forEach((item)=>{
    var column_letter = get_cell_column_letter(find_first_cell_by_value(sheet,item));
    comment_phrase += "$" + column_letter + "$" + ctes.EVALUATION_ITEMS_ROW  + "&\": \"&" + column_letter + row_number + "*$" + column_letter + "$" + last_row + "/SUM($" + evaluation_first_column_letter + "$" + last_row + ":$" + evaluation_last_column_letter + "$" + last_row + ") * " + max_mark +"/10 &\" punts.      Comentari: \"&" + nextChar(column_letter) + row_number + "&\"<br>\"&"; 
  });
  comment_phrase += "\"<br>\" &" + extra_comment_cell_column + row_number + "& \"</div>\";\"\")";
  sheet.getActiveCell().setValue(comment_phrase);
  fill_down(ss,mark_cell_column,2,comment_cell_column,last_row-1);
  sheet.getRange('A1:'+done_cell_column+last_row).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);  
  sheet.getRange('D1:'+done_cell_column+'1').activate();  
  sheet.getActiveRangeList().setBackground('#fff2cc').setHorizontalAlignment('center');
  sheet.getRange('A' + last_row + ':'+done_cell_column+last_row).activate();  
  sheet.getActiveRangeList().setBackground('#fff2cc');
  sheet.getRange(mark_cell_column+'2:'+comment_cell_column+String(last_row-1)).activate();  
  sheet.getActiveRangeList().setBackground('#d9ead3');
  sheet.getRange('A1').setValue('WEIGHTS');
  sheet.getRange(done_cell_column+'2:'+done_cell_column+String(last_row-1)).setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Yes', 'No','X'], true)
  .build());
  sheet.getRange(done_cell_column + "2").activate();
  for (var i = 2; i < last_row; i++) {
    sheet.getActiveCell().setValue('No');
    go_down_one_cell(sheet);
  }
  sheet.getRange(done_cell_column + last_row).setValue("=COUNTIF(" + done_cell_column + "2:" + done_cell_column + String(last_row-1) + ";\"Yes\")");

  var conditionalFormatRules = sheet.getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([sheet.getRange('D2:'+ extra_comment_cell_column + String(last_row - 1))])
  .whenFormulaSatisfied('=$'+ done_cell_column +'2="X"')
  .setBackground('#EA9999')
  .build());
  sheet.setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = sheet.getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([sheet.getRange('D2:'+ extra_comment_cell_column + String(last_row - 1))])
  .whenFormulaSatisfied('=$'+ done_cell_column +'2="Yes"')
  .setBackground('#CFE2F3')
  .build());
  sheet.setConditionalFormatRules(conditionalFormatRules);

  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);
  sheet.getRange('A1').activate();

  if (!spreadsheet_exists(ss,ctes.VARIABLES_SHEET_NAME)) {
    ss.insertSheet();
    ss.getActiveSheet().setName(ctes.VARIABLES_SHEET_NAME);
    }
  sheet = ss.getSheetByName(ctes.VARIABLES_SHEET_NAME);
  sheet.getRange('A1').activate();
  while ((sheet.getActiveCell().getValue() != '') && (sheet.getActiveCell().getValue() != exercise_short_name)) go_down_one_cell(sheet);
  if (sheet.getActiveCell().getValue() != exercise_short_name) {
  sheet.getActiveCell().setValue(exercise_short_name);
  go_right_one_cell(sheet);
  sheet.getActiveCell().setValue(ctes.ITEMS_NUMBER_VARIABLE_NAME);
  go_right_one_cell(sheet);
  sheet.getActiveCell().setValue(items.length);
  go_down_and_left_one_row(sheet);
  sheet.getActiveCell().setValue(exercise_short_name);
  go_right_one_cell(sheet);
  sheet.getActiveCell().setValue(ctes.DONE_COLUMN_VARIABLE_NAME);
  go_right_one_cell(sheet);
  sheet.getActiveCell().setValue(done_cell_column);
  go_down_and_left_one_row(sheet);
  sheet.getActiveCell().setValue(exercise_short_name);
  go_right_one_cell(sheet);
  sheet.getActiveCell().setValue(ctes.ROWS_NUMBER_VARIABLE_NAME);
  go_right_one_cell(sheet);
  sheet.getActiveCell().setValue(last_row-2);
  }
  sheet.hideSheet();
}

function fill_undone_rows(){
  const ctes = constants();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if (!spreadsheet_exists(ss,ctes.VARIABLES_SHEET_NAME))  {
    SpreadsheetApp.getUi().alert('Error. Variables sheet is not created. Please, talk to the administrator.');
    return;
    }  
  var variables_sheet = ss.getSheetByName(ctes.VARIABLES_SHEET_NAME);
  var variables_data = variables_sheet.getDataRange().getValues();
  variables_data = variables_data.filter((row) => {
    return row[0] == sheet.getName();
  });  
  if (variables_data.length == 0)  {
    SpreadsheetApp.getUi().alert('Error. Not information about this sheet is found in variables sheet. Please, talk to the administrator.');
    return;
    }

  var items_number = look_for_variable(variables_data,ctes.ITEMS_NUMBER_VARIABLE_NAME);
  var rows_number = look_for_variable(variables_data,ctes.ROWS_NUMBER_VARIABLE_NAME);
  var done_column = look_for_variable(variables_data,ctes.DONE_COLUMN_VARIABLE_NAME);

  if (items_number == '' || done_column == '' || rows_number == '')  {
    SpreadsheetApp.getUi().alert('Error. Items number or Done column letter not found in variables sheet. Please, talk to the administrator.');
    return;
    }  
  variables_sheet.hideSheet();
  sheet.getRange("D2").activate();
  for (var i = 2; i<rows_number+2; i++) {
    if (sheet.getRange(done_column + i).getValue() == "No") {
      for (var j=0 ; j<items_number * 2; j += 2 ) {
        sheet.getRange(i,j+4).setValue(10);
        sheet.getRange(i,j+5).setValue(ctes.CORRECT_VALUE)
      }
    }
  }
}
