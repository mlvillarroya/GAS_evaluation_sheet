function spreadsheet_exists(ss,sheetname){    
    if (ss.getSheetByName(sheetname)!= null) return true;
    return false;
}

function constants() {
  return {
    SS: SpreadsheetApp.getActiveSpreadsheet(),
    BASE_EVALUATION_SHEET_NAME: 'BaseEvaluation',
    STUDENT_DATA_SHEET_NAME: 'StudentData',    
    STUDENT_DATA_STUDENT_FIRST_NAME: 'First name',
    STUDENT_DATA_STUDENT_LAST_NAME: 'Last name',
    STUDENT_DATA_STUDENT_EMAIL: 'Email',
    EVALUATION_FIRST_COLUMN_NUMBER: 4,
    EVALUATION_FIRST_ROW_NUMBER: 2,
    EVALUATION_ITEMS_ROW: 1,
    EVALUATION_MAX_MARK_CELL: 'D6',
    PENALTY_COLUMN_TITLE: 'Penalty',
    EXTRA_COMMENT_COLUMN_TITLE: 'Extra comment',
    MARK_COLUMN_TITLE: 'Mark',
    COMMENT_COLUMN_TITLE: 'Comment',
    DONE_COLUMN_TITLE: 'Done',
    VARIABLES_SHEET_NAME: 'Variables',
    ITEMS_NUMBER_VARIABLE_NAME: 'Items number',
    DONE_COLUMN_VARIABLE_NAME: 'Done column',
    ROWS_NUMBER_VARIABLE_NAME: 'Rows number',
    CORRECT_VALUE: 'Correcte'
    // ...
  }
}

function create_student_sheet_content(ctes,sheet)
  {
    sheet.getRange(1,1).setValue(ctes.STUDENT_DATA_STUDENT_FIRST_NAME);  
    sheet.getRange(1,2).setValue(ctes.STUDENT_DATA_STUDENT_LAST_NAME);
    sheet.getRange(1,3).setValue(ctes.STUDENT_DATA_STUDENT_EMAIL);
    sheet.getRange('A1:C1').activate();
    sheet.getActiveRangeList().setBackground('#fff2cc');
  }

function create_evaluation_blank_sheet_content(ctes,sheet)
  {
    sheet.setName(ctes.BASE_EVALUATION_SHEET_NAME);
    sheet.getRange('A1').activate();
    sheet.getCurrentCell().setValue('Exercise name');
    sheet.getRange('A2').activate();
    sheet.getCurrentCell().setValue('Short name');
    sheet.getRange('B1').activate();
    sheet.setColumnWidth(2, 463);
    sheet.getRange('A1:B2').activate();
    sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange('A1:A2').activate();
    sheet.getActiveRangeList().setBackground('#fff2cc');
    sheet.getRange('D1').activate();
    sheet.getCurrentCell().setValue('Items');
    sheet.getRange('D2').activate();
    sheet.getCurrentCell().setValue('Item 1');
    sheet.getRange('E2').activate();
    sheet.getCurrentCell().setValue('Item 2');
    sheet.getRange('F2').activate();
    sheet.getCurrentCell().setValue('Item 3');
    sheet.getRange('G2').activate();
    sheet.getCurrentCell().setValue('Item 4');
    sheet.getRange('H2').activate();
    sheet.getCurrentCell().setValue('Item 5');
    sheet.getRange('I2').activate();
    sheet.getCurrentCell().setValue('Item 6');
    sheet.getRange('J2').activate();
    sheet.getCurrentCell().setValue('Item 7');
    sheet.getRange('K2').activate();
    sheet.getCurrentCell().setValue('Item 8');
    sheet.getRange('L2').activate();
    sheet.getCurrentCell().setValue('Item 9');
    sheet.getRange('M2').activate();
    sheet.getCurrentCell().setValue('Item 10');
    sheet.getRange('D3').activate();
    sheet.getCurrentCell().setValue('Weights');
    sheet.getRange('D4').activate();
    sheet.getCurrentCell().setValue('Weight 1');
    sheet.getRange('E4').activate();
    sheet.getCurrentCell().setValue('Weight 2');
    sheet.getRange('F4').activate();
    sheet.getCurrentCell().setValue('Weight 3');
    sheet.getRange('G4').activate();
    sheet.getCurrentCell().setValue('Weight 4');
    sheet.getRange('H4').activate();
    sheet.getCurrentCell().setValue('Weight 5');
    sheet.getRange('I4').activate();
    sheet.getCurrentCell().setValue('Weight 6');
    sheet.getRange('J4').activate();
    sheet.getCurrentCell().setValue('Weight 7');
    sheet.getRange('K4').activate();
    sheet.getCurrentCell().setValue('Weight 8');
    sheet.getRange('L4').activate();
    sheet.getCurrentCell().setValue('Weight 9');
    sheet.getRange('M4').activate();
    sheet.getCurrentCell().setValue('Weight 10');
    sheet.getRange('D1:M4').activate();
    sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange('D1:M1').activate().mergeAcross();
    sheet.getRange('D3:M3').activate().mergeAcross();
    sheet.getRange('D1:M2').activate();
    sheet.getActiveRangeList().setBackground('#d9d2e9');
    sheet.getRange('D3:M4').activate();
    sheet.getActiveRangeList().setBackground('#c9daf8');
    sheet.getRange('D5').activate();
    sheet.getCurrentCell().setValue('Max mark');
    sheet.getRange('D6').activate();
    sheet.getCurrentCell().setValue(10);
    sheet.getRange('D5:D6').activate();
    sheet.getActiveRangeList().setBackground('#b6d7a8').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  }
function get_cell_column_letter(cell){
  return cell.getA1Notation().match(/([A-Z]+)/)[0];
}
function get_active_cell_column_letter(sheet){
  return sheet.getActiveRange().getA1Notation().match(/([A-Z]+)/)[0];
}
function get_active_cell_row_number(sheet){
  return sheet.getActiveRange().getA1Notation().match(/([0-9]+)/)[0];
}
function go_down_one_cell(sheet) {
  sheet.getRange(sheet.getActiveRange().getRow()+1,sheet.getActiveRange().getColumn()).activate();
}
function go_down_and_left_one_row(sheet){
    sheet.getRange(sheet.getActiveRange().getRow()+1,1).activate();
}
function go_right_one_cell(sheet) {
  sheet.getRange(sheet.getActiveRange().getRow(),sheet.getActiveRange().getColumn()+1).activate();
}
function fill_down(ss,first_column,first_row,last_column,last_row){
  ss.getRange(first_column + first_row + ":" + last_column + first_row).activate();
  ss.getActiveRange().autoFill(ss.getRange(first_column + first_row + ":" + last_column + last_row),SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

function find_first_cell_by_value(sheet,value){
  var tf = sheet.createTextFinder(value);
  return tf.findAll()[0];
}
function nextChar(c) {
    var a = '';
    if ((c[c.length - 1] =='Z') && (c.split('').every(char => char === c[0]))) {
        for (i=0;i<c.length+1;i++) a += 'A';
        return a;
      }
    else {
            return c.slice(0,-1) + String.fromCharCode(c.charCodeAt(c.length-1) + 1);
      }
    }
function look_for_variable(sheet, start_cell, sheet_name, variable_name) {
  start_cell.activate();
  while ((sheet.getCurrentCell().getValue() == sheet_name) && (sheet.getRange(sheet.getCurrentCell().getRow(),sheet.getCurrentCell().getColumn()+1).getValue() != variable_name)) go_down_one_cell(sheet)
  if (sheet.getCurrentCell().getValue() == sheet_name) return sheet.getRange(sheet.getCurrentCell().getRow(),sheet.getCurrentCell().getColumn()+2).getValue();
  else return '';
}
