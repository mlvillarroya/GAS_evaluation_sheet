function look_for_name(name){
	var rows =
	document.getElementsByTagName('tbody')[0].querySelectorAll('tr[class^="user"]');
	for (i=0;i<rows.length;i++) {
	  var cells = rows[i].getElementsByTagName('td');
	  for (j=0;j<cells.length;j++) {
		if (cells[j].innerText == name) return cells[j].parentNode;
	  }
	}
}

function look_for_name_code(){
  var code = "function look_for_name(name){\nvar rows =	document.getElementsByTagName('tbody')[0].querySelectorAll('tr[class^=\"user\"]');\n for (i=0;i<rows.length;i++) {\n var cells = rows[i].getElementsByTagName('td');\n	  for (j=0;j<cells.length;j++) {\n		if (cells[j].innerText == name) return cells[j].parentNode;\n	  }\n	}\n}";

  return code
}

function generate_script(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var mark_column = data[0].indexOf(constants().MARK_COLUMN_TITLE);
  var comments_column = data[0].indexOf(constants().COMMENT_COLUMN_TITLE);
  var done_column = data[0].indexOf(constants().DONE_COLUMN_TITLE);
  data.shift();
  data.pop();
  var ui = SpreadsheetApp.getUi();
  
  // Mostrar un cuadro de diálogo de entrada numérica
  var response1 = ui.alert(
    'Grade export',
    'Your grades will be exported to a JavaScript function so that you can paste them into your Moodle grade page using the JavaScript console. Please ensure that the name column in your grades matches the name column in the Moodle grade page.',
    ui.ButtonSet.OK
  );
  if (response1 != ui.Button.OK) return;
  
  var file_content = '';
  file_content += look_for_name_code() + "\nvar moodle_row = ''\n";
  data.forEach((row)=>{
    if (row[mark_column] && row[done_column]=='Yes')
    {
    file_content += "moodle_row = look_for_name(\"" + row[0] + "\");\n"
    file_content += "moodle_row.getElementsByClassName('quickgrade')[0].value=" + row[mark_column] + ";\n"; 
    file_content += "moodle_row.getElementsByClassName('quickgrade')[1].value=\"" + row[comments_column].replace(/\n/g, "") + "\";\n"; 
    }
  });
  Logger.log(file_content);
  // Paso 2: Crear un archivo con el contenido en Google Drive
  var file_name = "moodle_script.js"; // Nombre del archivo con extensión ".js"
  var folder = DriveApp.getRootFolder(); // Carpeta raíz de Google Drive
  var file = DriveApp.createFile(file_name, file_content, MimeType.PLAIN_TEXT);

  // Paso 3: Obtener la URL de descarga del archivo
  var url_download = file.getUrl();

  // Paso 4: Redirigir al usuario a la URL de descarga del archivo
  var response4 = ui.alert(
    'File created',
    'Your file has been created: ' + url_download + "\nPlease, download the file before pressing OK, it will be deleted inmediatelly after",
    ui.ButtonSet.OK
  );
  file.setTrashed(true);
}
