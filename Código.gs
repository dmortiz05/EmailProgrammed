function myFunctionMail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Enviar"));
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("B3:E5"); // Aquí indicamos el rango de celdas de la tabla
  var data = dataRange.getValues();
  
  var htmlOutput = HtmlService.createHtmlOutputFromFile('CumpleAutomatico'); // CumpleAutomatico es el nombre del archivo HTML
  var message = htmlOutput.getContent()
  
  // obtenemos la fecha de hoy y le damos formato
  var hoy = new Date(); 
  var mes = Utilities.formatDate(hoy,Session.getTimeZone(), "MM");
  var dia = Utilities.formatDate(hoy,Session.getTimeZone(), "dd");
  var fecha = dia+"/"+mes;
  
  for (i in data) {
    var rowData = data[i];
    var cumpleanios = rowData[3]; 
    if (cumpleanios == fecha){ // comparamos la fecha de hoy, con la fecha indicada en la tabla
      var emailAddress = rowData[2];
      var nombre = rowData[0];
      var apellido = rowData[1];
      var mensaje = message;
      var asunto = 'FELIZ CUMPLEAÑOS ESTIMADO/A' + ' ' + nombre + ' ' + apellido;
      MailApp.sendEmail(emailAddress, asunto, mensaje, {htmlBody : message});
    }
  }
}

function onOpen(){
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Send Emails', functionName: 'myFunctionMail'}
  ];
  spreadsheet.addMenu('Enviar Emails', menuItems);
}