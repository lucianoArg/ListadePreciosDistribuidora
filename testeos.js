function HOY(){

  // var ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var hoyEnA1 = "Actualizado el " + Utilities.formatDate(new Date(), "Etc/GMT", "dd/MM/yyyy");
  console.log(hoyenA1);

}


function ultimafila(){

  var ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
  var Avals = ws.getRange("A1:A").getValues();
  var ultimafila = Avals.filter(String).length; 
  let fin = ws.getMaxRows() - ultimafila;

//Logger.log(ultimafila+1);
//Logger.log(fin);
//  var formulas = ws.getRange(5,4,1,4).getFormulasR1C1();
//  ws.getRange(6,4,1,4).setFontColor("red").setFormulaR1C1(formulas);

//  ws.getRange(6,4,ws.getLastRow()-5,4).setBackground("yellow");
  if (fin >0){ws.deleteRows(ultimafila+1,fin);}
  
}

function titulos(){

  var ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
/*  ws.setRowHeight(1,18);
  ws.setRowHeight(2,10);
  ws.setRowHeight(3,60);
  ws.setRowHeight(4,5);*/

  const colCodigoPrecios = 130;
  const colDesc = 500;
  ws.setColumnWidth(1,colCodigoPrecios);
  ws.setColumnWidth(2,colDesc);
  ws.setColumnWidth(3,colCodigoPrecios);



}


function formato(){

  
  var cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
  var ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var currentRow = SpreadsheetApp.getActiveSheet().getActiveCell().getColumn();
  var fila = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  var columna = SpreadsheetApp.getActiveSheet().getActiveCell().getColumn();

//  if (ws.getActiveCell().offset(1,1).getvalue === ""){
//    cell.setValue("borro la fila")
//  ws.deleteRow(fila);
//  }
//  cell.offset(2,2).setValue(currentRow);
//  ws.getRange("D4").setValue(fila);
  ws.setRowHeight(fila+1,5);
  ws.setRowHeight(fila+2,100);
  ws.setRowHeight(fila+3,5);
//  ws.deleteRow(fila);
  ws.getActiveCell().offset(2,2).activateAsCurrentCell();

}

function test3(){

  var spreadsheet = SpreadsheetApp.openById("1xm6ECeEnd2J-hRoEIgQK1jntjZ02p7F2IBoXOJMu6dY"); //EL MASTER

  var allSheets = spreadsheet.getSheets();
  var colCodigoPrecios = 130;
  var colDesc = 500;
  for (var i = 2; i < allSheets.length-1 ; i++ ) {
  Logger.log("Sheet " + i + " se llama " + allSheets[i].getName());
//  Logger.log(allSheets);
  }

}


function ultimacolumna(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ss.getRange(1,1).setValue(1);
//  var lastcol = ss.getLastColumn()
//  Logger.log(ss.getName)
//  ss.getRange(1,7).activate();
//  ss.getRange(1,7).activateAsCurrentCell();
  ss.deleteColumns(4,1); // borro columna D
}

function verprimerahoja(){

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheets()[0]);
}

function final(){

  var PUBLICO = SpreadsheetApp.openById('1zrBU5DIEkCv20ojlkc-cWgQ-6slkRSw3VRuWqJDMihg'); //PUBLIC
  var allSheets = PUBLICO.getSheets();

  //(ya no) HAY QUILOMBO ACÁ, NO ENCUENTRA UNA HOJA EN EL LOG (236) Y SE FRENA. 
  var primerahoja = 2
  var ultimahoja = 9
  Logger.log(ultimahoja);

  for (var i = primerahoja; i <= ultimahoja; i++ ){
    var hojaSINfunciones = PUBLICO.getSheets()[i];
    var nombreNuevo = hojaSINfunciones.getRange("B3").getValue();
    Logger.log("la Hoja " + i + ". Se llama <" + allSheets[i].getName() + "> y pasará a: -" + nombreNuevo + "-");
    Logger.log(hojaSINfunciones.getName());
    hojaSINfunciones.setName(nombreNuevo);
    Logger.log(hojaSINfunciones.getName());
  }
}

function showMessageBox() {
Browser.msgBox('You clicked it!');
SpreadsheetApp.getUi().prompt("Mensaje");
}

function NOonOpen() {
  var ui = SpreadsheetApp.getUi();
}

function showAlert() {
  var ui = SpreadsheetApp.getUi(); 

  var result = ui.alert("titulo",'pregunta',ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {
    ui.alert('puso OK');
  } else {
    return;
  }

  ui.alert('sigue la macro');

}



function showMessage(){

  SpreadsheetApp.getUi().prompt("Mensaje");
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  ss.Browser.msgBox("mandé el file 'Lista de precios actual' a la carpeta 'ArchivodeSheets' y otras cosas"); // EXPLICAR QUÉ MÁS HAGO Y SEGUIR EL TESTING!

}



 function logerHojas(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

var allSheets = spreadsheet.getSheets();
  for (i in allSheets) {
  Logger.log("Sheet " + i + " is named " + allSheets[i].getName());
}
  Logger.log("Después:");

//TODOS ORDENADOS 
  var hojaSistemasDeEscapesFun = spreadsheet.getSheetByName("TODOS ORDENADOS.");
  hojaSistemasDeEscapesFun.copyTo(spreadsheet).setName('TODOS ORDENADOS');
  var hojaSistemasDeEscapes = spreadsheet.getSheetByName("TODOS ORDENADOS");
  hojaSistemasDeEscapes.activate();
  spreadsheet.moveActiveSheet(3);
  hojaSistemasDeEscapes.getRange('A:H').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  hojaSistemasDeEscapes.getRange('D:F').activate();
  hojaSistemasDeEscapes.getActiveRange().clear();

  var allSheets = spreadsheet.getSheets();
  for (i in allSheets) {
  Logger.log("Sheet " + i + " is named " + allSheets[i].getName());
}
 }

function test(){

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaSistemasDeEscapes = spreadsheet.getSheetByName("SISTEMAS DE ESCAPES");
  hojaSistemasDeEscapes.activate();
  spreadsheet.toast('test');
  hojaSistemasDeEscapes.getRange('D:D').activate();
  hojaSistemasDeEscapes.getActiveRange().clear();

}