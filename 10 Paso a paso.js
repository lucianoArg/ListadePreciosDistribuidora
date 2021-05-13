//PARTE 1

function Preparaarchivos(){

  //1 MUEVE EL ARCHIVO "lista de precios actual" A LA CARPETA "ARCHIVO DE SHEETS" y RENOMBRA EL ARCHIVO CON LA FECHA DE CREACIÓN.
  var dApp = DriveApp;
  var folderIter = dApp.getFoldersByName("Precios actuales");
  var folder = folderIter.next();
  var fileIter = folder.getFilesByName("lista de precios actual");
  var file = fileIter.next();
  const carpetaArchivodeSheetsID = "1EeYn4xLw8yYPWyUCD4aiZ6OqYc9IdcIi";
   var carpetaArchivodeSheets = DriveApp.getFolderById(carpetaArchivodeSheetsID);
  file.moveTo(carpetaArchivodeSheets);
  var fecha = file.getDateCreated();
  var fechaCorta = Utilities.formatDate(fecha, 'Etc/GMT', 'yyyyMMdd');
  file.setName("lista de precios "+fechaCorta);

  //2 CONVIERTE EL XLS DE LA CARPETA EN GOOGLE SHEETS NAME "lista de precios actual"
  var fileIter2 = folder.getFilesByType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  var fileXLS = fileIter2.next();
  var mimeXLS = fileXLS.getMimeType();
  var resource = {
   title : "lista de precios actual",
   mimeType : MimeType.GOOGLE_SHEETS,
   parents: [{id : folder.getId()}],
                }

  Drive.Files.insert(resource, fileXLS.getBlob());

  //3 MANDA EL XLS A CARPETA "ARCHIVO DE EXCELS".
  var carpetaArchivodeExcels = DriveApp.getFolderById("1UMgZQiL9t5_flk-Z9md6ecGT-hF3JJTz");
  fileXLS.moveTo(carpetaArchivodeExcels);

}


function BorraPreciosViejos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TODOS ORDENADOS.");
  ss.getRange("A5:C").clear();
}


function CopiodataaMAster(){
  var dApp = DriveApp;
  var folderIter = dApp.getFoldersByName("Precios actuales");
  var folder = folderIter.next();
  var fileIter = folder.getFilesByName("lista de precios actual");
  var file = fileIter.next();
  var folderId = folder.getId();
  var listadepreciosactualID = file.getId();

  // CAPTURO DATA (abro archivo y copio data)
  var listadepreciosactual  = SpreadsheetApp.openById(listadepreciosactualID);
  var hojaorigendedatos = listadepreciosactual.getSheets()[0];

  //(antes de capturar, hago TRIM y convierto códigos numéricos a valor (T) )
  //hojaorigendedatos.getRange('E10').setFormula('=TEXT(TRIM(A10);"#")'); 
  // si solo hago TRIM, convierte / a fechas. 
  //si hago TEXT, rompe los códigos numéricos (1004 a -109571 !?).
  // si hago T, convierte las / a fechas
  //IF "/" usar TEXT else T
  //LO resolví poniendo en formato plain text en el google sheets la columna A de Todos Ordenados.
  hojaorigendedatos.getRange('E10').setFormula('=T(TRIM(A10))')
  hojaorigendedatos.getRange('F10').setFormula('=TRIM(B10)');
  hojaorigendedatos.getRange('G10').setFormula('=C10');
  var rangetocopy = hojaorigendedatos.getRange(10,5,1,3);
   var pegoEn = hojaorigendedatos.getRange(11,5,hojaorigendedatos.getLastRow()+1,3);
  rangetocopy.copyTo(pegoEn);


  var data = hojaorigendedatos.getRange("E10:G" + hojaorigendedatos.getLastRow()).getValues(); // OJO, HACER E10 AUTOMATIZADO! DEBERÍA TOMAR DE A:C, BORRAR LOS CODIGOS RAROS Y HACER TRIM Y VALUE TODO ADENTRO DEL ARRAY, SIN TOCAR EL ARCHIVO.

  // NUEVO ARRAY ORDENADO
  var dataordenado = data.sort(function(Fila1,Fila2){
    var a = Fila1[0].toString().toLowerCase();
    var b = Fila2[0].toString().toLowerCase();


  if (a > b){
    return 1;
  } else if (a<b){
    return -1;
  }
    return 0;
  });
  //  Logger.log(dataordenado)
  
  //PEGO en Master
  //  var formatos = [[ "@", "@", "'$'\ #,##0.00" ]];
  //  var formatos = [[ "@", "@", "null" ]];

  var ArchivoMaster = SpreadsheetApp.openById("1xm6ECeEnd2J-hRoEIgQK1jntjZ02p7F2IBoXOJMu6dY"); //CAMBIAR ESTO POR EL MASTER
  var archivoDestino = ArchivoMaster.getSheetByName("TODOS ORDENADOS.");
  archivoDestino.getRange(5, 1, data.length, 3).setValues(dataordenado).setBorder(true,true,true,true,true,true,null,SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontSize(10).setFontFamily("Arial").setHorizontalAlignment("center");//.setNumberFormats(formatos);

/*/Borro filas sobrantes abajo
  var Avals = archivoDestino.getRange("A1:A").getValues(); //array con valores en A:A
  var ultimafila = Avals.filter(String).length; //me da la ultima fila escrita de A (cuenta celdas con string)
  let fin = archivoDestino.getMaxRows() - ultimafila; //calculo cuantas filas borrar a partir de la primera fila en blanco
  if (fin > 0){archivoDestino.deleteRows(ultimafila+1,fin);} //borro solo si hay filas para borrar
*/

}


//PARTE 2

function pasaraPublic1(){
//1 en el MASTER hago copia de las hojas con precios y pinto todo y, (en las copias) pego como valor y borro contenido y formatos de D:D (E:E en ENGACHES, D:F en TODOS ORDENADOS) (chequeos invisibles). 
  var spreadsheet = SpreadsheetApp.openById("1xm6ECeEnd2J-hRoEIgQK1jntjZ02p7F2IBoXOJMu6dY"); //CAMBIAR ESTO POR EL MASTER

  var allSheets = spreadsheet.getSheets();
  const colCodigoPrecios = 130;
  const colDesc = 500;
  var hoyEnA1 = "Actualizado el " + Utilities.formatDate(new Date(), "Etc/GMT", "dd/MM/yyyy");
  //for (i in allSheets) {
  //Logger.log("Sheet " + i + " se llamaba " + allSheets[i].getName());
  //}
  //Logger.log("TRABAJANDO...")
  //for (var i = 2; i < allSheets.length-1 ; i++ ) {
  //Logger.log("Sheet " + i + " se llama " + allSheets[i].getName());}

 //SISTEMAS DE ESCAPE
  var hojaconfunciones = spreadsheet.getSheetByName("SISTEMAS DE ESCAPES.");
  hojaconfunciones.copyTo(spreadsheet).setName('SISTEMAS DE ESCAPES1');
  var hojaSINfunciones = spreadsheet.getSheetByName("SISTEMAS DE ESCAPES1");
  hojaSINfunciones.activate();
  spreadsheet.moveActiveSheet(3);
  hojaSINfunciones.getRange('A:C').copyTo(spreadsheet.getRange('A:H'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  //hojaSINfunciones.getRange('D:D').clear();
  hojaSINfunciones.deleteColumns(4,1); // borro columna D
  hojaSINfunciones.getRange('B3').activateAsCurrentCell
  hojaSINfunciones.setColumnWidth(1,colCodigoPrecios);
  hojaSINfunciones.setColumnWidth(2,colDesc);
  hojaSINfunciones.setColumnWidth(3,colCodigoPrecios);
  hojaSINfunciones.getRange('A1').setValue(hoyEnA1);

  //ACCESORIOS DE ESCAPE
  var hojaconfunciones = spreadsheet.getSheetByName("ACCESORIOS DE ESCAPE.");
  hojaconfunciones.copyTo(spreadsheet).setName('ACCESORIOS DE ESCAPE1');
  var hojaSINfunciones = spreadsheet.getSheetByName("ACCESORIOS DE ESCAPE1");
  hojaSINfunciones.activate();
  spreadsheet.moveActiveSheet(4);
  hojaSINfunciones.getRange('A:C').copyTo(spreadsheet.getRange('A:H'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  //hojaSINfunciones.getRange('D:D').clear();
  hojaSINfunciones.deleteColumns(4,1); // borro columna D
  hojaSINfunciones.getRange('B3').activateAsCurrentCell
  hojaSINfunciones.setColumnWidth(1,colCodigoPrecios);
  hojaSINfunciones.setColumnWidth(2,colDesc);
  hojaSINfunciones.setColumnWidth(3,colCodigoPrecios);
  hojaSINfunciones.getRange('A1').setValue(hoyEnA1);

  //ESCAPESSILENS
  var hojaconfunciones = spreadsheet.getSheetByName("ESCAPESSILENS.");
  hojaconfunciones.copyTo(spreadsheet).setName('ESCAPESSILENS1');
  var hojaSINfunciones = spreadsheet.getSheetByName("ESCAPESSILENS1");
  hojaSINfunciones.activate();
  spreadsheet.moveActiveSheet(5);
  hojaSINfunciones.getRange('A:C').copyTo(spreadsheet.getRange('A:H'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  //hojaSINfunciones.getRange('D:D').clear();
  hojaSINfunciones.deleteColumns(4,1); // borro columna D
  hojaSINfunciones.getRange('B3').activateAsCurrentCell
  hojaSINfunciones.setColumnWidth(1,colCodigoPrecios);
  hojaSINfunciones.setColumnWidth(2,colDesc);
  hojaSINfunciones.setColumnWidth(3,colCodigoPrecios);
  hojaSINfunciones.getRange('A1').setValue(hoyEnA1);

  //DEPORTIVOS
  var hojaconfunciones = spreadsheet.getSheetByName("DEPORTIVOS.");
  hojaconfunciones.copyTo(spreadsheet).setName('DEPORTIVOS1');
  var hojaSINfunciones = spreadsheet.getSheetByName("DEPORTIVOS1");
  hojaSINfunciones.activate();
  spreadsheet.moveActiveSheet(6);
  hojaSINfunciones.getRange('A:C').copyTo(spreadsheet.getRange('A:H'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  //hojaSINfunciones.getRange('D:D').clear();
  hojaSINfunciones.deleteColumns(4,1); // borro columna D
  hojaSINfunciones.getRange('B3').activateAsCurrentCell
  hojaSINfunciones.setColumnWidth(1,colCodigoPrecios);
  hojaSINfunciones.setColumnWidth(2,colDesc);
  hojaSINfunciones.setColumnWidth(3,colCodigoPrecios);
  hojaSINfunciones.getRange('A1').setValue(hoyEnA1);

  //ENGANCHES
  var hojaconfunciones = spreadsheet.getSheetByName("ENGANCHES.");
  hojaconfunciones.copyTo(spreadsheet).setName('ENGANCHES1');
  var hojaSINfunciones = spreadsheet.getSheetByName("ENGANCHES1");
  hojaSINfunciones.activate();
  spreadsheet.moveActiveSheet(7);
  hojaSINfunciones.getRange('A:D').copyTo(spreadsheet.getRange('A:H'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  hojaSINfunciones.deleteColumns(5,1); // elimina columna E
  hojaSINfunciones.getRange('B3').activateAsCurrentCell
  hojaSINfunciones.setColumnWidth(1,colCodigoPrecios);
  hojaSINfunciones.setColumnWidth(2,colDesc);
  hojaSINfunciones.setColumnWidth(3,colCodigoPrecios);
  hojaSINfunciones.setColumnWidth(4,colCodigoPrecios);
  hojaSINfunciones.getRange('A1').setValue(hoyEnA1);

  //TANQUES 
  var hojaconfunciones = spreadsheet.getSheetByName("TANQUES.");
  hojaconfunciones.copyTo(spreadsheet).setName('TANQUES1');
  var hojaSINfunciones = spreadsheet.getSheetByName("TANQUES1");
  hojaSINfunciones.activate();
  spreadsheet.moveActiveSheet(8);
  hojaSINfunciones.getRange('A:C').copyTo(spreadsheet.getRange('A:H'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  //hojaSINfunciones.getRange('D:D').clear();
  hojaSINfunciones.deleteColumns(4,1); // borro columna D
  hojaSINfunciones.getRange('B4').activateAsCurrentCell
  hojaSINfunciones.setColumnWidth(1,colCodigoPrecios);
  hojaSINfunciones.setColumnWidth(2,colDesc);
  hojaSINfunciones.setColumnWidth(3,colCodigoPrecios);
  hojaSINfunciones.getRange('A1').setValue(hoyEnA1);

  //EQUIPAMIENTOS 
  var hojaconfunciones = spreadsheet.getSheetByName("EQUIPAMIENTOS.");
  hojaconfunciones.copyTo(spreadsheet).setName('EQUIPAMIENTOS1');
  var hojaSINfunciones = spreadsheet.getSheetByName("EQUIPAMIENTOS1");
  hojaSINfunciones.activate();
  spreadsheet.moveActiveSheet(9);
  hojaSINfunciones.getRange('A:C').copyTo(spreadsheet.getRange('A:H'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  //hojaSINfunciones.getRange('D:D').clear();
  hojaSINfunciones.deleteColumns(4,1); // borro columna D
  hojaSINfunciones.getRange('B3').activateAsCurrentCell
  hojaSINfunciones.setColumnWidth(1,colCodigoPrecios);
  hojaSINfunciones.setColumnWidth(2,colDesc);
  hojaSINfunciones.setColumnWidth(3,colCodigoPrecios);
  hojaSINfunciones.getRange('A1').setValue(hoyEnA1);

  //TODOS ORDENADOS 
  var hojaconfunciones = spreadsheet.getSheetByName("TODOS ORDENADOS.");
  hojaconfunciones.copyTo(spreadsheet).setName('TODOS ORDENADOS1');
  var hojaSINfunciones = spreadsheet.getSheetByName("TODOS ORDENADOS1");
  hojaSINfunciones.activate();
  spreadsheet.moveActiveSheet(10);
  hojaSINfunciones.getRange("D:D").copyTo(spreadsheet.getRange("D:D"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  hojaSINfunciones.getRange("D4").clear();
  hojaSINfunciones.deleteColumns(5,3); //ELIMINA E:G
  hojaSINfunciones.getRange('B4').activateAsCurrentCell
  hojaSINfunciones.setColumnWidth(1,colCodigoPrecios);
  hojaSINfunciones.setColumnWidth(2,colDesc);
  hojaSINfunciones.setColumnWidth(3,colCodigoPrecios);
  hojaSINfunciones.getRange('A1').setValue(hoyEnA1);

//  var allSheets = spreadsheet.getSheets();
//  for (i in allSheets) {
//  Logger.log("Sheet " + i + " se llama ahora " + allSheets[i].getName());
//  }

}

function pasaraPublic2(){
//2 COPIO las copias del MASTER al PÚBLICO que es siempe el mismo ID: "1zrBU5DIEkCv20ojlkc-cWgQ-6slkRSw3VRuWqJDMihg" . Y las borro del MASTER.

  var MASTER = SpreadsheetApp.openById("1xm6ECeEnd2J-hRoEIgQK1jntjZ02p7F2IBoXOJMu6dY"); //CAMBIAR ESTO POR EL MASTER
  var primerahoja = 2
  var ultimahoja = 9
  for (var i = primerahoja; i <= ultimahoja; i++ ){
  var hojaSINfunciones = MASTER.getSheets()[i];
  //  Logger.log("voy a copiar la hoja "+ i + " y borrar la hoja "+i);
  var PUBLICO = SpreadsheetApp.openById('1zrBU5DIEkCv20ojlkc-cWgQ-6slkRSw3VRuWqJDMihg'); //PUBLIC
  hojaSINfunciones.copyTo(PUBLICO);
  }
  for (var i = ultimahoja; i >= primerahoja; i-- ){
  var hojaSINfunciones = MASTER.getSheets()[i];
  MASTER.deleteSheet(hojaSINfunciones)
  }
}

function pasaraPublic3(){
  //3 Borra las hojas viejas del PÚBLICO y renombrar las nuevas, sacandole "copia de " y el número 1 del nombre

  var PUBLICO = SpreadsheetApp.openById('1zrBU5DIEkCv20ojlkc-cWgQ-6slkRSw3VRuWqJDMihg'); //PUBLIC
  var allSheets = PUBLICO.getSheets();

 
  var primerahoja = 2
  var ultimahoja = 9
  for (var i = ultimahoja; i >= primerahoja; i-- ){
    var hojaSINfunciones = PUBLICO.getSheets()[i];
    //    Logger.log("Borré la Hoja " + i + ". Se llamaba " + allSheets[i].getName());
    PUBLICO.deleteSheet(hojaSINfunciones)
  }

  //  Logger.log("TRABAJANDO...")

  var allSheets = PUBLICO.getSheets();
  var primerahoja = 2
  var ultimahoja = 9
  //  Logger.log(ultimahoja);

  for (var i = primerahoja; i <= ultimahoja; i++ ){
    var hojaSINfunciones = PUBLICO.getSheets()[i];
    var nombreNuevo = hojaSINfunciones.getRange("B3").getValue();
  //    Logger.log("la Hoja " + i + ". Se llama <" + allSheets[i].getName() + "> y pasará a: -" + nombreNuevo + "-");
  //    Logger.log(hojaSINfunciones.getName());
    hojaSINfunciones.setName(nombreNuevo);
//    Logger.log(hojaSINfunciones.getName());

  }

  var hoyEnA1 = "Actualizado el " + Utilities.formatDate(new Date(), "Etc/GMT", "dd/MM/yyyy");
  PUBLICO.getSheetByName("TAPA").getRange('H3').setValue(Utilities.formatDate(new Date(), "Etc/GMT", "dd/MM/yyyy"));
  PUBLICO.getSheetByName("AYUDA").getRange('B4').setValue(hoyEnA1);

}