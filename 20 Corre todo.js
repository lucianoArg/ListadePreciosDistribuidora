function correPrimeraParte(){

  var ui = SpreadsheetApp.getUi(); 

  var result = ui.alert("Voy a preparar archivos y copiar precios al MASTER",'1) MUEVO EL ARCHIVO "lista de precios actual" A LA CARPETA "ARCHIVO DE SHEETS" y RENOMBRO EL ARCHIVO CON LA FECHA DE CREACIÓN.\r\n2)CONVIERTO EL XLS EN GOOGLE SHEETS NAME "lista de precios actual"\r\n3)MUEVO EL XLS A CARPETA "ARCHIVO DE EXCELS"\r\n4)BORRO LOS PRECIOS VIEJOS Y COPIO LOS NUEVOS EN MASTER',ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {
    SpreadsheetApp.getActive().toast("Arrancando","Primera Parte",-1);
    Preparaarchivos();
    SpreadsheetApp.getActive().toast("1 de 3 listo","Primera Parte",-1);
    BorraPreciosViejos();
    SpreadsheetApp.getActive().toast("2 de 3 listo","Primera Parte",-1);
    CopiodataaMAster();
    SpreadsheetApp.getActive().toast("3 de 3 listo","Primera Parte");
    var result = ui.alert('Terminado paso 1\r\nRevisar errores en el MASTER antes de correr el paso 2 (que publica los cambios en la lista)');
   } else {
    ui.alert('Cancelado');
    return;
  }

}

function correSegundaParte(){


  var ui = SpreadsheetApp.getUi(); 

//CAMBIAR MENSAJE EXPLICATIVO!
var result = ui.alert("Voy a copiar info al PUBLIC",'1) en el MASTER hago copia de las hojas con precios y en esas copias borro fórmulas y chequeos de errores\r\n2) COPIO las copias del MASTER al PÚBLICO. Y las borro del MASTER.\r\n3) Borro las hojas viejas del PÚBLICO y renombro las nuevas, sacandole el "Copia de" y el número 1 del nombre',ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {
    SpreadsheetApp.getActive().toast("Arrancando","Segunda Parte",-1);
    pasaraPublic1();
    SpreadsheetApp.getActive().toast("1 de 3 listo","Segunda Parte",-1);
    pasaraPublic2();
    SpreadsheetApp.getActive().toast("2 de 3 listo","Segunda Parte",-1);
    pasaraPublic3();
    SpreadsheetApp.getActive().toast("3 de 3 listo","Segunda Parte");
    var result = ui.alert('La lista de precios ya está actualizada en el archivo público "Lista de Precios - Distribuidora Sur (P)"\r\nAbrilo para revisarlo, si querés.');

   } else {
    ui.alert('Cancelado');
    return;
  }



}


/* PENDIENTES
A) Automatizar el range de captura del array, tomar todo el archivo de precios intacto

B) borrar códigos raros de la lista inicial (dentro del array DATOS?):
VER TABLA 1 EN EL MASTER.

D) hacer trim y value a los datos (dentro del array DATOS?)

E) Hacer un bucle para el segundo paso, reocrriendo todas las hojas con (un listado en array?), para borrar formulas, borrar chequedos de D:D, mover al PUBLICO, etc.

*/