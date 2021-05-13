function initMenu(){


  var ui = SpreadsheetApp.getUi();

  var menu = ui.createMenu("Lista de Precios");
  menu.addItem("Primer paso","correPrimeraParte");
  menu.addSeparator();
  menu.addItem("Segundo paso","correSegundaParte");
  menu.addToUi();
}

function onOpen(){

initMenu();

}