//This is loaded manually into the sheet
function auto_exec(){
  Jade.load_js("https://davidfuller.github.io/operations.js", "operations");
  Jade.open_automations();
}

async function lockColumns(excel){
  /*Jade.listing:{"name":"Lock columns","description":"This locks columns"}*/
  await jade_modules.operations.lockColumns(excel);
}

async function unlock(excel){
  /*Jade.listing:{"name":"Unprotect sheet","description":"This unlocks sheet"}*/
  await jade_modules.operations.unlock(excel);
}

async function applyFilter(excel){
  /*Jade.listing:{"name":"Apply filter","description":"Applies empty filter to sheet"}*/
  await jade_modules.operations.applyFilter(excel);
}

async function removeFilter(excel){
  /*Jade.listing:{"name": "Remove filter","description":"Removes filter from sheet"}*/
  await jade_modules.operations.removeFilter(excel);
}

async function findNextScene(excel){
  /*Jade.listing:{"name": "Find Next Scene","description":"Finds the start of the next scene"}*/
  await jade_modules.operations.findScene(excel, 1);
}

async function findPreviousScene(excel){
  /*Jade.listing:{"name": "Find Previous Scene","description":"Finds the start of the previous scene"}*/
  await jade_modules.operations.findScene(excel, -1);
}