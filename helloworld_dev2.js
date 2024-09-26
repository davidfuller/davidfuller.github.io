async function auto_exec(){
  await Jade.load_js("https://davidfuller.github.io/operations_dev2.js", "operations");
  await Jade.load_js("https://davidfuller.github.io/html_dev2.js", "html");
  await Jade.load_js("https://davidfuller.github.io/css_dev2.js", "css");
  await jade_modules.css.mainCSS();
  await jade_modules.html.mainHTML();
}

async function openUI(excel){
  /*Jade.listing:{"name":"Open UI","description":"Opens the user interface"}*/
  await jade_modules.css.mainCSS();
  await jade_modules.html.mainHTML();
}

async function lockColumns(excel){
  /*Jade.listing:{"name":"Lock columns","description":"This locks columns"}*/
  await jade_modules.operations.lockColumns();
}

async function unlock(excel){
  /*Jade.listing:{"name":"Unprotect sheet","description":"This unlocks sheet"}*/
  await jade_modules.operations.unlock();
}

async function applyFilter(excel){
  /*Jade.listing:{"name":"Apply filter","description":"Applies empty filter to sheet"}*/
  await jade_modules.operations.applyFilter();
}

async function removeFilter(excel){
  /*Jade.listing:{"name": "Remove filter","description":"Removes filter from sheet"}*/
  await jade_modules.operations.removeFilter();
}

async function findNextScene(excel){
  /*Jade.listing:{"name": "Find Next Scene","description":"Finds the start of the next scene"}*/
  await jade_modules.operations.findScene(1);
}

async function findPreviousScene(excel){
  /*Jade.listing:{"name": "Find Previous Scene","description":"Finds the start of the previous scene"}*/
  await jade_modules.operations.findScene(-1);
}

function test(){
  alert('test');
}
