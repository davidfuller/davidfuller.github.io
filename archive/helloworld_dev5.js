async function auto_exec(){
  console.log("The very beginning");
  await Jade.load_js("https://davidfuller.github.io/operations_dev5.js", "operations");
  console.log('After operations');
  await Jade.load_js("https://davidfuller.github.io/scheduling_dev5.js", "scheduling");
  console.log('After scheduling');
  await Jade.load_js("https://davidfuller.github.io/walla_import_dev5.js", "wallaimport");
  console.log('After wallaimport');
  await Jade.load_js("https://davidfuller.github.io/html_dev5.js", "html");
  console.log('After html');
  await Jade.load_js("https://davidfuller.github.io/css_dev5.js", "css");
  console.log('After css');
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