async function auto_exec(){
  console.log("The very beginning");
  await Jade.load_js("https://davidfuller.github.io/operations_dev14.js", "operations");
  console.log('After operations');
  await Jade.load_js("https://davidfuller.github.io/scheduling_dev14.js", "scheduling");
  console.log('After scheduling');
  await Jade.load_js("https://davidfuller.github.io/walla_import_dev14.js", "wallaimport");
  console.log('After wallaimport');
  await Jade.load_js("https://davidfuller.github.io/comparison_dev14.js", "comparison");
  console.log('After comparison');
  await Jade.load_js("https://davidfuller.github.io/wordsToNumbers.js", "wordtonumbers");
  console.log('After wordsToNumbers');
  await Jade.load_js("https://davidfuller.github.io/usScript_dev14.js", "usscript");
  console.log('After usScript');
  await Jade.load_js("https://davidfuller.github.io/actorMultiple_dev14.js", "actormultiple");
  console.log('After actormultiple');
  await Jade.load_js("https://davidfuller.github.io/html_dev14.js", "html");
  console.log('After html');
  await Jade.load_js("https://davidfuller.github.io/css_dev14.js", "css");
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
