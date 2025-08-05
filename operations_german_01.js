function auto_exec(){
}

const codeVersion = '01.01';
async function showMain(){
  let waitPage = tag('start-wait');
  let mainPage = tag('main-page');
  waitPage.style.display = 'none';
  mainPage.style.display = 'block';
  await showMainPage();
}
async function showMainPage(){
  console.log('Showing Main Page')
  const mainPage = tag('main-page');
  mainPage.style.display = 'block';
  const versionInfo = tag('sheet-version');
  let versionString = 'Version ' + ' Code: ' + codeVersion + ' Released: ' ;
  versionInfo.innerText = versionString;
}

async function processGerman(){
  alert("Hello World")
}