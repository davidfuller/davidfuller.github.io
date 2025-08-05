function auto_exec(){
}

const codeVersion = '01.01';
const germanProcessingSheetName = 'German Processing'

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
  const admin = tag('admin');
  admin.style.display = 'block';
}

async function processGerman(){
  await Excel.run(async function(excel){
    procSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let originalTextRange = procSheet.getRange('gpOriginal');
    await excel.sync();
    originalTextRange.load('values');
    await excel.sync();
    console.log(originalTextRange.values);
  })
}