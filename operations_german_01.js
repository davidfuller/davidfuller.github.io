function auto_exec(){
}

const codeVersion = '01.01';
const germanProcessingSheetName = 'German Processing'
const openSpeechChar = '»';
const closeSpeechChar = '«'

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
    let germanText = originalTextRange.values.map(x => x[0]);
    console.log(germanText);
    for(i = 0; i < germanText.length; i++){
      let startQuotes = locations(openSpeechChar, germanText[i])
      let endQuotes = locations(closeSpeechChar, germanText[i])
      console.log(i, ' - ', startQuotes, ',', endQuotes)
    }
  })
}

function locations(substring,string){
  var a=[],i=-1;
  while((i=string.indexOf(substring,i+1)) >= 0) a.push(i);
  return a;
}