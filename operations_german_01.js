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
    let results = []
    let totalDirectCopy = 0;
    let totalGood = 0;
    let totalWrong = 0;
    let totalUnequal = 0;
    for(let i = 0; i < germanText.length; i++){
      let result = {};
      let startQuotes = locations(openSpeechChar, germanText[i])
      let endQuotes = locations(closeSpeechChar, germanText[i])
      let directCopy;
      let goodSpeech = 0;
      let wrongSpeech = 0;
      let unequalQuotes = 0;
      if (startQuotes.length == endQuotes.length){
        if (startQuotes.length == 0){
          directCopy == true
        } else {
          directCopy == false
          for (let speechPart = 0; speechPart < startQuotes.length; speechPart++ ){
            if (endQuotes(speechPart) > startQuotes(speechPart)){
              goodSpeech += 1;
            } else {
              wrongSpeech += 1;
            }
          }
        }
      } else {
        directCopy = false
        unequalQuotes += 1;
      }
      result.directCopy = directCopy;
      if (directCopy){
        totalDirectCopy += 1;
      }
      result.goodSpeech = goodSpeech;
      totalGood = totalGood + goodSpeech
      result.wrongSpeech = wrongSpeech;
      totalWrong = totalWrong + wrongSpeech
      result.unequalQuotes = unequalQuotes
      totalUnequal = totalUnequal + unequalQuotes
      results.push(result)
      console.log(i, ' - ', startQuotes, ',', endQuotes, ":", result )
    }
    console.log('Results')
    console.log('Total Good', totalGood, 'Total Wrong', totalWrong, 'Total Unequal', totalUnequal)
  })
}


function locations(substring,string){
  var a=[],i=-1;
  while((i=string.indexOf(substring,i+1)) >= 0) a.push(i);
  return a;
}