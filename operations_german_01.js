function auto_exec(){
}

const codeVersion = '01.01';
const germanProcessingSheetName = 'German Processing'
const openSpeechChar = '»';
const closeSpeechChar = '«';
const eolChar = '|eol|'
const bannedOpeningChars = [',', '.'];

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
    const procSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let originalTextRange = procSheet.getRange('gpOriginal');
    await excel.sync();
    originalTextRange.load('values');
    await excel.sync();
    let germanText = trimEmptyEnd(originalTextRange.values.map(x => x[0]));
    let results = []
    let totalDirectCopy = 0;
    let totalGood = 0;
    let totalWrong = 0;
    let totalUnequal = 0;
    for(let i = 0; i < germanText.length; i++){
      let result = {};
      let myStrings = []
      let original = []
      let startQuotes = locations(openSpeechChar, germanText[i]);
      let eols = locations(eolChar, germanText[i]);
      let endQuotes = locations(closeSpeechChar, germanText[i])
      let directCopy;
      let goodSpeech = 0;
      let wrongSpeech = 0;
      let unequalQuotes = 0;
      if (startQuotes.length == endQuotes.length){
        let myIndex = 0;
        if (startQuotes.length == 0){
          if (eols.length == 0){
            directCopy = true
            myStrings[0] = germanText[i].trim();
            original[0] = germanText[i].trim();
          } else {
            directCopy = false;
            for (let eol = 0; eol < eols.length; eol++){
              if (eol == 0){
                myStrings[myIndex] = germanText[i].substring(0, eols[eol]).trim();
                myIndex += 1;
                original[0] = germanText[i].trim()
              } 
              if (eol == eols.length - 1){
                // last part
                myStrings[myIndex] = germanText[i].substr(eols[eol] + 5).trim();
                myIndex += 1;   
              } else {
                // Between two eols
                myStrings[myIndex] = germanText[i].substring(eols[eol] + 5, eols[eol + 1]).trim();
                myIndex += 1;
              }
            }
          }
        } else {
          directCopy = false
          for (let speechPart = 0; speechPart < startQuotes.length; speechPart++ ){
            if (endQuotes[speechPart] > startQuotes[speechPart]){
              goodSpeech += 1;
              if (speechPart == 0){
                if (startQuotes[speechPart] > 0){
                  myStrings[myIndex] = germanText[i].substring(0, startQuotes[speechPart]).trim();
                  myIndex += 1;
                }
                original[0] = germanText[i].trim();
              }
              myStrings[myIndex] = germanText[i].substring(startQuotes[speechPart] + 1 , endQuotes[speechPart]).trim();
              myIndex += 1;
              if (speechPart == (startQuotes.length - 1)){
                if (germanText[i].substring(endQuotes[speechPart]).trim().length > 1){
                  myStrings[myIndex] = removedBannedOpeningCharacters(germanText[i].substring(endQuotes[speechPart] + 1).trim());
                }
              } else {
                //The bit between the close quotes and the next open quotes
                myStrings[myIndex] = removedBannedOpeningCharacters(germanText[i].substring(endQuotes[speechPart] + 1, startQuotes[speechPart+1]).trim());
                myIndex += 1;
              }
           } else {
              wrongSpeech += 1;
              myStrings[myIndex] = germanText[i].trim();
              original[myIndex] = germanText[i].trim();
              myIndex += 1;
            }
          }
        }
      } else {
        directCopy = false;
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
      result.lines = myStrings;
      result.original = original;
      results.push(result)
      console.log(i, ' - ', startQuotes, ',', endQuotes, ":", result )
    }
    let resultLines = createLines(results);
    await fillRange(germanProcessingSheetName, 'gpOriginal_Spaced', resultLines.original, true);
    await fillRange(germanProcessingSheetName, 'gpProcessed', resultLines.processed, true);
    console.log('Results')
    console.log('Total Good', totalGood, 'Total Wrong', totalWrong, 'Total Unequal', totalUnequal, 'Total Direct Copy', totalDirectCopy)
  })
}

function locations(substring,string){
  var a=[],i=-1;
  while((i=string.indexOf(substring,i+1)) >= 0) a.push(i);
  return a;
}

function createLines(results){
  originalLines = [];
  processedLines = [];
  index = 0;
  for (let i = 0; i < results.length; i++){
    for (let line = 0; line < results[i].lines.length; line++){
      if (line == 0){
        originalLines[index] = results[i].original[0];
      } else {
        originalLines[index] = "";
      }
      processedLines[index] = results[i].lines[line]
      index += 1;
    }
  }
  console.log('Original', originalLines);
  console.log('Processed', processedLines);
  return {original: originalLines, processed: processedLines}
}

async function fillRange(sheetName, rangeName, dataArray, doClear){
 await Excel.run(async function(excel){
  const mySheet = excel.workbook.worksheets.getItem(sheetName);
  const myRange = mySheet.getRange(rangeName);
  myRange.load("rowIndex, columnIndex");
  if (doClear){
    myRange.clear("Contents")
  }
  await excel.sync();

  const destRange = mySheet.getRangeByIndexes(myRange.rowIndex, myRange.columnIndex, dataArray.length, 1)
  destRange.load('address');
  await excel.sync();
  console.log('address:', destRange.address);
  let temp = []
  for (let i = 0; i < dataArray.length; i++){
    temp[i] = [];
    temp[i][0] = dataArray[i]; 
  }
  console.log(temp)
  destRange.values = temp;
  await excel.sync();
 }) 
}

function trimEmptyEnd(dataArray){
  for(let i = dataArray.length - 1; i >= 0; i--){
    if (dataArray[i] != ''){
      return dataArray.slice(0, i + 1);
    }
  }
  return dataArray;
}

function removedBannedOpeningCharacters(text){
  let doAgain = true
  let temp = text.trim();
  for (let attempt = 0; attempt < 100; attempt++){
    if (!doAgain){break;}
    doAgain = false;
    for (let i = 0; i < bannedOpeningChars.length; i++){
      if (temp.substr(0,1) == bannedOpeningChars[i]){
        temp = temp.substring(1).trim();
        doAgain = true;
      }
    }
  }
  return temp;
}

async function doAReplacement(rowIndex, searchText, replaceText){
  // Gets the relevant cell/range (original column)
  // replaces the text
  // saves it
  console.log('rowIndex, searchText, replaceText', rowIndex, searchText, replaceText)
  await Excel.run(async function(excel){
    const procSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let originalTextRange = procSheet.getRange('gpOriginal');
    originalTextRange.load('columnIndex');
    await excel.sync();
    let targetRange = procSheet.getRangeByIndexes(rowIndex, originalTextRange.columnIndex, 1, 1);
    targetRange.load('values');
    await excel.sync();
    let targetText = targetRange.values[0][0];
    console.log('Before', targetText);
    targetText = targetText.replace(searchText, replaceText);
    console.log('After', targetText)
    targetRange.values = [[targetText]];
    await excel.sync();
  })
  
}