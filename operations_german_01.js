function auto_exec() {}

const codeVersion = '01.01';
const germanProcessingSheetName = 'German Processing'
const openSpeechChar = '»';
const closeSpeechChar = '«';
const eolChar = '|eol|'
const bannedOpeningChars = [',', '.'];

const loadMessageLabelName = 'load-message';

const gpUkScriptName = 'gpUKScript'
const gpUkCueName = 'gpUKCue';

let chapterMinMaxDetails = {};
let lineNoMinMaxDetails = {};

async function showMain() {
  let waitPage = tag('start-wait');
  let mainPage = tag('main-page');
  waitPage.style.display = 'none';
  mainPage.style.display = 'block';
  await showMainPage();
  console/log('Here');
  await calcAndDisplayMaxAndMin();
}

async function calcAndDisplayMaxAndMin(){
  chapterMinMaxDetails = await calcChapterMinAndMax();
  let ctrlChapterMinMax = tag('min-and-max-chapter');
  ctrlChapterMinMax.innerText = chapterMinMaxDetails.min.toString() + '..' + chapterMinMaxDetails.max.toString();
  lineNoMinMaxDetails = await calcLineNoMinAndMax();
  let ctrlLineNoMinMax = tag('min-and-max-lineNo');
  ctrlLineNoMinMax.innerText = lineNoMinMaxDetails.min.toString() + '..' + lineNoMinMaxDetails.max.toString();
}



async function showMainPage() {
  console.log('Showing Main Page')
  const mainPage = tag('main-page');
  mainPage.style.display = 'block';
  const versionInfo = tag('sheet-version');
  let versionString = 'Version ' + ' Code: ' + codeVersion + ' Released: ';
  versionInfo.innerText = versionString;
  const admin = tag('admin');
  admin.style.display = 'block';
}

async function processGerman() {
  jade_modules.preprocess.showMessage(loadMessageLabelName, 'Processing the german text');
  let hasEols = false;
  await Excel.run(async function(excel) {
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
    for (let i = 0; i < germanText.length; i++) {
      let result = {};
      let myStrings = []
      let original = []
      let startQuotes = locations(openSpeechChar, germanText[i]);
      let eols = locations(eolChar, germanText[i]);
      hasEols = eols.length > 0;
      let endQuotes = locations(closeSpeechChar, germanText[i]);
      let directCopy;
      let goodSpeech = 0;
      let wrongSpeech = 0;
      let unequalQuotes = 0;
      if (startQuotes.length == endQuotes.length) {
        let myIndex = 0;
        if (startQuotes.length == 0) {
          if (eols.length == 0) {
            directCopy = true
            myStrings[0] = germanText[i].trim();
            original[0] = germanText[i].trim();
          } else {
            directCopy = false;
            let myEols = getEols(eols, germanText[i]);
            for (let index = 0; index < myEols.myStrings.length; index++) {
              myStrings[myIndex] = myEols.myStrings[index];
              myIndex += 1;
            }
            for (let index = 0; index < myEols.original.length; index++) {
              original[index] = myEols.original[index];
            }
          }
        } else {
          directCopy = false
          for (let speechPart = 0; speechPart < startQuotes.length; speechPart++) {
            if (hasEols) { console.log('=================== speechpart', speechPart, startQuotes.length, eols) }
            if (endQuotes[speechPart] > startQuotes[speechPart]) {
              goodSpeech += 1;
              if (speechPart == 0) {
                if (startQuotes[speechPart] > 0) {
                  let tempText = germanText[i].substring(0, startQuotes[speechPart]).trim();
                  let tempEols = locations(eolChar, tempText);
                  if (tempEols.length > 0) {
                    let myEols = getEols(tempEols, tempText);
                    if (hasEols) { console.log('=================== eols', speechPart, myEols) }
                    if (myEols.myStrings.length == 0) {
                      myStrings[myIndex] = germanText[i].substring(0, startQuotes[speechPart]).trim();
                      myIndex += 1;
                    } else {
                      for (let index = 0; index < myEols.myStrings.length; index++) {
                        myStrings[myIndex] = myEols.myStrings[index];
                        myIndex += 1;
                      }
                    }
                  } else {
                    myStrings[myIndex] = germanText[i].substring(0, startQuotes[speechPart]).trim();
                    myIndex += 1;
                  }
                }
                original[0] = germanText[i].trim();
              }
              let tempText = germanText[i].substring(startQuotes[speechPart] + 1, endQuotes[speechPart]).trim();
              let tempEols = locations(eolChar, tempText);
              if (tempEols.length > 0) {
                let myEols = getEols(tempEols, tempText);
                if (hasEols) { console.log('=================== first tempText', speechPart, myEols, tempText) };
                if (myEols.myStrings.length == 0) {
                  myStrings[myIndex] = germanText[i].substring(startQuotes[speechPart] + 1, endQuotes[speechPart]).trim();
                  myIndex += 1;
                } else {
                  for (let index = 0; index < myEols.myStrings.length; index++) {
                    myStrings[myIndex] = myEols.myStrings[index];
                    myIndex += 1;
                  }
                }
              } else {
                myStrings[myIndex] = germanText[i].substring(startQuotes[speechPart] + 1, endQuotes[speechPart]).trim();
                myIndex += 1;
              }
              if (speechPart == (startQuotes.length - 1)) {
                if (germanText[i].substring(endQuotes[speechPart]).trim().length > 1) {
                  let tempText = removedBannedOpeningCharacters(germanText[i].substring(endQuotes[speechPart] + 1).trim());
                  let tempEols = locations(eolChar, tempText);
                  if (tempEols.length > 0) {
                    let myEols = getEols(tempEols, tempText);
                    if (hasEols) { console.log('=================== last tempText', speechPart, myEols, tempText) };
                    if (myEols.myStrings.length == 0) {
                      myStrings[myIndex] = removedBannedOpeningCharacters(germanText[i].substring(endQuotes[speechPart] + 1).trim());
                    } else {
                      for (let index = 0; index < myEols.myStrings.length; index++) {
                        myStrings[myIndex] = myEols.myStrings[index];
                        myIndex += 1;
                      }
                    }
                  } else {
                    myStrings[myIndex] = removedBannedOpeningCharacters(germanText[i].substring(endQuotes[speechPart] + 1).trim());
                  }
                }
              } else {
                //The bit between the close quotes and the next open quotes
                let tempText = removedBannedOpeningCharacters(germanText[i].substring(endQuotes[speechPart] + 1, startQuotes[speechPart + 1]).trim());
                let tempEols = locations(eolChar, tempText);
                if (tempEols.length > 0) {
                  let myEols = getEols(tempEols, tempText);
                  if (hasEols) { console.log('=================== middle tempText', speechPart, myEols, tempText) };
                  if (myEols.myStrings.length == 0) {
                    myStrings[myIndex] = removedBannedOpeningCharacters(germanText[i].substring(endQuotes[speechPart] + 1, startQuotes[speechPart + 1]).trim());
                    myIndex += 1;
                  } else {
                    for (let index = 0; index < myEols.myStrings.length; index++) {
                      myStrings[myIndex] = myEols.myStrings[index];
                      myIndex += 1;
                    }
                  }
                } else {
                  myStrings[myIndex] = removedBannedOpeningCharacters(germanText[i].substring(endQuotes[speechPart] + 1, startQuotes[speechPart + 1]).trim());
                  myIndex += 1;
                }
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
      if (directCopy) {
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
      console.log(i, ' - ', startQuotes, ',', endQuotes, ":", result)
    }
    let resultLines = createLines(results);
    await fillRange(germanProcessingSheetName, 'gpOriginal_Spaced', resultLines.original, true);
    await fillRange(germanProcessingSheetName, 'gpProcessed', resultLines.processed, true);
    console.log('Results')
    console.log('Total Good', totalGood, 'Total Wrong', totalWrong, 'Total Unequal', totalUnequal, 'Total Direct Copy', totalDirectCopy)
  })
  jade_modules.preprocess.hideMessage(loadMessageLabelName)
}

function locations(substring, string) {
  var a = [],
    i = -1;
  while ((i = string.indexOf(substring, i + 1)) >= 0) a.push(i);
  return a;
}

function createLines(results) {
  originalLines = [];
  processedLines = [];
  index = 0;
  for (let i = 0; i < results.length; i++) {
    for (let line = 0; line < results[i].lines.length; line++) {
      if (line == 0) {
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
  return { original: originalLines, processed: processedLines }
}

async function fillRange(sheetName, rangeName, dataArray, doClear) {
  await Excel.run(async function(excel) {
    const mySheet = excel.workbook.worksheets.getItem(sheetName);
    const myRange = mySheet.getRange(rangeName);
    myRange.load("rowIndex, columnIndex");
    if (doClear) {
      myRange.clear("Contents")
    }
    await excel.sync();
    console.log('parameters', myRange.rowIndex, myRange.columnIndex, dataArray.length, 1);
    const destRange = mySheet.getRangeByIndexes(myRange.rowIndex, myRange.columnIndex, dataArray.length, 1)
    destRange.load('address');
    await excel.sync();
    console.log('address:', destRange.address);
    let temp = []
    for (let i = 0; i < dataArray.length; i++) {
      temp[i] = [];
      temp[i][0] = dataArray[i];
    }
    console.log(temp)
    destRange.values = temp;
    await excel.sync();
  })
}

function trimEmptyEnd(dataArray) {
  for (let i = dataArray.length - 1; i >= 0; i--) {
    if (dataArray[i] != '') {
      return dataArray.slice(0, i + 1);
    }
  }
  return dataArray;
}

function removedBannedOpeningCharacters(text) {
  let doAgain = true
  let temp = text.trim();
  for (let attempt = 0; attempt < 100; attempt++) {
    if (!doAgain) { break; }
    doAgain = false;
    for (let i = 0; i < bannedOpeningChars.length; i++) {
      if (temp.substr(0, 1) == bannedOpeningChars[i]) {
        temp = temp.substring(1).trim();
        doAgain = true;
      }
    }
  }
  return temp;
}

async function doAReplacement(rowIndex, searchText, replaceText) {
  // Gets the relevant cell/range (original column)
  // replaces the text
  // saves it
  console.log('rowIndex, searchText, replaceText', rowIndex, searchText, replaceText)
  await Excel.run(async function(excel) {
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
    targetRange.values = [
      [targetText]
    ];
    await excel.sync();
  })

}

function getEols(eols, germanText) {
  let myStrings = [];
  let original = [];
  let myIndex = 0;
  for (let eol = 0; eol < eols.length; eol++) {
    if (eol == 0) {
      myStrings[myIndex] = germanText.substring(0, eols[eol]).trim();
      myIndex += 1;
      original[0] = germanText.trim()
    }
    if (eol == eols.length - 1) {
      // last part
      myStrings[myIndex] = germanText.substr(eols[eol] + eolChar.length).trim();
      myIndex += 1;
    } else {
      // Between two eols
      myStrings[myIndex] = germanText.substring(eols[eol] + eolChar.length, eols[eol + 1]).trim();
      myIndex += 1;
    }
  }
  return { myStrings: myStrings, original: original, myIndex: myIndex }
}

function showJump() {
  const jumpTag = tag('jump-buttons')
  if (jumpTag.style.display == 'block') {
    jumpTag.style.display = 'none';
  } else {
    jumpTag.style.display = 'block';
  }
}

function showProcessing() {
  const processingTag = tag('processing-group')
  if (processingTag.style.display == 'block') {
    processingTag.style.display = 'none';
  } else {
    processingTag.style.display = 'block';
  }
}

function showAdmin() {
  const processingTag = tag('admin-group')
  if (processingTag.style.display == 'block') {
    processingTag.style.display = 'none';
  } else {
    processingTag.style.display = 'block';
  }
}

async function getTargetChapter(){
  let ctrlChapter = tag('chapter');
  let chapter = parseInt(ctrlChapter.value);
  if (!isNaN(chapter)){
    await selectChapter(chapter);
  }
}

async function getTargetLineNo(){
  let ctrlLineNo = tag('lineNo')
  let lineNo = parseInt(ctrlLineNo.value);
  if (!isNaN(lineNo)){
    await selectLineNo(lineNo);
  }
}

async function selectChapter(chapterNumber){
  let foundRowIndex = -1;
  for (let i = 0; i < chapterMinMaxDetails.details.length; i++){
    if (chapterMinMaxDetails.details[i].chapterNumber == chapterNumber){
      foundRowIndex = chapterMinMaxDetails.details[i].rowIndex
    }
  }

  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let ukScriptRange = gpSheet.getRange(gpUkScriptName);
    ukScriptRange.load('rowIndex, columnIndex');
    await excel.sync();
    if (foundRowIndex > -1) {
      let chapterRange = gpSheet.getRangeByIndexes(foundRowIndex, ukScriptRange.columnIndex, 1, 1);
      chapterRange.select();
      await excel.sync();
    }
  });
}

async function selectLineNo(lineNo){
  let foundRowIndex;
  let rowIndex;
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let ukCueRange = gpSheet.getRange(gpUkCueName);
    ukCueRange.load('rowIndex, columnIndex, values');
    await excel.sync();
    ukCueValues = ukCueRange.values.map(x => parseInt(x[0]));
    let lineNumber = parseInt(lineNo);
    if (!isNaN(lineNumber)){
      foundRowIndex = ukCueValues.indexOf(lineNumber);
      rowIndex = foundRowIndex + ukCueRange.rowIndex
      if (foundRowIndex > -1){
        let selectRange = gpSheet.getRangeByIndexes(rowIndex, ukCueRange.columnIndex, 1, 1);
        selectRange.select();
        await excel.sync();
      }
    }
    
  })
  

}

async function calcChapterMinAndMax(){
  let chapterDetails = [];
  let minChapter = 1000;
  let maxChapter = 0;
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let ukScriptRange = gpSheet.getRange(gpUkScriptName);
    ukScriptRange.load('rowIndex, columnIndex, values');
    await excel.sync();
    let ukScriptValues = ukScriptRange.values.map(x => x[0].trim().toLowerCase());
    for (let chapterNumber = 1; chapterNumber < 100; chapterNumber++){
      let chapterText = 'chapter ' + number2words(chapterNumber);
      console.log('chapterText', chapterText)
      let foundIndex = ukScriptValues.indexOf(chapterText);
      console.log('foundIndex', foundIndex)
      let foundRowIndex = foundIndex + ukScriptRange.rowIndex
      if (foundIndex > -1) {
        if (chapterNumber < minChapter){minChapter = chapterNumber};
        if (chapterNumber > maxChapter){maxChapter = chapterNumber};
        let tempDetails = {};
        tempDetails.chapterNumber = chapterNumber;
        tempDetails.chapterText = chapterText;
        tempDetails.rowIndex = foundRowIndex;
        console.log('tempDetails', tempDetails);
        chapterDetails.push(tempDetails);
        console.log('chapterDetails', chapterDetails);
      } else {
        break;
      }
    }
  })
  console.log('chapterDetails', chapterDetails);
  console.log('minChapter', minChapter, 'maxChapter', maxChapter)
  return {min: minChapter, max: maxChapter, details: chapterDetails};
}

async function calcLineNoMinAndMax(){
  let minLineNo = 100000;
  let maxLineNo = 0;
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let ukCueRange = gpSheet.getRange(gpUkScriptName);
    ukCueRange.load('rowIndex, values');
    await excel.sync();
    let ukCueValues = ukCueRange.values.map(x => parseInt(x[0]));
    for (let i = 1; i < ukCueValues.length; i++){
      if (!isNaN(ukCueValues[i])) {
        if (ukCueValues[i] < minLineNo){minLineNo = ukCueValues[i]};
        if (ukCueValues[i] > maxLineNo){maxLineNo = ukCueValues[i]};
      }
    }
  })
  console.log('minLineNo', minLineNo, 'maxLineNo', maxLineNo)
  return {min: minLineNo, max: maxLineNo};
}

function number2words(n) {
  const num = "zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen nineteen".split(" ");
  const tens = "twenty thirty forty fifty sixty seventy eighty ninety".split(" ");
  if (n < 20) return num[n];
  var digit = n % 10;
  if (n < 100) return tens[~~(n / 10) - 2] + (digit ? "-" + num[digit] : "");
  if (n < 1000) return num[~~(n / 100)] + " hundred" + (n % 100 == 0 ? "" : " and " + number2words(n % 100));
  return number2words(~~(n / 1000)) + " thousand" + (n % 1000 != 0 ? " " + number2words(n % 1000) : "");
}

async function getUsedRowCount(sheetName, rangeName){
  //Returns the rowIndex of the last cell
  let rowCount
  await Excel.run(async function(excel) {
    const sheet = excel.workbook.worksheets.getItem(sheetName);
    let wholeRange = sheet.getRange(rangeName);
    let usedRange = wholeRange.getUsedRange(true);
    usedRange.load('rowCount, address')
    await excel.sync();
    rowCount = usedRange.rowCount;
    console.log(usedRange.address, usedRange.rowCount);
  })
  return rowCount;
}