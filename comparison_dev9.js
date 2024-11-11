const pdfComparisonSheetName = 'PDF Comparison';
const apostropheSheetName = 'Apostrophes';
const sourceColumnIndex = 3;
const chaptersColumnIndex = 14;
const startRowIndex = 10;
const linesColumnIndex = 5;

async function getRowColumnDetails(){
  let details = {};
  await Excel.run(async function(excel){ 
    const pdfSheet = excel.workbook.worksheets.getItem(pdfComparisonSheetName);
    const usedRange = pdfSheet.getUsedRange();
    usedRange.load('rowIndex, rowCount, columnIndex, columnCount');
    await excel.sync();
    details = {
      rowIndex: usedRange.rowIndex,
      rowCount: usedRange.rowCount,
      columnIndex: usedRange.columnIndex,
      columnCount: usedRange.columnCount
    }
  })
  return details;
}

async function getChapterData(){
  //The text is pasted into column B of PDF Comparison from the adobe conversion
  //Some text massaging happens with result in column D (index 3)
  // this routine takes the text from column D and turns it into text strings
  //Splitting it into chapters. The output is a string array with a chapter per index
  const details = await getRowColumnDetails();
  let chapters = [];
  let index = -1;
  textSoFar = '';
  await Excel.run(async function(excel){  
    const pdfSheet = excel.workbook.worksheets.getItem(pdfComparisonSheetName);
    const rowCount = details.rowCount - details.rowIndex + 1 - startRowIndex;
    const sourceRange = pdfSheet.getRangeByIndexes(startRowIndex, sourceColumnIndex, rowCount, 1);
    sourceRange.load('rowIndex, values');
    await excel.sync();
    sourceValues = sourceRange.values.map(x => x[0]);

    for (let i = 0; i < sourceValues.length; i++){
      //console.log('i', i, 'value', sourceValues[i]);
      let text = sourceValues[i].trim();
      if (text != ''){
        //Does the string include 'chapter'
        if (text.toLowerCase().includes('chapter')){
          //Finish last chapter and start new one.
          if (textSoFar != ''){
            index += 1;
            chapters[index] = textSoFar;
          }
          textSoFar = text;
        } else {
          //append to textSoFar
          textSoFar = textSoFar + ' ' + text;
        }
      }
    }
    index += 1;
    chapters[index] = textSoFar;
  })
  return chapters  
}

function chapterToLines(theChapter){
  // takes a chapter and splits it into lines
  // if begins with a single ' it is doubled ''
  // This is for excel display 
  
  let myLines = theChapter.split("\n");
  for (let i = 0; i < myLines.length; i++){
    if (myLines[i].startsWith("'") && (!myLines[i].startsWith["''"])){
      myLines[i] = "'" + myLines[i];
    }
  }
  return myLines; 
} 

function createQuoteData(myLines, apostrophes, noKeeps){
  
  //Now search for curly quotes within each line
  let quoteData = [];
  let quoteIndex = -1;
  for (let i = 0; i < myLines.length; i++){
    let openQuote = findCurlyQuote('‘', myLines[i], false, apostrophes);
    let closeQuote = findCurlyQuote('’', myLines[i], true, apostrophes);
    quoteIndex += 1;
    let done = false;
    if (noKeeps){
      let theData = {
        index: i,
        text: myLines[i],
        openQuote: {},
        closeQuote: {}
        }
      theData.subStrings = [];
      quoteData[quoteIndex] = theData;
    } else {
      if ((openQuote.length == 1) && closeQuote.length == 1){
        console.log('Zero Indexes ', openQuote[0], closeQuote[0], 'text', myLines[i])
        if ((openQuote[0] <= 1) && closeQuote[0] >= (myLines[i].length - 2)){
          console.log('Used zero index');
          let theData = {
            index: i,
            text: myLines[i],
            openQuote: {},
            closeQuote: {}
            }
          theData.subStrings = [];
          quoteData[quoteIndex] = theData;
          done = true;
        }
      }
      if (!done){
        if ((openQuote.length > 0) || (closeQuote.length > 0)){
          console.log('Some Indexes ', openQuote, closeQuote, 'text', myLines[i])
          let theData = {
            index: i,
            text: myLines[i],
            openQuote: openQuote,
            closeQuote: closeQuote
          }
          theData.subStrings = createQuoteStrings(theData);
          quoteData[quoteIndex] = theData;
        } else {
          let theData = {
            index: i,
            text: myLines[i],
            openQuote: {},
            closeQuote: {}
            }
          theData.subStrings = [];
          quoteData[quoteIndex] = theData;
        }
      }
    }
  }
  return quoteData;
  
}
async function createChapters(){
  
  const apostrophes = await apostropheWords();
  const chapters = await getChapterData();
  let myLines = chapterToLines(chapters[4]);
  let quoteData = createQuoteData(myLines, apostrophes, false);  
  await displayDecision(quoteData, true);
}

async function createResult(){
  const apostrophes = await apostropheWords();
  let myLines = await readDecisionData();
  let quoteData = createQuoteData(myLines, apostrophes, true);  
  await displayDecision(quoteData, false);
}

function findCurlyQuote(character, myString, doApostropheCheck, words){
  let index = 0
  let result = []
  let position = myString.indexOf(character, index)
  while (position != -1){
    if (doApostropheCheck){
      if (!containsApostropheWord(myString, position, words)){
        result.push(position);
      }
    } else {
      result.push(position);  
    }
    index = position + 1;
    position = myString.indexOf(character, index)
  }
  
  return result;
}
/*
‘Lying there with their eyes wide open! Cold as ice! Still in their dinner things!’
*/
function createQuoteStrings(theData){
  // theData includes text and openQuote and closeQuote
  // returns an array of objects
  //  start, stop, substring
  
  let result = [];
  let index = -1;
  
  //loop through openQuote
  for (let i = 0; i < theData.openQuote.length; i++){
    //loop through closeQuote
    for (let j = 0; j < theData.closeQuote.length; j++){
      // if closeQuote > openQuote create substring object
      let doIt = false;
      if (theData.closeQuote[j] > theData.openQuote[i]){
        if ((i + 1) < theData.openQuote.length){
          if (theData.openQuote[i + 1] > theData.closeQuote[j]){
            doIt = true;
          }
        } else {
          doIt = true;
        }
        if (doIt){
          index += 1;
          result[index] = {
            start: theData.openQuote[i] + 1,
            stop: theData.closeQuote[j],
            subString: theData.text.substring(theData.openQuote[i] + 1, theData.closeQuote[j])
          }
        }
      }
    }
  }
  return result;
}
async function findCommonContractions(myText){
  const test = ['I’m', 'You’re', 'He’s', 'She’s', 'It’s', 'We’re', 'They’re', 'Can’t', 'Don’t', 'Won’t', 'Shouldn’t', 'Wouldn’t', 'Couldn’t', 'Isn’t', 'Aren’t', 'Haven’t', 'Hasn’t', 'Hadn’t', 'Wasn’t', 'Weren’t', 'I’ve', 'didn’t']

}
//'‘Always thought he was odd,’ she told the eagerly listening villagers, after her fourth sherry. ‘Unfriendly, like. I’m sure if I’ve offered him a cuppa once, I’ve offered it a hundred times. Never wanted to mix, he didn’t.’'


async function apostropheWords(){
  let values = [];
  await Excel.run(async function(excel){
    const range = excel.workbook.worksheets.getItem(apostropheSheetName).getRange('apWords');
    range.load('values')
    await excel.sync();
    values = range.values.map(x => x[0]).filter(x => x != '');
  })
  let results = [];
  for (let i = 0; i < values.length; i++){
    results[i] = {
      word: values[i],
      position: values[i].indexOf('’'),
      length: values[i].length
    }
  }
  return results;
}

function containsApostropheWord(text, position, words){
  // text is the full line of text, position is the pos of the apostrophe.
  // words contains the list of test apostrophe words and apostrophe position/length.
  
  
  for (let i = 0; i < words.length; i++){
    let start = position - words[i].position;
    let stop = start + words[i].length;
    if (start < 0){start = 0}
    if (stop >= text.length){stop = text.length}
    let test = text.substring(start, stop).toLowerCase();
    console.log()
    if (test == words[i].word){
      return true;
    }
  }
  return false;
}

async function displayDecision(quoteData, doDecision){
  const lineIndex = 0;
  const textIndex = 1;
  const subNoIndex = 2
  const substringIndex = 3;
  const startIndex = 4;
  const endIndex = 5;

  await Excel.run(async function(excel){
    let displayRange
    if (doDecision){
      displayRange = excel.workbook.worksheets.getItem('Decision').getRange('deTable');
    } else {
      displayRange = excel.workbook.worksheets.getItem('Result').getRange('reTable');
    }
    displayRange.clear('Contents');
    displayRange.load('rowIndex, rowCount, columnIndex, columnCount');
    await excel.sync();
    let rowIndex = -1;
    let display = [];
    for (let i = 0; i < quoteData.length; i++){
      rowIndex += 1;
      console.log('quotedata', quoteData[i]);
      if (quoteData[i].subStrings.length > 0){
        display[rowIndex] = [];
        display[rowIndex][lineIndex] = i;
        display[rowIndex][textIndex] = quoteData[i].text;
        display[rowIndex][subNoIndex] = 0;
        display[rowIndex][substringIndex] = quoteData[i].subStrings[0].subString;
        display[rowIndex][startIndex] = quoteData[i].subStrings[0].start;
        display[rowIndex][endIndex] = quoteData[i].subStrings[0].stop;
        for (let j = 1; j < quoteData[i].subStrings.length; j++){
          rowIndex += 1
          display[rowIndex] = [];
          display[rowIndex][lineIndex] = i;
          display[rowIndex][textIndex] = '';
          display[rowIndex][subNoIndex] = j;
          display[rowIndex][substringIndex] = quoteData[i].subStrings[j].subString;
          display[rowIndex][startIndex] = quoteData[i].subStrings[j].start;
          display[rowIndex][endIndex] = quoteData[i].subStrings[j].stop;
        }  
      } else {
        display[rowIndex] = [];
        display[rowIndex][lineIndex] = i;
        display[rowIndex][textIndex] = quoteData[i].text;
        display[rowIndex][subNoIndex] = '';
        display[rowIndex][substringIndex] = '';
        display[rowIndex][startIndex] = '';
        display[rowIndex][endIndex] = '';
      }
    }
    console.log('Display', display);
    console.log(displayRange.rowIndex, displayRange.columnIndex, display.length, displayRange.columnCount);
    let tempRange;
    if (doDecision){
      tempRange = excel.workbook.worksheets.getItem('Decision').getRangeByIndexes(displayRange.rowIndex, displayRange.columnIndex, display.length, displayRange.columnCount);  
    } else {
      tempRange = excel.workbook.worksheets.getItem('Result').getRangeByIndexes(displayRange.rowIndex, displayRange.columnIndex, display.length, displayRange.columnCount);
    }
    tempRange.values = display;
  })
}

async function readDecisionData(){
  const lineIndex = 0;
  const textIndex = 1;
  const startIndex = 4;
  const endIndex = 5;
  const decisionIndex = 6;
  let myLines = [];

  await Excel.run(async function(excel){
    let displayRange = excel.workbook.worksheets.getItem('Decision').getRange('deTableDecision');
    
    displayRange.load('rowIndex, values');
    await excel.sync();
    let prevLine = -1;
    let original;
    for (let i = 0; i < displayRange.values.length; i++){
      let line = parseInt(displayRange.values[i][lineIndex]);
      //console.log('line', line);
      if (!isNaN(line)){
        if (line != prevLine){
          original = displayRange.values[i][textIndex]
          let offset = 0;
          let nextLine = parseInt(displayRange.values[i + offset][lineIndex]);
          let myDecision = [];
          if (!isNaN(nextLine)){
            while (nextLine == line){
              myDecision[offset] = {
                decision: displayRange.values[i + offset][decisionIndex],
                start: displayRange.values[i + offset][startIndex],
                end: displayRange.values[i + offset][endIndex]
              }
              offset += 1;
              nextLine = parseInt(displayRange.values[i + offset][lineIndex]);
              if (isNaN(nextLine)){break}
            }
          }
          let tempLines = doSplit(original, myDecision);
          if (tempLines.length > 1){
            console.log('line', line, 'original', original, 'the lines', tempLines);
          }
          myLines = myLines.concat(tempLines);
          prevLine = line; 
        }
      }
    } 
    console.log('myLines', myLines);
  })
  return myLines;
}
function doSplit(original, decisions){
  let indexes = []
  //console.log('decisions', decisions);
  for (let i = 0; i < decisions.length; i++){
    if (decisions[i].decision.toLowerCase() == 'split'){
      indexes.push(decisions[i].start);
      indexes.push(decisions[i].end);
    }
  }
  if (indexes.length == 0){
    indexes = [0, original.length]
  }
  //console.log('indexes', indexes)
  let duplicatesRemoved = Array.from(new Set(indexes));
  //console.log('duplicated removed', duplicatesRemoved)
  let sortedIndexes = duplicatesRemoved.sort((a,b) => a - b);

  if (sortedIndexes.slice(-1) < (original.length -2)){
    //we need to add a last index
    sortedIndexes.push(original.length);
  }

  if (sortedIndexes[0] > 1){
    sortedIndexes.push(0)
    sortedIndexes = sortedIndexes.sort((a,b) => a - b);
  }

  let item = -1;
  let myLines = [];
  for (let i = 0; i < (sortedIndexes.length - 1); i++){
    item += 1;
    myLines[item] = removeAndTrim(original.substring(sortedIndexes[i], sortedIndexes[i + 1]));
  }
  if (myLines.length > 0){
    console.log('Original', original, 'sorted Indexes', sortedIndexes, 'split Lines', myLines);
  }

  return myLines;
}

function removeAndTrim(theText){
  //removes a ’ if first character and ‘ if it's last character.
  let temp = theText.trim();
  //console.log('temp(0)', temp[0])
  if (temp[0] == '’'){
    temp = temp.substring(1);
    //console.log('removed first', temp)
  }
  //console.log('temp.slice(-1)', temp.slice(-1))
  if (temp.slice(-1) == '‘'){
    temp = temp.slice(0, -1);
    //console.log('removed last', temp)
  }
  return temp.trim();

}

async function copySheets(){
  let chapterCompareSelect = tag('chapter-compare-select');
  let myChapter = chapterCompareSelect.value;
  console.log(myChapter);
  
  await Excel.run(async (excel) => {
    let myWorkbook = excel.workbook;
    let decisionSheet = myWorkbook.worksheets.getItem('Decision');
    let copiedSheet = decisionSheet.copy("End")

    decisionSheet.load("name");
    copiedSheet.load("name");

    await excel.sync();

    console.log("'" + decisionSheet.name + "' was copied to '" + copiedSheet.name + "'")
    copiedSheet.name = decisionSheet.name + ' Chapter ' + myChapter;

    let resultSheet = myWorkbook.worksheets.getItem('Result');
    copiedSheet = resultSheet.copy("End")

    resultSheet.load("name");
    copiedSheet.load("name");

    await excel.sync();

    console.log("'" + resultSheet.name + "' was copied to '" + copiedSheet.name + "'")
    copiedSheet.name = resultSheet.name + ' Chapter ' + myChapter;


  });

}

async function fillChapter(){
  const minAndMax = await jade_modules.operations.getChapterMaxAndMin();
  
  let chapterCompareSelect = tag('chapter-compare-select');
  chapterCompareSelect.innerHTML = '';
  chapterCompareSelect.add(new Option('Please select chapter', ''));
  for (let i = minAndMax.min; i <= minAndMax.max; i++){
    chapterCompareSelect.add(new Option('Chapter ' + i, i));
  }
}

async function clearDecisionAndResult(){
  await Excel.run(async (excel) => {
    let decisionSheet = excel.workbook.worksheets.getItem('Decision');
    let resultSheet = excel.workbook.worksheets.getItem('Result');
    let decisionTable = decisionSheet.getRange('deTable');
    let keepRange = decisionSheet.getRange('deKeep');
    let resultTable = resultSheet.getRange('reTable');

    decisionTable.clear('Contents');
    keepRange.clear('Contents');
    resultTable.clear('Contents');
  })


}