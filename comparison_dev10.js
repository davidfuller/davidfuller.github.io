const pdfComparisonSheetName = 'PDF Comparison';
const apostropheSheetName = 'Apostrophes';
const sourceColumnIndex = 3;
const chaptersColumnIndex = 14;
const startRowIndex = 10;
const linesColumnIndex = 5;

const myTypes = {
  chapter: 'Chapter',
  scene: 'Scene',
  line: 'Line',
  sceneBlock: 'Scene Block',
  wallaScripted: 'Walla Scripted',
  wallaBlock: 'Walla Block'
}

const issueType = {
  finished: 'finished',
  fixLf: 'fix LF',
  fixSpaceQuote: 'fix space quote'
}

function auto_exec(){
}

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
    const rowCount = details.rowCount;
    //console.log(startRowIndex, sourceColumnIndex, rowCount, 1)
    const sourceRange = pdfSheet.getRangeByIndexes(startRowIndex, sourceColumnIndex, rowCount, 1);
    sourceRange.load('rowIndex, values');
    await excel.sync();
    sourceValues = sourceRange.values.map(x => x[0]);

    for (let i = 0; i < sourceValues.length; i++){
      //console.log('i', i, 'value', sourceValues[i]);
      let text = sourceValues[i].trim();
      if (text != ''){
        //Does the string include 'chapter'
        if (text.toLowerCase().includes('— chapter')){
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
        //console.log('Zero Indexes ', openQuote[0], closeQuote[0], 'text', myLines[i])
        if ((openQuote[0] <= 1) && closeQuote[0] >= (myLines[i].length - 2)){
          //console.log('Used zero index');
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
          //console.log('Some Indexes ', openQuote, closeQuote, 'text', myLines[i])
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
  let chapterCompareSelect = tag('chapter-compare-select');
  let myChapter = chapterCompareSelect.value;
  const apostrophes = await apostropheWords();
  const chapters = await getChapterData();
  //console.log('Chapters', chapters);
  let myLines = chapterToLines(chapters[myChapter - 1]);
  //console.log('myLines', myLines);
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
    //console.log()
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
      //console.log('quotedata', quoteData[i]);
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
    //console.log('Display', display);
    //console.log(displayRange.rowIndex, displayRange.columnIndex, display.length, displayRange.columnCount);
    let tempRange;
    if (doDecision){
      tempRange = excel.workbook.worksheets.getItem('Decision').getRangeByIndexes(displayRange.rowIndex, displayRange.columnIndex, display.length, displayRange.columnCount);  
    } else {
      tempRange = excel.workbook.worksheets.getItem('Result').getRangeByIndexes(displayRange.rowIndex, displayRange.columnIndex, display.length, displayRange.columnCount);
      excel.workbook.worksheets.getItem('Result').activate();
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
            //console.log('line', line, 'original', original, 'the lines', tempLines);
          }
          myLines = myLines.concat(tempLines);
          prevLine = line; 
        }
      }
    } 
    //console.log('myLines', myLines);
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
    //console.log('Original', original, 'sorted Indexes', sortedIndexes, 'split Lines', myLines);
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

    resultSheet.activate();
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
    let manualRange = decisionSheet.getRange('deManual');

    decisionTable.clear('Contents');
    keepRange.clear('Contents');
    resultTable.clear('Contents');
    manualRange.clear('Contents');
  })
}

async function selectResultLowestTrue(){
  await Excel.run(async (excel) => {
    let resultSheet = excel.workbook.worksheets.getItem('Result');
    let compRange = resultSheet.getRange('reComparison');
    compRange.load('values, rowIndex, columnIndex');
    await excel.sync();
    
    let compValues = compRange.values.map(x => x[0]);
    let highest = 0
    for (i = 0; i < compValues.length; i++){
      if (compValues[i]){
        highest = i;
      }
    }
    let cellToSelect = resultSheet.getRangeByIndexes(highest + compRange.rowIndex, compRange.columnIndex, 1, 1);
    cellToSelect.select();
  })
}
async function correctTextReplaceLF(doReplace){
  if (doReplace){
    await putSelectedCellInTextArea();
  }
  let replaceColumnIndex = 1;
  let searchDetails = await findSearchTextInPDF();
  let success = false;

  await Excel.run(async (excel) => {
    let pdfSheet = excel.workbook.worksheets.getItem('PDF Comparison');
    let indexes = searchDetails.indexes;
    //console.log('indexes', indexes);
    if (indexes.length == 1){
      let foundText = searchDetails.bookText[indexes[0]];
      let index = foundText.toLowerCase().indexOf(searchDetails.mySearch.toLowerCase());
      let position = index + searchDetails.mySearch.length;
      let char = foundText.substr(position, 1);
      let twoChars = foundText.substr(position, 2);
      let threeChars = foundText.substr(position, 3);
      let newText;
      //console.log('the char', char, 'the area', foundText.substr(position - 5, 10));
      if (char == '\n'){
        newText = foundText.substring(0, position) + ' ' + foundText.substr(position + 1);
        //console.log('newText', newText);
        //now lets put it back in the pdf sheet.
        
        let rowIndex = indexes[0] + searchDetails.rowIndex;
        let replaceRange = pdfSheet.getRangeByIndexes(rowIndex, replaceColumnIndex, 1, 1);
        replaceRange.load('address');
        if (doReplace){
          console.log('Replace LF. Doing It');
          replaceRange.values = [[newText]];
          success = true;
        } else {
          console.log('This looks good');
        }
        await excel.sync();
        //console.log('address', replaceRange.address);
        if (doReplace){
          await createChapters();
          await doKeepsAndManuals();
          await createResult();
        }
      } else if ((twoChars == '\r\n') || (twoChars == '’\n')){
        newText = foundText.substring(0, position) + ' ' + foundText.substr(position + 2);
        let rowIndex = indexes[0] + searchDetails.rowIndex;
        let replaceRange = pdfSheet.getRangeByIndexes(rowIndex, replaceColumnIndex, 1, 1);
        replaceRange.load('address');
        if (doReplace){
          console.log('Replace LF. Doing It');
          replaceRange.values = [[newText]];
          success = true;
        } else {
          console.log('This looks good');
        }
        await excel.sync();
        //console.log('address', replaceRange.address);
        if (doReplace){
          await createChapters();
          await doKeepsAndManuals();
          await createResult();
        }
      } else if (threeChars == '’\r\n'){
        newText = foundText.substring(0, position) + ' ' + foundText.substr(position + 3);
        let rowIndex = indexes[0] + searchDetails.rowIndex;
        let replaceRange = pdfSheet.getRangeByIndexes(rowIndex, replaceColumnIndex, 1, 1);
        replaceRange.load('address');
        if (doReplace){
          console.log('Replace LF. Doing It');
          replaceRange.values = [[newText]];
          success = true;
        } else {
          console.log('This looks good');
        }
        await excel.sync();
        //console.log('address', replaceRange.address);
        if (doReplace){
          await createChapters();
          await doKeepsAndManuals();
          await createResult();
        }
      } else {
        console.log('Replace LF. Not the expected LF', char);
        for (let c = -5; c <= 5; c++){
          console.log(c, foundText.substr(position + c, 1));
        }
      }
    } else {
      console.log('Replace LF. Too many posibilities', indexes);
    }
  })
  console.log('Replace LF success.', success);
  return success;
}
async function correctTextSpaceQuotes(doReplace){
  if (doReplace){
    await putSelectedCellInTextArea();
  }
  let replaceColumnIndex = 1;
  let searchDetails = await findSearchTextInPDF();
  let success = false;

  await Excel.run(async (excel) => {
    let pdfSheet = excel.workbook.worksheets.getItem('PDF Comparison');
    let indexes = searchDetails.indexes;
    //console.log('indexes', indexes);
    if (indexes.length == 1){
      let foundText = searchDetails.bookText[indexes[0]];
      let index = foundText.toLowerCase().indexOf(searchDetails.mySearch.toLowerCase());
      let position = index + searchDetails.mySearch.length + 1; //+1 to get pat closing quote
      let char = foundText.substr(position, 1);
      //console.log('the char', char, 'the area', foundText.substr(position - 5, 10));
      let newText
      let rowIndex = indexes[0] + searchDetails.rowIndex;
      if (char == ' '){
        newText = foundText.substring(0, position) + '\n' + foundText.substr(position + 1);
        //console.log('newText', newText);
        let replaceRange = pdfSheet.getRangeByIndexes(rowIndex, replaceColumnIndex, 1, 1);
        replaceRange.load('address');
        if (doReplace){
          console.log('Doing it')
          replaceRange.values = [[newText]];
          success = true;
        } else {
          console.log('This looks good');
        }
        await excel.sync();
        //console.log('address', replaceRange.address);
      } else {
        if ((searchDetails.isEnds[0]) && (doReplace)) {
          success = await fixEndOfCellSpaceQuotes(rowIndex)
        } else {
          console.log('A space was expected here, but we got:', char);
          for (let c = -5; c <= 5; c++){
            console.log(c, foundText.substr(c + position, 1));
          } 
        }
      }
      if ((success) && (doReplace)){
        await createChapters();
        await doKeepsAndManuals();
        await createResult();
      }
    } else {
      console.log('Too many possibilities to accurately guess', indexes)
    }
  })
  console.log('Text Qute Space Success', success);
  return success;
}

async function fixEndOfCellSpaceQuotes(rowIndex){
  //rowIndex is the row which has the text that ends the cell
  //Find next row with text in. (within next 10)
  let replaceColumnIndex = 1;
  let nextOne = -1;
  let success = false
  await Excel.run(async (excel) => {
    let pdfSheet = excel.workbook.worksheets.getItem('PDF Comparison');
    let testRange = pdfSheet.getRangeByIndexes(rowIndex, replaceColumnIndex, 10, 1);
    testRange.load('rowIndex, values');
    await excel.sync();
    for (i = 1; i < testRange.values.length; i++){
      if (testRange.values[i][0].toString().trim() != ''){
        nextOne = i;
        break;
      }
    }
    if (nextOne != -1){
      //find the first space
      let firstSpace = testRange.values[nextOne][0].indexOf(' ');
      if (firstSpace != -1){
        //create the bitToMove
        let bitToMove = testRange.values[nextOne][0].substring(0, firstSpace).trim();
        //Create new texts
        let newFirstText = (testRange.values[0][0] + '\n' + bitToMove).trim();
        let newNextText = testRange.values[nextOne][0].substr(firstSpace).trim();
        //
        let firstText = pdfSheet.getRangeByIndexes(rowIndex, replaceColumnIndex, 1, 1);
        let nextText = pdfSheet.getRangeByIndexes(testRange.rowIndex + nextOne, replaceColumnIndex, 1, 1)
        firstText.values =[[newFirstText]];
        nextText.values = [[newNextText]];
        success = true
        console.log('new first text', newFirstText);
        console.log('new next text', newNextText);
      } else {
        console.log('Failed to find first space')
      }
    } else {
      console.log('Failed to find next text')
    }
  })
  return success;
}

async function findSearchTextInPDF(){
  //Takes the text from the textArea and finds every occurace in the pdf book.
  let searchText = tag('search-text');
  //console.log('searchText', searchText.value);
  let mySearch = searchText.value;
  let firstRowIndex = 9;
  let lastRowIndex = 1300;
  let columnIndex = 3;
  let bookText;
  let bookRange;
  let indexes = [];
  let isEnds = [];
  await Excel.run(async (excel) => {
    let pdfSheet = excel.workbook.worksheets.getItem('PDF Comparison');
    bookRange = pdfSheet.getRangeByIndexes(firstRowIndex, columnIndex, (lastRowIndex - firstRowIndex + 1), 1);
    bookRange.load('values, rowIndex');
    await excel.sync();
    bookText = bookRange.values.map(x => x[0]);
    //console.log('bookText', bookText);
    
    for (i = 0; i < bookText.length; i ++){
      let foundText = bookText[i];
      let isEnd = false;
      if (foundText.toLowerCase().includes(mySearch.toLowerCase())){
        let index = foundText.toLowerCase().indexOf(mySearch.toLowerCase());
        if (index != -1){
          let endPosition = index + mySearch.length;
          isEnd = (Math.abs(endPosition - foundText.length) < 2)
        }
        indexes.push(i);
        isEnds.push(isEnd);
      }
    }
  })
  return {
    indexes: indexes,
    bookText: bookText,
    rowIndex: bookRange.rowIndex,
    mySearch: mySearch,
    isEnds: isEnds
  }
}

async function putSelectedCellInTextArea(){
  await Excel.run(async (excel) => {
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load('values');
    await excel.sync();
    let searchText = tag('search-text');
    searchText.value = activeCell.values[0][0];
  })
}

async function findSearchTextInDecision(){
  await putSelectedCellInTextArea();
  let searchText = tag('search-text');
  console.log('searchText', searchText.value);
  let mySearch = searchText.value;
  let textArrayIndex = 1;
  let textColumnIndex = 2;
  await Excel.run(async (excel) => {
    let decisionSheet = excel.workbook.worksheets.getItem('Decision');
    let tableRange = decisionSheet.getRange('deTable');
    tableRange.load('values, rowIndex');
    await excel.sync();
    let textValues = tableRange.values.map(x => x[textArrayIndex]);
    let found = false;
    for (let i = 0; i < textValues.length; i++){
      if (textValues[i].toLowerCase().includes(mySearch.toLowerCase())){
        let rowIndex = i + tableRange.rowIndex;
        let selectedRange = decisionSheet.getRangeByIndexes(rowIndex, textColumnIndex, 1, 1);
        selectedRange.select();
        decisionSheet.activate();
        await excel.sync();
        found = true;
        break;
      }
    }
    console.log('Found', found);
  })
}
async function getLinksToTextFromChapter(){
  let chapterCompareSelect = tag('chapter-compare-select');
  let myChapter = parseInt(chapterCompareSelect.value);
  let formulaPrefix = '=Script!K'
  console.log(myChapter);
  await Excel.run(async (excel) => {
    const chapterRange = await jade_modules.operations.getChapterRange(excel);
    chapterRange.load('values, rowIndex');
    const typeCodeRange = await jade_modules.operations.getTypeCodeRange(excel);
    typeCodeRange.load('values, rowIndex');
    await excel.sync();
    let chapterRowIndexes = [];
    console.log('Chapter Range Values', chapterRange.values, 'Typecode Range values', typeCodeRange.values)
    for (let i = 0; i < chapterRange.values.length; i++){
      if ((chapterRange.values[i][0] == myChapter) && ((typeCodeRange.values[i][0] == myTypes.line) || typeCodeRange.values[i][0] == myTypes.scene)){
        chapterRowIndexes.push(i + chapterRange.rowIndex);
      }
    }
    console.log('chapterRowIndexes', chapterRowIndexes);
    let formulas = []
    for (let i = 0; i < chapterRowIndexes.length; i++){
      formulas[i] = [formulaPrefix + (chapterRowIndexes[i] + 1)];
    }
    console.log('formulas', formulas);

    const resultSheet = excel.workbook.worksheets.getItem('Result')
    let scriptRange = resultSheet.getRange('reScript');
    scriptRange.clear('Contents');
    scriptRange.load('rowIndex, columnIndex');
    await excel.sync();
    let tempRange = resultSheet.getRangeByIndexes(scriptRange.rowIndex, scriptRange.columnIndex, formulas.length, 1);
    tempRange.formulas = formulas
  })  
}
async function createChaptersAndResults(){
  await createChapters();
  await doKeepsAndManuals();
  await createResult();
}

async function findInPDF(){
  await putSelectedCellInTextArea();
  let searchDetails = await findSearchTextInPDF();
  let results = [];
  let columnIndex = 1;
  await Excel.run(async (excel) => {
    let pdfSheet = excel.workbook.worksheets.getItem('PDF Comparison');
    let indexes = searchDetails.indexes;
    console.log('indexes', indexes);
    
    for (let i = 0; i < indexes.length; i++){
      let foundText = searchDetails.bookText[indexes[i]];
      let index = foundText.toLowerCase().indexOf(searchDetails.mySearch.toLowerCase());
      if (index != -1){
        let rowIndex = indexes[i] + searchDetails.rowIndex;
        let endPosition = index + searchDetails.mySearch.length;
        let isEnd = (Math.abs(endPosition - foundText.length) < 2)
        let myResult = {
          message: searchDetails.mySearch + ' found at position ' + index + ' to ' + (index + searchDetails.mySearch.length) + ' in row ' + (rowIndex + 1) + ' isEnd ' + isEnd,
          searchText: searchDetails.mySearch,
          startPosition: index,
          endPosition: endPosition,
          rowIndex: rowIndex,
          isEnd: isEnd
        }
        results.push(myResult);
      }
    }
    if (results.length > 0){
      let selectedCell = pdfSheet.getRangeByIndexes(results[0].rowIndex, columnIndex, 1, 1);
      pdfSheet.activate();
      selectedCell.select();
    }
    console.log('Results: ', results.map(x => x.message));
  })
}

async function findRed(){
  //returns rowIndex of first row where charDiff > 5
  let result = -1;
  await Excel.run(async (excel) => {
    const resultSheet = excel.workbook.worksheets.getItem('Result');
    const charDiff = resultSheet.getRange('reCharDiff');
    charDiff.load('rowIndex, values');
    await excel.sync();
    let values = charDiff.values.map(x => x[0]);
    for (let i = 0; i < values.length; i++){
      if (values[i] > 5){
        let doneCell = charDiff.getCell(i, 1);
        doneCell.load('values, address');
        await excel.sync();
        console.log('donecell', doneCell.address);
        console.log(doneCell.values[0][0])
        if (doneCell.values[0][0].toLocaleLowerCase() != 'done'){
          result = i + charDiff.rowIndex;
          break
        }
      }
    }
  });
  return result;
}

async function findEmpty() {
  let result = -1;
  await Excel.run(async (excel) => {
    const resultSheet = excel.workbook.worksheets.getItem('Result');
    let bookRange = resultSheet.getRange('reBook');
    bookRange.load('rowIndex, values');
    await excel.sync();
    let values = bookRange.values.map(x => x[0]);
    for (let i = 0; i < values.length; i++){
      if (values[i].trim() == ''){
        result = i + bookRange.rowIndex - 1;
        // Check line number
        let lineCell = bookRange.getCell(result, -1);
        lineCell.load('values, address')
        await excel.sync();
        console.log('lineCell',lineCell.address);
        console.log('lineCell',lineCell.values);
        if (isNaN(parseInt(lineCell.values[0][0]))){
          result = -1;
        }
        break;
      }
    }
  })
  return result;
}

async function fixNextIssue() {
  //finds next issue, be it red or empty line, and attepts to fix it.
  const redLine = await findRed();
  console.log('redLine', redLine);
  const empty = await findEmpty();
  console.log('empty', empty)
  
  /*
  finished
  fixLf
  fixSpaceQuote
  */
  let rowIndex;
  let issue;
  if (empty == -1){
    rowIndex = redLine;
    if (redLine == -1){
      issue = issueType.finished;
    } else {
      issue = issueType.fixLf
    }
  } else if (redLine == -1){
    rowIndex = empty;
    issue = issueType.fixSpaceQuote;
  } else if (empty < redLine) {
    rowIndex = empty;
    issue = issueType.fixSpaceQuote;
  } else {
    rowIndex = redLine;
    issue = issueType.fixLf
  }

  if (rowIndex > -1) {
    await selectIssueCell(rowIndex);
  }
  
  return {
    rowIndex: rowIndex,
    issue: issue
  }
}

async function selectIssueCell(rowIndex){
  await Excel.run(async (excel) => {
    const resultSheet = excel.workbook.worksheets.getItem('Result');
    let bookRange = resultSheet.getRange('reBook');
    bookRange.load('rowIndex');
    await excel.sync();
    let selectCell = bookRange.getCell(rowIndex - bookRange.rowIndex, 0);
    selectCell.select();
  })  
}

async function comparisonLoop(){
  //This will loop through next fixes.
  
  let finished = false;
  let counter = 0;
  let theIssue;
  while (!finished){
    counter += 1;
    if (counter > 10){finished = true}
    theIssue = await fixNextIssue()
    if ((theIssue.issue == issueType.finished) || (theIssue.rowIndex == -1)){
      finished = true;
    } else if (theIssue.issue == issueType.fixLf) {
      console.log(counter + ': Doing ' + theIssue.issue + ' on rowIndex ' + theIssue.rowIndex);
      let success = await correctTextReplaceLF(true);
      console.log('success', success)
      finished = !success; 
    } else if (theIssue.issue == issueType.fixSpaceQuote) {
      console.log(counter + ': Doing ' + theIssue.issue + ' on rowIndex ' + theIssue.rowIndex);
      let success = await correctTextSpaceQuotes(true);
      finished = !success;
    } else {
      console.log(counter + ': Unexpected issue: ' + theIssue.issue + ' on rowIndex ' + theIssue.rowIndex);
      finished = true;
    }
  }
}

async function autoSelectChapter(){
  await Excel.run(async (excel) => {
    const resultSheet = excel.workbook.worksheets.getItem('Result');
    const chapterRange = resultSheet.getRange('reChapter');
    chapterRange.load('values');
    await excel.sync();
    console.log('chapter', chapterRange.values[0][0]);
    const chapter = chapterRange.values[0][0];
    let chapterNum = parseInt(jade_modules.wordtonumbers.text2num(chapter.toLowerCase()));
    console.log('Chapter Num', chapterNum);
    if (!isNaN(chapterNum)){
      let chapterCompareSelect = tag('chapter-compare-select');
      chapterCompareSelect.value = chapterNum;
    }
  })
}

async function doKeepsAndManuals(){
  await Excel.run(async (excel) => {
    const decisionSheet = excel.workbook.worksheets.getItem('Decision');
    let keepRange = decisionSheet.getRange('deKeep');
    let keepCalulationRange = decisionSheet.getRange('deKeepCalculation');
    let manualRange = decisionSheet.getRange('deManual');
    keepCalulationRange.load('values, rowIndex');
    manualRange.load('values, rowIndex');
    keepRange.load('rowIndex, columnIndex');
    keepRange.clear('contents');
    await excel.sync();
    let indexes = []
    for (let i = 0; i < keepCalulationRange.values.length; i++){
      if (keepCalulationRange.values[i][0] == 'Keep'){
        if (manualRange.values[i][0] != 'Override'){
          indexes.push(i + keepCalulationRange.rowIndex);
        }
      } else if (manualRange.values[i][0] == 'Manual'){
        indexes.push(i + keepCalulationRange.rowIndex);
      } else if (keepCalulationRange.values[i][0] == 'End'){
        break;
      }
    }
    console.log('Indexes', indexes);
    let tempRange = [];
    for (let i = 0; i < indexes.length; i++){
      tempRange[i] = decisionSheet.getRangeByIndexes(indexes[i], keepRange.columnIndex, 1, 1);
      tempRange[i].values = [['Keep']];
    }
    await excel.sync();
  })
}

