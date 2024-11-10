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

async function createChapters(){
  const details = await getRowColumnDetails();
  let chapters = [];
  let index = -1;
  textSoFar = '';
  const apostrophes = await apostropheWords();
  await Excel.run(async function(excel){  
    const pdfSheet = excel.workbook.worksheets.getItem(pdfComparisonSheetName);
    const rowCount = details.rowCount - details.rowIndex + 1 - startRowIndex;
    const sourceRange = pdfSheet.getRangeByIndexes(startRowIndex, sourceColumnIndex, rowCount, 1);
    sourceRange.load('rowIndex, values');
    await excel.sync();
    //console.log('sourceRange', sourceRange.values)
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
    chapterValues = chapters.map(x => [x]);
    console.log('chapterValues', chapterValues);
    /*
    console.log(startRowIndex, chaptersColumnIndex, chapterValues.length, 1);
    let chapterRange = pdfSheet.getRangeByIndexes(startRowIndex, chaptersColumnIndex, chapterValues.length, 1);
    chapterRange.clear('Contents');
    chapterRange.values = chapterValues;
*/

    let myLines = chapters[0].split("\n");
    for (let i = 0; i < myLines.length; i++){
      if (myLines[i].startsWith("'") && (!myLines[i].startsWith["''"])){
        myLines[i] = "'" + myLines[i];
      }
    }
    lineValues = myLines.map(x => [x]);
    console.log(myLines);

    let lineRange = pdfSheet.getRangeByIndexes(startRowIndex, linesColumnIndex, lineValues.length, 1);
    lineRange.clear('Contents');
    lineRange.values = lineValues;
    
    //Now search for curly quotes within each line
    let quoteData = [];
    let quoteIndex = -1;
    for (let i = 0; i < myLines.length; i++){
      let openQuote = findCurlyQuote('‘', myLines[i], false, apostrophes);
      let closeQuote = findCurlyQuote('’', myLines[i], true, apostrophes);
      quoteIndex += 1;
      if ((openQuote.length > 0) || (closeQuote.length > 0)){
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
    console.log('quoteData', quoteData);
    for (let i = 0; i < quoteData.length; i++){
      if (quoteData[i].subStrings.length > 0){
        console.log('i',i, quoteData[i].subStrings)
      }
    }
    await displayDecision(quoteData);
  })
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

async function displayDecision(quoteData){
  const lineIndex = 0;
  const textIndex = 1;
  const subNoIndex = 2
  const substringIndex = 3;
  const startIndex = 4;
  const endIndex = 5;

  await Excel.run(async function(excel){
    let displayRange = excel.workbook.worksheets.getItem('Decision').getRange('deTable');
    displayRange.clear('Contents');
    displayRange.load('rowIndex, rowCount, columnIndex, columnCount');
    await excel.sync();
    let rowIndex = -1;
    let display = [];
    for (let i = 0; i < quoteData.length; i++){
      rowIndex += 1;
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
          display[rowIndex][lineIndex] = '';
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
    let tempRange = excel.workbook.worksheets.getItem('Decision').getRangeByIndexes(displayRange.rowIndex, displayRange.columnIndex, display.length, displayRange.columnCount);
    tempRange.values = display;
    
  })
  
}