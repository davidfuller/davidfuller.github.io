const pdfComparisonSheetName = 'PDF Comparison';
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
      let openQuote = findCurlyQuote('‘', myLines[i]);
      let closeQuote = findCurlyQuote('’', myLines[i]);
      if ((openQuote.length > 0) || (closeQuote.length > 0)){
        quoteIndex += 1;
        let theData = {
          index: i,
          text: myLines[i],
          openQuote: openQuote,
          closeQuote: closeQuote
        }
        quoteData[quoteIndex] = theData;
      }
    }
    console.log('quoteData', quoteData);
  })
}

function findCurlyQuote(character, myString){
  let index = 0
  let result = []
  let position = myString.indexOf(character, index)
  while (position != -1){
    result.push(position);
    index = position + 1;
    position = myString.indexOf(character, index)
  }
  
  return result;
}
/*
‘Lying there with their eyes wide open! Cold as ice! Still in their dinner things!’
*/