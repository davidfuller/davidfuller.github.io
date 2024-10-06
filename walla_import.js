const wallaSheetName = 'Walla Import';
const sourceTextRangeName = 'wiSource';
const namedCharacters = 'Named Characters - For reaction sounds and walla';
const wallaTableName = 'wiTable';

function auto_exec(){
}

async function parseSource(){
  await Excel.run(async (excel) => {
    let wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
    let sourceRange = wallaSheet.getRange(sourceTextRangeName);
    sourceRange.load('values')
    await excel.sync();
    let mySourceText = sourceRange.values[0][0];
    let theLines = mySourceText.split('\n');
    let theResults = [];
    for (let i = 1; i < theLines.length; i++){
      theResults[i - 1] = splitLine(theLines[i]);
    }
    console.log('theResults', theResults)
    await doWallaTable(theLines[0], theResults)
  })
}

/* where to put the data
    Walla cue Number - auto calculated starting at W00001

    Walla line range - either 'whole scene' or the best guess of line range
    Type of Walla - top line
    Walla characters - character
    Walla description - the description
    Num of characters - count the characters
    walla original - All
*/
function splitLine(theLine){
  //first split with '-'
  let theSections = theLine.split('-');
  let theCharacter = theSections[0].trim();
  let individualCharacters = theCharacter.split(',')
  
  let thePosition = theSections[1].trim()
  let wholeScene = thePosition.toLowerCase().indexOf('whole scene')
  let firstLine = thePosition.toLowerCase().indexOf('line')
  let lineNo = -1;
  if (firstLine != -1){
    lineNo = parseInt(thePosition.substring(firstLine + 4));
  }
  let theRestPosition = theLine.toLowerCase().indexOf(thePosition.toLowerCase());
  let theRest = '';
  if (theRestPosition != -1){
    theRest = theLine.substring(theRestPosition);
  }
  let lastBit = theSections[theSections.length - 1];
  

  let theDescription;
  let lastBitPosition;
  let lineRange;
  if (isNaN(parseInt(lastBit))){
    theDescription = lastBit.trim();
    lastBitPosition = theLine.toLowerCase().indexOf(lastBit.toLowerCase());
    lineRange = theLine.substring(theRestPosition, lastBitPosition - 2).trim() ;
  } else {
    theDescription = 'N/A'
    lineRange = theRest;
  }
  
  result = {
    all: theLine,
    character: theCharacter,
    wholeScene: (wholeScene != -1),
    line: lineNo,
    rest: theRest,
    description: theDescription,
    lineRange:  lineRange,
    numCharacters: individualCharacters.length
  }
  return result;

}

async function doWallaTable(typeWalla, theResults){
  await Excel.run(async (excel) => {
    let wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
    let wallaTable = wallaSheet.getRange(wallaTableName);
    wallaTable.load('rowIndex, rowCount, columnIndex, columnCount, address');
    wallaTable.clear("Contents");
    await excel.sync();
    console.log(wallaTable.address, wallaTable.rowCount);
    console.log(typeWalla, theResults);
    let resultArray = []
    for (let i = 0; i < theResults.length; i++){
      resultArray[i] = []
      resultArray[i][0] = theResults[i].all;
      resultArray[i][1] = theResults[i].lineRange;
      resultArray[i][2] = typeWalla;
      resultArray[i][3] = theResults[i].character;
      resultArray[i][4] = theResults[i].description;
      resultArray[i][5] = theResults[i].numCharacters;
      resultArray[i][6] = theResults[i].line;
    }
    let displayRange = wallaSheet.getRangeByIndexes(wallaTable.rowIndex, wallaTable.columnIndex, resultArray.length, wallaTable.columnCount);
    displayRange.load('rowCount, columnCount');
    await excel.sync();
    console.log(resultArray)
    console.log('Display Range rows: ', displayRange.rowCount, 'columns: ', displayRange.columnCount);

    displayRange.values = resultArray;
    await excel.sync()
  })
}

async function loadIntoScriptSheet(){
  await Excel.run(async (excel) => {
    const lineNoArrayColumn = 6;
    const lineRangeArrayColumn = 1;
    const typeOfWallaArrayColumn = 2;
    const characterArrayColumn = 3;
    const descriptionArrayColumn = 4;
    const numCharactersArrayColumn = 5;
    const allArrayColumn = 0;
    let wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
    let wallaTableRange = wallaSheet.getRange(wallaTableName);
    wallaTableRange.load('rowIndex');
    wallaTableRange.load('rowCount');
    wallaTableRange.load('values');
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load("rowIndex");
    activeCell.load(("columnIndex"))
    await excel.sync();

    let arrayRow = activeCell.rowIndex - wallaTableRange.rowIndex
    console.log('array row: ', arrayRow, 'data: ', wallaTableRange.values[arrayRow]);

    let lineNo = wallaTableRange.values[arrayRow][lineNoArrayColumn];
    let myRowIndex = await jade_modules.operations.getLineNoRowIndex(lineNo)
    console.log('row Index', myRowIndex);
    let wallaData = {
      wallaLineRange: wallaTableRange.values[arrayRow][lineRangeArrayColumn],
      typeOfWalla: wallaTableRange.values[arrayRow][typeOfWallaArrayColumn],
      characters: wallaTableRange.values[arrayRow][characterArrayColumn],
      description: wallaTableRange.values[arrayRow][descriptionArrayColumn],
      numCharacters: wallaTableRange.values[arrayRow][numCharactersArrayColumn],
      all: wallaTableRange.values[arrayRow][allArrayColumn]
    }
    console.log('Walla Data', wallaData);
    await jade_modules.operations.createWalla(wallaData, myRowIndex, false, true)
  })
}