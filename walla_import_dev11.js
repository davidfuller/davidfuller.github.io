const wallaSheetName = 'Walla Import';
const sourceTextRangeName = 'wiSource';
const namedCharacters = 'Named Characters - For reaction sounds and walla';
const wallaTableName = 'wiTable';
const tableCols ={
  wallaOriginal: 0,
  lineRange: 1,
  typeOfWalla: 2,
  character: 3,
  description: 4,
  numCharacters: 5,
  lineNo: 6
}


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
  const textLines = ['lines', 'line']
  let firstLine;
  let lineNo = -1;
  let theSections = theLine.split('-');
  let theCharacter = theSections[0].trim();
  let individualCharacters = theCharacter.split(',')
  let thePosition = '';
  if (!(theSections[1] === undefined)){
    thePosition = theSections[1].trim()
  }
  let wholeScene = thePosition.toLowerCase().indexOf('whole scene')
  for (let i = 0; i < textLines.length; i++){
    if (thePosition.toLowerCase().includes(textLines[i])){
      firstLine = thePosition.toLowerCase().indexOf(textLines[i])
      if (firstLine != -1){
        lineNo = parseInt(thePosition.substring(firstLine + textLines[i].length));
        break;
      }
    }
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
    await excel.sync();

    const sortFields = [
      {
        key: 6, //Line No
        ascending: true
      },
      {
        key: 0, // Walla Original
        ascending: true
      }
    ]
    displayRange.sort.apply(sortFields);
  })
}

async function loadIntoScriptSheet(){
  await Excel.run(async (excel) => {
    let loadMessage = tag('load-message');
    loadMessage.style.display = 'none';
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

    let lineNo = wallaTableRange.values[arrayRow][tableCols.lineNo];
    let myRowIndex = await jade_modules.operations.getLineNoRowIndex(lineNo)
    console.log('row Index', myRowIndex);
    let wallaData = {
      wallaLineRange: wallaTableRange.values[arrayRow][tableCols.lineRange],
      typeOfWalla: wallaTableRange.values[arrayRow][tableCols.typeOfWalla],
      characters: wallaTableRange.values[arrayRow][tableCols.character],
      description: wallaTableRange.values[arrayRow][tableCols.description],
      numCharacters: wallaTableRange.values[arrayRow][tableCols.numCharacters],
      all: wallaTableRange.values[arrayRow][tableCols.wallaOriginal]
    }
    console.log('Walla Data', wallaData);
    await jade_modules.operations.createWalla(wallaData, myRowIndex, false, true)
  })
}

async function loadMultipleIntoScriptSheet(doAll){
  let wallaData = [];
  await Excel.run(async (excel) => {
    const wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
    const wallaTableRange = wallaSheet.getRange(wallaTableName);
    wallaTableRange.load('rowIndex, rowCount, values');
    await excel.sync()
    const tableRowFirst = wallaTableRange.rowIndex;
    const tableRowLast = wallaTableRange.rowIndex + wallaTableRange.rowCount - 1;
    let rowIndexes = [];
    if (doAll){
      for (let i = tableRowFirst; i <= tableRowLast; i++){
        rowIndexes.push(i);
      }
    } else {
      const selectedRanges = excel.workbook.getSelectedRanges();
      selectedRanges.load('address');
      selectedRanges.areas.load('items');
      await excel.sync();
      console.log('selectedRange address', selectedRanges.address)
      let ranges = selectedRanges.areas.items;
      console.log(ranges)
      
      for (i = 0; i < ranges.length; i++){
        ranges[i].load('address', 'rowIndex', 'rowCount')
        await excel.sync();
        console.log(ranges[i].address);
        for (let j = 0; j < ranges[i].rowCount; j++){
          rowIndexes.push(ranges[i].rowIndex + j);
        }
      }
    }

    for (let i = 0; i < rowIndexes.length; i++){
      if ((rowIndexes[i] >= tableRowFirst) && (rowIndexes[i] <= tableRowLast)){
        let tableRow = rowIndexes[i] - tableRowFirst;
        let lineNo = wallaTableRange.values[tableRow][tableCols.lineNo];
        if (lineNo > 0){
          let data = {};
          data.rowIndex = await jade_modules.operations.getLineNoRowIndex(lineNo);
          data.wallaLineRange = wallaTableRange.values[tableRow][tableCols.lineRange];
          data.typeOfWalla = wallaTableRange.values[tableRow][tableCols.typeOfWalla];
          data.characters =wallaTableRange.values[tableRow][tableCols.character];
          data.description = wallaTableRange.values[tableRow][tableCols.description];
          data.numCharacters = wallaTableRange.values[tableRow][numCharacters];
          data.all = wallaTableRange.values[tableRow][tableCols.wallaOriginal];
          data.lineNo = lineNo;
          wallaData.push(data); 
        }
      }
    }
    console.log('wallaData', wallaData);
    await jade_modules.operations.createMultipleWallas(wallaData, false, true);
  })
}

async function showWallaLineNo(){
  await Excel.run(async (excel) => {
    let lineNo;
    const wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
    const wallaTableRange = wallaSheet.getRange(wallaTableName);
    wallaTableRange.load('rowIndex, rowCount, values');
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex');
    await excel.sync();
    if (isRowWithinTable(activeCell.rowIndex, wallaTableRange.rowIndex, wallaTableRange.rowCount)){
      lineNo = wallaTableRange.values[activeCell.rowIndex - wallaTableRange.rowIndex][tableCols.lineNo];
      await jade_modules.operations.showWallaLine(lineNo);
    } else {
      alert('please select a valid line number')
    }
  })
}

function isRowWithinTable(rowIndex, tableRowIndex, tableRowCount){
  return (rowIndex >= tableRowIndex) && (rowIndex < (tableRowIndex + tableRowCount))
}