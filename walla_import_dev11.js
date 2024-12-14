const wallaSheetName = 'Walla Import';
const sourceTextRangeName = 'wiSource';

const wallaTableName = 'wiTable';
const wallaSourceSheetName = 'Walla Script'
const wallaSourceUKScriptColumnIndex = 9;
const wallaImportName = 'Walla Import';

const tableCols ={
  wallaOriginal: 0,
  lineRange: 1,
  typeOfWalla: 2,
  character: 3,
  description: 4,
  numCharacters: 5,
  lineNo: 6,
  rowIndex: 7,
  scene: 8
}

const wallaTypes = {
  named: 'named',
  unNamed: 'unNamed',
  general: 'general'
}


const namedCharacters = ['Named Characters - For reaction sounds and walla', 'Named Characters - For reaction sounds and walla:', 'Named Characters Reactions and Walla', 'Named character walla', 'Named - Character & Reactions', 
  'Named character walla:', 'Named character walla']
let displayWallaName = 'Named Characters Reactions and Walla:'
const unnamedCharacters = ['Un-named Character Walla','Un-named Character Walla:'];
let displayWallaUnNamed = 'Un-named Character Walla:';
const generalWalla = ['General Walla', 'General Walla:']
let displayGeneralWalla = 'General Walla:';

function isNamedWalla(theType){
  for (text of namedCharacters){
    if (theType.trim().toLowerCase() == text.trim().toLowerCase()){
      return true;
    }
  }
  return false;
}
function isUnamedWalla(theType){
  for (text of unnamedCharacters){
    if (theType.trim().toLowerCase() == text.trim().toLowerCase()){
      return true;
    }
  }
  return false;
}

function isGeneralWalla(theType){
  for (text of generalWalla){
    if (theType.trim().toLowerCase() == text.trim().toLowerCase()){
      return true;
    }
  }
  return false;
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
    let scenes = [];
    let anyNonScenes = false;
    for (let i = 0; i < theResults.length; i++){
      let rowAndScene = await jade_modules.operations.getLineNoRowIndexAndScene(theResults[i].line);
      console.log(i, 'rowAndScene', rowAndScene);
      if (rowAndScene.scene == -1){
        anyNonScenes = true
      } else {
        scenes.push(rowAndScene.scene)
      }
      resultArray[i] = []
      resultArray[i][0] = theResults[i].all;
      resultArray[i][1] = theResults[i].lineRange;
      resultArray[i][2] = getDisplayWallaName(typeWalla);
      resultArray[i][3] = theResults[i].character;
      resultArray[i][4] = theResults[i].description;
      resultArray[i][5] = theResults[i].numCharacters;
      resultArray[i][6] = theResults[i].line;
      resultArray[i][7] = rowAndScene.rowIndex;
      resultArray[i][8] = rowAndScene.scene;
    }

    scenes = [...new Set(scenes)]
    console.log('anyNonScenes', anyNonScenes, 'scenes', scenes)
    if ((anyNonScenes) && (scenes.length == 1)){
      rowLineDetails = await jade_modules.operations.getRowIndexLineNoFirstLineScene(scenes[0])
      if ((rowLineDetails.lineNo != -1) && (rowLineDetails.rowIndex != -1)){
        for (let i = 0; i < resultArray.length; i++){
          if (resultArray[i][6] == -1){
            resultArray[i][6] = rowLineDetails.lineNo;
          }
          if (resultArray[i][7] == -1){
            resultArray[i][7] = rowLineDetails.rowIndex;
          }
          if (resultArray[i][8] == -1){
            resultArray[i][8] = scenes[0];
          }
        }
      }
    }
    if (resultArray.length > 0){
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
    }
    if (scenes.length == 1){
      const sceneWallaIndexColumn = 4;
      const sourceRowIdRange = wallaSheet.getRange('wiSourceRowId');
      const indexTableRange = wallaSheet.getRange('wiWallaIndexTable');
      sourceRowIdRange.load('values');
      indexTableRange.load('rowIndex, columnIndex')
      await excel.sync();
      let sourceRowId = sourceRowIdRange.values[0][0]
      console.log('sourceRow', sourceRowId, 'scenes', scenes[0]);
      console.log(indexTableRange.rowIndex + sourceRowId - 1, indexTableRange.columnIndex + sceneWallaIndexColumn, 1, 1)
      let sceneRange = wallaSheet.getRangeByIndexes(indexTableRange.rowIndex + sourceRowId - 1, indexTableRange.columnIndex + sceneWallaIndexColumn, 1, 1)
      sceneRange.load('address');
      await excel.sync();
      console.log('address', sceneRange.address)
      sceneRange.values = [[scenes[0]]];
      await excel.sync();
    }
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
          data.rowIndex = wallaTableRange.values[tableRow][tableCols.rowIndex];
          data.wallaLineRange = wallaTableRange.values[tableRow][tableCols.lineRange];
          data.typeOfWalla = wallaTableRange.values[tableRow][tableCols.typeOfWalla];
          data.characters =wallaTableRange.values[tableRow][tableCols.character];
          data.description = wallaTableRange.values[tableRow][tableCols.description];
          data.numCharacters = wallaTableRange.values[tableRow][tableCols.numCharacters];
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

function getDisplayWallaName(theType){
  if (isNamedWalla(theType)){
    return displayWallaName;
  } else if (isUnamedWalla(theType)){
    return displayWallaUnNamed;
  } else if (isGeneralWalla(theType)){
    return displayGeneralWalla;
  } else {
    return theType;
  }
}

async function getTheWallaSourceIndecies(){
  let wallaIndexes = []
  let named = 0;
  let unNamed = 0;
  let general = 0;
  let isGood = true;
  await Excel.run(async (excel) => {
    const sourceSheet = excel.workbook.worksheets.getItem(wallaSourceSheetName);
    const usedRange = sourceSheet.getUsedRange();
    usedRange.load('rowIndex, rowCount');
    await excel.sync();
    console.log('rowIndex', usedRange.rowIndex, 'rowCount', usedRange.rowCount);
    let scriptRange = sourceSheet.getRangeByIndexes(usedRange.rowIndex, wallaSourceUKScriptColumnIndex, usedRange.rowCount, 1)
    scriptRange.load('values');
    await excel.sync()
    for (let i = 0; i < scriptRange.values.length; i++){
      let raw = scriptRange.values[i][0];
      let lines = raw.split('\n');
      console.log(i, lines[0]);
      let wallaData = null;
      if (isNamedWalla(lines[0])){
        wallaData = {
          type: wallaTypes.named,
          rowIndex: i + usedRange.rowIndex
        }
        named += 1;
      } else if (isUnamedWalla(lines[0])){
        wallaData = {
          type: wallaTypes.unNamed,
          rowIndex: i + usedRange.rowIndex
        }
        unNamed += 1;
      } else if (isGeneralWalla(lines[0])){
        wallaData = {
          type: wallaTypes.general,
          rowIndex: i + usedRange.rowIndex
        }
        general += 1;
      }
      if (wallaData != null){
        wallaIndexes.push(wallaData);
      }
    }
    console.log('Walla Idndexes', wallaIndexes);
    console.log('named', named, 'unNamed', unNamed, 'general', general);
    for (i = 0; i < wallaIndexes.length; i++){
      if (i % 3 == 0){
        if (wallaIndexes[i].type != wallaTypes.named){
          console.log(i, 'Named', wallaIndexes[i].rowIndex, wallaIndexes[i].type);
          isGood = false;
          break;
        }
      }
      if (i % 3 == 1){
        if (wallaIndexes[i].type != wallaTypes.unNamed){
          console.log(i, 'Unnamed', wallaIndexes[i].rowIndex, wallaIndexes[i].type);
          isGood = false;
          break;
        }
      }
      if (i % 3 == 2){
        if (wallaIndexes[i].type != wallaTypes.general){
          console.log(i, 'General', wallaIndexes[i].rowIndex, wallaIndexes[i].type);
          isGood = false;
          break;
        }
      }
    }
  })
  if (isGood){
    await displayWallaIndexes(wallaIndexes);
    return wallaIndexes;
  }
}

async function displayWallaIndexes(wallaIndexes){
  await Excel.run(async (excel) => {
    const wallaSheet = excel.workbook.worksheets.getItem(wallaImportName);
    const indexTableRange = wallaSheet.getRange('wiWallaIndexTable');
    indexTableRange.load('rowIndex, rowCount, columnIndex, columnCount');
    indexTableRange.clear('Contents');
    await excel.sync();
    let num = 0;
    let results = []
    for (i = 0; i < wallaIndexes.length; i = i + 3){
      num += 1;
      let myRow = [num, wallaIndexes[i].rowIndex, wallaIndexes[i + 1].rowIndex, wallaIndexes[i + 2].rowIndex]
      results.push(myRow)
    }
    console.log('results', results)
    let tempRange = wallaSheet.getRangeByIndexes(indexTableRange.rowIndex, indexTableRange.columnIndex, results.length, indexTableRange.columnCount);
    tempRange.values = results;
  })
}

async function loadSelectedCellIntoTextBox(){
  await Excel.run(async (excel) => {
    const wallaSheet = excel.workbook.worksheets.getItem(wallaImportName);
    const indexTableRange = wallaSheet.getRange('wiWallaIndexTable');
    indexTableRange.load('rowIndex, rowCount, columnIndex, columnCount');
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex, columnIndex, values');
    await excel.sync();
    let topRow = indexTableRange.rowIndex;
    let bottomRow = indexTableRange.rowIndex + indexTableRange.rowCount - 1;
    let leftColumn = indexTableRange.columnIndex + 1 // not the first column
    let rightColumn = indexTableRange.columnIndex + indexTableRange.columnCount - 1; 
    if ((activeCell.rowIndex >= topRow) && (activeCell.rowIndex <= bottomRow) && (activeCell.columnIndex >= leftColumn) && (activeCell.columnIndex <= rightColumn)){
      let testRowIndex = activeCell.values[0][0];
      await loadTextBox(testRowIndex);
      const rowIdRange = wallaSheet.getRange('wiSourceRowId');
      rowIdRange.values = [[activeCell.rowIndex - indexTableRange.rowIndex + 1]];
    } else {
      alert('Not in table');
    }
  })
}
async function loadTextBox(rowIndex){
  await Excel.run(async (excel) => {
    if (!isNaN(parseInt(rowIndex))){
      const wallaSourceSheet = excel.workbook.worksheets.getItem(wallaSourceSheetName);
      const wallaSheet = excel.workbook.worksheets.getItem(wallaImportName);
      let testRange = wallaSourceSheet.getRangeByIndexes(rowIndex, wallaSourceUKScriptColumnIndex, 1, 1);
      testRange.load('values');
      const sourceRowIndexRange = wallaSheet.getRange('wiSourceRowIndex');
      await excel.sync();
      console.log(testRange.values[0][0]);
      let wallaText = testRange.values[0][0];
      let textRange = wallaSheet.getRange('wiSource');
      textRange.values = [[wallaText.trim()]];  
      sourceRowIndexRange.values = [[rowIndex]];
    }
  })
}
async function loopThroughTheIndexes(){
  await Excel.run(async (excel) => {
    const wallaSheet = excel.workbook.worksheets.getItem(wallaImportName);
    const indexTableRange = wallaSheet.getRange('wiWallaIndexTable');
    indexTableRange.load('rowIndex, values');
    await excel.sync();
    for (let i = 0; i < indexTableRange.values.length; i++){
      let rowIndex = indexTableRange.values[i][1];
      console.log(i, 'rowIndex', rowIndex);
      await loadTextBox(rowIndex);
      const rowIdRange = wallaSheet.getRange('wiSourceRowId');
      rowIdRange.values = [[indexTableRange.values[i][0]]];
      await excel.sync();
      await parseSource();
    }
  })
}
