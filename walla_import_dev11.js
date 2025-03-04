const wallaSheetName = 'Walla Import';
const sourceTextRangeName = 'wiSource';

const wallaTableName = 'wiTable';
const wallaSourceSheetName = 'Walla Script'
const wallaSourceUKScriptColumnIndex = 9;
const wallaImportName = 'Walla Import';

const wallaSourceWallaColumn = [
  {
    book: 'Book 1',
    column: 5,
    scriptColumn: 9
  },
  {
    book: 'Book 2',
    column: 5,
    scriptColumn: 9
  },
  {
    book: 'Book 3',
    column: 5,
    scriptColumn: 9
  },
  {
    book: 'Book 4',
    column: 9,
    scriptColumn: 9
  }
]

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

const wallaScriptColumns = {
  cue: 4,
  character: 6,
  presentCharacters: 7,
  stageDirection: 8,
  columnCount: 6
}

const wallaScriptingNames = ['WALLA SCRIPTED', 'WALLA SCRIPTED LINES', 'WALLA SCRIPTED LINES - SCOLDING CARRYING ON', 'WALLA SCRIPTING', 'WALLA SCRIPTING', 'WALLA SCRIPTING - lines to lead into the scripted argument', 'SCRIPTED WALLA LINES', 'WALLA SCEIPTED LINES']

const namedCharacters = ['Named Characters - For reaction sounds and walla', 'Named Characters - For reaction sounds and walla:', 'Named Characters Reactions and Walla', 'Named character walla', 'Named - Character & Reactions', 
  'Named character walla:', 'Named character walla', 'Named Characters Reactions and Walla:', 'Named Characters for Reaction Sounds & Walla:', 'Named Characters for reaction sounds and Walla:', 'Named Chsracters for reaction sounds and walla:', 
  'Named characters for reaction sounds and walla', 'Characters for reaction sounds and walla', 'Named characters for. reaction sounds and walla:']
let displayWallaName = 'Named Characters Reactions and Walla:'
const unnamedCharacters = ['Un-named Character Walla','Un-named Character Walla:', 'Un-named Character Walla: None', 'Unnamed Character Walla:', 'Un-named characters:', 'Un-named character walla :', 'Un-named character walla : none', 'Un-named characters: none'];
let displayWallaUnNamed = 'Un-named Character Walla:';
const generalWalla = ['General Walla', 'General Walla:', 'General Walla: None']
let displayGeneralWalla = 'General Walla:';

function isNamedWalla(theType){
  console.log(theType)
  for (text of namedCharacters){
    //console.log(text);
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

async function parseSource(tableRowIndex = -1){
  const replacements = await wallaReplacementWords();
  console.log('replacements', replacements)
  await Excel.run(async (excel) => {
    let wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
    let sourceRange = wallaSheet.getRange(sourceTextRangeName);
    sourceRange.load('values')
    await excel.sync();
    let mySourceText = replaceReplacements(sourceRange.values[0][0],replacements);
    let theLines = mySourceText.split('\n');
    let theResults = [];
    for (let i = 1; i < theLines.length; i++){
      if ((theLines[i].trim() != '') && (theLines[i].trim().toLowerCase() != 'none') && (theLines[i].trim().toLowerCase() != 'n/a')){
        theResults.push(splitLine(theLines[i]));
      }
    }
    console.log('theResults', theResults)
    await doWallaTable(theLines[0], theResults, tableRowIndex);
    //await doWallaTableV2()
  })
}

async function parseSourceText(sourceText){
  console.log('sourceText', sourceText);
  let mySourceText = sourceText;
  let theLines = mySourceText.split('\n');
  let theResults = [];
  for (let i = 1; i < theLines.length; i++){
    let trimmedLine = theLines[i].trim();
    if (trimmedLine != ''){
      if (!(trimmedLine.toLowerCase().startsWith('none'))){
        if (!(trimmedLine.toLowerCase().startsWith('n/a'))){
          theResults.push(splitLine(theLines[i]));
        }
      }
    }
  }
  return {
    typeWalla: theLines[0],
    theResults: theResults
  }
  
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
  let stageOneUsed = false;
  // Deal with a raw line number being in first section
  console.log('Stage 1 theSections:', theSections);
  if (!isNaN(parseInt(theSections[0]))){
    if (!isNaN(parseInt(theSections[1]))){
      let tempArray = [];
      tempArray[0] = theSections[2]
      tempArray[1] = "Line " + parseInt(theSections[0]) + '-' + parseInt(theSections[1]);
      for (let i = 3; i < theSections.length; i++){
        tempArray[i - 1] = theSections[i];
      }
      console.log('tempArray', tempArray, 'theSections', theSections);
      theSections = tempArray;
    } else {
      let tempLineNo = parseInt(theSections[0]);
      if (theSections.length > 1){
        theSections[0] = theSections[1];
        theSections[1] = "Line " + tempLineNo;
      } else {
        theSections[0] = "Line " + tempLineNo;
      }
    }
    stageOneUsed =true;
  }
  console.log('Stage 2 theSections:', theSections, 'stageOneUsed', stageOneUsed);
  if (!stageOneUsed){
    // Deal with a line number being in first section
    for (let i = 0; i < textLines.length; i++){
      if (theSections[0].toLowerCase().includes(textLines[i])){
        if (theSections.length > 1){
          let theSwap = theSections[0];
          theSections[0] = theSections[1];
          theSections[1] = theSwap;
        }
        break;
      }
    }
  }
  console.log('Stage 3 theSections:', theSections)
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
  
  console.log('At this point, theLine', theLine, 'thePosition', thePosition, 'wholeScene', wholeScene, 'firstLine', firstLine, 'lineNo', lineNo, 'theRestPosition', theRestPosition);

  let theRest = '';
  if (theRestPosition != -1){
    theRest = theLine.substring(theRestPosition);
  }
  console.log('theSections', theSections);



  let lastBit = theSections[theSections.length - 1];
  
  console.log('At this point 2, theLine', theLine, 'theRest', theRest, 'lastBit', lastBit);

  let theDescription;
  let lastBitPosition;
  let lineRange;
  if (isNaN(parseInt(lastBit))){
    if ((lineNo == -1) && (wholeScene != -1)){
      console.log('I am here in first if');
      if (lastBit.trim().toLowerCase() != 'whole scene'){
        console.log('I am here in second if');
        theDescription = lastBit.trim();
      } else {
        console.log('I am here in second else');
        theDescription = '';
      }
      lineRange = 'whole scene'
    } else if ((lineNo > 0) && (wholeScene == -1)){
      if (lastBit.trim().toLowerCase() == thePosition.trim().toLowerCase()){
        console.log('I am here in third if');
        theDescription = '';
        lineRange = thePosition;
      } else {
        console.log('I am here in third else');
        theDescription = lastBit.trim();
        if (theDescription == ''){
          lineRange = theRest.trim()
          if (lineRange.endsWith('-')){
            lineRange = lineRange.substring(0, lineRange.length - 2).trim();
          }
        } else {
          lastBitPosition = theLine.toLowerCase().indexOf(lastBit.toLowerCase());
          console.log('theRestPosition', theRestPosition, 'lastBitPosition', lastBitPosition);
          lineRange = theLine.substring(theRestPosition, lastBitPosition - 1).trim();
        }
      }
      
    } else {
      console.log('I am here in final else');
      theDescription = lastBit.trim();
      lastBitPosition = theLine.toLowerCase().indexOf(lastBit.toLowerCase());
      console.log('theRestPosition', theRestPosition, 'lastBitPosition', lastBitPosition);
      lineRange = theLine.substring(theRestPosition, lastBitPosition - 1).trim();
      if (!isNaN(parseInt(lineRange))){
        lineNo = parseInt(lineRange)
        lineRange = "Line " + lineRange
      }
    }
    console.log('At this point 3, theLine', theLine, 'theDescription', theDescription, 'lineRange', lineRange);
  } else {
    theDescription = ''
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

async function doWallaTable(typeWalla, theResults, tableRowIndex = -1){
  await Excel.run(async (excel) => {
    let wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
    let wallaTable = wallaSheet.getRange(wallaTableName);
    wallaTable.load('rowIndex, rowCount, columnIndex, columnCount, address');
    wallaTable.clear("Contents");
    await excel.sync();
    let sourceRowId = tableRowIndex;
    if (tableRowIndex == -1){
      const sourceRowIdRange = wallaSheet.getRange('wiSourceRowId');
      sourceRowIdRange.load('values');
      await excel.sync();
      sourceRowId = sourceRowIdRange.values[0][0];
    }
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
      console.log(i, 'line range', theResults[i].lineRange);
      if (theResults[i].lineRange.trim() == ''){
        theResults[i].lineRange = 'whole scene';
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
    if (theResults.length == 0){
      let display = getDisplayWallaName(typeWalla);
      scenes[0] = await getScene(sourceRowId, false);
      console.log('sourceRowId', sourceRowId, 'scene', scenes);
      anyNonScenes = true;
      resultArray[0] = [];
      resultArray[0][0] = display;
      resultArray[0][1] = 'Whole Scene';
      resultArray[0][2] = display;
      resultArray[0][3] = '';
      resultArray[0][4] = ''
      resultArray[0][5] = 0;
      resultArray[0][6] = -1;
      resultArray[0][7] = -1;
      resultArray[0][8] = scenes[0];
    }

    scenes = [...new Set(scenes)]
    if (scenes.length == 0){
      scenes[0] = await getScene(sourceRowId, true) + 1;
      console.log('In doWallaTable scenes[0]', scenes[0], 'sourceRowId', sourceRowId);
      if ((scenes[0] == 0) && (sourceRowId == 1)){
        let firstScene = parseInt(await jade_modules.operations.getFirstScene()) + 1;
        console.log('In doWallaTable firstScene', firstScene);
        if (!isNaN(firstScene)){
          scenes[0] = firstScene
        }
      }
    }
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
    if ((scenes.length == 1) && (isNamedWalla(typeWalla))){
      const sceneWallaIndexColumn = 4;
      const indexTableRange = wallaSheet.getRange('wiWallaIndexTable');
      indexTableRange.load('rowIndex, columnIndex')
      await excel.sync();
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

async function doWallaTableV2(typeWalla, theResults, scene){
  let wallaData = [];
  let resultArray = [];
  if (theResults.length > 0){
    await Excel.run(async (excel) => {
      let wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
      let wallaTable = wallaSheet.getRange(wallaTableName);
      wallaTable.load('rowIndex, rowCount, columnIndex, columnCount, address');
      wallaTable.clear("Contents");
      await excel.sync();
      
      //console.log(wallaTable.address, wallaTable.rowCount);
      console.log('typeWalla',typeWalla, 'theResults', theResults);
      let scenes = [];
      let anyNonScenes = false;
      for (let i = 0; i < theResults.length; i++){
        let rowAndScene = await jade_modules.operations.getLineNoRowIndexAndScene(theResults[i].line);
        //console.log(i, 'rowAndScene', rowAndScene);
        if (rowAndScene.scene == -1){
          anyNonScenes = true
        } else {
          scenes.push(rowAndScene.scene)
        }
        console.log(i, 'line range', theResults[i].lineRange);
        if (theResults[i].lineRange.trim() == ''){
          theResults[i].lineRange = 'whole scene';
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
      if (scenes.length == 0){
        scenes[0] = scene;
      }
      //console.log('anyNonScenes', anyNonScenes, 'scenes', scenes)
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
        //console.log(resultArray)
        //console.log('Display Range rows: ', displayRange.rowCount, 'columns: ', displayRange.columnCount);
    
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
    })  
    for (let i = 0; i < resultArray.length; i++){
      let lineNo = resultArray[i][tableCols.lineNo];
      if (lineNo > 0){
        let data = {};
        data.rowIndex = resultArray[i][tableCols.rowIndex];
        data.wallaLineRange = resultArray[i][tableCols.lineRange];
        data.typeOfWalla = resultArray[i][tableCols.typeOfWalla];
        data.characters = resultArray[i][tableCols.character];
        data.description = resultArray[i][tableCols.description];
        data.numCharacters = resultArray[i][tableCols.numCharacters];
        data.all = resultArray[i][tableCols.wallaOriginal];
        data.lineNo = lineNo;
        wallaData.push(data); 
      }
    }
    wallaData.sort(mySortCompare);
  }
  return wallaData;
}

function mySortCompare(a, b){
  if (a.lineNo == b.lineNo){
    if (a.all > b.all){
      return 1
    }
    if (a.all < b.all){
      return -1
    }
    return 0
  } else {
    if (a.lineNo > b.lineNo){
      return 1
    }
    if (a.lineNo < b.lineNo){
      return -1
    }
    return 0
  }
}
async function getScene(sourceRowId, doPrevious){
  let scene = -1;
  await Excel.run(async (excel) => {
    let wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
    const sceneWallaIndexColumn = 4;
    const indexTableRange = wallaSheet.getRange('wiWallaIndexTable');
    indexTableRange.load('rowIndex, columnIndex')
    await excel.sync();
    console.log('sourceRow', sourceRowId);
    let theRow = indexTableRange.rowIndex + sourceRowId - 1;
    if (doPrevious){
      theRow = theRow - 1;
    }
    console.log(theRow, indexTableRange.columnIndex + sceneWallaIndexColumn, 1, 1)
    let sceneRange = wallaSheet.getRangeByIndexes(theRow, indexTableRange.columnIndex + sceneWallaIndexColumn, 1, 1)
    sceneRange.load('values');
    await excel.sync();
    let temp = sceneRange.values[0][0];
    if (!isNaN(parseInt(temp))){
      scene = parseInt(temp);
    }
  })
  return scene;
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
      //console.log('selectedRange address', selectedRanges.address)
      let ranges = selectedRanges.areas.items;
      //console.log(ranges)
      
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
    //console.log('wallaData', wallaData);
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

async function getWallaSourceWallaColumn(isScript){
  const book = await jade_modules.operations.getBook();
  let wallaColumn = wallaSourceUKScriptColumnIndex;
  for (let i = 0; i < wallaSourceWallaColumn.length; i++){
    if (wallaSourceWallaColumn[i].book == book){
      if (isScript){
        wallaColumn = wallaSourceWallaColumn[i].scriptColumn;  
      } else {
        wallaColumn = wallaSourceWallaColumn[i].column;
      }
      console.log('Got walla column', wallaColumn);
      break
    }
  }
  return wallaColumn;
}
async function getTheWallaSourceIndecies(){
  let wallaIndexes = []
  let named = 0;
  let unNamed = 0;
  let general = 0;
  let isGood = true;
  const wallaColumn = await getWallaSourceWallaColumn(false)
  await Excel.run(async (excel) => {
    const sourceSheet = excel.workbook.worksheets.getItem(wallaSourceSheetName);
    const usedRange = sourceSheet.getUsedRange();
    usedRange.load('rowIndex, rowCount');
    await excel.sync();
    console.log('rowIndex', usedRange.rowIndex, 'rowCount', usedRange.rowCount);
    let scriptRange = sourceSheet.getRangeByIndexes(usedRange.rowIndex, wallaColumn, usedRange.rowCount, 1)
    scriptRange.load('values');
    await excel.sync()
    for (let i = 0; i < scriptRange.values.length; i++){
      let raw = scriptRange.values[i][0].toString();
      console.log('raw', raw)
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
    console.log('wallaIndexes', wallaIndexes)
    if (wallaIndexes.length > 0){
      for (i = 0; i < wallaIndexes.length; i = i + 3){
        num += 1;
        let myRow = [num, wallaIndexes[i].rowIndex, wallaIndexes[i + 1].rowIndex, wallaIndexes[i + 2].rowIndex, '']
        results.push(myRow)
      }
      console.log('results', results)
      let tempRange = wallaSheet.getRangeByIndexes(indexTableRange.rowIndex, indexTableRange.columnIndex, results.length, indexTableRange.columnCount);
      tempRange.values = results;
    }
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
      let tableRowIndex = activeCell.rowIndex - indexTableRange.rowIndex + 1;
      rowIdRange.values = [[tableRowIndex]];
      await parseSource(tableRowIndex);
    } else {
      alert('Not in table');
    }
  })
}
async function loadTextBox(rowIndex){
  let sourceText;
  let wallaColumn = await getWallaSourceWallaColumn(false);
  const replacements = await wallaReplacementWords();
  await Excel.run(async (excel) => {
    if (!isNaN(parseInt(rowIndex))){
      const wallaSourceSheet = excel.workbook.worksheets.getItem(wallaSourceSheetName);
      const wallaSheet = excel.workbook.worksheets.getItem(wallaImportName);
      let testRange = wallaSourceSheet.getRangeByIndexes(rowIndex, wallaColumn, 1, 1);
      testRange.load('values');
      const sourceRowIndexRange = wallaSheet.getRange('wiSourceRowIndex');
      await excel.sync();
      console.log(testRange.values[0][0]);
      let wallaText = testRange.values[0][0];
      let textRange = wallaSheet.getRange('wiSource');
      textRange.values = [[wallaText.trim()]];  
      sourceText = replaceReplacements(wallaText.trim(), replacements);
      sourceRowIndexRange.values = [[rowIndex]];
      await excel.sync();
    }
  })
  return sourceText;
}
async function loopThroughTheIndexes(){
  await Excel.run(async (excel) => {
    const wallaSheet = excel.workbook.worksheets.getItem(wallaImportName);
    const indexTableRange = wallaSheet.getRange('wiWallaIndexTable');
    indexTableRange.load('rowIndex, values');
    await excel.sync();
    for (let i = 0; i < indexTableRange.values.length; i++){
      let rowIndex = indexTableRange.values[i][1];
      console.log(i, 'rowIndex', rowIndex, 'of', indexTableRange.values.length);
      if (!isNaN(parseInt(rowIndex))){
        await loadTextBox(rowIndex);
        const rowIdRange = wallaSheet.getRange('wiSourceRowId');
        rowIdRange.values = [[indexTableRange.values[i][0]]];
        await excel.sync();
        await parseSource();
      }
    }
  })
}

async function findFirstWallaOriginal(){
  await Excel.run(async (excel) => {
    const wallaSheet = excel.workbook.worksheets.getItem(wallaImportName);
    const tableRange = wallaSheet.getRange(wallaTableName);
    tableRange.load('values');
    await excel.sync();
    let textToSearch = tableRange.values[0][0];
    if (textToSearch.trim() != ''){
      const wallaSourceSheet = excel.workbook.worksheets.getItem(wallaSourceSheetName);
      const usedRange = wallaSourceSheet.getUsedRange();
      let found = usedRange.findOrNullObject(textToSearch);
      await excel.sync()
      if (!found.isNullObject){
        wallaSourceSheet.activate();
        found.select();
        await excel.sync();
      }
    }
  })
}

async function completeProcess(){
  let progressPanel = tag('walla-progress');
  progressPanel.style.display = 'block';
  let textArea = tag('walla-text');
  textArea.value = 'Starting \n';
  let startRow = 36;
  let endRow = 40;
  
  let startTextBox = tag('walla-start');
  let endTextBox = tag('walla-end');
  let start = parseInt(startTextBox.value);
  let end = parseInt(endTextBox.value);

  if ((isNaN(start)) || (isNaN(end))){
    textArea.value += 'Incorrect row values. Stopping \n'
  } else {
    startRow = start;
    endRow = end;
    /*
    textArea.value += 'Clearing Walla from Script \n';
    await jade_modules.operations.clearWalla();
    textArea.value += 'Clearing Walla Blocks from Script \n';
    await jade_modules.operations.deleteAllWallaBlocks(false);
    
    textArea.value += 'Getting Walla Data \n';
    await getTheWallaSourceIndecies();
    textArea.value += 'Getting Scene Data \n';
    await loopThroughTheIndexes();
    textArea.value += 'Checking all scenes \n';
    */
    let good = await checkWeHaveAllScenes();
    if (good){
      textArea.value += 'All scenes OK \nDoing parsing \n'
      await putDataInScript(startRow, endRow)
    }
  }
  textArea.value += 'Done \n';
}

async function checkWeHaveAllScenes(){
  let allGood = true;
  await Excel.run(async (excel) => {
    const sceneColumnIndex = 4;
    const wallaSheet = excel.workbook.worksheets.getItem(wallaImportName);
    const indexTableRange = wallaSheet.getRange('wiWallaIndexTable');
    indexTableRange.load('rowIndex, values');
    await excel.sync();
    let lastValue, thisValue;
    let allGood = true;
    for (let i = 0; i < indexTableRange.values.length; i++){
      if (i == 0){
        lastValue = parseInt(indexTableRange.values[i][sceneColumnIndex]);
      } else {
        thisValue = parseInt(indexTableRange.values[i][sceneColumnIndex]);
        if (!isNaN(thisValue)){
          if (lastValue + 1 == thisValue){
            lastValue = thisValue;
          } else {
            allGood = false;
            break;
          }
        } else {
          //No more scenes
        }
      }
    }
    if (allGood){
      console.log('All present up to: ' + lastValue);
    } else {
      console.log('Inconsitancy at: ' + thisValue);
    }
  })
  return allGood;
}

async function putDataInScript(startRow, endRow){
  let textArea = tag('walla-text');
  let baseText = textArea.value
  await Excel.run(async (excel) => {
    const wallaSheet = excel.workbook.worksheets.getItem(wallaImportName);
    const indexTableRange = wallaSheet.getRange('wiWallaIndexTable');
    indexTableRange.load('rowIndex, values');
    await excel.sync();
    for (let i = startRow; i < endRow; i++){
      textArea.value = baseText + 'Doing row: ' + (i + 1) + ' \n';
      let sceneNo = indexTableRange.values[i][4]

      textArea.value += 'Clearing Walla Block from Script \n';
      console.log('sceneNo', sceneNo);
      await jade_modules.operations.deleteWallaBlock(sceneNo, true);
      await jade_modules.operations.clearWallaForScene(sceneNo);

      let namedRowIndex = indexTableRange.values[i][1];
      await doTheRowIndex(namedRowIndex, sceneNo);
      let unNamedRowIndex  = indexTableRange.values[i][2];
      await doTheRowIndex(unNamedRowIndex, sceneNo);
      let generalRowIndex = indexTableRange.values[i][3];
      await doTheRowIndex(generalRowIndex, sceneNo);
      textArea.value += 'Calculating walla cues \n'
      await jade_modules.operations.calculateWallaCues();
      textArea.value += 'Named Character Walla to script \n'
      await jade_modules.operations.getSceneWallaInformation(1, sceneNo, i == 0);
      textArea.value += 'Un-named Walla to script \n'
      await jade_modules.operations.getSceneWallaInformation(2, sceneNo, i == 0);
      textArea.value += 'General Walla to script \n'
      await jade_modules.operations.getSceneWallaInformation(3, sceneNo, i == 0);
    }
  }) 
}

async function doTheWallaBlocks(){
  let textArea = tag('walla-text');
  let wallaProgress = tag('walla-progress')
  wallaProgress.style.display = 'block'
  let baseText = textArea.value
  let sceneStart = parseInt(tag('walla-block-start').value);
  let sceneEnd = parseInt(tag('walla-block-end').value);
  if ((!isNaN(sceneStart)) && (!isNaN(sceneEnd))){
    for (let sceneNo = sceneStart; sceneNo < sceneEnd; sceneNo++){
      textArea.value = baseText + 'Doing scene: ' + sceneNo + ' \n';
      await jade_modules.operations.getSceneWallaInformation(1, sceneNo);
      await jade_modules.operations.getSceneWallaInformation(2, sceneNo);
      await jade_modules.operations.getSceneWallaInformation(3, sceneNo);
    }
  }
  textArea.value += 'Done!';
}

async function doTheRowIndex(theRowIndex, sceneNo){
  let sourceText = await loadTextBox(theRowIndex);
  let details = await parseSourceText(sourceText);
  let wallaData = await doWallaTableV2(details.typeWalla, details.theResults, sceneNo);
  console.log('wallaData', wallaData);
  if (wallaData.length > 0){
    await jade_modules.operations.createMultipleWallas(wallaData, false, true, false);
  }
}
async function wallaReplacementWords(){
  let data = [];
  await Excel.run(async (excel) => {
    const settingsSheet = excel.workbook.worksheets.getItem('Settings');
    const replacementRange = settingsSheet.getRange('seWallaReplace');
    replacementRange.load('values');
    await excel.sync();
    for (let i = 0; i < replacementRange.values.length; i++){
      if (replacementRange.values[i][0] != ''){
        data.push({original: replacementRange.values[i][0], replacement: replacementRange.values[i][1]})        
      }
    }
  })
  return data;
}

function replaceReplacements(theLine, replacements){
  let result = theLine;
  for (let i = 0; i < replacements.length; i++){
    let original = new RegExp(replacements[i].original, 'gi')
    result = result.replace(original, replacements[i].replacement);
  }
  return result
}

function isWallaScripted(theText){
  return wallaScriptingNames.includes(theText.trim())
}

async function getWallaScriptingRowIndexes(){
  let indexes = []
  let scriptColumnIndex = await getWallaSourceWallaColumn(true);
  await Excel.run(async (excel) => {
    const sourceSheet = excel.workbook.worksheets.getItem(wallaSourceSheetName);
    const used = sourceSheet.getUsedRange();
    used.load('rowIndex, rowCount');
    await excel.sync();
    const scriptRange = sourceSheet.getRangeByIndexes(used.rowIndex, scriptColumnIndex, used.rowCount, 1);
    scriptRange.load('rowIndex, values');
    await excel.sync();
    const scriptValues = scriptRange.values.map(x => x[0])
    for (let i = 0; i < scriptValues.length; i++){
      if (isWallaScripted(scriptValues[i])){
        indexes.push(i + scriptRange.rowIndex);
      }
    }
  })
  return indexes;
}

async function wallaScriptDetails(indexes){
  let details = [];
  const lookAhead = 20;
  let valueIndexes = {
    character: wallaScriptColumns.character - wallaScriptColumns.cue,
    presentCharacters: wallaScriptColumns.presentCharacters - wallaScriptColumns.cue,
    stageDirection: wallaScriptColumns.stageDirection - wallaScriptColumns.cue,
  }
  let scriptValueIndex = await getWallaSourceWallaColumn(true) - wallaScriptColumns.cue
  await Excel.run(async (excel) => {
    const sourceSheet = excel.workbook.worksheets.getItem(wallaSourceSheetName);
    let range = []
    for (let i = 0; i < indexes.length; i++){
      range[i] = sourceSheet.getRangeByIndexes(indexes[i], wallaScriptColumns.cue, lookAhead, wallaScriptColumns.columnCount)
      range[i].load('values');
    }
    await excel.sync();
    for (let i = 0; i < indexes.length; i++){
      let character = range[i].values[0][valueIndexes.character];
      let presentCharacters = range[i].values[0][valueIndexes.presentCharacters];
      let stageDirection  = range[i].values[0][valueIndexes.stageDirection];
      let script = range[i].values[0][scriptValueIndex];
      let nextCue = -1;
      for (let j = 1; j < lookAhead; j++ ){
        let test = parseInt(range[i].values[j][0])
        if (!isNaN(test)){
          nextCue = test;
          break
        }
      }
      details.push({
        character: character,
        presentCharacters: presentCharacters,
        stageDirection: stageDirection,
        script: script,
        nextCue: nextCue
      })
    }
    console.log('details', details)
  })
  return details;
}

async function insertIntoMainScript(details){
  const start = 0
  const end = details.length;
  for (let i = start; i < end; i++){
    let rowIndex = await jade_modules.operations.findCueRowIndex(details[i].nextCue)
    rowIndex = await jade_modules.operations.findPreviousTypeLineRowIndex(rowIndex);
    console.log('Walla Script', i, 'of', details.length, 'rowIndex', rowIndex);
    let newIndex = await jade_modules.operations.insertRowV2(rowIndex, false, true)
    console.log('newIndex', newIndex);
    await jade_modules.operations.insertWallaScript(newIndex, details[i], true);
  }
}

async function doWallaScripting(doInsert){
  const indexes = await getWallaScriptingRowIndexes();
  const details = await wallaScriptDetails(indexes);
  if (doInsert){
    await insertIntoMainScript(details);
  }
  console.log('Done');
}
async function deleteAllWallaScripting(){
  let indexes = await jade_modules.operations.findAllWallaScripted();
  await jade_modules.operations.deleteRowsFromIndexes(indexes, true);
}
