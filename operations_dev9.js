const codeVersion = '9.0';
const firstDataRow = 3;
const lastDataRow = 29999;
const scriptSheetName = 'Script';
const characterListName = 'Character List';
const newTextSheetName = 'NewText'
const settingsSheetName = 'Settings';
const forDirectorName = 'For Directors';
const forActorsName = 'For Actors';
const forSchedulingName = 'For Scheduling';
const wallaImportName = 'Walla Import';
const locationSheetName = 'Locations';
const comparisonSheetName = 'Comparison';
const columnsToLock = "A:T";
const sceneBlockRows = 4;
const namedCharacters = 'Named Characters - For reaction sounds and walla';
const namedCharactersColon = 'Named Characters - For reaction sounds and walla:';
const unnamedCharacters = 'Un-named Character Walla';
const unnamedCharactersColon = 'Un-named Character Walla:';
const generalWalla = 'General Walla';
const generalWallaColon = 'General Walla:';
const actorScriptName = 'Actor Script';
let sceneBlockColumns = 9; //Can be changed in add scene block
let wallaBlockColumns = 8;

let sceneIndex, numberIndex, cueIndex, characterIndex, locationIndex, chapterIndex, lineIndex;
let totalTakesIndex, ukTakesIndex, ukTakeNoIndex, ukDateIndex, ukStudioIndex, ukEngineerIndex, ukMarkUpIndex, usCueIndex;
let usTakesIndex, usTakeNoIndex, usDateIndex, usStudioIndex, usEngineerIndex, usMarkUpIndex, usScriptColumnIndex;
let wallaTakesIndex, wallaTakeNoIndex, wallaDateIndex, wallaStudioIndex, wallaEngineerIndex, wallaMarkUpIndex; 
let wallaLineRangeIndex, numberOfPeoplePresentIndex, wallaOriginalIndex, wallaCueIndex, typeOfWallaIndex, typeCodeIndex;
let mySheetColumns, ukScriptIndex, otherNotesIndex, sceneWordCountCalcIndex, bookIndex, lineWordCountIndex, sceneLineNumberRangeIndex, chapterCalculationIndex;
let scriptSheet;

let sceneInput, lineNoInput, chapterInput;
let typeCodeValues, addSelectList;

let scriptHiddenRows = [];
let myTypes = {
  chapter: 'Chapter',
  scene: 'Scene',
  line: 'Line',
  sceneBlock: 'Scene Block',
  wallaScripted: 'Walla Scripted',
  wallaBlock: 'Walla Block'
}

let myFormats = {
  purple: '#f3d1f0',
  green: '#daf2d0',
  wallaGreen: '#b5e6a2',
  black: '#000000',
  white: '#ffffff',
  chapterGreen: '#00B060'
}

let screenColours = {
  main: {
    background:'#d8dfe5',
    fontColour: '#46656F'
  },
  actor: {
    background: '#fbe2d5',
    fontColour: '#592509'
  },
  director: {
    background: '#caedfb',
    fontColour: '#06394d'
  },
  scheduling: {
    background: '#daf2d0',
    fontColour: '#1d3a10'
  },
  walla: {
    background: '#f2ceef',
    fontColour: '#481343'
  },
  location: {
    background: '#c1f0c8',
    fontColour: '#0d3714'
  },
  actorScript: {
    background: '#ffffff',
    fontColour: '#333333'
  }
}

let choiceType ={
  list: 'List Search',
  text: 'Text Search'
}

function auto_exec(){
}

async function showMain(){
  let waitPage = tag('start-wait');
  let mainPage = tag('main-page');
  waitPage.style.display = 'none';
  mainPage.style.display = 'block';
  await showMainPage();
}

async function getColumnData(sheetName, rangeName){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(rangeName);
    range.load("values");
    await excel.sync();
    console.log(range.values);
    let result = [];
    for (let i = 0; i < range.values.length; i++){
      if (range.values[i][0] != ""){
        let temp = {};
        temp.name = range.values[i][0];
        temp.number = range.values[i][1];
        temp.column = range.values[i][2];
        temp.index = range.values[i][3];
        temp.width = range.values[i][4];
        result.push(temp);
      }
    }
    console.log(result);
    mySheetColumns = result;
  })
}

async function initialiseVariables(){
  sceneIndex = findColumnIndex('Scene')
  numberIndex = findColumnIndex("Number");
  chapterIndex = findColumnIndex('Chapter')
  totalTakesIndex = findColumnIndex('Total Takes');
  sceneLineNumberRangeIndex = findColumnIndex('Scene Line Number Range');
  cueIndex = findColumnIndex('Cue');
  stageDirectionWallaDescriptionIndex = findColumnIndex("Stage Direction/ Walla description") //J

  characterIndex = findColumnIndex('Character');
  locationIndex = findColumnIndex('Location');
  lineIndex = findColumnIndex('Line');
  ukScriptIndex = findColumnIndex('UK script');
  otherNotesIndex = findColumnIndex('Other notes');
  usCueIndex = findColumnIndex('US Cue');
  usScriptColumnIndex = findColumnIndex('US Script');  
  ukTakesIndex = findColumnIndex('UK No of takes');
  ukTakeNoIndex = findColumnIndex('UK Take No')
  ukDateIndex = findColumnIndex("UK Date Recorded");
  ukStudioIndex = findColumnIndex("UK Studio");
  ukEngineerIndex = findColumnIndex("UK Engineer");
  ukMarkUpIndex = findColumnIndex("UK Broadcast Assistant Markup");

  usTakesIndex = findColumnIndex('US No of takes');
  usTakeNoIndex = findColumnIndex('US Take No');
  usDateIndex = findColumnIndex("US Date Recorded");
  usStudioIndex = findColumnIndex("US Studio");
  usEngineerIndex = findColumnIndex("US Engineer");
  usMarkUpIndex = findColumnIndex("US Broadcast Assistant Markup");

  wallaTakesIndex = findColumnIndex('Walla No Of takes');
  wallaTakeNoIndex = findColumnIndex('Walla Take No');
  wallaDateIndex = findColumnIndex("Walla Date Recorded");
  wallaStudioIndex = findColumnIndex("Walla Studio");
  wallaEngineerIndex = findColumnIndex("Walla Engineer");
  wallaMarkUpIndex = findColumnIndex("Walla Broadcast Assistant Markup");

  wallaLineRangeIndex = findColumnIndex('Walla Line Range');
  typeOfWallaIndex = findColumnIndex('Type Of Walla')
  numberOfPeoplePresentIndex = findColumnIndex('Number of people present');
  typeCodeIndex = findColumnIndex('Type Code');
  wallaOriginalIndex = findColumnIndex('Walla Original');  
  wallaCueIndex = findColumnIndex('Walla Cue No')

  chapterCalculationIndex = findColumnIndex('Chapter Calculation');
  lineWordCountIndex = findColumnIndex('Line Word Count');
  bookIndex = findColumnIndex('Book');
  sceneWordCountCalcIndex = findColumnIndex('Scene word count calc'); 

  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    await excel.sync();
  });
}

async function getMySheetColumns(){
  return mySheetColumns
}

function findColumnIndex(name){
  return mySheetColumns.find((col) => col.name === name).index;
}

function findColumnLetter(name){
  return mySheetColumns.find((col) => col.name === name).column;
}

async function lockColumns(){
  await Excel.run(async function(excel){
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    scriptSheet.protection.load('protected');
    await excel.sync();
    if (!scriptSheet.protection.protected){
      scriptSheet.protection.protect({ selectionMode: "Normal", allowAutoFilter: true });
      await excel.sync();    
    }
    let protectionText = tag('lockMessage')
    scriptSheet.protection.load('protected');
    await excel.sync();
    if (scriptSheet.protection.protected){
      protectionText.innerText = 'Sheet locked'
    } else {
      protectionText.innerText = 'Sheet unlocked'
    }
  })
}

async function unlock(){
  await Excel.run(async function(excel){
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    scriptSheet.protection.load('protected');
    await excel.sync();
    if (scriptSheet.protection.protected){
      scriptSheet.protection.unprotect("")
      await excel.sync();
    }
    scriptSheet.protection.load('protected');
    await excel.sync();
    let protectionText = tag('lockMessage')
    if (scriptSheet.protection.protected){
      protectionText.innerText = 'Sheet locked'
    } else {
      protectionText.innerText = 'Sheet unlocked'
    }
  })
}

async function applyFilter(){
  /*Jade.listing:{"name":"Apply filter","description":"Applies empty filter to sheet"}*/
  await Excel.run(async function(excel){
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    
    const myRange = await getDataRange(excel);
    
    scriptSheet.autoFilter.apply(myRange, 0, { criterion1: "*", filterOn: Excel.FilterOn.custom});
    scriptSheet.autoFilter.clearCriteria();
    await excel.sync();
    if (isProtected){
      await lockColumns();
    }
  });
}

async function unlockIfLocked(){
  // returns true if it was locked
  let isProtected;
  await Excel.run(async function(excel){
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    scriptSheet.protection.load('protected');
    await excel.sync();
    
    isProtected = scriptSheet.protection.protected
    if (isProtected){
      await unlock();
    }
  })
  return isProtected
}

async function removeFilter(){
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    scriptSheet.autoFilter.load('enabled')
    await excel.sync()
    if (scriptSheet.autoFilter.enabled){
      let isProtected = await unlockIfLocked();
      scriptSheet.autoFilter.remove();
      await excel.sync();
      if (isProtected){
        await lockColumns();
      }
    }  
  });
}

async function findScene(offset){
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load("rowIndex");
    activeCell.load(("columnIndex"))
    await excel.sync()
    const startRow = activeCell.rowIndex;
    const startColumn = activeCell.columnIndex
    let range = await getSceneRange(excel);
    range.load("values");
    await excel.sync();
    console.log("Scene range");
    console.log(range.values);
    
    console.log("Start Row");
    console.log(startRow);

    let currentValue = range.values[startRow - 2][0];
    console.log("Current Value");
    console.log(currentValue);

    let myIndex = -1;

    if (currentValue + offset > 0){
      myIndex = range.values.findIndex(a => a[0] == (currentValue + offset));
    }
    
    console.log("Found Index");
    console.log(myIndex);
    
    if (myIndex == -1){
      if (offset == 1){
        alert('This is the final scene')
      } else if (offset == -1){
        await firstScene();
      }
    } else {
      const myTarget = scriptSheet.getRangeByIndexes(myIndex + 2, startColumn, 1, 1);
      await excel.sync();
      myTarget.select();
      await excel.sync();
    }
  })
}

async function findSceneNo(sceneNo){
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load("rowIndex");
    activeCell.load(("columnIndex"))
    await excel.sync()
    const startRow = activeCell.rowIndex;
    const startColumn = activeCell.columnIndex
    let range = await getSceneRange(excel);
    range.load("values");
    await excel.sync();
    console.log("Scene range");
    console.log(range.values);
    
    console.log("Start Row");
    console.log(startRow);

    const minAndMax = await getSceneMaxAndMin();
    console.log("Min and Max");
    console.log(minAndMax);

    if (sceneNo > minAndMax.max){
      sceneNo = minAndMax.max;
    }

    if (sceneNo < minAndMax.min){
      sceneNo = minAndMax.min
    }

    const myIndex = range.values.findIndex(a => a[0] == (sceneNo));

    console.log("Found Index");
    console.log(myIndex);
    
    if (myIndex == -1){
      alert('Invalid Scene Number');
    } else {
      scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      const myTarget = scriptSheet.getRangeByIndexes(myIndex + 2, startColumn, 1, 1);
      myTarget.select();
      await excel.sync();
    }
  })
}

async function findLineNo(lineNo){
  let theRowIndex = -1;
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load("rowIndex");
    activeCell.load(("columnIndex"))
    await excel.sync()
    const startRow = activeCell.rowIndex;
    const startColumn = activeCell.columnIndex
    let range = await getLineRange(excel);
    range.load("values");
    range.load('rowIndex')
    await excel.sync();
    console.log("Line range");
    console.log(range.values);
    
    console.log("Start Row");
    console.log(startRow);

    const minAndMax = await getLineNoMaxAndMin();
    console.log("Min and Max");
    console.log(minAndMax);

    if (lineNo > minAndMax.max){
      lineNo = minAndMax.max;
    }

    if (lineNo < minAndMax.min){
      lineNo = minAndMax.min
    }

    const myIndex = range.values.findIndex(a => a[0] == (lineNo));
    theRowIndex = range.rowIndex + myIndex

    console.log("Found Index, Range Value, Calculated Row Index");
    console.log(myIndex, range.values[myIndex], theRowIndex, lineNo);
    
    if (myIndex == -1){
      alert('Invalid Line Number');
    } else {
      scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      const myTarget = scriptSheet.getRangeByIndexes(theRowIndex, startColumn, 1, 1);
      myTarget.select();
      await excel.sync();
    }
  })
  return theRowIndex
}

async function getLineNoRowIndex(lineNo){
  let myRowIndex;
  await Excel.run(async function(excel){
    let range = await getLineRange(excel);
    range.load("values");
    range.load('rowIndex')
    await excel.sync();
    const myIndex = range.values.findIndex(a => a[0] == (lineNo));
    console.log(myIndex, range.rowIndex);
    myRowIndex = myIndex + range.rowIndex;
  })
  return myRowIndex;
}

async function findChapter(chapter){
  await Excel.run(async function(excel){
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load("rowIndex, columnIndex");
    await excel.sync()
    const startRow = activeCell.rowIndex;
    const startColumn = activeCell.columnIndex
    let range = await getChapterRange(excel);
    range.load("values, rowIndex");
    await excel.sync();
    console.log("Chapter range");
    console.log(range.values);
    
    console.log("Start Row");
    console.log(startRow);

    const minAndMax = await getChapterMaxAndMin();
    console.log("Min and Max");
    console.log(minAndMax);

    if (chapter > minAndMax.max){
      chapter = minAndMax.max;
    }

    if (chapter < minAndMax.min){
      chapter = minAndMax.min
    }

    const myIndex = range.values.findIndex(a => a[0] == (chapter));

    console.log("Found Index");
    console.log(myIndex);
    
    if (myIndex == -1){
      alert('Invalid Line Number');
    } else {
      let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      let targetRowIndex = range.rowIndex + myIndex;
      let rowOffset = 5
      let theOffsets = [-rowOffset, rowOffset, 0];
      let myTarget;
      for (let i = 0; i < theOffsets.length; i++){
        let tempRowIndex = targetRowIndex + theOffsets[i];
        console.log('i', i, 'tempRowIndex', tempRowIndex)
        if (tempRowIndex < 0){ tempRowIndex = 0};
        myTarget = scriptSheet.getRangeByIndexes(tempRowIndex, startColumn, 1, 1);
        myTarget.select();
        await excel.sync();
      }
    }
  })
}

async function firstScene(){
  await Excel.run(async function(excel){
    const minAndMax = await getSceneMaxAndMin();
    console.log("Min and Max");
    console.log(minAndMax);
    if (minAndMax.min < 1){
      await findSceneNo(1);
    } else {
      await findSceneNo(minAndMax.min);
    }
    
  })
}

async function lastScene(){
  await Excel.run(async function(excel){
    const minAndMax = await getSceneMaxAndMin();
    console.log("Min and Max");
    console.log(minAndMax);
    await findSceneNo(minAndMax.max);
  })
}

async function getSceneRange(excel){
  scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
  const endRow = scriptSheet.getUsedRange().getLastRow();
  endRow.load("rowIndex");
  await excel.sync();
  range = scriptSheet.getRangeByIndexes(2, sceneIndex, endRow.rowIndex, 1);
  await excel.sync();
  return range;
}

async function getLineRange(excel){
  scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
  const endRow = scriptSheet.getUsedRange().getLastRow();
  endRow.load("rowIndex");
  await excel.sync();
  range = scriptSheet.getRangeByIndexes(2, lineIndex, endRow.rowIndex, 1);
  await excel.sync();
  return range;
}

async function getChapterRange(excel){
  let details = await getFirstLastIndex();
  let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
  console.log('details', details);
  console.log(details.rowIndex + 2 , chapterCalculationIndex, details.rowCount - (2 - details.rowIndex), 1);
  let range = scriptSheet.getRangeByIndexes(details.rowIndex + 2 , chapterCalculationIndex, details.rowCount - (2 - details.rowIndex), 1);
  await excel.sync();
  return range;
}

async function getDataRange(excel){
  let range;
  let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
  const myLastRow = scriptSheet.getUsedRange().getLastRow();
  const myLastColumn = scriptSheet.getUsedRange().getLastColumn();
  myLastRow.load("rowIndex");
  myLastColumn.load("columnIndex")
  await excel.sync();
  range = scriptSheet.getRangeByIndexes(1,0, myLastRow.rowIndex, myLastColumn.columnIndex + 1);
  await excel.sync();
  return range
}

async function getTargetSceneNumber(){
  const textValue = sceneInput.value;
  const sceneNumber = parseInt(textValue);
  console.log(textValue, sceneNumber);
  if (!isNaN(sceneNumber)){
    console.log(sceneNumber);
    await findSceneNo(sceneNumber);
  }  else {
    alert("Please enter a number")
  }  
}

async function getTargetLineNo(){
  const textValue = lineNoInput.value;
  const lineNumber = parseInt(textValue);
  if (!isNaN(lineNumber)){
    console.log(lineNumber);
    await findLineNo(lineNumber);
  }  else {
    alert("Please enter a number")
  }  
}

async function getTargetChapter(){
  const textValue = chapterInput.value;
  const chapterNumber = parseInt(textValue);
  if (!isNaN(chapterNumber)){
    console.log(chapterNumber);
    await findChapter(chapterNumber);
  }  else {
    alert("Please enter a number")
  }  
}


async function getSceneMaxAndMin(){
  let result = {};
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const min = scriptSheet.getRange("minScene");
    await excel.sync();
    min.load("values");
    await excel.sync();
    const max = scriptSheet.getRange("maxScene");
    await excel.sync();
    max.load("values")
    await excel.sync();
    console.log('scene min:', min.values[0][0]);
    console.log('scene max: ',max.values[0][0]);
    
    result.min = min.values[0][0];
    result.max = max.values[0][0];
    console.log(result);
  })
  return result;
}

async function getLineNoMaxAndMin(){
  let result = {};
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const min = scriptSheet.getRange("minLine");
    await excel.sync();
    min.load("values");
    await excel.sync();
    const max = scriptSheet.getRange("maxLine");
    await excel.sync();
    max.load("values")
    await excel.sync();
    console.log(min.values[0][0]);
    console.log(max.values[0][0]);
    
    result.min = min.values[0][0];
    result.max = max.values[0][0];
    console.log(result);
  })
  return result;
}

async function getChapterMaxAndMin(){
  let result = {};
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const min = scriptSheet.getRange("minChapter");
    await excel.sync();
    min.load("values");
    await excel.sync();
    const max = scriptSheet.getRange("maxChapter");
    await excel.sync();
    max.load("values")
    await excel.sync();
    console.log(min.values[0][0]);
    console.log(max.values[0][0]);
    
    result.min = min.values[0][0];
    result.max = max.values[0][0];
    console.log(result);
  })
  return result;
}



async function fill(country){
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const studioText = tag("studio-select").value;
    const engineerText = tag("engineer-select").value;
    const markupText = tag('markup').value
    let dateIndex, studioIndex, engineerIndex, markUpIndex;
    if (country == 'UK'){
      markUpIndex = ukMarkUpIndex;
      dateIndex = ukDateIndex;
      studioIndex = ukStudioIndex;
      engineerIndex = ukEngineerIndex;
    } else if ( country == 'US'){
      markUpIndex = usMarkUpIndex;
      dateIndex = usDateIndex;
      studioIndex = usStudioIndex;
      engineerIndex = usEngineerIndex;
    } else if ( country == 'Walla'){
      markUpIndex = wallaMarkUpIndex;
      dateIndex = wallaDateIndex;
      studioIndex = wallaStudioIndex;
      engineerIndex = wallaEngineerIndex;
    }
    
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load("rowIndex");
    await excel.sync();
    const myRow = activeCell.rowIndex;    
    console.log("Row Index");
    console.log(myRow)
    const markupRange = scriptSheet.getRangeByIndexes(myRow, markUpIndex, 1, 1);
    const dateRange = scriptSheet.getRangeByIndexes(myRow, dateIndex, 1, 1);
    const studioRange = scriptSheet.getRangeByIndexes(myRow, studioIndex, 1, 1);
    const engineerRange = scriptSheet.getRangeByIndexes(myRow, engineerIndex, 1, 1);
    await excel.sync();
    let isProtected = await unlockIfLocked();
    
    console.log(studioRange);
    markupRange.values = [[markupText]]
    dateRange.values = [[dateInFormat()]];
    studioRange.values = [[studioText]];
    engineerRange.values = [[engineerText]];
    await excel.sync();
    if (isProtected){
      await lockColumns();
    }
    engineerRange.select();
    await excel.sync();
    markupRange.select();
    await excel.sync();
  })
}

function dateInFormat(){
	var nowDate = new Date(); 
  let myMonth = (nowDate.getMonth()+1)
  if (myMonth < 10){
    myMonth = "0" + myMonth;
  } else {
    myMonth = myMonth.toString();
  }
  let myDay = nowDate.getDate();
  if (myDay < 10){
    myDay = "0" + myDay;
  } else {
    myDay = myDay.toString();
  }

	return nowDate.getFullYear().toString().substring(2) + myMonth + myDay; 
}
async function getDataFromSheet(sheetName, rangeName, selectTag){
  console.log('sheet', sheetName, rangeName, selectTag);
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(rangeName);
    range.load("values");
    await excel.sync();
    console.log(range.values);
    console.log(range.values.length)
    let result = [];
    for (let i = 0; i < range.values.length; i++){
      if (range.values[i][0] != ""){
        result.push(range.values[i][0]);
      }
    }
    console.log(result); 
    var studioSelect = tag(selectTag);

    for (let i = 0; i < result.length; i++){
      studioSelect.add(new Option(result[i], result[i]));
    }
      
  })
}

function getColumnFormulae(firstRow, firstRestRow, lastRow){
  const cueColumn = findColumnLetter("Cue") //F
  const sceneWordCountCalcColumn = findColumnLetter("Scene word count calc"); //CC
  const sceneLineNumberRangeColumn = findColumnLetter("Scene Line Number Range"); //C
  const positionMinusColumn = findColumnLetter("Position -"); //BT
  const startLineColumn = findColumnLetter("Start Line"); //BU
  const positionEndSqaureBracketColumn = findColumnLetter("Position ]"); //BV
  const endLineColumn = findColumnLetter("End Line"); //BW
  const numberColumn = findColumnLetter("Number"); //G
  const ukTakeNoColumn = findColumnLetter("UK Take No"); //V
  const UKScriptColumn = findColumnLetter("UK script"); //K
  const sceneBordersColumn = findColumnLetter("Scene Borders"); //CH
  const sceneColumn = findColumnLetter("Scene"); //CB
  const wordCountToThisLineColumn = findColumnLetter("Word count to this line"); //CB
  const lineWordCountColumn = findColumnLetter("Line Word Count") //BY
  const lineColumn = findColumnLetter("Line"); // CA
  const stageDirectionWallaDescriptionColumn = findColumnLetter("Stage Direction/ Walla description") //J
  const positionChapterColumn = findColumnLetter("Position Chapter"); //CF
  const chapterCalculationColumn = findColumnLetter("Chapter Calculation"); //CG
  const alphaLineRangeColumn = findColumnLetter('Alpha Line Range') //CJ
  const sceneLineCountCalculationColumn = findColumnLetter("Scene Line Count Calculation"); //CH
  const bookColumn = findColumnLetter("Book"); //CK

  const columnFormulae = [
    {
      columnName: "Scene Word Count", //A
      formulaFirst: '=""',
      formulaRest: '=IF(' + cueColumn + firstRestRow + '="","",' + sceneWordCountCalcColumn + firstRestRow + ')'
    },
    {
      columnName: "Position -", //BV
      formulaFirst: '=IF(' + sceneLineNumberRangeColumn + firstRow + '="",0,FIND("-",' + sceneLineNumberRangeColumn + firstRow + '))',
      formulaRest: '=IF(' + sceneLineNumberRangeColumn + firstRestRow + '="",0,FIND("-",' + sceneLineNumberRangeColumn + firstRestRow + '))'
    },
    {
      columnName: "Start Line", //BW
      formulaFirst: 0,
      formulaRest: "=IF(" + positionMinusColumn + firstRestRow + "=0," + startLineColumn + firstRow + ",VALUE(MID(" + sceneLineNumberRangeColumn + firstRestRow + ",2," + positionMinusColumn + firstRestRow + "-2)))"
    },
    {
      columnName: "Position ]", //BX
      formulaFirst: '=IF(' + sceneLineNumberRangeColumn + firstRow + '="",0,FIND("]",' + sceneLineNumberRangeColumn + firstRow + '))',
      formulaRest: '=IF(' + sceneLineNumberRangeColumn + firstRestRow + '="",0,FIND("]",' + sceneLineNumberRangeColumn + firstRestRow + '))'
    },
    {
      columnName: "End Line",
      formulaFirst: 0,
      formulaRest: "=IF(" + positionEndSqaureBracketColumn + firstRestRow + "=0," + endLineColumn + firstRow + ",VALUE(MID(" + sceneLineNumberRangeColumn + firstRestRow + "," + positionMinusColumn + firstRestRow + "+1," + positionEndSqaureBracketColumn + firstRestRow + "-" + positionMinusColumn + firstRestRow + "-1)))"
    },
    {
      columnName: "Valid Line Number", //BZ
      formulaFirst:  "=AND(" + numberColumn + firstRow + ">=" + startLineColumn + firstRow + ", " + numberColumn + firstRow + "<=" + endLineColumn + firstRow + ")",
      formulaRest: "=AND(" + numberColumn + firstRestRow + ">=" + startLineColumn + firstRestRow + ", " + numberColumn + firstRestRow + "<=" + endLineColumn + firstRestRow + ")"
    },
    {
      columnName: "Line Word Count", //CA
      formulaFirst:  0,
      formulaRest: '=IF(NOT(OR(' + ukTakeNoColumn + firstRestRow + '="",' + ukTakeNoColumn + firstRestRow + '=1)), 0, LEN(TRIM(' + UKScriptColumn + firstRestRow + ')) - LEN(SUBSTITUTE(' + UKScriptColumn + firstRestRow + ', " ", "")) + 1)'
    },
    {
      columnName: "Scene", //CB
      formulaFirst:  '=seFirstScene',
      formulaRest: '=IF(OR(' + sceneBordersColumn + firstRestRow + '="Copy",' + sceneBordersColumn + firstRestRow + '=""),' + sceneColumn + firstRow + ',' + sceneColumn + firstRow + '+1)'
    },
    {
      columnName: "Line", //CC
      formulaFirst:  0,
      formulaRest: "=" + numberColumn + firstRestRow + ""
    },
	  {
	    columnName: "Word count to this line", //CD
      formulaFirst:  0,
      formulaRest: "=IF(" + sceneColumn + firstRestRow + "=" + sceneColumn + firstRow + "," + wordCountToThisLineColumn + firstRow + "+" + lineWordCountColumn + firstRestRow + "," + lineWordCountColumn + firstRestRow + ")"
  	},
	  {
	    columnName: "Scene word count calc", //CE
      formulaFirst:  0,
      formulaRest: "=VLOOKUP(" + endLineColumn + firstRestRow + "," + "$" + lineColumn + "$" + (firstDataRow + 1) + ":$" + wordCountToThisLineColumn + "$" + lastDataRow + ",2,FALSE)"
  	},
    {
      columnName: "Position Chapter", //CF
      formulaFirst: '=IF(' + stageDirectionWallaDescriptionColumn + firstRow + '="","",IF(ISERROR(FIND("Chapter",' + stageDirectionWallaDescriptionColumn + firstRow + ')),"",FIND("Chapter",' + stageDirectionWallaDescriptionColumn + firstRow + ')))',
      formulaRest: '=IF('+ stageDirectionWallaDescriptionColumn + firstRestRow + '="","",IF(ISERROR(FIND("Chapter",' + stageDirectionWallaDescriptionColumn + firstRestRow + ')),"",FIND("Chapter",' + stageDirectionWallaDescriptionColumn + firstRestRow + ')))'
    },
    {
      columnName: "Chapter Calculation", //CG
      formulaFirst: '=VALUE(IF(' + positionChapterColumn + firstRow + '="","",MID(' + stageDirectionWallaDescriptionColumn + firstRow + ',' + positionChapterColumn + firstRow + '+7,99)))',
      formulaRest: '=VALUE(IF(' + positionChapterColumn + firstRestRow + '="",' + chapterCalculationColumn + firstRow + ',MID(' + stageDirectionWallaDescriptionColumn + firstRestRow + ',' + positionChapterColumn + firstRestRow + '+7,99)))'
    },
    {
      columnName: "Chapter", //E
      formulaFirst: '=IF(' + cueColumn + firstRow + '="", "","Chapter " & TEXT(' + chapterCalculationColumn + firstRow + ', "0"))',
      formulaRest: '=IF(' + cueColumn + firstRestRow + '="", "","Chapter " & TEXT(' + chapterCalculationColumn + firstRestRow + ', "0"))'
    },
    {
      columnName: "Scene Borders", //CI
      formulaFirst: "Start",
      formulaRest: '=IF(' + cueColumn + firstRestRow + '="", IF(' + sceneBordersColumn + firstRow + '="Start",' + sceneBordersColumn + firstRow + ',""),IF(' + alphaLineRangeColumn + firstRestRow + '=' + alphaLineRangeColumn + firstRow + ',"Copy","Original"))'
    },
    {
      columnName: "Scene Line Count Calculation", //CH
      formulaFirst: 0,
      formulaRest: '=' + endLineColumn + firstRestRow + '-' + startLineColumn + firstRestRow + '+1'
    },
    {
      columnName: "Scene Line Count", //B
      formulaFirst: 0,
      formulaRest: '=IF(' + cueColumn + firstRestRow + '="","",' + sceneLineCountCalculationColumn + firstRestRow + ')'
    },
    {
      columnName: "Scene Number", //D
      formulaFirst: '=IF(' + sceneColumn + firstRow + '=0,"",' + sceneColumn + firstRow + ')',
      formulaRest: '=IF(' + sceneColumn + firstRestRow + '=0,"",' + sceneColumn + firstRestRow + ')',
    },
    {
      columnName: "Alpha Line Range", //CJ
      formulaFirst: '=' + startLineColumn + firstRow + '&' + endLineColumn + firstRow,
      formulaRest: '=' + startLineColumn + firstRestRow + '&' + endLineColumn + firstRestRow
    },
    {
      columnName: "Book", //CK
      formulaFirst: '=IF(' + positionChapterColumn + firstRow + '="","",LEFT(' + stageDirectionWallaDescriptionColumn + firstRow + ', ' + positionChapterColumn + firstRow + '-3))',
      formulaRest: '=IF(' + positionChapterColumn + firstRestRow + '="",' + bookColumn + firstRow + ',LEFT(' + stageDirectionWallaDescriptionColumn + firstRestRow + ',' + positionChapterColumn + firstRestRow + '-3))'
    }
  ]
  return columnFormulae;
}

async function theFormulas(actualFirstRow, actualLastRow){
  let waitLabel = tag('formula-wait');
  waitLabel.style.display = 'block';
  let firstRow = "" + firstDataRow;
  let firstRestRow = "4";
  let lastRow = "" + lastDataRow;
  let doTopRow = false;
  
  await Excel.run(async function(excel){ 
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    //console.log('actual', actualFirstRow, actualLastRow);
    if ((actualFirstRow === undefined) || (actualFirstRow == firstRow)) {
      doTopRow = true
      if ((actualLastRow === undefined)||(actualLastRow == lastRow)){
      } else {
        lastRow = actualLastRow;
      }
    } else {
      firstRow = "" + (actualFirstRow - 1);
      firstRestRow = "" + actualFirstRow;
      if ((actualLastRow === undefined)||(actualLastRow == lastRow)){
      } else {
        lastRow = actualLastRow;
      }
    }
    //console.log('firstRow: ', firstRow, "firstRestRow", firstRestRow, "lastRow", lastRow, 'doToprow', doTopRow);
    let columnFormulae = getColumnFormulae(firstRow, firstRestRow, lastRow);
    for (let columnFormula of columnFormulae){
      const columnLetter = findColumnLetter(columnFormula.columnName);
      let myTopRow;
      let topRowRange;
      let myRange;
      let range;
      if (doTopRow) {
        //console.log('Doing top row');
        myTopRow = columnLetter + firstRow;
        topRowRange = scriptSheet.getRange(myTopRow);
        topRowRange.formulas = columnFormula.formulaFirst;
      } 
      
      myRange = columnLetter + firstRestRow + ":" + columnLetter + lastRow;
      range = scriptSheet.getRange(myRange);
      range.formulas = columnFormula.formulaRest;
    
      //console.log(myRange + "  " + myTopRow);
      //console.log(columnFormula.formulaRest + "   " + columnFormula.formulaFirst);
      await excel.sync();
      //console.log(range.formulas + "   " + topRowRange.formulas);
    }
    if (isProtected){
      await lockColumns();
    } 
  });
  waitLabel.style.display = 'none';
}

async function insertRowV2(currentRowIndex, doCopy, doFullFormula){
  let newRowIndex;
  await Excel.run(async function(excel){
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    const dataRange = await getDataRange(excel);
    const myLastColumn = dataRange.getLastColumn();
    myLastColumn.load("columnindex")
    await excel.sync();
    const myRow = scriptSheet.getRangeByIndexes(currentRowIndex, 0, 1, myLastColumn.columnIndex+1);
    const newRow = myRow.insert("Down");
    await excel.sync();
    if (doCopy){
      newRow.copyFrom(myRow, "All");
      await excel.sync(); 
    }
    if (doFullFormula){
      //console.log('Doing full formulas');
      await theFormulas((currentRowIndex + 1), (currentRowIndex + 1));
    } else {
      await correctFormulas(currentRowIndex + 1);  
    }
    
    newRow.load('rowIndex');
    await excel.sync();
    newRowIndex = newRow.rowIndex;
    if (isProtected){
      await lockColumns();
    }
  });
  return newRowIndex;
}

async function deleteRow(){
  await Excel.run(async function(excel){
    const activeCell = excel.workbook.getActiveCell();
    const selectCell = activeCell.getOffsetRange(-1, 0);
    activeCell.load('rowIndex');
    await excel.sync();
    console.log(activeCell.rowIndex);
  
    const myRow = activeCell.getEntireRow();
    myRow.load('address');
    await excel.sync();
    console.log(myRow.address);
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    myRow.delete("Up");
    myRow.load('address');
    await excel.sync();
    console.log(myRow.address);
    await correctFormulas(activeCell.rowIndex);
    await doTakesAndNumTakes(activeCell.rowIndex - 1, 'UK', false, false, false, false);
    selectCell.select();
    await excel.sync();
    if (isProtected){
      await lockColumns();
    }
  });
}
async function correctFormulas(firstRow){
  const sceneLineNumberRangeColumn = findColumnLetter("Scene Line Number Range"); //C
  const sceneNumberColumn = findColumnLetter("Scene Number"); //D
  const stageDirectionWallaDescriptionColumn = findColumnLetter("Stage Direction/ Walla description") //J
  const positionMinusColumn = findColumnLetter("Position -"); //BU
  const startLineColumn = findColumnLetter("Start Line"); //BV
  const positionEndSqaureBracketColumn = findColumnLetter("Position ]"); //BW
  const endLineColumn = findColumnLetter("End Line"); //BX
  const lineWordCountColumn = findColumnLetter("Line Word Count") //BZ
  const sceneColumn = findColumnLetter("Scene"); //CA
  const wordCountToThisLineColumn = findColumnLetter("Word count to this line"); //CC
  const positionChapterColumn = findColumnLetter("Position Chapter"); //CE
  const chapterCalculationColumn = findColumnLetter("Chapter Calculation"); //CF
  const sceneBordersColumn = findColumnLetter("Scene Borders"); //CH
  const cueColumn = findColumnLetter('Cue') //F
  const alphaLineRangeColumn = findColumnLetter('Alpha Line Range') //CJ
  const bookColumn = findColumnLetter("Book"); //CK
  
  const columnFormulae = [
    {
      columnName: "Start Line", //BV
      formulaRest: "=IF(" + positionMinusColumn + firstRow + "=0," + startLineColumn + (firstRow - 1) + ",VALUE(MID(" + sceneLineNumberRangeColumn + firstRow + ",2," + positionMinusColumn + firstRow + "-2)))"
    },
    {
      columnName: "End Line", //BX
      formulaRest: "=IF(" + positionEndSqaureBracketColumn + firstRow + "=0," + endLineColumn + (firstRow - 1) + ",VALUE(MID(" + sceneLineNumberRangeColumn + firstRow + "," + positionMinusColumn + firstRow + "+1," + positionEndSqaureBracketColumn + firstRow + "-" + positionMinusColumn + firstRow + "-1)))"
    },
    {
      columnName: "Scene", //CB
      formulaRest: '=IF(OR(' + sceneBordersColumn + firstRow + '="Copy",' + sceneBordersColumn + firstRow + '=""),' + sceneColumn + (firstRow - 1) + ',' + sceneColumn + (firstRow - 1) + '+1)'
    },
    {
	    columnName: "Word count to this line", //CC
      formulaRest: "=IF(" + sceneColumn + firstRow + "=" + sceneColumn + (firstRow - 1) + "," + wordCountToThisLineColumn + (firstRow -1) + "+" + lineWordCountColumn + firstRow + "," + lineWordCountColumn + firstRow + ")"
  	}    ,
    {
      columnName: "Chapter Calculation", //CF
      formulaRest: '=VALUE(IF(' + positionChapterColumn + firstRow + '="",' + chapterCalculationColumn + (firstRow - 1) + ',MID(' + stageDirectionWallaDescriptionColumn + firstRow + ',' + positionChapterColumn + firstRow + '+7,99)))'
    },
    {
      columnName: "Scene Borders", //CI
      formulaRest: '=IF(' + cueColumn + firstRow + '="", IF(' + sceneBordersColumn + (firstRow - 1) + '="Start",' + sceneBordersColumn + (firstRow - 1) + ',""),IF(' + alphaLineRangeColumn + firstRow + '=' + alphaLineRangeColumn + (firstRow - 1) + ',"Copy","Original"))'
    },
    {
      columnName: "Book", //CK
      formulaRest: '=IF(' + positionChapterColumn + firstRow + '="",' + bookColumn + (firstRow - 1) + ',LEFT(' + stageDirectionWallaDescriptionColumn + firstRow + ',' + positionChapterColumn + firstRow + '-3))'
    }
    
  ]
  await Excel.run(async function(excel){ 
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    for (let columnFormula of columnFormulae){
      const columnLetter = findColumnLetter(columnFormula.columnName);
      const myRange = columnLetter + firstRow + ":" + columnLetter + (firstRow +1) ;
      //console.log("Range to replace: " + myRange);
      const range = scriptSheet.getRange(myRange);
      //console.log("Formula: " + columnFormula.formulaRest);
      range.formulas = columnFormula.formulaRest;
      await excel.sync();
      //console.log("Formula after sync: " + range.formulas);
    }
    if (isProtected){
      await lockColumns();
    }
  })
}

function zeroElement(value){
  return value[0];
}

async function addTakeDetails(country, doDate){
  let myAction = radioButtonChoice();
  console.log('The action: ', myAction);

  await Excel.run(async function(excel){ 
    const activeCell = excel.workbook.getActiveCell();
    let selectCell = activeCell.getOffsetRange(1, 0);
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    let lineDetails =  await findDetailsForThisLine();
    console.log(lineDetails);
    let takeNoIndex, dateRecordedIndex, markUpIndex, studioIndex, engineerIndex, countryTakes;
    let newLine;
    let newLineIndex;
    if (country == 'UK'){
      takeNoIndex = ukTakeNoIndex;
      dateRecordedIndex = ukDateIndex;
      markUpIndex = ukMarkUpIndex;
      studioIndex = ukStudioIndex;
      engineerIndex = ukEngineerIndex;
      countryTakes = lineDetails.ukTakes
      newLine = lineDetails.ukTakes + 1;
    } else if (country == 'US'){
      takeNoIndex = usTakeNoIndex;
      dateRecordedIndex = usDateIndex;
      markUpIndex = usMarkUpIndex;
      studioIndex = usStudioIndex;
      engineerIndex = usEngineerIndex;
      countryTakes = lineDetails.usTakes
      newLine = lineDetails.usTakes + 1;
    }else if (country == 'Walla'){
      takeNoIndex = wallaTakeNoIndex;
      dateRecordedIndex = wallaDateIndex;
      markUpIndex = wallaMarkUpIndex;
      studioIndex = wallaStudioIndex;
      engineerIndex = wallaEngineerIndex;
      countryTakes = lineDetails.wallaTakes
      newLine = lineDetails.wallaTakes + 1;
    }
    if (lineDetails.totalTakes == 0){
      let currentRowIndex = lineDetails.indicies[0];
      console.log('Current Row Index');
      console.log(currentRowIndex);
      newLineIndex = currentRowIndex;
      lineDetails.totalTakes = 1;
      console.log('Added row');
      console.log(lineDetails);
      selectCell = activeCell.getOffsetRange(0, 0);
    } else if (lineDetails.totalTakes == countryTakes){
      let currentRowIndex = lineDetails.indicies[countryTakes - 1];
      console.log('Current Row Index');
      console.log(currentRowIndex);
      await insertRowV2(currentRowIndex, true, false)
      newLineIndex = currentRowIndex + 1;
      lineDetails.indicies.push(newLineIndex);
      lineDetails.totalTakes += 1;
      console.log('Added row');
      console.log(lineDetails);
    } else {
      newLineIndex = lineDetails.indicies[newLine - 1];
      //Need to copy from the row above
      console.log('New Line Index: ', newLineIndex);
      scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      let newRange = scriptSheet.getRangeByIndexes(newLineIndex, markUpIndex, 1, (engineerIndex - markUpIndex + 1));
      let copyRange = scriptSheet.getRangeByIndexes(newLineIndex - 1, markUpIndex, 1, (engineerIndex - markUpIndex + 1));
      newRange.copyFrom(copyRange, "All");
      await excel.sync();
    }
    let takeNoRange = scriptSheet.getRangeByIndexes(newLineIndex, takeNoIndex, 1, 1)
    takeNoRange.values = newLine;
    if (country == 'UK'){
      lineDetails.ukTakes = newLine;
    } else if (country == 'US'){
      lineDetails.usTakes = newLine;
    } else if (country == 'Walla'){
      lineDetails.wallaTakes = newLine;
    }
    console.log("New Line");
    console.log(newLine);
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    if (doDate){
      let dateRange = scriptSheet.getRangeByIndexes(newLineIndex, dateRecordedIndex, 1, 1);
      let theDate = dateInFormat();
      console.log('The Date:', theDate);
      dateRange.values = theDate;
    }
    let markUpRange = scriptSheet.getRangeByIndexes(newLineIndex, markUpIndex, 1, 1);
    let studioRange = scriptSheet.getRangeByIndexes(newLineIndex, studioIndex, 1, 1);
    let engineerRange = scriptSheet.getRangeByIndexes(newLineIndex, engineerIndex, 1, 1);
    if ((myAction == 'justDate') || (myAction == 'detailsBelow')){
      markUpRange.clear("Contents");
      studioRange.clear("Contents");
      engineerRange.clear("Contents");
    }
    if (myAction == 'detailsBelow'){
      const studioText = tag("studio-select").value;
      const engineerText = tag("engineer-select").value;
      const markupText = tag("markup").value;

      markUpRange.values = markupText;
      studioRange.values = studioText;
      engineerRange.values = engineerText;
    }

    selectCell.select();
    await excel.sync();
  
    console.log("Line Details")
    console.log(lineDetails);
    await doTheTidyUp(lineDetails)
    if (isProtected){
      await lockColumns();
    } 
  });
}


async function findDetailsForThisLine(){
  let result = {};
  await Excel.run(async function(excel){ 
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex')
    await excel.sync();
    const currentRowIndex = activeCell.rowIndex
    let myIndecies = await getAllLinesWithThisNumber(excel, currentRowIndex);
    console.log("myIndecies");
    console.log(myIndecies);

    const totalTakesCell = scriptSheet.getRangeByIndexes(myIndecies[0], totalTakesIndex, 1, 1);
    const ukTakesCell = scriptSheet.getRangeByIndexes(myIndecies[0], ukTakesIndex, 1, 1);
    const usTakesCell = scriptSheet.getRangeByIndexes(myIndecies[0], usTakesIndex, 1, 1);
    const wallaTakesCell = scriptSheet.getRangeByIndexes(myIndecies[0], wallaTakesIndex, 1, 1);

    totalTakesCell.load('values');
    ukTakesCell.load('values');
    usTakesCell.load('values');
    wallaTakesCell.load('values');

    await excel.sync();
    result.totalTakes = cleanTakes(totalTakesCell.values);
    result.ukTakes = cleanTakes(ukTakesCell.values);
    result.usTakes = cleanTakes(usTakesCell.values);
    result.wallaTakes = cleanTakes(wallaTakesCell.values);
    result.indicies = myIndecies;
    result.currentRowIndex = currentRowIndex;

    console.log('Result');
    console.log(result);
    if (isProtected){
      await lockColumns();
    }
  })
  /*
  Find total number of takes.
  If 0/blank - use this line
    Make No of takes 1, total number of takes 1, takeNo = 1, Other coutries take N/A
  if 1 or more...
    Find number of takes for each country.
    If the country has a N/A use the lowest one
    If not add a row.  

  */
  return result;

}
function cleanTakes(values){
  let temp = parseInt(values);
  if (isNaN(temp)){
      return 0;
    } else {
      return temp;
    }
}

async function removeTake(country){
  let markUpIndex, engineerIndex, takeNoIndex, countryTakes;
  await Excel.run(async function(excel){
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    // get the lineDetails
    const activeCell = excel.workbook.getActiveCell();
    const selectCell = activeCell.getOffsetRange(-1, 0);
    let lineDetails =  await findDetailsForThisLine();
    console.log(lineDetails);
    let foundTake = 0;
    for (let i = 0; i < lineDetails.indicies.length; i++){
      if (lineDetails.indicies[i] == lineDetails.currentRowIndex){
        foundTake = i + 1
      }
    }
    if (country == 'UK'){
      markUpIndex = ukMarkUpIndex;
      console.log('Mark Up Index', markUpIndex);
      engineerIndex = ukEngineerIndex;
      console.log('Engineer Index', engineerIndex);
      takeNoIndex = ukTakeNoIndex;
      countryTakes = lineDetails.ukTakes;
    } else if (country == 'US'){
      markUpIndex = usMarkUpIndex;
      console.log('Mark Up Index', markUpIndex);
      engineerIndex = usEngineerIndex;
      console.log('Engineer Index', engineerIndex);
      takeNoIndex = usTakeNoIndex;
      countryTakes = lineDetails.usTakes;
    } else if (country == 'Walla'){
      markUpIndex = wallaMarkUpIndex;
      console.log('Mark Up Index', markUpIndex);
      engineerIndex = wallaEngineerIndex;
      console.log('Engineer Index', engineerIndex);
      takeNoIndex = wallaTakeNoIndex;
      countryTakes = lineDetails.wallaTakes;
    }   
    if (foundTake > 0){
      // Is this the last take for this country...
      console.log('Found take: ', foundTake);
      if (lineDetails.totalTakes == 1){
        console.log('Only 1 total takes, which we cannot delete, so we clear the relevant area')
        console.log('currentRowIndex: ', lineDetails.currentRowIndex);
        let clearRange = scriptSheet.getRangeByIndexes(lineDetails.currentRowIndex, markUpIndex, 1, (engineerIndex - markUpIndex + 1));
        clearRange.load('address');
        await excel.sync();
        console.log("Clear range: ", clearRange.address)
        clearRange.clear("Contents");
        let takeNoRange = scriptSheet.getRangeByIndexes(lineDetails.currentRowIndex, takeNoIndex, 1, 1);
        takeNoRange.values = "N/A"
        if (country == 'UK'){
          lineDetails.ukTakes -= 1;
        } else if (country == 'US'){
          lineDetails.usTakes -= 1;
        } else if (country == 'Walla'){
          lineDetails.wallaTakes -= 1;
        }
        if ((lineDetails.ukTakes == 0) && (lineDetails.usTakes == 0) && (lineDetails.wallaTakes == 0)){
          lineDetails.totalTakes = 0;
        }
        await excel.sync();
      } else {
        if (foundTake == countryTakes){
        console.log('Found take is countries final take')
          // Yes => is it on the final totaltakes?
          if (countryTakes == lineDetails.totalTakes){
            //Yes - Are there any other countries on this take?
            let otherCountriesOnThisTake;
            if (country == 'UK'){
              otherCountriesOnThisTake = (lineDetails.totalTakes == lineDetails.usTakes) || (lineDetails.totalTakes == lineDetails.wallaTakes);
            } else if (country == 'US'){
              otherCountriesOnThisTake = (lineDetails.totalTakes == lineDetails.ukTakes) || (lineDetails.totalTakes == lineDetails.wallaTakes);
            } else if (country == 'Walla'){
              otherCountriesOnThisTake = (lineDetails.totalTakes == lineDetails.ukTakes) || (lineDetails.totalTakes == lineDetails.usTakes);
            }
            if (otherCountriesOnThisTake){
              // test country is on final take as is another country
              //Yes - just clear the relevant cells and adjust that countries numbers.
              console.log(country, "is on the final take, but another one also");
              console.log('currentRowIndex: ', lineDetails.currentRowIndex);
              console.log('markUpIndex', markUpIndex);
              console.log('Diff: ', (engineerIndex - markUpIndex + 1));
              let clearRange = scriptSheet.getRangeByIndexes(lineDetails.currentRowIndex, markUpIndex, 1, (engineerIndex - markUpIndex + 1));
              clearRange.load('address');
              await excel.sync();
              console.log("Clear range: ", clearRange.address)
              clearRange.clear("Contents");
              let takeNoRange = scriptSheet.getRangeByIndexes(lineDetails.currentRowIndex, takeNoIndex, 1, 1);
              takeNoRange.values = "N/A"
              if (country == 'UK'){
                lineDetails.ukTakes -= 1;
              } else if (country == 'US') {
                lineDetails.usTakes -= 1;
              } else if (country == 'Walla') {
                lineDetails.wallaTakes -= 1;
              }
              await excel.sync();
            } else {
              // test country is on final take and it's the only one
              //No - Delete the row and update the total and country numbers
              console.log(country, " is on the final take and its the only one.");
              console.log('currentRowIndex: ', lineDetails.currentRowIndex);
              scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
              let deleteRange = scriptSheet.getRangeByIndexes(lineDetails.currentRowIndex, 0, 1, 1).getEntireRow();
              deleteRange.load('address');
              await excel.sync();
              console.log("Delete range address: ", deleteRange.address);
              deleteRange.delete("Up");
              selectCell.select();
              await excel.sync();
              await correctFormulas(lineDetails.currentRowIndex);
              lineDetails.totalTakes = lineDetails.totalTakes - 1;
              if (country == 'UK'){
                lineDetails.ukTakes -= 1;
              } else if (country == 'US') {
                lineDetails.usTakes -= 1;
              } else if (country == 'Walla') {
                lineDetails.wallaTakes -= 1;
              }
              lineDetails.currentRowIndex -= 1;
              lineDetails.indicies.pop();
            }
          } else {
            // Test country is not the final total take, but we are deleting the final test country take
            //No - just clear the relevant cells and adjust that countries numbers.
            console.log(country, " is not the final total take, but we are deleting the final take");
            console.log('currentRowIndex: ', lineDetails.currentRowIndex);
            console.log('markUpIndex', markUpIndex);
            console.log('Diff: ', (engineerIndex - markUpIndex + 1));
            let clearRange = scriptSheet.getRangeByIndexes(lineDetails.currentRowIndex, markUpIndex, 1, (engineerIndex - markUpIndex + 1));
            clearRange.load('address');
            await excel.sync();
            console.log("Clear range: ", clearRange.address)
            clearRange.clear("Contents");
            if (country == 'UK'){
              lineDetails.ukTakes -= 1;
            } else if (country == 'US') {
              lineDetails.usTakes -= 1;
            } else if (country == 'Walla') {
              lineDetails.wallaTakes -= 1;
            }
            await excel.sync();
          } 
        } else {
          // Test country is not the final one of UK and so....
          // No - so here we need to
            // 1. remove the one to be deleted.
            // 2. move the one below up
            // 3. if we now have a totally empty row - delete it
            // 4. Adjust the details
            // UK is not the final total take, but we are deleting the final UK take
            //No - just clear the relevant cells and adjust that countries numbers.
          console.log(country, " is not the final take of ", country);
          console.log('currentRowIndex: ', lineDetails.currentRowIndex);
          console.log('markUpIndex', markUpIndex);
          console.log('Diff: ', (engineerIndex - markUpIndex + 1)); 
          let firstItem  = lineDetails.currentRowIndex;
          let lastItem;
          if (country == 'UK'){
            lastItem = lineDetails.indicies[lineDetails.ukTakes - 1];
          } else if (country = 'US') {
            lastItem = lineDetails.indicies[lineDetails.usTakes - 1];
          } else if (country = 'Walla') {
            lastItem = lineDetails.indicies[lineDetails.wallaTakes - 1];
          }
          console.log('First/Last item', firstItem, lastItem);
          for (let item = firstItem; item < lastItem; item++){
            let currentRange = scriptSheet.getRangeByIndexes(item, markUpIndex, 1, (engineerIndex - markUpIndex + 1));
            let nextRowRange = scriptSheet.getRangeByIndexes(item + 1, markUpIndex, 1, (engineerIndex - markUpIndex + 1));
            currentRange.copyFrom(nextRowRange, "All");
            await excel.sync();
          }
          let lastRowRange = scriptSheet.getRangeByIndexes(lastItem, markUpIndex, 1, (engineerIndex - markUpIndex + 1));
          lastRowRange.clear("Contents");
          await excel.sync();
          if (country == 'UK'){
            lineDetails.ukTakes -= 1;
          } else if (country == 'US') {
            lineDetails.usTakes -= 1;
          } else if (country == 'Walla') {
            lineDetails.wallaTakes -= 1;
          }        
          if (!((lineDetails.totalTakes == lineDetails.ukTakes) || (lineDetails.totalTakes == lineDetails.usTakes) || (lineDetails.totalTakes == lineDetails.wallaTakes))){
            // if we get here then we need to delete a row because we now have an empty row.
            console.log('We now have to delete a row')
            scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
            let lastItem = lineDetails.indicies[lineDetails.indicies.length - 1]
            let deleteRange = scriptSheet.getRangeByIndexes(lastItem, 0, 1, 1).getEntireRow();
            deleteRange.load('address');
            await excel.sync();
            console.log("Delete range address: ", deleteRange.address);
            deleteRange.delete("Up");
            selectCell.select();
            await excel.sync();
            await correctFormulas(lineDetails.currentRowIndex);
            lineDetails.totalTakes = lineDetails.totalTakes - 1;
            lineDetails.currentRowIndex -= 1;
            lineDetails.indicies.pop();
          }
        }
      }
    } else {
      console.log('Take not found')
    }
    console.log("Line Details")
    console.log(lineDetails);
    await doTheTidyUp(lineDetails)
    if (isProtected){
      await lockColumns();
    }
  });
}

async function getAllLinesWithThisNumber(excel, currentRowIndex){
  //returns an array of indexes
  scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
  const numberColumn = findColumnLetter("Number");
  console.log('Current Row Index: ', currentRowIndex);
  console.log('numberIndex:', numberIndex)
  let currentNumberCell = scriptSheet.getRangeByIndexes(currentRowIndex, numberIndex, 1, 1)
  currentNumberCell.load('values');
  currentNumberCell.load('address')
  await excel.sync();
  console.log('currentNumberCell address: ', currentNumberCell.address);
  let numberData = scriptSheet.getRange(numberColumn + firstDataRow + ":" + numberColumn + lastDataRow);
  numberData.load('values');
  await excel.sync();
  let targetValue = currentNumberCell.values
  console.log("Target Value:" + targetValue);
  let myData = numberData.values.map(x => x[0]);
  console.log("Raw values");
  console.log(numberData.values);
  console.log("Mapped values");
  console.log(myData)
  const myIndecies = myData.map((x, i) => [x, i]).filter(([x, i]) => x == targetValue).map(([x, i]) => i + firstDataRow - 1);
  console.log("Found Index");
  console.log(myIndecies);
  return myIndecies;
}

async function doTheTidyUp(lineDetails){
  await Excel.run(async function(excel){ 
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let item = 0;
    for (let index of lineDetails.indicies){
      item += 1;
      let totalTakesRange = scriptSheet.getRangeByIndexes(index, totalTakesIndex, 1, 1)
      totalTakesRange.values = lineDetails.totalTakes;
      let ukTakesRange = scriptSheet.getRangeByIndexes(index, ukTakesIndex, 1, 1);
      let ukTakeNoRange = scriptSheet.getRangeByIndexes(index, ukTakeNoIndex, 1, 1);
      ukTakesRange.values = lineDetails.ukTakes;
      if (item > lineDetails.ukTakes){
        ukTakeNoRange.values = 'N/A';
      } else {
        ukTakeNoRange.values = item;
      }
      
      let usTakesRange = scriptSheet.getRangeByIndexes(index, usTakesIndex, 1, 1);
      let usTakeNoRange = scriptSheet.getRangeByIndexes(index, usTakeNoIndex, 1, 1);
      usTakesRange.values = lineDetails.usTakes;
      if (item > lineDetails.usTakes){
        usTakeNoRange.values = 'N/A';
      } else {
        usTakeNoRange.values = item;
      }

      let wallaTakesRange = scriptSheet.getRangeByIndexes(index, wallaTakesIndex, 1, 1);
      let wallaTakeNoRange = scriptSheet.getRangeByIndexes(index, wallaTakeNoIndex, 1, 1);
      wallaTakesRange.values = lineDetails.wallaTakes;
      if (item > lineDetails.wallaTakes){
        wallaTakeNoRange.values = 'N/A';
      } else {
        wallaTakeNoRange.values = item;
      }
    }
    await excel.sync();
  });
}


async function doTakesAndNumTakes(currentRowIndex, country, doDate, doAdditional, includeMarkUp, includeStudio, includeEngineer){
  const numberColumn = findColumnLetter("Number");
  let noOfTakesIndex, dateRecordedIndex, markUpIndex, studioIndex, engineerIndex;

  if (country == "UK"){
    noOfTakesIndex = ukTakesIndex;
    dateRecordedIndex = ukDateIndex
    markUpIndex = ukMarkUpIndex;
    studioIndex = ukStudioIndex;
    engineerIndex = ukEngineerIndex;
  }
  await Excel.run(async function(excel){ 
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    let currentNumberCell = scriptSheet.getRangeByIndexes(currentRowIndex, numberIndex, 1, 1)
    currentNumberCell.load('values')
    let numberData = scriptSheet.getRange(numberColumn + firstDataRow + ":" + numberColumn + lastDataRow);
    numberData.load('values');
    await excel.sync();
    let targetValue = currentNumberCell.values
    console.log("Target Value:" + targetValue);
    let myData = numberData.values.map(x => x[0]);
    console.log("Raw values");
    console.log(numberData.values);
    console.log("Mapped values");
    console.log(myData)
    const myIndecies = myData.map((x, i) => [x, i]).filter(([x, i]) => x == targetValue).map(([x, i]) => i);
    console.log("Found Index");
    console.log(myIndecies);
    if (myIndecies.length > 0){
      let firstIndex = myIndecies[0] + firstDataRow - 1
      console.log("First Index: " + firstIndex )
      let numTakesRange = scriptSheet.getRangeByIndexes(firstIndex, noOfTakesIndex, myIndecies.length, 2)
      numTakesRange.load('address');
      await excel.sync();
      console.log("Target address: " + numTakesRange.address)
      let newValues = [];
      if (myIndecies.length == 1){
        newValues = [[1, 1]];
      } else {
        for (i = 0; i < myIndecies.length; i++){
          newValues.push([myIndecies.length, i + 1]);
        }
      }
      console.log("New values");
      console.log(newValues)
      numTakesRange.values = newValues;
      await excel.sync();
      if ((myIndecies.length > 1) && (doAdditional)){
        let rowIndex = firstIndex + myIndecies.length - 1;
        console.log("Row index: " + rowIndex);
        if (doDate){
          let dateRange = scriptSheet.getRangeByIndexes(rowIndex, dateRecordedIndex, 1, 1);
          let theDate = dateInFormat();
          dateRange.values = theDate;
        }
        if (!includeMarkUp){
          let markUpRange = scriptSheet.getRangeByIndexes(rowIndex, markUpIndex, 1, 1);
          markUpRange.clear("Contents");
        }
        if (!includeStudio){
          console.log('Studio');
          let studioRange = scriptSheet.getRangeByIndexes(rowIndex, studioIndex, 1, 1);
          studioRange.clear("Contents");
        }
        if(!includeEngineer){
          let engineerRange = scriptSheet.getRangeByIndexes(rowIndex, engineerIndex, 1, 1);
          engineerRange.clear("Contents");
        }
      }
      await excel.sync();
    }
    if (isProtected){
      await lockColumns();
    }
  });
}
async function hideRows(visibleType, country){
  let noOfTakesColumn;
  let takeNumberColumn;
  let isProtected = await unlockIfLocked();
  if (country == "UK"){
    noOfTakesColumn = findColumnLetter("UK No of takes");
    takeNumberColumn = findColumnLetter("UK Take No")
  }
  await Excel.run(async function(excel){ 
    let myMessage = tag('takeMessage')
    myMessage.innerText = "Please wait...";
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    
    /*let myRange = scriptSheet.getRange(noOfTakesColumn + firstDataRow + ":" + takeNumberColumn + lastDataRow);
    myRange.load('values')
    await excel.sync();
    console.log(myRange.values)
    console.log(myRange.values.length)
    console.log(myRange.values[0].length)
    */
    //First unhide all

    
    let tempRange = [];
    for (let i = 0; i < scriptHiddenRows.length; i++){
      tempRange[i] = scriptSheet.getRange(scriptHiddenRows[i]);
      tempRange[i].rowHidden = false;
    }
    await excel.sync();
    //scriptSheet.getUsedRange().rowHidden = false;
    myMessage.innerText = "Showing all takes";
    /*
    let hideRange = scriptSheet.getRangeByIndexes(firstDataRow - 1, 0, lastDataRow - 2, 1);
    hideRange.load('address');
    hideRange.rowHidden = false;
    await excel.sync();
    console.log(hideRange.address);
    */
    

    if (visibleType == 'last'){
      for (i = 0; i < myRange.values.length; i++){
        if (myRange.values[i][0] != ""){
          if (myRange.values[i][0] != myRange.values[i][1]){
            console.log(myRange.values[i][0]);
            console.log(myRange.values[i][1]);
            let hideRange = scriptSheet.getRangeByIndexes(i + firstDataRow - 1, 0, 1, 1);
            hideRange.load('address');
            hideRange.rowHidden = true;
            await excel.sync();
            console.log(hideRange.address);
          }
        }
      }
      myMessage.innerText = "Showing last takes"
    }
    
    if (visibleType == 'first'){
      for (i = 0; i < myRange.values.length; i++){
        if (myRange.values[i][0] != ""){
          if (myRange.values[i][1] != 1){
            console.log(myRange.values[i][0]);
            console.log(myRange.values[i][1]);
            let hideRange = scriptSheet.getRangeByIndexes(i + firstDataRow - 1, 0, 1, 1);
            hideRange.load('address');
            hideRange.rowHidden = true;
            await excel.sync();
            console.log(hideRange.address);
          }
        }
      }
      myMessage.innerText = "Showing first takes"
    }
  })
  if (isProtected){
    await lockColumns();
  }
}

async function hiddenRows(){
  await Excel.run(async function(excel){ 
    const details = await getFirstLastIndex();
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let myRows = [];
    let address = [];
    for (let i = details.rowIndex; i <= (details.rowCount - details.rowIndex + 1); i++){
      address[i] = (i+1).toString() + ':' + (i+1).toString();
      myRows[i] = scriptSheet.getRange(address[i]);
      myRows[i].load('rowHidden');
    }
    await excel.sync();
    let result = [];
    let index = -1;
    for (let i = 0; i < myRows.length; i++){
      if (myRows[i].rowHidden){
        index += 1;
        result[index] = address[i];
      }
    }
    console.log('Hidden rows:', result)
    scriptHiddenRows = result;
  })
}

async function showHideColumns(columnType){
  const sheetName = "Settings"
  const rangeName = "columnHide"
  let columnMessage = tag('columnMessage')
  let hideUnedited = tag('hideUnedited').checked;
  console.log('Hide Unedited', hideUnedited);
  await Excel.run(async function(excel){ 
    let app = excel.workbook.application;
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    app.suspendScreenUpdatingUntilNextSync();
    const settingsSheet = excel.workbook.worksheets.getItem(sheetName);
    const range = settingsSheet.getRange(rangeName);
    range.load('values');
    await excel.sync();
    app.suspendScreenUpdatingUntilNextSync();
    console.log(range.values);
    let allIndex = range.values.findIndex(x => x[0] == 'All');
    console.log(allIndex);
    let unhideColumns = range.values[allIndex][1]
    console.log(unhideColumns);
    
    const unhideColumnsRange = scriptSheet.getRange(unhideColumns);
    unhideColumnsRange.columnHidden = false;
    //await excel.sync();
    if (columnType == 'UK Script'){
      let ukIndex = range.values.findIndex(x => x[0] == 'UK Script');
      console.log(ukIndex);
      let hideUKColumns = range.values[ukIndex][2].split(",")
      console.log(hideUKColumns);
      for (let hide of hideUKColumns){
        let hideUKColumnsRange = scriptSheet.getRange(hide);
        /*
        hideUKColumnsRange.load('address');
        await excel.sync();
        console.log(hideUKColumnsRange.address);
        */
        hideUKColumnsRange.columnHidden = true;
        //await excel.sync();  
      }
    }
    if (columnType == 'US Script'){
      let usIndex = range.values.findIndex(x => x[0] == 'US Script');
      console.log(usIndex);
      let hideUSColumns = range.values[usIndex][2].split(",")
      console.log(hideUSColumns);
      for (let hide of hideUSColumns){
        let hideUSColumnsRange = scriptSheet.getRange(hide);
        /*
        hideUSColumnsRange.load('address');
        await excel.sync();
        console.log(hideUSColumnsRange.address);
        */
        hideUSColumnsRange.columnHidden = true;
        //await excel.sync();  
      }
    }
    if (columnType == 'Walla Script'){
      let wallaIndex = range.values.findIndex(x => x[0] == 'Walla Script');
      console.log(wallaIndex);
      let hideWallaColumns = range.values[wallaIndex][2].split(",")
      console.log(hideWallaColumns);
      for (let hide of hideWallaColumns){
        let hideWallaColumnsRange = scriptSheet.getRange(hide);
        /*
        hideWallaColumnsRange.load('address');
        await excel.sync();
        console.log(hideWallaColumnsRange.address);
        */
        hideWallaColumnsRange.columnHidden = true;
        //await excel.sync();  
      }
    }

    if (hideUnedited){
      let uneditedIndex = range.values.findIndex(x => x[0] == 'Unedited Script');
      console.log('Unedited column', uneditedIndex);
      let hideUneditedColumns = range.values[uneditedIndex][2].split(",")
      console.log(hideUneditedColumns);
      for (let hide of hideUneditedColumns){
        let hideUneditedColumnsRange = scriptSheet.getRange(hide);
        /*
        hideUneditedColumnsRange.load('address');
        await excel.sync();
        console.log(hideUneditedColumnsRange.address);
        */
        hideUneditedColumnsRange.columnHidden = true;
        //await excel.sync();  
      }
    } else {
      console.log('Not hiding');
    }
    await excel.sync();
    if (isProtected){
      await lockColumns();
    }
  })  
  console.log(columnMessage.innerText, columnType);
  columnMessage.innerText = 'Showing ' + columnType
}


  /* ​
  0: Array(10) [ '=IF(C3="",0,FIND("-",C3))', 0, '=IF(C3="",0,FIND("]",C3))', … ]
  ​​
  0: '=IF(C3="",0,FIND("-",C3))'
  ​​
  1: 0
  ​​
  2: '=IF(C3="",0,FIND("]",C3))'
  ​​
  3: 0
  ​​
  4: "=AND(F3>=BU3, F3<=BW3)"
  ​​
  5: 0
  ​​
  6: 1
  ​​
  7: 0
  ​​
  8: 0
  ​​
  9: 0
  ​​
  length: 10
  ​​
  <prototype>: Array []
  ​
  1: Array(10) [ '=IF(C4="",0,FIND("-",C4))', "=IF(BT4=0,BU3,VALUE(MID(C4,2,BT4-2)))", '=IF(C4="",0,FIND("]",C4))', … ]
  ​​
  0: '=IF(C4="",0,FIND("-",C4))'
  ​​
  1: "=IF(BT4=0,BU3,VALUE(MID(C4,2,BT4-2)))"
  ​​
  2: '=IF(C4="",0,FIND("]",C4))'
  ​​
  3: "=IF(BV4=0,BW3,VALUE(MID(C4,BT4+1,BV4-BT4-1)))"
  ​​
  4: "=AND(F4>=BU4, F4<=BW4)"
  ​​
  5: '= LEN(TRIM(J4)) - LEN(SUBSTITUTE(J4, " ", "")) + 1'
  =IF(NOT(OR(U4="",U4=1)), 0, LEN(TRIM(J4)) - LEN(SUBSTITUTE(J4, " ", "")) + 1)
  ​​
  6: '=IF(D4="",BZ3,VALUE(D4))'
  ​​
  7: "=F4"
  ​​
  8: "=IF(BZ4=BZ3,CB3+BY4,BY4)"
  ​​
  9: "=VLOOKUP(BW4,CA4:CB99999,2,FALSE)"
  ​​
  length: 10
  ​​
  <prototype>: Array []
  ​
  length: 2
  ​
  <prototype>: Array []
  index.html:448:13
  ]
  */ 
  /*
Scene Line Number Range	3	C	2
Scene Number	4	D	3
Cue	5	E	4
Number	6	F	5
UK Script without dialog tags	10	J	9
UK Date Recorded	22	V	21
UK Studio	23	W	22
UK Engineer	24	X	23
US Date Recorded	27	AA	26
US Studio	28	AB	27
US Engineer	29	AC	28
Walla Date Recorded	45	AS	44
Walla Studio	46	AT	45
Walla Engineer	47	AU	46
Position -	72	BT	71
Start Line	73	BU	72
Position ]	74	BV	73
End Line	75	BW	74
Valid Line Number	76	BX	75
Line Word Count	77	BY	76
Scene	78	BZ	77
Line	79	CA	78
Word count to this line	80	CB	79
Scene word count calc	81	CC	80
*/


/*
Delete row 13
Rebuild BU12, BU13, BW12, BW13, BZ12, BZ13, CB12, CB13, 
*/

function checkboxChecked(name){
  console.log(radioButtonChoice());
}

function radioButtonChoice(){
  let justDate = tag('radJustDate');
  let detailsAbove = tag('radAboveDetails');
  let detailsBelow = tag('radBelowDetails');
  const textValue = tag("markup").value;
  console.log('Markup: ', textValue);

  if (justDate.checked){
    return 'justDate';
  } else if (detailsAbove.checked){
    return 'detailsAbove';
  } else if (detailsBelow.checked){
    return 'detailsBelow'
  } else {
    return NaN;
  }
}

async function displayMinAndMax(){
  const minAndMax = await getSceneMaxAndMin();
  let display = tag('min-and-max');
  display.innerText = "(" + minAndMax.min + ".." + minAndMax.max + ")";
  const lineMinAndMax = await getLineNoMaxAndMin();
  let lineDisplay = tag('min-and-max-lineNo');
  lineDisplay.innerText = "(" + lineMinAndMax.min + ".." + lineMinAndMax.max + ")";
  const chapterMinAndMax = await getChapterMaxAndMin();
  let chapterDisplay = tag('min-and-max-chapter');
  chapterDisplay.innerText = "(" + chapterMinAndMax.min + ".." + chapterMinAndMax.max + ")";
}

async function fillSceneNumber(startRow, endRow){
  let waitLabel = tag('scene-wait');
  waitLabel.style.display = 'block';
  await Excel.run(async function(excel){ 
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let app = excel.workbook.application;
    app.suspendScreenUpdatingUntilNextSync();
    app.suspendApiCalculationUntilNextSync();
    
    const sceneBordersColumn = findColumnLetter('Scene Borders');
    const sceneLineNumberRangeColumn = findColumnLetter('Scene Line Number Range')

    console.log('startRow', startRow, 'endRow', endRow);
    if (startRow === undefined){
      startRow = firstDataRow;
    }
    if (endRow === undefined){
      endRow = lastDataRow;
    }

    let borderRange = scriptSheet.getRange(sceneBordersColumn + startRow + ":" +  sceneBordersColumn + endRow);
    let lineNoRange = scriptSheet.getRange(sceneLineNumberRangeColumn + startRow + ':' + sceneLineNumberRangeColumn + endRow);
    borderRange.load('values');
    lineNoRange.load('values')
    await excel.sync();
    console.log(lineNoRange.values)
    
    app.suspendScreenUpdatingUntilNextSync();
    app.suspendApiCalculationUntilNextSync();
    let borderValues = borderRange.values.map(x => x[0]);
    let lineNoValues = lineNoRange.values

    let currentLineNo = '';
    for (let i = 0; i < borderValues.length; i++){
      if (borderValues[i] == 'Original'){
        currentLineNo = lineNoValues[i][0];
      } else if (borderValues[i] == 'Copy'){
        lineNoValues[i][0] = currentLineNo;
      } else if(borderValues[i] == ''){
        lineNoValues[i][0] = '';
      }
    }

    lineNoRange.values = lineNoValues;
    await excel.sync();
  })
  waitLabel.style.display = 'none';
}

async function setDefaultColumnWidths(){
  await Excel.run(async function(excel){ 
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    let app = excel.workbook.application;
    app.suspendScreenUpdatingUntilNextSync();
    app.suspendApiCalculationUntilNextSync();
    for (let i = 0; i < mySheetColumns.length; i++){
      if (mySheetColumns[i].width != ''){
        let myColumn = scriptSheet.getRange(mySheetColumns[i].column + ':' + mySheetColumns[i].column);
        let myFormat = myColumn.format
        myFormat.columnWidth = (mySheetColumns[i].width * 7)
      }
    }
    await excel.sync();
    if (isProtected){
      await lockColumns();
    }
  });
}

async function setUpEvents(){
  sceneInput = tag('scene');
  lineNoInput = tag('lineNo');
  chapterInput = tag('chapter');
  sceneInput.addEventListener('keypress',async function(event){
    if (event.key === 'Enter'){
      event.preventDefault();
      await getTargetSceneNumber();
    }
  })
  lineNoInput.addEventListener('keypress',async function(event){
    if (event.key === 'Enter'){
      event.preventDefault();
      await getTargetLineNo();
    }
  })
  chapterInput.addEventListener('keypress',async function(event){
    if (event.key === 'Enter'){
      event.preventDefault();
      await getTargetChapter();
    }
  })
  console.log('Events set up')
}

function showAdmin(){
  let admin = tag('admin')
  if (admin.style.display === 'block'){
    admin.style.display = 'none';
  } else {
    admin.style.display = 'block';
  }
}

async function getCharacters(sheetName, charIndex){
  let characters
  if (charIndex === null){
    charIndex = characterIndex;
  }
  await Excel.run(async function(excel){
    sheet = excel.workbook.worksheets.getItem(sheetName); 
    let characterRange = sheet.getRangeByIndexes(firstDataRow, charIndex, lastDataRow - firstDataRow, 1);
    characterRange.load('values');
    await excel.sync()
    characters = characterRange.values;
  })
  console.log(characters);
  return characters;
}

async function filterOnCharacter(characterName){
  await Excel.run(async function(excel){
    let myRange = await getDataRange(excel);
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const myCriteria = {
      filterOn: Excel.FilterOn.custom,
      criterion1: characterName
    }
    scriptSheet.autoFilter.apply(myRange, characterIndex, myCriteria);
    myRange.load('address');
    await excel.sync();
    console.log('My range address:', myRange.address)
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let filteredRange = myRange;
    filteredRange.load('values');
    filteredRange.load('address')
    await excel.sync();
    console.log('Filtered');
    console.log(filteredRange.address)
    console.log(filteredRange.values);
  })
}

async function filterOnLocation(locationText){
  await Excel.run(async function(excel){
    let myRange = await getDataRange(excel);
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const myCriteria = {
      filterOn: Excel.FilterOn.custom,
      criterion1: '=*' + locationText +'*'
    }
    scriptSheet.autoFilter.apply(myRange, locationIndex, myCriteria);
    myRange.load('address');
    await excel.sync();
    console.log('My range address:', myRange.address)
    let filteredRange = myRange;
    filteredRange.load('values');
    filteredRange.load('address')
    await excel.sync();
    console.log('Filtered');
    console.log(filteredRange.address)
    console.log(filteredRange.values);
  })
}

async function doChunkedFilter(character, sheetName){
  let myAddresses;
  let hiddenColumnAddresses = await getHiddenColumns();
  //This does a filter based on a parameter criteria and a chunking of lineNo
  //It returns a string array of addresses
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(sheetName);
    let usedRange = await getDataRange(excel);
    usedRange.load('address');
    usedRange.columnHidden = false;
    await excel.sync()

    //find the min and max for column G
    let minAndMax = await getLineNoMaxAndMin();

    // set up loop variables
    let chunkLength = 500;
    let startChunk = minAndMax.min;
    let endChunk = startChunk + chunkLength;

    //set up the character criteria
    let myCriteria = {};
    
    console.log('Character Object: ', character);
    if (character.type == choiceType.text){
      myCriteria ={
        filterOn: Excel.FilterOn.custom,
        criterion1: '=*' + character.name +'*'
      }
    } else {
      myCriteria = {
        filterOn: Excel.FilterOn.custom,
        criterion1: character.name
      }
    }
    let myNumberCriteria = {};

    //set up loop boolean
    let doChunk = true;

    //start the loop
    let tempArray = [];
    while(doChunk){
      // set up end loop
      endChunk = startChunk + chunkLength;
      //Check end condition
      if (endChunk > minAndMax.max){
        endChunk = minAndMax.max;
        doChunk = false;
      }
      console.log('StartChunk: ', startChunk, 'After endChunk:', endChunk, 'doChunk:', doChunk);
      // set up loop criteria
      myNumberCriteria = {
        filterOn: Excel.FilterOn.custom,
        criterion1: '>=' + startChunk,
        criterion2: '<' + endChunk,
        operator: 'And'
      }
      //remove the autofilter
      scriptSheet.autoFilter.remove();
      //apply filters to both columns
      console.log('Used Range: ', usedRange.address, 'characterIndex:', characterIndex, 'myCriteria: ', myCriteria);
      scriptSheet.autoFilter.apply(usedRange, characterIndex, myCriteria);
      scriptSheet.autoFilter.apply(usedRange, numberIndex, myNumberCriteria);
      // get the formula range for this chunk
      let formulaRanges = usedRange.getSpecialCellsOrNullObject('Visible');
      formulaRanges.load('address, areas');
      await excel.sync();
      if (formulaRanges.isNullObject){
        console.log('Null object detected');
      } else {
        console.log('Null object NOT detected');
      }
      console.log('Range areas', formulaRanges.address, 'Areas object', formulaRanges.areas.items);
      let formulaRangeItems = formulaRanges.areas.items;
      let theAddresses = []
      for (let i = 0; i < formulaRangeItems.length; i++){
        theAddresses[i] = formulaRangeItems[i].address;
      }
      console.log('theAddresses', theAddresses);
      tempArray = tempArray.concat(theAddresses);
      /*
      if (!(formulaRanges === null)){
        tempArray = tempArray.concat(formulaRanges.address.split(','));
        console.log('concataned:', tempArray)
      }
      */
      //increment the loop
      startChunk += chunkLength;
    }
    myAddresses = [...new Set(tempArray)];
    scriptSheet.autoFilter.remove();
    for (let col of hiddenColumnAddresses){
      let tempRange = scriptSheet.getRange(col);
      tempRange.columnHidden = true;
    }
  })
  return myAddresses;
}

async function getDirectorData(character){
  let myData = [];
  //let hiddenColumnAddresses = await getHiddenColumns();
  
	await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    /*
		let usedRange = await getDataRange(excel);


    let myFilter = scriptSheet.autoFilter
    myFilter.load('criteria');
    await excel.sync();
    
    console.log('The criteria: ', myFilter.criteria)

    usedRange.load('address');
    usedRange.columnHidden = false;
    await excel.sync()
    */
    let app = excel.workbook.application;
    app.suspendScreenUpdatingUntilNextSync();
    //console.log('Used range address', usedRange.address)
    /*
    let doChunking = false;
    //find the min and max for column G
    let minAndMax = getLineNoMaxAndMin();
    if ((minAndMax.max - minAndMax.min) > 10000){
      doChunking = true;
    }

    let chunkLength = 1000;
    let startChunk = minAndMax.min;
    let endChunk = startChunk + chunkLength;

    


    let myCriteria
    if (character.type == choiceType.list){
      myCriteria = {
        filterOn: Excel.FilterOn.custom,
        criterion1: character.name
      }
    } else {
      if (doChunking){
        myNumberCriteria = {
          filterOn: Excel.FilterOn.custom,
          criterion1: '>=' + startNumber,
          criterion2: '<=' + endNumber,
          operator: 'And'
        }
        myCharacterCriteria = {
          filterOn: Excel.FilterOn.custom,
          criterion1: '=*' + character.name +'*'
        } 
      } else {
        myCriteria = {
          filterOn: Excel.FilterOn.custom,
          criterion1: '=*' + character.name +'*'
        }  
      }
      
    }
    
    
    scriptSheet.autoFilter.apply(usedRange, characterIndex, myCriteria);
		let formulaRanges = usedRange.getSpecialCells("Visible");
    formulaRanges.load('address');
    formulaRanges.load('cellCount');
    formulaRanges.load('areaCount');
    formulaRanges.load('areas')
		await excel.sync();
    //app.suspendScreenUpdatingUntilNextSync();
    console.log('Areas:', formulaRanges.areas.toJSON());
    console.log('Range areas', formulaRanges.address);
    console.log('Cell count', formulaRanges.cellCount);
    console.log('Area count', formulaRanges.areaCount);
    
    let myAddresses = formulaRanges.address.split(",");
    console.log('myAddresses', myAddresses);
    
    scriptSheet.autoFilter.remove();
    for (let col of hiddenColumnAddresses){
      let tempRange = scriptSheet.getRange(col);
      tempRange.columnHidden = true;
    }
    */

    let myAddresses = await doChunkedFilter(character, scriptSheetName);
    console.log('myAddresses', myAddresses);

    let startIndex = 0;
    let stopIndex = 200;
    let doIt = true;
    let theRanges = [];

    while (doIt){
      if (stopIndex > myAddresses.length){
        stopIndex = myAddresses.length;
        doIt = false;
      }
      console.log('startIndex', startIndex, 'stopIndex', stopIndex);
    
      for (let i = startIndex; i < stopIndex; i++){
        theRanges[i] = scriptSheet.getRange(myAddresses[i]);
        theRanges[i].load('values');
        theRanges[i].load('rowIndex');
        theRanges[i].load('rowCount');
      }
      await excel.sync();
      startIndex = startIndex + 200;
      stopIndex = stopIndex + 200;
    }
    console.log('theRanges', theRanges)
    /*
    for (let i = 0; i < theRanges.length; i++){
      console.log('Range items', i, theRanges[i].values);
    }
    */
    let results = [];
    console.log('The Ranges length', theRanges.length);
    for (let i = 0; i < theRanges.length; i++){
      console.log('i', i, 'numRows', theRanges[i].values.length, theRanges[i].rowCount);
      for (let myRow = 0; myRow < theRanges[i].values.length; myRow++){
        //console.log(i, myRow, theRanges[i].rowIndex, theRanges[i].rowCount, theRanges[i].values[myRow]);
        let newItem = {
          rowIndex: theRanges[i].rowIndex,
          myItems: theRanges[i].values[myRow]
        }
        results.push(newItem);
      }
    }
    console.log('Results', results);

    let headings = results.find(head => head.rowIndex == 1);
    console.log('Headings', headings);

    let sceneArrayIndex = headings.myItems.findIndex(x => x == 'Scene Number');
    let numberArrayIndex = headings.myItems.findIndex(x => x == 'Number');
    let numUkTakesArrayIndex = headings.myItems.findIndex(x => x == 'UK No of takes');
    let ukTakeNumArrayIndex = headings.myItems.findIndex(x => x == 'UK Take No');
    let ukDateArrayIndex = headings.myItems.findIndex(x => x == "UK Date Recorded");
    let lineWordCountArrayIndex = headings.myItems.findIndex(x => x == 'Line Word Count');
    let sceneWordCountArrayIndex = headings.myItems.findIndex(x => x == 'Scene Word Count');
    let characterArrayIndex = headings.myItems.findIndex(x => x == 'Character')
    console.log('Scene Index', sceneArrayIndex, 'Number Index', numberArrayIndex);

    for (let result of results){
      if (result.rowIndex != 1){
        if(result.myItems[sceneArrayIndex] != ""){
          let theData = {
            sceneNumber: result.myItems[sceneArrayIndex],
            lineNumber: result.myItems[numberArrayIndex],
            ukNumTakes: result.myItems[numUkTakesArrayIndex],
            ukTakeNum: result.myItems[ukTakeNumArrayIndex],
            ukDateRecorded: result.myItems[ukDateArrayIndex],
            lineWordCount: result.myItems[lineWordCountArrayIndex],
            sceneWordCount: result.myItems[sceneWordCountArrayIndex],
            character: result.myItems[characterArrayIndex]
          }
          myData.push(theData);  
        }
      }
    }
    console.log('myData', myData);
    if (isProtected){
      await lockColumns();
    }
  })
  console.log('directors myData', myData);
  return myData;
}

async function gatherActorsforScene(sceneNumberArray){
  let myData = [];
  let hiddenColumnAddresses = await getHiddenColumns();
  
	await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
		let usedRange = await getDataRange(excel);
    usedRange.load('address');
    usedRange.columnHidden = false;
    await excel.sync()
    
    let startIndex = firstDataRow - 1;
    let rowCount = lastDataRow - firstDataRow + 1;
    let sceneRange = scriptSheet.getRangeByIndexes(startIndex, sceneIndex, rowCount, 1);
    let characterRange = scriptSheet.getRangeByIndexes(startIndex, characterIndex, rowCount, 1);
    sceneRange.load('values, rowIndex');
    characterRange.load('values, rowIndex');
    await excel.sync();
    //console.log('Scene:', sceneRange.values, 'rowIndex', sceneRange.rowIndex);
    //console.log('Character:', characterRange.values, 'rowIndex', characterRange.rowIndex);
    let sceneValues = sceneRange.values.map(x => x[0]);
    let characterValues = characterRange.values.map(x => x[0]);
    //console.log('Scene Values', sceneValues);

    for (let a = 0; a < sceneNumberArray.length; a++){
      let myIndecies = sceneValues.map((x, i) => [x, i]).filter(([x, i]) => x == sceneNumberArray[a]).map(([x, i]) => i);
      //console.log('Scene Indecies', myIndecies);
      let characterArray = [];
      let characterIndex = -1;
      for (let s = 0; s < myIndecies.length; s++){
        let thisCharacter = characterValues[myIndecies[s]];
        if (thisCharacter != ''){
          characterIndex += 1
          characterArray[characterIndex] = thisCharacter;
        }
      }
      //console.log('Character Array', characterArray);
      let sortedArray = Array.from(new Set(characterArray)).sort();
      //console.log('Sorted array', sortedArray);
      let newData = {
        index: a,
        rowIndex: a + sceneRange.rowIndex,
        scene: sceneNumberArray[a],
        characters: sortedArray
      }
      myData.push(newData);
    }
    for (let col of hiddenColumnAddresses){
      let tempRange = scriptSheet.getRange(col);
      tempRange.columnHidden = true;
    }
    await excel.sync();
    if (isProtected){
      await lockColumns();
    }
  })
  //console.log('myData about to return', myData)
  return myData;
}

async function getLocationData(locationText){
  let myData = [];
  let hiddenColumnAddresses = await getHiddenColumns();
  
	await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
		let usedRange = await getDataRange(excel);
    usedRange.load('address');
    usedRange.columnHidden = false;
    await excel.sync()
    let app = excel.workbook.application;
    app.suspendScreenUpdatingUntilNextSync();
    console.log('Used range address', usedRange.address)
    const myCriteria = {
      filterOn: Excel.FilterOn.custom,
      criterion1: '=*' + locationText +'*'
    }
    scriptSheet.autoFilter.apply(usedRange, locationIndex, myCriteria);
		let formulaRanges = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.visible);
    formulaRanges.load('address');
    formulaRanges.load('cellCount');
    formulaRanges.load('areaCount');
    //formulaRanges.load('areas')
		await excel.sync();
    app.suspendScreenUpdatingUntilNextSync();
    console.log('Range areas', formulaRanges.address);
    console.log('Cell count', formulaRanges.cellCount);
    console.log('Area count', formulaRanges.areaCount);
    
    let myAddresses = formulaRanges.address.split(",");
    console.log('myAddresses', myAddresses);
    
    
    let theRanges = [];
    for (let i = 0; i < myAddresses.length; i++){
      theRanges[i] = scriptSheet.getRange(myAddresses[i]);
      theRanges[i].load('values');
      theRanges[i].load('rowIndex');
      theRanges[i].load('rowCount');
    }
    
    scriptSheet.autoFilter.remove();
    for (let col of hiddenColumnAddresses){
      let tempRange = scriptSheet.getRange(col);
      tempRange.columnHidden = true;
    }
    await excel.sync();
    console.log(theRanges)
    /*
    for (let i = 0; i < theRanges.length; i++){
      console.log('Range items', i, theRanges[i].values);
    }
    */
    let results = [];

    for (let i = 0; i < theRanges.length; i++){
      //console.log('numRows', theRanges[i].values.length, theRanges[i].rowCount);
      for (let myRow = 0; myRow < theRanges[i].values.length; myRow++){
        //console.log(i, myRow, theRanges[i].rowIndex, theRanges[i].rowCount, theRanges[i].values[myRow]);
        let newItem = {
          rowIndex: theRanges[i].rowIndex,
          myItems: theRanges[i].values[myRow]
        }
        results.push(newItem);
      }
    }
    console.log('Results', results);

    let headings = results.find(head => head.rowIndex == 1);
    console.log('Headings', headings);

    let sceneArrayIndex = headings.myItems.findIndex(x => x == 'Scene Number');
    let numberArrayIndex = headings.myItems.findIndex(x => x == 'Number');
    let locationArrayIndex = headings.myItems.findIndex(x => x == 'Location');
    
    console.log('Scene Index', sceneArrayIndex, 'Number Index', numberArrayIndex, 'Location index', locationArrayIndex);

    for (let result of results){
      if (result.rowIndex != 1){
        if(result.myItems[sceneArrayIndex] != ""){
          let theData = {
            sceneNumber: result.myItems[sceneArrayIndex],
            lineNumber: result.myItems[numberArrayIndex],
            location: removeDoubleLf(result.myItems[locationArrayIndex])
          }
          myData.push(theData);  
        }
      }
    }
    console.log('myData', myData);
    if (isProtected){
      await lockColumns();
    }
  })
  console.log('directors myData', myData);
  return myData;
};

async function getLocations(){
  let myData = [];
  let hiddenColumnAddresses = await getHiddenColumns();
	await Excel.run(async (excel) => {
		let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
		let usedRange = await getDataRange(excel);
    usedRange.load('address');
    usedRange.columnHidden = false;
    await excel.sync()
    let app = excel.workbook.application;
    app.suspendScreenUpdatingUntilNextSync();
    console.log('Used range address', usedRange.address)
    const myCriteria = {
      filterOn: Excel.FilterOn.custom,
      criterion1: '<>'
    }
    scriptSheet.autoFilter.apply(usedRange, locationIndex, myCriteria);
		let formulaRanges = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.visible);
    formulaRanges.load('address');
    formulaRanges.load('cellCount');
    formulaRanges.load('areaCount');
    //formulaRanges.load('areas')
		await excel.sync();
    app.suspendScreenUpdatingUntilNextSync();
    console.log('Range areas', formulaRanges.address);
    console.log('Cell count', formulaRanges.cellCount);
    console.log('Area count', formulaRanges.areaCount);
    
    let myAddresses = formulaRanges.address.split(",");
    console.log('myAddresses', myAddresses);
    
    let theRanges = [];
    for (let i = 0; i < myAddresses.length; i++){
      theRanges[i] = scriptSheet.getRange(myAddresses[i]);
      theRanges[i].load('values');
      theRanges[i].load('rowIndex');
      theRanges[i].load('rowCount');
    }
    
    scriptSheet.autoFilter.remove();
    for (let col of hiddenColumnAddresses){
      let tempRange = scriptSheet.getRange(col);
      tempRange.columnHidden = true;
    }
    await excel.sync();
    console.log(theRanges)
    /*
    for (let i = 0; i < theRanges.length; i++){
      console.log('Range items', i, theRanges[i].values);
    }
    */
    let results = [];
    for (let i = 0; i < theRanges.length; i++){
      //console.log('numRows', theRanges[i].values.length, theRanges[i].rowCount);
      for (let myRow = 0; myRow < theRanges[i].values.length; myRow++){
        //console.log(i, myRow, theRanges[i].rowIndex, theRanges[i].rowCount, theRanges[i].values[myRow]);
        let newItem = {
          rowIndex: theRanges[i].rowIndex,
          myItems: theRanges[i].values[myRow]
        }
        results.push(newItem);
      }
    }
    console.log('Results', results);

    let headings = results.find(head => head.rowIndex == 1);
    console.log('Headings', headings);

    let sceneArrayIndex = headings.myItems.findIndex(x => x == 'Scene Number');
    let locationArrayIndex = headings.myItems.findIndex(x => x == 'Location');
    
    for (let result of results){
      if (result.rowIndex != 1){
        let theData = {
          sceneNumber: result.myItems[sceneArrayIndex],
          location: result.myItems[locationArrayIndex],
        }
        myData.push(theData);
      }
    }
    console.log('myData', myData);
    if (isProtected){
      await lockColumns();
    }
  })
  console.log('get Locations myData', myData);
  return myData;
}

async function getHiddenColumns(){
  let results = [];
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const myUsedRange = scriptSheet.getUsedRange();
    myUsedRange.load('columnIndex, columnCount, address');
    await excel.sync();
    console.log('columnIndex', myUsedRange.columnIndex, 'columnCount', myUsedRange.columnCount);
    let temps = [];
    for (let i = myUsedRange.columnIndex; i < myUsedRange.columnCount; i++){
      temps[i] = myUsedRange.getCell(0,i);
      temps[i].load('address, columnHidden');
    }
    await excel.sync();
    for (let temp of temps){
      if (temp.columnHidden){
        results.push(temp.address);
      }
    }
  });
  console.log('hidden columns', results);
  return results;
}
async function showForDirector(){
  const mainPage = tag('main-page');
  mainPage.style.display = 'none';
  const forDirectorPage = tag('for-director-page');
  forDirectorPage.style.display = 'block';
  const forActorsPage = tag('for-actor-page');
  forActorsPage.style.display = 'none';
  const forSchedulingPage = tag('for-scheduling-page');
  forSchedulingPage.style.display = 'none';
  const wallaImportPage = tag('walla-import-page');
  wallaImportPage.style.display = 'none';
  const locationPage = tag('location-page');
  locationPage.style.display = 'none';
  const actorScriptPage = tag('script-page');
  actorScriptPage.style.display = 'none';
  styleScriptController('director');
  await Excel.run(async function(excel){
    let ForDirectorSheet = excel.workbook.worksheets.getItem(forDirectorName);
    ForDirectorSheet.activate();
  })
}
async function showWallaImportPage(){
  const mainPage = tag('main-page');
  mainPage.style.display = 'none';
  const forDirectorPage = tag('for-director-page');
  forDirectorPage.style.display = 'none';
  const forActorsPage = tag('for-actor-page');
  forActorsPage.style.display = 'none';
  const forSchedulingPage = tag('for-scheduling-page');
  forSchedulingPage.style.display = 'none';
  const wallaImportPage = tag('walla-import-page');
  wallaImportPage.style.display = 'block';
  let loadMessage = tag('load-message');
  loadMessage.style.display = 'none';
  const locationPage = tag('location-page');
  locationPage.style.display = 'none';
  const actorScriptPage = tag('script-page');
  actorScriptPage.style.display = 'none';
  styleScriptController('walla');
  await Excel.run(async function(excel){
    let wallaImportSheet = excel.workbook.worksheets.getItem(wallaImportName);
    wallaImportSheet.activate();
  })
}
async function showActorScript(){
  console.log('Start of show actor script')
  const mainPage = tag('main-page');
  mainPage.style.display = 'none';
  const forDirectorPage = tag('for-director-page');
  forDirectorPage.style.display = 'none';
  const forActorsPage = tag('for-actor-page');
  forActorsPage.style.display = 'none';
  const forSchedulingPage = tag('for-scheduling-page');
  forSchedulingPage.style.display = 'none';
  const wallaImportPage = tag('walla-import-page');
  wallaImportPage.style.display = 'none';
  let loadMessage = tag('load-message');
  loadMessage.style.display = 'none';
  const locationPage = tag('location-page');
  locationPage.style.display = 'none';
  const actorScriptPage = tag('script-page');
  actorScriptPage.style.display = 'block';
  styleScriptController('actorScript');
  await Excel.run(async function(excel){
    let actorScriptSheet = excel.workbook.worksheets.getItem(actorScriptName);
    actorScriptSheet.activate();
  })
}

async function showMainPage(){
  console.log('Showing Main Page')
  const mainPage = tag('main-page');
  mainPage.style.display = 'block';
  const forDirectorPage = tag('for-director-page');
  forDirectorPage.style.display = 'none';
  const forActorsPage = tag('for-actor-page');
  forActorsPage.style.display = 'none';
  const forSchedulingPage = tag('for-scheduling-page');
  forSchedulingPage.style.display = 'none';
  const wallaImportPage = tag('walla-import-page');
  wallaImportPage.style.display = 'none';
  const locationPage = tag('location-page');
  locationPage.style.display = 'none';
  const versionInfo = tag('sheet-version');
  const actorScriptPage = tag('script-page');
  actorScriptPage.style.display = 'none';
  styleScriptController('main')
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    scriptSheet.activate();
    settingsSheet = excel.workbook.worksheets.getItem(settingsSheetName);
    let versionRange = settingsSheet.getRange('seVersion');
    await excel.sync();
    let dateRange = settingsSheet.getRange('seDate')
    await excel.sync();
    versionRange.load('values');
    await excel.sync();
    dateRange.load('text');
    await excel.sync();
    console.log(dateRange.text);
    let versionString = 'Version ' + versionRange.values + ' Code: ' + codeVersion + ' Released: ' + dateRange.text;
    versionInfo.innerText = versionString;
    scriptSheet.protection.load('protected');
    await excel.sync();
    let protectionText = tag('lockMessage')
    if (scriptSheet.protection.protected){
      protectionText.innerText = 'Sheet locked'
    } else {
      protectionText.innerText = 'Sheet unlocked'
    }
  })
}
async function showForActorsPage(){
  const mainPage = tag('main-page');
  mainPage.style.display = 'none';
  const forDirectorPage = tag('for-director-page');
  forDirectorPage.style.display = 'none';
  const forActorsPage = tag('for-actor-page');
  forActorsPage.style.display ='block';
  const forSchedulingPage = tag('for-scheduling-page');
  forSchedulingPage.style.display = 'none';
  const wallaImportPage = tag('walla-import-page');
  wallaImportPage.style.display = 'none';
  const locationPage = tag('location-page');
  locationPage.style.display = 'none';
  const actorScriptPage = tag('script-page');
  actorScriptPage.style.display = 'none';
  styleScriptController('actor');
  await Excel.run(async function(excel){
    let actorsSheet = excel.workbook.worksheets.getItem(forActorsName);
    actorsSheet.activate();
  })
  await jade_modules.scheduling.displayScenes();
}

function styleScriptController(theme){
  console.log('Theme', theme, screenColours[theme]);
  const scriptController = tag('Script-Controller');
  scriptController.style.backgroundColor = screenColours[theme].background;
  scriptController.style.height = '100vh';
  console.log('Color before:', scriptController.style.color);
  scriptController.style.color = screenColours[theme].fontColour;
  console.log('Color after:', scriptController.style.color);
}

async function showForSchedulingPage(){
  const mainPage = tag('main-page');
  mainPage.style.display = 'none';
  const forDirectorPage = tag('for-director-page');
  forDirectorPage.style.display = 'none';
  const forActorsPage = tag('for-actor-page');
  forActorsPage.style.display = 'none';
  const forSchedulingPage = tag('for-scheduling-page');
  forSchedulingPage.style.display = 'block';
  const wallaImportPage = tag('walla-import-page');
  wallaImportPage.style.display = 'none';
  const locationPage = tag('location-page');
  locationPage.style.display = 'none';
  const actorScriptPage = tag('script-page');
  actorScriptPage.style.display = 'none';
  styleScriptController('scheduling');
  await Excel.run(async function(excel){
    let schedulingSheet = excel.workbook.worksheets.getItem(forSchedulingName);
    schedulingSheet.activate();
  })
}
async function showLocation(){
  const mainPage = tag('main-page');
  mainPage.style.display = 'none';
  const forDirectorPage = tag('for-director-page');
  forDirectorPage.style.display = 'none';
  const forActorsPage = tag('for-actor-page');
  forActorsPage.style.display = 'none';
  const forSchedulingPage = tag('for-scheduling-page');
  forSchedulingPage.style.display = 'none';
  const wallaImportPage = tag('walla-import-page');
  wallaImportPage.style.display = 'none';
  const locationPage = tag('location-page');
  locationPage.style.display = 'block';
  const locationWait = tag('location-wait');
  locationWait.style.display = 'none';
  const actorScriptPage = tag('script-page');
  actorScriptPage.style.display = 'none';
  styleScriptController('location');
  await Excel.run(async function(excel){
    let locationSheet = excel.workbook.worksheets.getItem(locationSheetName);
    locationSheet.activate();
  })
}

async function registerExcelEvents(){
  await Excel.run(async (excel) => {
    const directorSheet = excel.workbook.worksheets.getItem(forDirectorName);
    directorSheet.onChanged.add(handleChange);
    const actorsSheet = excel.workbook.worksheets.getItem(forActorsName);
    actorsSheet.onChanged.add(handleActor); 
    actorsSheet.onSelectionChanged.add(actorSelectionChange);
    await excel.sync();
    const schedulingSheet = excel.workbook.worksheets.getItem(forSchedulingName);
    schedulingSheet.onChanged.add(handleScheduling);
    const locationSheet = excel.workbook.worksheets.getItem(locationSheetName);
    locationSheet.load('name');
    await excel.sync();
    console.log('Sheet name', locationSheet)
    locationSheet.onChanged.add(handleLocation);
    locationSheet.onSelectionChanged.add(handleSelection)
    await excel.sync();
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    scriptSheet.onRowHiddenChanged.add(handleRowHide)
    console.log("Event handler successfully registered for onChanged event for four sheets.");
  });
}

async function handleRowHide(event){
  await Excel.run(async (excel) => {
    await excel.sync();
    console.log('Row hide', event);
    if (event.source == 'local'){
      if(event.changeType == 'Unhidden'){
        console.log('script Hidden Rows', scriptHiddenRows);
        let index = scriptHiddenRows.findIndex(event.address);
        if (index != -1){
          scriptHiddenRows.splice(index, 1);
          console.log('Index removed:', index, 'address', event.address, ' scriptHiddenRows', scriptHiddenRows);
        }
      }
    }
  })
}

async function handleChange(event) {
  await Excel.run(async (excel) => {
      await excel.sync();        
      if ((event.address == 'C6') && event.source == 'Local'){
        await jade_modules.scheduling.getDirectorInfo();
      }
  }).catch(errorHandlerFunction(e));
}
async function handleActor(event) {
  await Excel.run(async (excel) => {
      await excel.sync();        
      if ((event.address == 'D6') && event.source == 'Local'){
        await jade_modules.scheduling.searchCharacter();
      }
      if((event.address == 'D4') && event.source == 'Local'){
        await jade_modules.scheduling.searchCharacter();
      }
  }).catch(errorHandlerFunction(e));
}

async function handleScheduling(event) {
  await Excel.run(async (excel) => {
      await excel.sync();        
      if (((event.address == 'D4')||(event.address == 'D6')) && event.source == 'Local'){
        await jade_modules.scheduling.getForSchedulingInfo();
      }
  }).catch(errorHandlerFunction(e));
}

async function handleLocation(event) {
  await Excel.run(async (excel) => {
      await excel.sync();        
      if ((event.address == 'C6') && event.source == 'Local'){
        console.log('I got here')
        await jade_modules.scheduling.getLocationInfo();
      }
  }).catch(errorHandlerFunction(e));
}

async function handleSelection(event) {
  await Excel.run(async (excel) => {
      await excel.sync();        
      console.log('I got here, event', event)
  }).catch(errorHandlerFunction(e));
}

async function actorSelectionChange(event){
  await Excel.run(async (excel) => {
    await excel.sync();
    console.log('actorSelectionChange 1', event);
    await jade_modules.scheduling.displayScenes();
  }).catch(errorHandlerFunction(e))
}
function errorHandlerFunction(e){
  console.log('I have an error')
  console.log(e)
}

async function createTypeCodes(){
  let isProtected = await unlockIfLocked();
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const positionChapterColumn = findColumnLetter("Position Chapter"); 
    let chapterIndicies = await getIndices(positionChapterColumn, '<>', '');
    let resultArray = initialiseMyArray();
    resultArray = addValuesToArray(resultArray, chapterIndicies, myTypes.chapter, true);

    let sceneBordersColumn = findColumnLetter('Scene Borders');
    let sceneIndicies = await getIndices(sceneBordersColumn, 'textEquals', 'Original');
    resultArray = addValuesToArray(resultArray, sceneIndicies, myTypes.scene, false);
    const typeCodeColumn = findColumnLetter("Type Code"); 
    let sceneBlockIndicies = await getIndices(typeCodeColumn, 'textEquals', myTypes.sceneBlock)
    resultArray = addValuesToArray(resultArray, sceneBlockIndicies, myTypes.sceneBlock, true);

    let ukScriptColumn = findColumnLetter('UK script');
    let wallaScriptIndicies = await getIndices(ukScriptColumn, "textEquals", 'WALLA SCRIPTED LINES');
    resultArray = addValuesToArray(resultArray, wallaScriptIndicies, myTypes.wallaScripted, false);

    wallaScriptIndicies = await getIndices(ukScriptColumn, "textEquals", 'WALLA SCRIPTED LINES?');
    resultArray = addValuesToArray(resultArray, wallaScriptIndicies, myTypes.wallaScripted, false);

    let cueColumn = findColumnLetter('Cue');
    let cueIndicies = await getIndices(cueColumn, '<>', '');
    resultArray = addValuesToArray(resultArray, cueIndicies, myTypes.line, false);

    let typeCodeRange = scriptSheet.getRange(typeCodeColumn + firstDataRow + ":" + typeCodeColumn + lastDataRow);
    typeCodeRange.values = resultArray;
    await excel.sync();
  })
  if (isProtected){
    await lockColumns();
  }
}

async function getIndices(theColumn, test, testValue){
  let results = [];
  console.log('testValue', testValue, theColumn);
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    console.log(theColumn + firstDataRow + ":" + theColumn + lastDataRow);
    let theRange = scriptSheet.getRange(theColumn + firstDataRow + ":" + theColumn + lastDataRow);
    theRange.load('values');
    await excel.sync();
    let theValues = theRange.values.map(x => x[0]);
    console.log('The Values', theValues);
    let index = -1;
    for (let i = 0; i< theValues.length; i++){
      let doIt = false;
      if (test == "textEquals"){
        doIt = theValues[i].toLowerCase() == testValue.toLowerCase();
      } else if (test == "textNotEquals"){
        doIt = theValues[i].toLowerCase() != testValue.toLowerCase();
      } else if (test == "equals"){
        doIt = theValues[i] == testValue;
      } else if (test == "<>"){
        doIt = theValues[i] != testValue;
      }
      if (doIt){
        index += 1
        results[index] = i;
      }
    }
    console.log('initial results', results)
  })
  return results
}

function initialiseMyArray(){
  let resultArray = []
  for (let i = 0; i <= lastDataRow - firstDataRow; i++ ){
    resultArray[i] = [''];
  } 
  return resultArray;
}

function addValuesToArray(myArray, myIndicies, theValue, replaceExisting){
  console.log('My indicies', myIndicies);
  for (let i = 0; i < myIndicies.length; i++){
    console.log('i', i, 'myIndicies[i]', myIndicies[i], 'myArray', myArray[myIndicies[i]][0]);
    if (myArray[myIndicies[i]][0] == ''){
      myArray[myIndicies[i]][0] = theValue;
    } else if (replaceExisting){
      myArray[myIndicies[i]][0] = theValue;
    }
    console.log('After => myArray', myArray[myIndicies[i]][0]);
  }
  console.log('myArray', myArray)
  return myArray;
}
async function selectChapterCellAtRowIndex(excel, sheet, rowIndex, isScene){
  const rowOffset = 5;
  let myCell = sheet.getRangeByIndexes(rowIndex + rowOffset, chapterIndex, 1, 1)
  myCell.select()
  await excel.sync();
  if (rowIndex - rowOffset > 0){
    myCell = sheet.getRangeByIndexes(rowIndex - rowOffset, chapterIndex, 1, 1)
    myCell.select()
    await excel.sync();
  }
  if (isScene){
    myCell = sheet.getRangeByIndexes(rowIndex, chapterIndex, 1, 1)
  } else {
    myCell = sheet.getRangeByIndexes(rowIndex, chapterIndex, 1, 1)
  }
  myCell.select()
  await excel.sync();
}
async function goSceneChapter(){
  const addChapterValue = tag("chapter-scene-select").value;
  let chapterSceneID = parseInt(addChapterValue);
  if (!isNaN(chapterSceneID)){
    let sceneListData = addSelectList[chapterSceneID]
    await Excel.run(async (excel) => {
      let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      await selectChapterCellAtRowIndex(excel, scriptSheet, addSelectList[chapterSceneID].rowIndex, false);
    });
  }   
}

async function goWallaScene(){
  const addSceneValue = tag("walla-scene").value;
  let sceneID = parseInt(addSceneValue);
  if (!isNaN(sceneID)){
    await Excel.run(async (excel) => {
      let range = await getSceneRange(excel);
      range.load("values");
      const activeCell = excel.workbook.getActiveCell();
      activeCell.load("rowIndex");
      await excel.sync();
      const startRow = activeCell.rowIndex;
      let currentValue = range.values[startRow - 2][0];
      myIndex = range.values.findIndex(a => a[0] == (currentValue));
      let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      await selectChapterCellAtRowIndex(excel, scriptSheet, myIndex, false);
    });
  }   
}

async function doTheActualSceneBlock(chapterSceneID){
  if (!isNaN(chapterSceneID)){
    let sceneListData = addSelectList[chapterSceneID]
    //console.log('typeCodeValues', typeCodeValues, 'addSelectList', addSelectList);
    //console.log('Item', sceneListData.display, sceneListData.rowIndex);
    await Excel.run(async (excel) => {
      let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      await selectChapterCellAtRowIndex(excel, scriptSheet, addSelectList[chapterSceneID].rowIndex, (addSelectList[chapterSceneID].type == myTypes.scene))
      let cueColumnIndex = findColumnIndex('Cue');
      let usScriptColumnIndex = findColumnIndex('US Script');
      sceneBlockColumns =  usScriptColumnIndex - cueColumnIndex + 1
      let theRowIndex = sceneListData.rowIndex
      let nextIndex = sceneListData.arrayIndex + 1;
      let previousIndex = sceneListData.arrayIndex - 1;
      
      //console.log('The Row Index', theRowIndex, 'nextIndex (of array)', nextIndex, 'previous', previousIndex)
        
      let nextRowType = typeCodeValues.typeCodes.values[nextIndex];
      let previousRowType = typeCodeValues.typeCodes.values[previousIndex];
      //console.log('Found: rowIndex', theRowIndex, 'Next code:', nextRowType);
      let newRowIndex;
      sceneBlockColumns =  usScriptColumnIndex - cueColumnIndex + 1
      if (sceneListData.type == myTypes.scene){
        let sceneDataArray;
        if (previousRowType == myTypes.sceneBlock){
          //check there are 4 of them
          let numActualSceneBlockRows = 0;
          for (i = previousIndex; i > previousIndex - 30; i--){
            //console.log(i, typeCodeValues.typeCodes.values[i]);
            if (typeCodeValues.typeCodes.values[i] == myTypes.sceneBlock){
              numActualSceneBlockRows += 1;
            } else {
              break;
            }
          }
          let sceneDataArray = await getSceneBlockData(theRowIndex, numActualSceneBlockRows);
          //console.log('numActualSceneBlockRows', numActualSceneBlockRows)
          if (numActualSceneBlockRows == sceneBlockRows){
            newRowIndex = theRowIndex - sceneBlockRows;
            let myMergeRange = scriptSheet.getRangeByIndexes(newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            myMergeRange.load('address');
            myMergeRange.clear("Contents");
            let mergedAreas = myMergeRange.getMergedAreasOrNullObject();
            mergedAreas.load("cellCount");
    
            await excel.sync();
            if (!(mergedAreas.cellCount == (sceneBlockRows * sceneBlockColumns))){
              //console.log('Not merged')
              myMergeRange.merge(true);
            }
            myMergeRange.values = sceneDataArray;
            myMergeRange = await formatSceneBlock(excel, scriptSheet, myMergeRange, newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            await excel.sync()
          } else if (numActualSceneBlockRows < sceneBlockRows){
            let topRowIndex = theRowIndex - numActualSceneBlockRows;
            //console.log('topRowIndex', topRowIndex);
            for (let i = numActualSceneBlockRows; i < sceneBlockRows; i++){
              //console.log('i', i);
              newRowIndex = await insertRowV2(topRowIndex, false, true);
              //console.log('newRowIndex', newRowIndex);
              let newTypeRange = scriptSheet.getRangeByIndexes(newRowIndex, typeCodeValues.typeCodes.columnIndex, 1, 1);
              newTypeRange.values = myTypes.sceneBlock;
              await excel.sync();
            }
            await sortOutSceneLineNumberRange(newRowIndex + 1, newRowIndex + sceneBlockRows);
            let myMergeRange = scriptSheet.getRangeByIndexes(newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            myMergeRange.merge(true);
            myMergeRange.values = sceneDataArray;
            myMergeRange = await formatSceneBlock(excel, scriptSheet, myMergeRange, newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            await excel.sync();
          } else if (numActualSceneBlockRows > sceneBlockRows){
            newRowIndex = theRowIndex - 1;
            console.log('newRowIndex', newRowIndex);        
            for (let i = sceneBlockRows; i < numActualSceneBlockRows; i++){
              //console.log('i', i , 'newRowIndex', newRowIndex);
              await deleteThisRow(excel, newRowIndex);
              theRowIndex -= 1;
            }
            let topRow = theRowIndex - sceneBlockRows;
            console.log('topRow', topRow, 'theRowIndex', theRowIndex);
            let myMergeRange = scriptSheet.getRangeByIndexes(topRow, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            myMergeRange.load('address');
            myMergeRange.clear("Contents");
            let mergedAreas = myMergeRange.getMergedAreasOrNullObject();
            mergedAreas.load("cellCount");
            await excel.sync();
            if (!(mergedAreas.cellCount == (sceneBlockRows * sceneBlockColumns))){
              //console.log('Not merged')
              myMergeRange.merge(true);
            }
            myMergeRange.values = sceneDataArray;
            myMergeRange = await formatSceneBlock(excel, scriptSheet, myMergeRange, newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            await excel.sync();
          }
        } else if ((nextRowType == myTypes.line) || (nextRowType == myTypes.wallaScripted)){
          sceneDataArray = await getSceneBlockData(theRowIndex, 0);
          
          for (let i = 0; i < sceneBlockRows; i++){
            newRowIndex = await insertRowV2(theRowIndex, false, true);
            //console.log('newRowIndex', newRowIndex);
            let newTypeRange = scriptSheet.getRangeByIndexes(newRowIndex, typeCodeValues.typeCodes.columnIndex, 1, 1);
            newTypeRange.values = myTypes.sceneBlock;
            await excel.sync();
          }
          await sortOutSceneLineNumberRange(newRowIndex + 1, newRowIndex + sceneBlockRows);
          let myMergeRange = scriptSheet.getRangeByIndexes(newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
          myMergeRange.merge(true);
          myMergeRange.values = sceneDataArray;
          myMergeRange = await formatSceneBlock(excel, scriptSheet, myMergeRange, newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
          await excel.sync();
        }
      } else if (sceneListData.type == myTypes.chapter){
        if ((nextRowType == myTypes.line) || (nextRowType == myTypes.scene)){
          let sceneDataArray = await getSceneBlockData(theRowIndex, 0);
          for (let i = 0; i < sceneBlockRows; i++){
            newRowIndex = await insertRowV2(theRowIndex + 1, false, true);
            //console.log('newRowIndex', newRowIndex);
            let newTypeRange = scriptSheet.getRangeByIndexes(newRowIndex, typeCodeValues.typeCodes.columnIndex, 1, 1);
            newTypeRange.values = myTypes.sceneBlock;
            await excel.sync();
          }
          await sortOutSceneLineNumberRange(newRowIndex + 1, newRowIndex + sceneBlockRows);
          let myMergeRange = scriptSheet.getRangeByIndexes(newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
          myMergeRange.merge(true);
          myMergeRange.values = sceneDataArray;
          myMergeRange = await formatSceneBlock(excel, scriptSheet, myMergeRange, newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
          await excel.sync();
        } else if (nextRowType == myTypes.sceneBlock){
          //check there are 4 of them
          let numActualSceneBlockRows = 0;
          for (i = nextIndex; i < nextIndex + 30; i++){
            //console.log(i, typeCodeValues.typeCodes.values[i]);
            if (typeCodeValues.typeCodes.values[i] == myTypes.sceneBlock){
              numActualSceneBlockRows += 1;
            } else {
              break;
            }
          }
          sceneDataArray = await getSceneBlockData(theRowIndex, numActualSceneBlockRows);
          //console.log('numActualSceneBlockRows', numActualSceneBlockRows)
          if (numActualSceneBlockRows == sceneBlockRows){
            newRowIndex = theRowIndex + 1;
            let myMergeRange = scriptSheet.getRangeByIndexes(newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            myMergeRange.load('address');
            myMergeRange.clear("Contents");
            let mergedAreas = myMergeRange.getMergedAreasOrNullObject();
            mergedAreas.load("cellCount");
            await excel.sync();
            if (!(mergedAreas.cellCount == (sceneBlockRows * sceneBlockColumns))){
              //console.log('Not merged')
              myMergeRange.merge(true);
            }
            myMergeRange.values = sceneDataArray;
            myMergeRange = await formatSceneBlock(excel, scriptSheet, myMergeRange, newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            await excel.sync()
          } else if (numActualSceneBlockRows < sceneBlockRows){
            for (let i = numActualSceneBlockRows; i < sceneBlockRows; i++){
              console.log('i', i);
              newRowIndex = await insertRowV2(theRowIndex + 1, false, true);
              //console.log('newRowIndex', newRowIndex);
              let newTypeRange = scriptSheet.getRangeByIndexes(newRowIndex, typeCodeValues.typeCodes.columnIndex, 1, 1);
              newTypeRange.values = myTypes.sceneBlock;
              await excel.sync();
            }
            await sortOutSceneLineNumberRange(newRowIndex + 1, newRowIndex + sceneBlockRows);
            let myMergeRange = scriptSheet.getRangeByIndexes(newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            myMergeRange.merge(true);
            myMergeRange.values = sceneDataArray;
            myMergeRange = await formatSceneBlock(excel, scriptSheet, myMergeRange, newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            await excel.sync();
          } else if (numActualSceneBlockRows > sceneBlockRows){
            newRowIndex = theRowIndex + 1;
            for (let i = sceneBlockRows; i < numActualSceneBlockRows; i++){
              //console.log('i', i , 'newRowIndex', newRowIndex);
              await deleteThisRow(excel, newRowIndex);
            }
            let myMergeRange = scriptSheet.getRangeByIndexes(newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            myMergeRange.load('address');
            myMergeRange.clear("Contents");
            let mergedAreas = myMergeRange.getMergedAreasOrNullObject();
            mergedAreas.load("cellCount");
            await excel.sync();
            if (!(mergedAreas.cellCount == (sceneBlockRows * sceneBlockColumns))){
              //console.log('Not merged')
              myMergeRange.merge(true);
            }
            myMergeRange.values = sceneDataArray;
            myMergeRange = await formatSceneBlock(excel, scriptSheet, myMergeRange, newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns);
            await excel.sync()
          }
        }    
      }
      await fillChapterAndScene();
    });
  } else {
    alert("Please enter a number")
  }
}

async function addSceneBlock(){
  let myWait = tag('scene-add-wait');
  myWait.style.display = 'block'
  const addChapterValue = tag("chapter-scene-select").value;
  console.log('Chapter/Scene', addChapterValue);
  let chapterSceneID = parseInt(addChapterValue);
  await doTheActualSceneBlock(chapterSceneID);
  await fillChapterAndScene();
  myWait.style.display = 'none';
}

async function sortOutSceneLineNumberRange(startRow, endRow){
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let sceneLineNumberColumn = findColumnLetter('Scene Line Number Range');
    let sourceAddress = sceneLineNumberColumn + (startRow - 1);
    let sourceRange = scriptSheet.getRange(sourceAddress);
    sourceRange.load('values');
    await excel.sync();
    let myValue = sourceRange.values[0][0];
    console.log(myValue);
    let destRangeAddress = sceneLineNumberColumn + startRow + ':' + sceneLineNumberColumn + endRow;
    let destRange = scriptSheet.getRange(destRangeAddress);
    destRange.values = myValue;
  })
}

async function deleteThisRow(excel, rowIndex){
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let myRow = scriptSheet.getRangeByIndexes(rowIndex, 1, 1, 1).getEntireRow();
    let isProtected = await unlockIfLocked();
    myRow.delete("Up");
    await excel.sync();
    await correctFormulas(rowIndex);
    if (isProtected){
      await lockColumns();
    }
}

function doBorder(borders, edge){
  let myEdge = borders.getItem(edge);
  myEdge.color = '#000000';
  myEdge.style = 'Continuous';
  myEdge.weight = 'Thin'
}

async function formatSceneBlock(excel, sheet, theRange, newRowIndex, cueColumnIndex, sceneBlockRows, sceneBlockColumns){
  theRange.format.font.name = 'Courier New';
  theRange.format.font.size = 12;
  theRange.format.font.bold = true;
  theRange.format.fill.color = myFormats.purple;
  theRange.format.horizontalAlignment = 'Center';
  theRange.format.verticalAlignment = 'Top';
  
  let myBorders = theRange.format.borders;
  doBorder(myBorders, 'EdgeTop');
  doBorder(myBorders, 'EdgeBottom');
  doBorder(myBorders, 'EdgeLeft');
  doBorder(myBorders, 'EdgeRight');
  doBorder(myBorders, 'InsideHorizontal');
  await excel.sync();

  
  myBorders.load('items');
  await excel.sync()
  //console.log('Border count', myBorders.count);
  for (let i = 0; i < myBorders.items.length; i++){
    //console.log('Borders', i, myBorders.items[i].color, myBorders.items[i].id, myBorders.items[i].sideIndex, myBorders.items[i].style, myBorders.items[i].weight)
  }
  
  for (let i = 0; i < sceneBlockRows; i++){
    let tempRange = sheet.getRangeByIndexes(newRowIndex + i, cueColumnIndex, 1, sceneBlockColumns);
    await mergedRowAutoHeight(excel, sheet, tempRange);
  }
}
async function formatWallaBlock(excel, sheet, theRange, newRowIndex, leftMostColumn, blockRows, numColumns){
  theRange.format.font.name = 'Courier New';
  theRange.format.font.size = 12;
  theRange.format.font.bold = true;
  theRange.format.fill.color = myFormats.green;
  theRange.format.horizontalAlignment = 'Left';
  theRange.format.verticalAlignment = 'Top';
  await excel.sync()
  for (let i = 0; i < blockRows; i++){
    let tempRange = sheet.getRangeByIndexes(newRowIndex + i, leftMostColumn, 1, numColumns);
    await mergedRowAutoHeight(excel, sheet, tempRange);
  }
}
async function formatWallaBlockCue(excel, theRange){
  theRange.format.font.name = 'Courier New';
  theRange.format.font.size = 12;
  theRange.format.font.bold = true;
  theRange.format.fill.color = myFormats.green;
  theRange.format.horizontalAlignment = 'Center';
  theRange.format.verticalAlignment = 'Top';
  await excel.sync();
  
}


function removeDoubleLf(myText){
  let mySplit = myText.split('\n');
  let result = []
  let resultIndex = -1;
  for (let i = 0; i < mySplit.length; i++){
    if (mySplit[i] != ''){
      resultIndex += 1
      result[resultIndex] = mySplit[i];
    }
  }
  return result.join('\n');
}

async function getSceneBlockData(myRowIndex, numSceneBlockLines){
  // returns a formatted array suitable for the merged cells
  let sceneDataArray;
  await Excel.run(async (excel) => {
    let sceneNumberIndex = findColumnIndex('Scene Number');
    let otherNotesIndex = findColumnIndex('Other notes');
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    //console.log ('Indexes', myRowIndex, sceneNumberIndex, 2 + numSceneBlockLines, otherNotesIndex - sceneNumberIndex + 1);
    let myDataRange = scriptSheet.getRangeByIndexes(myRowIndex, sceneNumberIndex, 2 + numSceneBlockLines, otherNotesIndex - sceneNumberIndex + 1);
    myDataRange.load('values');
    await excel.sync();
    
    let sceneData = {}
    sceneData.scene = myDataRange.values[0][0]
    sceneData.location = '';
    sceneData.beasts ='';
    sceneData.otherNotes = '';
    for (let row = 0; row < myDataRange.values.length; row++){
      //console.log('Row', row);
      if (sceneData.scene == ''){
        sceneData.scene = myDataRange.values[row][0]
      }
      if (sceneData.location == ''){
        sceneData.location = removeDoubleLf(myDataRange.values[row][11]);
      }
      if (sceneData.beasts == ''){
        sceneData.beasts = removeDoubleLf(myDataRange.values[row][15]);
      }
      if (sceneData.otherNotes == ''){
        sceneData.otherNotes = removeDoubleLf(myDataRange.values[row][16]);
      }
    }

    sceneDataArray = Array(sceneBlockRows).fill().map(() => Array(sceneBlockColumns).fill(''));
    sceneDataArray[0][0] = "Scene " + sceneData.scene;
    sceneDataArray[1][0] = 'Scene Location: ' + sceneData.location;
    sceneDataArray[2][0] = 'Beasts/Animals: ' + sceneData.beasts;
    sceneDataArray[3][0] = 'Other notes: ' + sceneData.otherNotes;
});
return sceneDataArray;
}
function createChapterIndecies(theTypeCodesValues){
  let chapterIndecies = []
  let chapterIndex = -1
  for (let i = 0; i < theTypeCodesValues.length; i++){
    if (theTypeCodesValues[i] == myTypes.chapter){
      chapterIndex += 1;
      chapterIndecies[chapterIndex] = i;
    }
  }
  return chapterIndecies;
}

function createChapterAndSceneList(theTypeCodeValues){
  let theList = [];
  let listIndex = -1;
  let tcv = theTypeCodeValues.typeCodes.values;
  for (let i = 0; i < tcv.length; i++){
    if (tcv[i] == myTypes.chapter){
      listIndex += 1;
      let item = {
        arrayIndex: i,
        rowIndex: theTypeCodeValues.typeCodes.rowIndex + i,
        type: myTypes.chapter,
        number: theTypeCodeValues.chapters.values[i],
        display: 'Chapter ' + theTypeCodeValues.chapters.values[i]
      }
      theList[listIndex] = item;
    } else if (tcv[i] == myTypes.scene){
      listIndex += 1;
      let item = {
        arrayIndex: i,
        rowIndex: theTypeCodeValues.typeCodes.rowIndex + i,
        type: myTypes.scene,
        number: theTypeCodeValues.scenes.values[i],
        display: 'Scene ' + theTypeCodeValues.scenes.values[i]
      }
      theList[listIndex] = item;
    }
  }
  //console.log('The list', theList);
  return theList;
}

async function getTypeCodes(){
  let theValues = {};
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const typeCodeColumn = findColumnLetter("Type Code"); 
    const chapterNoColumn = findColumnLetter('Chapter Calculation');
    const sceneColumn = findColumnLetter('Scene')
    let typeCodeRange = scriptSheet.getRange(typeCodeColumn + firstDataRow + ":" + typeCodeColumn +lastDataRow);
    let chapterNoRange = scriptSheet.getRange(chapterNoColumn + firstDataRow + ":" + chapterNoColumn +lastDataRow);
    let sceneRange = scriptSheet.getRange(sceneColumn + firstDataRow + ":" + sceneColumn +lastDataRow);
    typeCodeRange.load('values');
    typeCodeRange.load('rowIndex');
    typeCodeRange.load('rowCount');
    typeCodeRange.load('address');
    typeCodeRange.load('columnIndex');
    chapterNoRange.load('values');
    chapterNoRange.load('rowIndex');
    chapterNoRange.load('rowCount');
    chapterNoRange.load('address');
    chapterNoRange.load('columnIndex');
    sceneRange.load('values');
    sceneRange.load('rowIndex');
    sceneRange.load('rowCount');
    sceneRange.load('address');
    sceneRange.load('columnIndex');
    await excel.sync();
    theValues.typeCodes = {
      values: typeCodeRange.values.map(x => x[0]),
      rowIndex: typeCodeRange.rowIndex,
      rowCount: typeCodeRange.rowCount,
      columnIndex: typeCodeRange.columnIndex,
      address: typeCodeRange.address
    }
    theValues.chapters = {
      values: chapterNoRange.values.map(x => x[0]),
      rowIndex: chapterNoRange.rowIndex,
      rowCount: chapterNoRange.rowCount,
      columnIndex: chapterNoRange.columnIndex,
      address: chapterNoRange.address
    }
    theValues.scenes = {
      values: sceneRange.values.map(x => x[0]),
      rowIndex: sceneRange.rowIndex,
      rowCount: sceneRange.rowCount,
      columnIndex: sceneRange.columnIndex,
      address: sceneRange.address
    }
  });
  return theValues;
}

/* Do a Chapter
  Find the next chapter
  Is the next row a scene
    Insert 3 lines
    Mark them, as Scene Block
  Is the next row a line
    Insert 3 lines
  Is the next line a Scene block
    Check there are three of them    
  

  Now do the filling in
      
*/

async function mergedRowAutoHeight(excel, theSheet, theRange){
  let app = excel.workbook.application;
  app.suspendScreenUpdatingUntilNextSync();
  theRange.load('columnCount');
  theRange.load('columnIndex');
  theRange.load('rowIndex');
  theRange.load('rowCount');
  theRange.format.load('rowHeight')
  await excel.sync()
  if (theRange.rowCount == 1){
    let totalcolumnWidth = 0;
    let thisCol = []
    for (let i = 0; i < theRange.columnCount; i++){
      thisCol[i] = theRange.getCell(0,i);
      thisCol[i].format.load('columnWidth');
    }
    await excel.sync();
    let columnOneWidth = thisCol[0].format.columnWidth;
    for (let i = 0; i < theRange.columnCount; i++){
      totalcolumnWidth = totalcolumnWidth + thisCol[i].format.columnWidth
    }
    app.suspendScreenUpdatingUntilNextSync();
    theRange.unmerge();
    let tempRange = theSheet.getRangeByIndexes(theRange.rowIndex, theRange.columnIndex, 1, 1);
    tempRange.format.wrapText = false;
    await excel.sync()
    if (totalcolumnWidth > 1300){totalcolumnWidth = 1300}
    //console.log('totalcolumnWidth', totalcolumnWidth);
    tempRange.format.columnWidth = totalcolumnWidth
    await excel.sync()
    tempRange.format.wrapText = true;
    await excel.sync()
    tempRange.format.autofitRows();
    await excel.sync()
    tempRange.format.load('rowHeight')
    await excel.sync()
    app.suspendScreenUpdatingUntilNextSync();
    let finalRowHeight = tempRange.rowHeight;
    tempRange.format.columnWidth = columnOneWidth;
    tempRange.format.rowHeight = finalRowHeight;
    theRange.merge(true);
    await excel.sync();
  }
}
async function fillChapterAndScene(){
  addSelectList = await createSceneList();
  let chapterAddSelect = tag('chapter-scene-select');
  let selected = chapterAddSelect.selectedIndex;
  //console.log('Selected index:', chapterAddSelect.selectedIndex);
  chapterAddSelect.innerHTML = '';
  chapterAddSelect.add(new Option('Please select', ''));
  for (let i = 0; i < addSelectList.length; i++){
    chapterAddSelect.add(new Option(addSelectList[i].display, i));
  }
  chapterAddSelect.selectedIndex = selected;
}

async function createSceneList(){
  typeCodeValues = await getTypeCodes();
  return createChapterAndSceneList(typeCodeValues);
}

async function createWalla(wallaData, rowIndex, doReplace, doNext){
  await Excel.run(async (excel) => {
    let loadMessage = tag('load-message');
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let numberColumns = numberOfPeoplePresentIndex - wallaLineRangeIndex + 1
    let firstWallaRange = scriptSheet.getRangeByIndexes(rowIndex, wallaLineRangeIndex, 1, numberColumns);
    let wallaOriginalRange = scriptSheet.getRangeByIndexes(rowIndex, wallaOriginalIndex, 1 , 1)
    firstWallaRange.load('address');
    firstWallaRange.load('values');
    wallaOriginalRange.load('address')
    await excel.sync();
    console.log(firstWallaRange.address, wallaOriginalRange.address);

    let dataArray = [
      wallaData.wallaLineRange,
      wallaData.typeOfWalla,
      wallaData.characters,
      wallaData.description,
      wallaData.numCharacters
    ]

    if (firstWallaRange.values[0][1] != ''){
      if (doReplace){
        firstWallaRange.clear("Contents");
        wallaOriginalRange.clear("Contents");
      }
      if (doNext){
        if (!isDataTheSame(dataArray, firstWallaRange.values[0])){
          for (let i = rowIndex + 1; i < rowIndex + 100; i++){
            console.log(i)
            firstWallaRange = scriptSheet.getRangeByIndexes(i, wallaLineRangeIndex, 1, numberColumns);
            wallaOriginalRange = scriptSheet.getRangeByIndexes(i, wallaOriginalIndex, 1 , 1)
            firstWallaRange.load('values');
            await excel.sync();
            console.log('Testing row: ', i, 'Row data: ', firstWallaRange.values[0]);
            if (!isDataTheSame(dataArray, firstWallaRange.values[0])){
              if (firstWallaRange.values[0][1] == ''){
                rowIndex = i;
                break;
              }
            } else {
              console.log('Already there');
              loadMessage.style.display = 'block'
              return null;
            }
          }
          console.log('New row index', rowIndex)
        } else {
          console.log('Already there')
          loadMessage.style.display = 'block'
          return null;
        }
      }
    }

    firstWallaRange.values = [dataArray];
    wallaOriginalRange.values = [[wallaData.all]]
    firstWallaRange.select();   
    await excel.sync();
    await showMainPage();

  })

}
function isDataTheSame(newData, currentData){
  if (newData.length == currentData.length){
    for (let i = 0; i < newData.length; i++){
      if (newData[i] != currentData[i]){
        console.log('Not the same');
        return false
      }
    }
    console.log('The same')
    return true
  } else {
    console.log('Different dimensions')
    return null;
  }
}

async function calculateWallaCues(){
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let numberColumns = numberOfPeoplePresentIndex - wallaLineRangeIndex + 1
    let wallaRange = scriptSheet.getRangeByIndexes(firstDataRow - 1, wallaLineRangeIndex, (lastDataRow - firstDataRow), numberColumns);
    wallaRange.load('rowIndex');
    wallaRange.load('values');
    await excel.sync();
    console.log ('rowIndex', wallaRange.rowIndex, 'values', wallaRange.values);
    let rowsToDo = []
    let rowIndex = -1;
    for (let i = 0; i < wallaRange.values.length; i++){
      if ((wallaRange.values[i][1].toLowerCase() == namedCharacters.toLowerCase()) || (wallaRange.values[i][1].toLowerCase() == namedCharactersColon.toLowerCase())){
        rowIndex += 1;
        rowsToDo[rowIndex] = i
      }
    }
    console.log('Rows to do: ', rowsToDo);
    let wallaCueColumn = scriptSheet.getRangeByIndexes(firstDataRow - 1, wallaCueIndex, (lastDataRow - firstDataRow), 1);
    wallaCueColumn.clear("Contents")
    await excel.sync();

    let wallaNumber = 0
    let theCells = []
    for (let i = 0; i < rowsToDo.length; i++){
      wallaNumber += 1
      wallaCue = "W" + String(wallaNumber).padStart(5, 0);
      console.log(wallaCue)
      theCells[i] = scriptSheet.getRangeByIndexes(rowsToDo[i] + wallaRange.rowIndex, wallaCueIndex, 1, 1);
      theCells[i].values = [[wallaCue]]
    }
    await excel.sync();
  })

}

function allEmpty(theArray){
  for (let i = 0; i < theArray.length; i++){
    if (theArray[i] != ''){
      return false;
    }
  }
  return true;
}

async function getSceneWallaInformation(typeNo){
  let wallaScene = tag('walla-scene').value;
  sceneNo = parseInt(wallaScene);
  let doNamed, doUnnamed, doGeneral;
  if (typeNo == 1){
    doNamed = true;
    doUnnamed = false;
    doGeneral = false;
  } else if (typeNo == 2){
    doNamed = false;
    doUnnamed = true;
    doGeneral = false;
  } else if (typeNo == 3){
    doNamed = false;
    doUnnamed = false;
    doGeneral = true;
  } else {
    alert('Invalid choice');
    return null;
  }

  if (!isNaN(sceneNo)){
    await Excel.run(async (excel) => {
      const firstRowIndex = firstDataRow - 1;
      const lastRowIndex = lastDataRow - firstDataRow;
      let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      let typeOfWallaRange = scriptSheet.getRangeByIndexes(firstRowIndex, typeOfWallaIndex, lastRowIndex, 1);
      let sceneRange = scriptSheet.getRangeByIndexes(firstRowIndex, sceneIndex, lastRowIndex, 1);
      let wallaCueRange = scriptSheet.getRangeByIndexes(firstRowIndex, wallaCueIndex, lastRowIndex, 1);
      let wallaOriginalRange = scriptSheet.getRangeByIndexes(firstRowIndex, wallaOriginalIndex, lastRowIndex, 1);
      let typeCodeRange = scriptSheet.getRangeByIndexes(firstRowIndex, typeCodeIndex, lastRowIndex, 1);
  
      typeOfWallaRange.load('rowIndex');
      typeOfWallaRange.load('values');
      sceneRange.load('values');
      wallaCueRange.load('values');
      wallaOriginalRange.load('values');
      typeCodeRange.load('values')
      
      await excel.sync();
      let myIndecies = [];
      let theIndex = - 1;
      
      if (doNamed){
        for (let i = 0; i < typeOfWallaRange.values.length; i++){
          console.log('Scene: ', sceneRange.values[i][0]);
          if (isNamedWalla(typeOfWallaRange.values[i][0])){
            if (sceneRange.values[i][0] == sceneNo){
              theIndex += 1;
              myIndecies[theIndex] = i;
            } else if (sceneRange.values[i][0] > sceneNo){
              break;
            }
          } 
        }
      }
      
      if (doUnnamed){
        for (let i = 0; i < typeOfWallaRange.values.length; i++){
          console.log('Unnamed Scene: ', sceneRange.values[i][0]);
          if (isUnamedWalla(typeOfWallaRange.values[i][0])){
            if (sceneRange.values[i][0] == sceneNo){
              theIndex += 1;
              myIndecies[theIndex] = i;
            }
          } else if (sceneRange.values[i][0] > sceneNo){
              break; 
          }
        }
      }

      if (doGeneral){
        for (let i = 0; i < typeOfWallaRange.values.length; i++){
          console.log('General Scene: ', sceneRange.values[i][0]);
          if (isGeneralWalla(typeOfWallaRange.values[i][0])){
            if (sceneRange.values[i][0] == sceneNo){
              theIndex += 1;
              myIndecies[theIndex] = i;
            }
          } else if (sceneRange.values[i][0] > sceneNo){
              break; 
          }
        }
      }
 
      console.log(myIndecies, theIndex);

      let cues = [''];
      let details = []
      if (doNamed){
        if (myIndecies.length == 0){
          details = [namedCharactersColon + ' None'];
        } else {
          details = [namedCharactersColon];
        }
      }
      if (doUnnamed){
        if (myIndecies.length == 0){
          details = [unnamedCharactersColon + ' None'];  
        } else {
          details = [unnamedCharactersColon];  
        }
      }
      
      if (doGeneral){
        if (myIndecies.length == 0){
          details = [generalWallaColon + ' None'];  
        } else {
          details = [generalWallaColon];  
        }
      }
      let item = 0;

      for (let i = 0; i < myIndecies.length; i++){
        item += 1;
        console.log(item)
        cues[item] = wallaCueRange.values[myIndecies[i]][0];
        details[item] = wallaOriginalRange.values[myIndecies[i]][0];
      }

      console.log(details);
      console.log(cues.join('\n'));
      console.log(details.join('\n'));
      
      let sceneRowIndex = -1; 
      let doIt = false;
      for (let i = 0; i < typeCodeRange.values.length; i++){
        if (typeCodeRange.values[i][0] == myTypes.scene){
          if (sceneRange.values[i][0] == sceneNo){
            sceneRowIndex = i + typeOfWallaRange.rowIndex;
            console.log('Scene route sceneRowIndex', sceneRowIndex);
            doIt = true;
            break;
          }
        } else if (typeCodeRange.values[i][0] == myTypes.chapter){
          if (sceneRange.values[i][0] == sceneNo){
            if (doNamed){
              sceneRowIndex = i + typeOfWallaRange.rowIndex + sceneBlockRows + 1;  
            } else if (doUnnamed){
              sceneRowIndex = i + typeOfWallaRange.rowIndex + sceneBlockRows + 2;  
            } else if (doGeneral){
              sceneRowIndex = i + typeOfWallaRange.rowIndex + sceneBlockRows + 3;  
            }
            console.log('Chapter Route sceneRowIndex', sceneRowIndex);
            doIt = true;
            break;
          }
        }
      }
      if (doIt){
        let selectCell = scriptSheet.getRangeByIndexes(sceneRowIndex, cueIndex, 1, 1);
        selectCell.select();
        await insertRowV2(sceneRowIndex, false, false);
        let typeCodeCell = scriptSheet.getRangeByIndexes(sceneRowIndex, typeCodeIndex, 1, 1);
        let wallaCueCell = scriptSheet.getRangeByIndexes(sceneRowIndex, cueIndex, 1, 1);
        let wallaDetailsCell = scriptSheet.getRangeByIndexes(sceneRowIndex, numberIndex, 1, 1);
        typeCodeCell.values =[[myTypes.wallaBlock]];
        wallaCueCell.values = [[cues.join('\n')]];
        wallaDetailsCell.values = [[details.join('\n')]]
        let wallaDetailsMergeRange = scriptSheet.getRangeByIndexes(sceneRowIndex, numberIndex, 1, wallaBlockColumns);
        wallaDetailsMergeRange.merge(true);
        await formatWallaBlockCue(excel, wallaCueCell);
        await formatWallaBlock(excel, scriptSheet, wallaDetailsMergeRange, sceneRowIndex, numberIndex, 1, wallaBlockColumns);
      }
    })
  } else {
    alert('Enter a valid scene number')
  }
}
function isNamedWalla(theType){
  return ((theType.trim().toLowerCase() == namedCharacters.trim().toLowerCase()) || (theType.trim().toLowerCase() == namedCharactersColon.trim().toLowerCase()));
}

function isUnamedWalla(theType){
  return ((theType.trim().toLowerCase() == unnamedCharacters.trim().toLowerCase()) || (theType.trim().toLowerCase() == unnamedCharactersColon.trim().toLowerCase()));
}

function isGeneralWalla(theType){
  return ((theType.trim().toLowerCase() == generalWalla.trim().toLowerCase()) || (theType.trim().toLowerCase() == generalWallaColon.trim().toLowerCase()));
}

async function deleteAllSceneAndWallaBlocks(){
  let isProtected = await unlockIfLocked();
  await Excel.run(async (excel) => {
    for (let myDelete = 0; myDelete < 1000; myDelete++){

      let myTypeCodes = await getTypeCodes();
      console.log(myTypeCodes);
      let theIndexes = [];
      let theIndex = -1
      for (let i = 0; i < myTypeCodes.typeCodes.values.length;i++){
        //console.log(i, myTypeCodes.typeCodes.values[i]);
        if ((myTypeCodes.typeCodes.values[i] == myTypes.sceneBlock)||(myTypeCodes.typeCodes.values[i] == myTypes.wallaBlock)){
            theIndex += 1
            theIndexes[theIndex] = i + myTypeCodes.typeCodes.rowIndex;
            break;
        }
      }
      console.log('The Indexes',theIndexes);
      
      if (theIndexes.length == 0){
        break;
      }
      
      let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      let thisRow = [];
      for (let i = 0 ; i < theIndexes.length; i++){
        thisRow[i] = scriptSheet.getRangeByIndexes(theIndexes[i],1,1,1).getEntireRow();
        thisRow[i].select();
        await excel.sync();
        thisRow[i].delete("Up");
        console.log('Num: ', myDelete)
        console.log('Before sync');
        await excel.sync();
        console.log('After sync');  
      }
    }
    const firstRowIndex = firstDataRow - 1;
    const lastRowIndex = lastDataRow - firstDataRow;
    
  })
  if (isProtected){
    await lockColumns();
  }
}

async function clearWalla(){
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let wallaOrigninalRange = scriptSheet.getRangeByIndexes(firstDataRow - 1, wallaOriginalIndex, lastDataRow - firstDataRow, 1);
    wallaOrigninalRange.clear("Contents");
    let wallaDetails = scriptSheet.getRangeByIndexes(firstDataRow - 1, wallaCueIndex, lastDataRow - firstDataRow, numberOfPeoplePresentIndex - wallaCueIndex + 1);
    wallaDetails.clear('Contents');
  })
}


async function getRowIndeciesForScene(sceneNumber){
  let myIndecies, newIndexes;
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let sceneRange = scriptSheet.getRangeByIndexes(firstDataRow, sceneIndex, lastDataRow - firstDataRow, 1);
    sceneRange.load('values, rowIndex');
    await excel.sync();
    myIndecies = sceneRange.values.map((x, i) => [x, i]).filter(([x, i]) => x == sceneNumber).map(([x, i]) => i + sceneRange.rowIndex);
    let typeCodeRange = scriptSheet.getRangeByIndexes(myIndecies[0], typeCodeIndex, myIndecies[myIndecies.length-1] - myIndecies[0] + 1, 1);
    typeCodeRange.load('values, rowIndex');
    await excel.sync();
    let dodgyIndexes = typeCodeRange.values.map((x, i) => [x, i]).filter(([x, i]) => x == myTypes.sceneBlock).map(([x, i]) => i + typeCodeRange.rowIndex);
    console.log('dodgy', dodgyIndexes);
    newIndexes = [];
    let newIndex = - 1;
    for (let i = 0; i < myIndecies.length; i++){
      if (!(dodgyIndexes.includes(myIndecies[i]))){
        newIndex += 1;
        newIndexes[newIndex] = myIndecies[i];
      }
    }
    console.log('newIndexes', newIndexes)

  })
  return newIndexes;
}

async function getSceneBlockNear(index){
  let startOffset = -12;
  let endOffset = + 6;
  let sceneBlockText = [];
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    console.log(index + startOffset, typeCodeIndex, endOffset - startOffset, 1)
    let startRowIndex = index + startOffset
    if (startRowIndex < 1){
      startRowIndex = 1;
    }
    let typeCodeRange = scriptSheet.getRangeByIndexes(startRowIndex, typeCodeIndex, endOffset - startOffset, 1);
    typeCodeRange.load('values, rowIndex');
    await excel.sync();
    let indexes = []
    let theIndex = -1;
    for (let i = 0; i < typeCodeRange.values.length; i++){
      if (typeCodeRange.values[i][0] == myTypes.sceneBlock){
        theIndex += 1;
        indexes[theIndex] = i + typeCodeRange.rowIndex; 
      }
    }
    console.log('indexes', indexes);

    if (indexes.length > 0){
      console.log(indexes[0], cueIndex, indexes.length)
      let sceneBlockRange= scriptSheet.getRangeByIndexes(indexes[0], cueIndex, indexes.length, 1);
      sceneBlockRange.load('values');
      await excel.sync();
      sceneBlockText = sceneBlockRange.values.map(x => x[0])
    }
  })
  return sceneBlockText;
}

async function getActorScriptDetails(indexes){
  let details = {};
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let columnCount = ukScriptIndex - cueIndex + 1;
    console.log(indexes[0],cueIndex, indexes.length,columnCount)
    let dataRange = scriptSheet.getRangeByIndexes(indexes[0], cueIndex, indexes.length, columnCount);
    dataRange.load('values');
    let rangeProperties = dataRange.getCellProperties({
      format: {
        font: {
          color: true,
          name: true,
          size: true,
          bold: true,
          italic: true
        }
      }
    })
    await excel.sync();
    console.log(dataRange.values);
    console.log(rangeProperties.value);
  })
}

async function getBook(){
  let book;
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let bookRange = scriptSheet.getRangeByIndexes(firstDataRow - 1, bookIndex, 50, 1);
    bookRange.load('values');
    await excel.sync();
    for (let i = 0; i < bookRange.values.length; i++){
      console.log(i);
      if (bookRange.values[i][0].trim().toLowerCase().startsWith('book')){
        book = bookRange.values[i][0].trim();
        break;
      }
    }
  })
  console.log('Book: ', book);
  return book;
}

async function getActorScriptRanges(indexes, startRowIndex){
  let rangeBounds = []
  let rangeIndex = 0;
  let actorCueColumnIndex = 0;
  let actorCharacterColumnIndex = 1;  
  let actorDirectionColumnIndex = 2;
  let actorUkScriptColumnIndex = 3;
  for (let i = 0; i < indexes.length; i++){
    if (i == 0){
      rangeBounds[rangeIndex] = {};
      rangeBounds[rangeIndex].start = indexes[i];
    }else if (i == indexes.length -1){
      rangeBounds[rangeIndex].end = indexes[i]
    } else {
      if (indexes[i] == (indexes[i-1] + 1)){
        //Do Nothing
      } else {
        rangeBounds[rangeIndex].end = indexes[i - 1]
        rangeIndex += 1
        rangeBounds[rangeIndex] = {};
        rangeBounds[rangeIndex].start = indexes[i];
      }
    }
  }
  let cueRange, characterRange, directionRange, ukScriptRange;
  console.log('Rangebound length', rangeBounds.length);
  let rowIndexes = []
  let item = - 1;
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let actorScriptSheet = excel.workbook.worksheets.getItem(actorScriptName);
    for (let i = 0; i< rangeBounds.length; i++){
      let rowCount = rangeBounds[i].end - rangeBounds[i].start + 1
      item += 1;
      rowIndexes[item] = {
        startRow: startRowIndex,
        rowCount: rowCount
      }
      cueRange = scriptSheet.getRangeByIndexes(rangeBounds[i].start, cueIndex, rowCount, 1);
      characterRange = scriptSheet.getRangeByIndexes(rangeBounds[i].start, characterIndex, rowCount, 1);
      directionRange = scriptSheet.getRangeByIndexes(rangeBounds[i].start, stageDirectionWallaDescriptionIndex, rowCount, 1);
      ukScriptRange = scriptSheet.getRangeByIndexes(rangeBounds[i].start, ukScriptIndex, rowCount, 1);
      
      
      console.log('start row', startRowIndex)
      let actorCueRange = actorScriptSheet.getRangeByIndexes(startRowIndex, actorCueColumnIndex, 1, 1);
      let actorCharacterRange = actorScriptSheet.getRangeByIndexes(startRowIndex, actorCharacterColumnIndex, 1, 1);
      let actorDirectionRange  = actorScriptSheet.getRangeByIndexes(startRowIndex, actorDirectionColumnIndex, 1, 1);
      let actorUkScriptRange  = actorScriptSheet.getRangeByIndexes(startRowIndex, actorUkScriptColumnIndex, 1, 1);
      actorCueRange.copyFrom(cueRange, 'Values', false, false);
      actorCueRange.copyFrom(cueRange, 'Formats', false, false);
      actorCharacterRange.copyFrom(characterRange, 'Values', false, false);
      actorCharacterRange.copyFrom(characterRange, 'Formats', false, false);
      actorDirectionRange.copyFrom(directionRange, 'Values', false, false);
      actorDirectionRange.copyFrom(directionRange, 'Formats', false, false);
      actorUkScriptRange.copyFrom(ukScriptRange, 'Values', false, false);
      actorUkScriptRange.copyFrom(ukScriptRange, 'Formats', false, false);
      await excel.sync();
      startRowIndex = startRowIndex + rowCount
    }
  })
  console.log('Row Indexes', rowIndexes);
  return rowIndexes;
}

async function fillColorLinesAndScriptedWalla(){
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    let typeCodeRange = scriptSheet.getRangeByIndexes(firstDataRow, typeCodeIndex, lastDataRow - firstDataRow, 1);
    typeCodeRange.load('values, rowIndex');
    await excel.sync();
    let typeCodes = typeCodeRange.values.map(x => x[0]);
    const theRowIndex = typeCodeRange.rowIndex;
    const wallaScriptedIndexes = typeCodes.map((x, i) => [x, i]).filter(([x, i]) => x == myTypes.wallaScripted).map(([x, i]) => i + theRowIndex);
    console.log('typeCodes: ', typeCodes);
    console.log('theRowIndex:', theRowIndex);
    console.log('walla scripted indexes:', wallaScriptedIndexes);
    const lineIndexes = typeCodes.map((x, i) => [x, i]).filter(([x, i]) => x == myTypes.line).map(([x, i]) => i + theRowIndex);
    const sceneAndChapterIndexes = typeCodes.map((x, i) => [x, i]).filter(([x, i]) => (x == myTypes.chapter)||(x == myTypes.scene)).map(([x, i]) => i + theRowIndex);
    console.log('line indexes: ', lineIndexes);
    console.log('sceneAndChapter', sceneAndChapterIndexes);
    const columnCount = otherNotesIndex - cueIndex + 1;
   
    
    
    let sceneIndexes = []
    let chapterIndexes = []
    let mySceneIndex = -1;
    let myChapterIndex = -1

    for (let i = 0; i < sceneAndChapterIndexes.length; i++){
      let thisIndex  = sceneAndChapterIndexes[i] - theRowIndex;
      let tc = typeCodes[thisIndex];
      console.log("I", i, 'Type Code', tc, ' index ', thisIndex);
      if (tc == myTypes.chapter){
        myChapterIndex += 1;
        chapterIndexes[myChapterIndex] = sceneAndChapterIndexes[i];
      } else if (tc == myTypes.scene){
        let found = false;
        for (let j = thisIndex - 1; j >= thisIndex - 5; j--){
          let testIndex = sceneAndChapterIndexes[j] - theRowIndex; 
          console.log("J", j, ' typeCode ', typeCodes[testIndex], ' index ', testIndex);
          if (typeCodes[testIndex] == myTypes.chapter){
            myChapterIndex += 1;
            chapterIndexes[myChapterIndex] = sceneAndChapterIndexes[i];
            found = true;
            //console.log('Found')
            break;
          }
        }
        if (!found){
          //console.log('Not found')
          mySceneIndex += 1;
          sceneIndexes[mySceneIndex] = sceneAndChapterIndexes[i];
        }
      }
    }
    console.log ('Scene Indexes: ', sceneIndexes, 'Chapter Indexes: ', chapterIndexes);

    console.log('Doing Walla Formatting')
    await doSomeFormatting(excel, wallaScriptedIndexes, scriptSheet, columnCount, false, myFormats.wallaGreen, myFormats.black)
    console.log('Doing Line Formatting')
    await doSomeFormatting(excel, lineIndexes, scriptSheet, columnCount, true, myFormats.white, myFormats.black)
    console.log('Doing Scene Formatting')
    await doSomeFormatting(excel, sceneIndexes, scriptSheet, columnCount, true, myFormats.white, myFormats.black)
    console.log('Doing Chapter Formatting')
    await doSomeFormatting(excel, chapterIndexes, scriptSheet, columnCount, false, myFormats.chapterGreen, myFormats.black)

    if (isProtected){
      await lockColumns();
    }
  })
  
}

async function getSceneWordCount(){
  let myData = []
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let startRowIndex = firstDataRow - 1;
    let rowCount = lastDataRow - firstDataRow + 1;
    let sceneRange = scriptSheet.getRangeByIndexes(startRowIndex, sceneIndex, rowCount, 1);
    let countRange = scriptSheet.getRangeByIndexes(startRowIndex, sceneWordCountCalcIndex, rowCount, 1);
    sceneRange.load('values');
    countRange.load('values');
    await excel.sync();

    for (let i = 0; i < sceneRange.values.length; i++){
      if (myData.length == 0){
        if (sceneRange.values[i][0] > 0) {
          let thisData = { scene: sceneRange.values[i][0], wordCount: countRange.values[i][0]}
          myData.push(thisData);
        } 
      } else {
        if (i == 0){
          let thisData = { scene: sceneRange.values[i][0], wordCount: countRange.values[i][0]}
          myData.push(thisData);
        } else {
          if (sceneRange.values[i][0] != sceneRange.values[i - 1][0]){
            let thisData = { scene: sceneRange.values[i][0], wordCount: countRange.values[i][0]}
          myData.push(thisData);
          }
        }
      }
    }
    console.log('sceneWordCount data:', myData);
  })
  return myData; 
}

async function getDirectorDataV2(character){
  let myData = [];
  await Excel.run(async (excel) => {
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let isProtected = await unlockIfLocked();
    let app = excel.workbook.application;
    app.suspendScreenUpdatingUntilNextSync();
    let indexDetails = await getFirstLastIndex();
    //get character column range
    let characterRange = scriptSheet.getRangeByIndexes(indexDetails.rowIndex, characterIndex, indexDetails.rowCount, 1);
    let sceneRange = scriptSheet.getRangeByIndexes(indexDetails.rowIndex, sceneIndex, indexDetails.rowCount, 1);
    let numberRange = scriptSheet.getRangeByIndexes(indexDetails.rowIndex, numberIndex, indexDetails.rowCount, 1);
    let ukNumTakesRange = scriptSheet.getRangeByIndexes(indexDetails.rowIndex, ukTakesIndex, indexDetails.rowCount, 1);
    let ukTakeNumRange = scriptSheet.getRangeByIndexes(indexDetails.rowIndex, ukTakeNoIndex, indexDetails.rowCount, 1);
    let ukDateRecordedRange = scriptSheet.getRangeByIndexes(indexDetails.rowIndex, ukDateIndex, indexDetails.rowCount, 1); 
    let lineWordCountRange  = scriptSheet.getRangeByIndexes(indexDetails.rowIndex, lineWordCountIndex, indexDetails.rowCount, 1); 
    let sceneWordCountRange = scriptSheet.getRangeByIndexes(indexDetails.rowIndex, sceneWordCountCalcIndex, indexDetails.rowCount, 1); 
    
    characterRange.load('values');
    sceneRange.load('values');
    numberRange.load('values');
    ukNumTakesRange.load('values');
    ukTakeNumRange.load('values');
    ukDateRecordedRange.load('values');
    lineWordCountRange.load('values');
    sceneWordCountRange.load('values')
    await excel.sync();

    let myIndexes = [];
    let index = -1;
    if (character.type == choiceType.text){
      for (let i = 0; i < characterRange.values.length; i++){
        let theValue = '' + characterRange.values[i][0].toLowerCase();
        if (theValue.includes(character.name.toLowerCase())){
          index += 1;
          myIndexes[index] = i;
        }
      }
    } else {
      for (let i = 0; i < characterRange.values.length; i++){
        let theValue = '' + characterRange.values[i][0].toLowerCase();
        if (theValue == character.name.toLowerCase()){
          index += 1;
          myIndexes[index] = i;
        }
      }
    }

    //remove duplicate values

    let uniqueIndexes = [...new Set(myIndexes)];
    myData = [];
    
    for (let i = 0; i < uniqueIndexes.length; i++){
      let theData = {
        character: characterRange.values[uniqueIndexes[i]][0],
        sceneNumber: sceneRange.values[uniqueIndexes[i]][0],
        lineNumber: numberRange.values[uniqueIndexes[i]][0],
        ukNumTakes: ukNumTakesRange.values[uniqueIndexes[i]][0],
        ukTakeNum: ukTakeNumRange.values[uniqueIndexes[i]][0],
        ukDateRecorded: ukDateRecordedRange.values[uniqueIndexes[i]][0],
        lineWordCount: lineWordCountRange.values[uniqueIndexes[i]][0],
        sceneWordCount: sceneWordCountRange.values[uniqueIndexes[i]][0]
      }
      myData.push(theData);    
    }
    console.log('myData', myData);
    if (isProtected){
      await lockColumns();
    }
  })
  return myData;
}

async function getFirstLastIndex(){
  let details = {};
  await Excel.run(async (excel) => {
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let usedRange = scriptSheet.getUsedRange()
    usedRange.load('rowIndex, rowCount, columnIndex, columnCount');
    await excel.sync()
    details = {
      rowIndex: usedRange.rowIndex,
      rowCount: usedRange.rowCount,
      columnIndex: usedRange.columnIndex,
      columnCount:  usedRange.columnCount
    }
  })
  return details;
}

async function checkAllTheSceneBreaks(){
  let details = await getFirstLastIndex();
  console.log('details', details)
  await Excel.run(async (excel) => {
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let columnCount = numberIndex - sceneLineNumberRangeIndex + 1;
    console.log(details.rowIndex, sceneLineNumberRangeIndex, details.rowCount, columnCount);
    let testRange = scriptSheet.getRangeByIndexes(details.rowIndex, sceneLineNumberRangeIndex, details.rowCount, columnCount);
    testRange.load('address, values, rowIndex');
    await excel.sync();
    console.log('address:', testRange.address, 'Values: ', testRange.values);
    let previousValue = '';
    let index = -1;
    let results = [];
    for (let i = 0; i < testRange.values.length; i++){
      if (testRange.values[i][0] !== previousValue){
        if (testRange.values[i][0].trim() != ''){
          previousValue = testRange.values[i][0];
          index += 1;
          results[index] = {
            index: i,
            sceneLineNumberRange: testRange.values[i][0],
            cue: testRange.values[i][3],
            previousCue: testRange.values[i - 1][3],
            number: testRange.values[i][4],
            previousNumber: testRange.values[i - 1][4]
          }
        }
      }
    }
    console.log('results', results)
    let maxGap = 40;
    let endMessage = []
    for (let i = 0; i < results.length; i++){
      let hasIssue = false;
      let textBit = results[i].sceneLineNumberRange;
      let lineValues = (textBit).substr(1, textBit.length - 1).split('-')
      let message = '';
      if (lineValues.length == 2){
        let start = parseInt(lineValues[0]);
        let end = parseInt(lineValues[1]);
        if ((start == results[i].cue) && (start == results[i].number)){
          message = results[i].sceneLineNumberRange + ' is good so far';
          if (results[i].cue == results[i].previousCue){
            message += ' ============> But fails on cue';
            hasIssue = true;
          } else {
            message += ' And good on cue';
          }
          if (results[i].number == results[i].previousNumber){
            message += ' ============> But fails on number';
            hasIssue = true;
          } else {
            message += ' And good on number';
          }
        } else {
          let first = -1;
          message += results[i].sceneLineNumberRange + ' does not match cue: ' + results[i].cue + ' or number: ' + results[i].number;
          console.log('Loop start:', results[i].index - maxGap,'Loop end', results[i].index);
          for (let test = (results[i].index - maxGap); test < results[i].index; test++){
            if (test >= 0){
              if ((start == testRange.values[test][3]) && (start == testRange.values[test][4])){
                first = test;
                break;
              }
            }
          }
          if (first != -1){
            console.log('Index', first + testRange.rowIndex, 'value', results[i].sceneLineNumberRange, 'rows', results[i].index - first);
            let tempRange = scriptSheet.getRangeByIndexes(first + testRange.rowIndex, sceneLineNumberRangeIndex, results[i].index - first, 1);
            tempRange.values = results[i].sceneLineNumberRange;
            await excel.sync();
            hasIssue = true;
            message += '. Hopefully fixed'
          } else {
            message += '. Not within ' + maxGap + ' rows'
            hasIssue = true;
          }
        }
      } else {
        message += results[i].sceneLineNumberRange + ' is not a valid line number range';
      }
      console.log(message);
      if (hasIssue){
        endMessage.push(message);
      }
    }
    if (endMessage.length == 0){
      console.log('No issues');
    } else {
      console.log(endMessage);
    }
  });
}

async function copyNewText(){
  let details = await getFirstLastIndex();
  let isProtected = await unlockIfLocked();
    
  await Excel.run(async (excel) => {
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const newTextSheet = excel.workbook.worksheets.getItem(newTextSheetName) ;
    let startRowIndexCurrent = 2;
    let rowCount = details.rowCount - (startRowIndexCurrent - details.rowIndex);
    let columnCount = usScriptColumnIndex - cueIndex + 1;
    let currentRange = scriptSheet.getRangeByIndexes(startRowIndexCurrent, cueIndex, rowCount, columnCount);
    currentRange.columnHidden = false;
    await excel.sync();
    
    currentRange.clear('Contents');
    currentRange.clear('Formats');
    await excel.sync();

    let newUsedRange = newTextSheet.getUsedRange();
    newUsedRange.load('address, rowIndex, rowCount');
    await excel.sync();
    console.log('address', newUsedRange.address, 'rowCount:', newUsedRange.rowCount)
    let newTextRowIndex = 1;
    let newTextRowCount = newUsedRange.rowCount - (newTextRowIndex - newUsedRange.rowIndex);
    let newTextColumnIndexOne = 0;
    let newTextColumnCountOne = 3;
    let newTextColumnIndexTwo = 3;
    let newTextColumnCountTwo = 5;
    
    let newTextRangePartOne = newTextSheet.getRangeByIndexes(newTextRowIndex, newTextColumnIndexOne,  newTextRowCount, newTextColumnCountOne);
    let newTextRangePartTwo = newTextSheet.getRangeByIndexes(newTextRowIndex, newTextColumnIndexTwo,  newTextRowCount, newTextColumnCountTwo);
    
    let myNewRangePartOne = scriptSheet.getRangeByIndexes(startRowIndexCurrent, cueIndex, 1, 1);
    let myNewRangePartTwo = scriptSheet.getRangeByIndexes(startRowIndexCurrent, stageDirectionWallaDescriptionIndex, 1, 1);
    
    myNewRangePartOne.copyFrom(newTextRangePartOne, "All");
    myNewRangePartTwo.copyFrom(newTextRangePartTwo, "All");
    

    await excel.sync();
  })
  if (isProtected){
    await lockColumns();
  }
}
async function reconcileLocations(doCorrection){
  let details = await getFirstLastIndex();
  let isProtected = await unlockIfLocked();
  
  await Excel.run(async (excel) => {
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let startRowIndex = details.rowIndex + 2;
    let rowCount = details.rowCount - startRowIndex + 1;
    let typeCodeRange = scriptSheet.getRangeByIndexes(startRowIndex, typeCodeIndex, rowCount, 1);
    let locationRange = scriptSheet.getRangeByIndexes(startRowIndex, locationIndex, rowCount, 1);
    typeCodeRange.load('values, rowIndex');
    locationRange.load('values, rowIndex');
    await excel.sync();
    let indexes = [];
    let index = - 1
    for (let i = 0; i < locationRange.values.length; i++){
      if (locationRange.values[i][0].trim() != ''){
        index += 1;
        indexes[index] = {
          index: i,
          rowIndex: i + locationRange.rowIndex,
          location: locationRange.values[i][0]
        }
      }
    }
    let chapterAndScenes = [];
    index = -1;
    for (let i = 0; i < typeCodeRange.values.length; i++){
      let typeCode = typeCodeRange.values[i][0];
      if ((typeCode == myTypes.chapter)||(typeCode == myTypes.scene)){
        index += 1
        chapterAndScenes[index] = {
          index: i,
          rowIndex: i + typeCodeRange.rowIndex,
          typeCode: typeCodeRange.values[i][0]
        }
      }
    }
    let result = [];
    console.log('locations', indexes, 'typeCodes', chapterAndScenes);
    for (let i = 0; i < indexes.length; i++){
      for (let j = 0; j < chapterAndScenes.length; j++){
        if (chapterAndScenes[j].rowIndex == indexes[i].rowIndex){
          result[i] = {
            location: indexes[i],
            typeCode: chapterAndScenes[j],
            match: 0
          }
          break;
        } else if (chapterAndScenes[j].rowIndex > indexes[i].rowIndex){
          //We are using the previous scene/chapter
          result[i] = {
            location: indexes[i],
            typeCode: chapterAndScenes[j - 1],
            match: (chapterAndScenes[j - 1].rowIndex - indexes[i].rowIndex)
          }
          break;
        }
      }
    }
    console.log('Result ', result);
    duplicates = [];
    index = -1;
    for (let i = 1; i < result.length; i++){
      if (result[i - 1].typeCode.rowIndex == result[i].typeCode.rowIndex){
        index += 1
        duplicates[index] = {
          first: result[i - 1],
          second: result[i]
        }
      }
    }
    console.log('Duplicates:', duplicates);
    if (doCorrection){
      let myColumnCount = otherNotesIndex - locationIndex + 1;
      let myLocationRange = [];
      let myTargetRange = [];
      let myIndex = -1;
      let targetRowIndex;
      for (let i = 0; i < result.length; i++){
        if(result[i].match < -1){
          //So we need to move the details
          myIndex += 1
          console.log('Match', result[i].match)
          console.log(i, myIndex, 'location params', result[i].location.rowIndex, locationIndex, 1, myColumnCount);
          myLocationRange[myIndex] = scriptSheet.getRangeByIndexes(result[i].location.rowIndex, locationIndex, 1, myColumnCount);
          if (result[i].typeCode.typeCode == myTypes.scene){
            targetRowIndex = result[i].typeCode.rowIndex
          } else {
            targetRowIndex = result[i].typeCode.rowIndex + 1
          }
          console.log(i, myIndex, 'target params', targetRowIndex, locationIndex, 1, 1);
          myTargetRange[myIndex] = scriptSheet.getRangeByIndexes(targetRowIndex, locationIndex, 1, 1);
          myTargetRange[myIndex].copyFrom(myLocationRange[myIndex], 'Values');
          myTargetRange[myIndex].copyFrom(myLocationRange[myIndex], 'Formats');
          myLocationRange[myIndex].clear('Contents');
          myLocationRange[myIndex].clear('Formats');
        }
        await excel.sync();
      }  
    }
  })
  if (isProtected){
    await lockColumns();
  }
}


async function autoSceneBlockCreation(){
  let myList = await createSceneList();
  //for (let i = 2; i < 5; i++){
  for (let i = 2; i < myList.length; i++){
    console.log('Doing: ', i, 'of', myList.length, 'Name, ', myList[i].display)
    await doTheActualSceneBlock(i);    
  }

}

async function testFontColor(){
  await Excel.run(async (excel) => {
    const activeCell = excel.workbook.getActiveCell();
    let myFont = activeCell.format.font;
    myFont.load('color');
    await excel.sync();
    console.log(myFont.color);
  })
}


async function loadNewSheetCharacters(){
  await Excel.run(async function(excel){ 
    let characterlistSheet = excel.workbook.worksheets.getItem(characterListName);
    let characters = await getCharacters(newTextSheetName, 2);
    console.log('the characters', characters);
    let characterRange = characterlistSheet.getRange('clNewCharacters');
    characterRange.clear("Contents");
    characterRange.load('values');
    await excel.sync();
    console.log(characterRange.values);
    characterRange.values = characters;
    await excel.sync();
    characterRange.removeDuplicates([0], false);
    await excel.sync();
    const sortFields = [
      {
        key: 0,
        ascending: true
      }
    ]
    characterRange.sort.apply(sortFields);
    await excel.sync();
  })  
}

async function newCharacters(){
  await Excel.run(async function(excel){ 
    let characterlistSheet = excel.workbook.worksheets.getItem(characterListName);
    let currentCharacterRange = characterlistSheet.getRange('clCharacters');
    let newCharacterRange = characterlistSheet.getRange('clNewCharacters')
    currentCharacterRange.load('values');
    newCharacterRange.load('values');
    await excel.sync();

    let missingInNew = [];
    let index = -1;
    for (let i = 0; i < currentCharacterRange.values.length; i++){
      let currentChar = currentCharacterRange.values[i][0];
      if (currentChar.trim() != ''){
        let found = false;
        for (let j = 0; j < newCharacterRange.values.length; j++){
          if (currentChar.toLowerCase() == newCharacterRange.values[j][0].toLowerCase()){
            found = true;
            break;
          }
        }
        if (!found){
          index += 1
          missingInNew[index] = currentChar
        }
        //console.log('i', i, 'Character', currentChar, 'Found:', found, 'index', index, 'missing', missingInNew);
      }
    }
    console.log('Missing In New: ', missingInNew);
    let missingInCurrent = [];
    index = -1;
    for (let i = 0; i < newCharacterRange.values.length; i++){
      let currentChar = newCharacterRange.values[i][0];
      if (currentChar.trim() != ''){
        let found = false;
        for (let j = 0; j < currentCharacterRange.values.length; j++){
          if (currentChar.toLowerCase() == currentCharacterRange.values[j][0].toLowerCase()){
            found = true;
            break;
          }
        }
        if (!found){
          index += 1
          missingInCurrent[index] = currentChar
        }
        //console.log('i', i, 'Character', currentChar, 'Found:', found, 'index', index, 'missing', missingInCurrent);
      }
    }
    console.log('Missing In Current: ', missingInCurrent);
    let missingInNewWalla = await missingInNewStatus(excel, missingInNew);
    await displayMissingCharacters(excel, missingInNew, missingInCurrent, missingInNewWalla);
  })
}

async function missingInNewStatus(excel, missingInNew){
  let details = await getFirstLastIndex();
  let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
  let characterRange = scriptSheet.getRangeByIndexes(details.rowIndex, characterIndex, details.rowCount, 1);
  let typeCodeRange = scriptSheet.getRangeByIndexes(details.rowIndex, typeCodeIndex, details.rowCount, 1);
  characterRange.load('values');
  typeCodeRange.load('values');
  await excel.sync()
  let characterValues = characterRange.values.map(x => x[0]);
  let typeCodeValues = typeCodeRange.values.map(x => x[0]);
  console.log('Character Values', characterValues);
  result = [];
  for (let i = 0; i < missingInNew.length; i++){
    let testIndex = characterValues.findIndex(x => x == missingInNew[i]);
    if (testIndex != -1){
      if (typeCodeValues[testIndex] = myTypes.wallaScripted){
        result[i] = myTypes.wallaScripted;
      } else {
        result[i] = ''
      }
    }
  }
  return result
}

async function displayMissingCharacters(excel, missingInNew, missingInCurrent, missingInNewStatus){
  const comparisonSheet = excel.workbook.worksheets.getItem(comparisonSheetName);
  
  await showArrayInRange(excel, comparisonSheet, 'coNotInNew', missingInNew);
  await showArrayInRange(excel, comparisonSheet, 'coNotInCurrent', missingInCurrent);
  await showArrayInRange(excel, comparisonSheet, 'coCurrentStatus', missingInNewStatus);
  
}

async function showArrayInRange(excel, sheet, rangeName, values){
  const range = sheet.getRange(rangeName);
  range.load('rowIndex, columnIndex');
  range.clear("Contents");
  await excel.sync();
  
  if (values.length > 0){
    let tempRange = sheet.getRangeByIndexes(range.rowIndex, range.columnIndex, values.length, 1);
    let tempValues = values.map(x => [x]);
    
    tempRange.values = tempValues;
    const sortFields = [{ key: 0, ascending: true }]
    tempRange.sort.apply(sortFields);
    await excel.sync();
  }
}

async function copyTextV2(doTheCopy, doScriptDifferences){
  let details = await getFirstLastIndex()
  await Excel.run(async function(excel){ 
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const newSheet = excel.workbook.worksheets.getItem(newTextSheetName);
    let firstRowIndex = details.rowIndex + 2;
    let rowCount = details.rowCount - firstRowIndex + 1
    let typeCodeRange = scriptSheet.getRangeByIndexes(firstRowIndex, typeCodeIndex, rowCount, 1);
    let cueRange = scriptSheet.getRangeByIndexes(firstRowIndex, cueIndex, rowCount, 1);
    typeCodeRange.load('values, rowIndex');
    cueRange.load('values, rowIndex');
    let newUsedRange = newSheet.getUsedRange();
    newUsedRange.load('values, rowIndex')
    await excel.sync();
    let myTypeCodes = typeCodeRange.values.map(x => x[0]);
    let myCues = cueRange.values.map(x => x[0]);
    let newSheetCues = newUsedRange.values.map(x => x[0]);
    console.log('current sheet types:', myTypeCodes, 'rowIndex', typeCodeRange.rowIndex);
    console.log('newData', newUsedRange.values, 'Row Index', newUsedRange.rowIndex);

    let start = -1;
    let stop = -1;
    let startStopRowIndecies = []
    let index = -1
    for (let i = 0; i < myTypeCodes.length; i++){
      if (start == -1){
        if (codeSuitableForStart(myTypeCodes[i])){
         start = i
        }
      } else {
        if (codeForcesStop(myTypeCodes[i])){
          stop = i;
          index += 1
          startStopRowIndecies[index] = {
            startRowIndex: start + typeCodeRange.rowIndex,
            rowCount: stop - start
          }
          start = -1;
        }
      }
    }
    console.log('Start stop indecies', startStopRowIndecies);
    
    let rowDetails = [];
    
    for (let i = 0; i < startStopRowIndecies.length; i++){
    //for (let i = 0; i < 10; i++){
      let index = startStopRowIndecies[i].startRowIndex - cueRange.rowIndex
      let cue = {
        value: myCues[index],
        previousCue: getPreviousCue(index, myCues),
        nextCue: getNextCue(index, myCues)
      }
      console.log('Index ', index, 'Cue ', cue);
      let newSheetRowIndex = -1;
      let newSheetIndex = findNewSheetCue(cue, newSheetCues);
      if (newSheetIndex != -1){
        newSheetRowIndex = newSheetIndex + newUsedRange.rowIndex 
        rowDetails[i] = {
          currentRowIndex: startStopRowIndecies[i].startRowIndex,
          newSheetRowIndex: newSheetRowIndex,
          rowCount: startStopRowIndecies[i].rowCount
        }
      } else {
        rowDetails[i] = {
          currentRowIndex: startStopRowIndecies[i].startRowIndex,
          newSheetRowIndex: -1,
          rowCount: startStopRowIndecies[i].rowCount
        }
      }
    }
    console.log('Row details', rowDetails)

    //so now we go through and process each of them
    let errors = 0
    const newCueIndex = 0;
    const newCharacterIndex = 2;
    const newScriptIndex = 4;
    let scriptDifferences = [];
    let diffIndex = - 1
    for (let i = 0; i < rowDetails.length; i++){
      //check the cues are the same...
      let currentCue = scriptSheet.getRangeByIndexes(rowDetails[i].currentRowIndex, cueIndex, rowDetails[i].rowCount, 1)
      let currentCharacter = scriptSheet.getRangeByIndexes(rowDetails[i].currentRowIndex, characterIndex, rowDetails[i].rowCount, 1);
      let currentScript = scriptSheet.getRangeByIndexes(rowDetails[i].currentRowIndex, ukScriptIndex, rowDetails[i].rowCount, 1);
      let newCue= newSheet.getRangeByIndexes(rowDetails[i].newSheetRowIndex, newCueIndex, rowDetails[i].rowCount, 1);
      let newCharacter = newSheet.getRangeByIndexes(rowDetails[i].newSheetRowIndex, newCharacterIndex, rowDetails[i].rowCount, 1);
      let newScript = newSheet.getRangeByIndexes(rowDetails[i].newSheetRowIndex, newScriptIndex, rowDetails[i].rowCount, 1);
      currentCue.load('values, rowIndex');
      currentCharacter.load('values, rowIndex');
      currentScript.load('values, rowIndex');
      newCue.load('values, rowIndex');
      newCharacter.load('values, rowIndex');
      newScript.load('values, rowIndex');
      await excel.sync();
      let currentCueValues = currentCue.values.map(x => x[0]);
      let newCueValues = newCue.values.map(x => x[0]);
      let currentCharacterValues = currentCharacter.values.map(x => x[0]);
      let currentScriptValues = currentScript.values.map(x => x[0]);
      let newCharacterValues = newCharacter.values.map(x => x[0]);
      let newScriptValues = newScript.values.map(x => x[0]);
      //console.log("I", i, 'currentCue.rowIndex', currentCue.rowIndex, 'currentCue.values', currentCueValues, 'newCue.rowIndex', newCue.rowIndex, 'newRow.values', newCueValues);

      if (currentCueValues.length == newCueValues.length){
        for (let j = 0; j < currentCueValues.length; j++){
          if(currentCueValues[j] != newCueValues[j]){
            console.log("I", i, 'currentCue.rowIndex', currentCue.rowIndex, 'currentCue.values', currentCueValues, 'newCue.rowIndex', newCue.rowIndex, 'newRow.values', newCueValues);
            console.log('Row ', j, ' is different. Current Cue: ', currentCueValues[j], 'rowIndex: ', (j + currentCue.rowIndex), 'New cue: ' + newCueValues[j] + ' rowIndex: ' + (j + newCue.rowIndex))
            errors += 1
          }
          if ((currentCharacterValues[j] != newCharacterValues[j])||(currentScriptValues[j] != newScriptValues[j])){
            diffIndex += 1;
            scriptDifferences[diffIndex] = {
              currentCue: currentCueValues[j],
              currentRowIndex: (j + currentCue.rowIndex),
              newCue: newCueValues[j],
              newRowIndex: (j + newCue.rowIndex),
              currentScript: currentScriptValues[j],
              newScript: newScriptValues[j],
              currentCharacter: currentCharacterValues[j],
              newCharacter: newCharacterValues[j],
              scriptDiff: (currentScriptValues[j] != newScriptValues[j]),
              characterDiff: (currentCharacterValues[j] != newCharacterValues[j])
            }
          }
        }
      } else {
        console.log('Ranges are diiferent length');
        errors += 1
      }
    }
    console.log('Script differences: ', scriptDifferences);
    if (doScriptDifferences){
      await displayScriptDifferences(excel, scriptDifferences);
    }

    console.log('No of errors: ', errors);  
    if ((errors == 0) && (doTheCopy)){
      //We can continue
      let isProtected = await unlockIfLocked();
      const cueToCharacterColumns = 3;
      const newStageDirectionsIndex = 3;
      const stageToUsScriptColumns = 5;
      for (let i = 0; i < rowDetails.length; i++){
        let sourceRangeCueToCharacter = newSheet.getRangeByIndexes(rowDetails[i].newSheetRowIndex, newCueIndex, rowDetails[i].rowCount, cueToCharacterColumns);
        sourceRangeCueToCharacter.load('address')
        let destinationCueToCharacter = scriptSheet.getRangeByIndexes(rowDetails[i].currentRowIndex, cueIndex, 1, 1);
        destinationCueToCharacter.load('address');
        destinationCueToCharacter.copyFrom(sourceRangeCueToCharacter, "All", false, false);
        let sourceRangeStageToUsScript = newSheet.getRangeByIndexes(rowDetails[i].newSheetRowIndex, newStageDirectionsIndex, rowDetails[i].rowCount, stageToUsScriptColumns);
        sourceRangeStageToUsScript.load('address');
        let destinationStageToUsScript = scriptSheet.getRangeByIndexes(rowDetails[i].currentRowIndex, stageDirectionWallaDescriptionIndex, 1, 1);
        destinationStageToUsScript.load('address');
        destinationStageToUsScript.copyFrom(sourceRangeStageToUsScript, "All", false, false);
        await excel.sync();
        console.log (i, ' of ', rowDetails.length + ' completed ', sourceRangeCueToCharacter.address, ' to ', destinationCueToCharacter.address, 'and', sourceRangeStageToUsScript.address, 'to', destinationStageToUsScript.address);
      }
      if (isProtected){
        await lockColumns();
      }
    }
  })  
}

async function displayScriptDifferences(excel, differences){
  let comparisonSheet = excel.workbook.worksheets.getItem(comparisonSheetName);
  let tableRange = comparisonSheet.getRange('coTable');
  tableRange.load('rowIndex, columnIndex, columnCount');
  tableRange.clear('Contents');
  await excel.sync();
  let display = [];
  for (let i = 0; i < differences.length; i++){
    display[i] = [differences[i].currentCue, differences[i].currentCharacter, differences[i].currentScript, differences[i].newCharacter, differences[i].newScript, differences[i].characterDiff, differences[i].scriptDiff, differences[i].currentRowIndex, differences[i].newRowIndex];
  }
  let tempRange = comparisonSheet.getRangeByIndexes(tableRange.rowIndex, tableRange.columnIndex, differences.length, tableRange.columnCount);
  tempRange.values = display;
  await excel.sync();
}

function getPreviousCue(index, theCues){
  //returns the previous numeric cue, or -1 if fails
  let previousCue = -1;
  let possible;
  for (let i = index - 1; i >= 0; i--){
    possible = parseInt(theCues[i])
    if (!isNaN(possible)){
      previousCue = possible;
      break;
    }
  }
  return previousCue; 
}

function getNextCue(index, theCues){
  //returns next numeric cue, or -1 if fails
  let nextCue = -1;
  let possible;
  for (let i = index + 1; i < theCues.length; i++){
    possible = parseInt(theCues[i])
    if (!isNaN(possible)){
      nextCue = possible;
      break;
    }
  }
  return nextCue;
}

function findNewSheetCue(theCue, newCues){
  for (let i = 0; i < newCues.length; i++){
    if (newCues[i] == theCue.value){
      //found it 
      //if cue.prev and cue.next are both not equal to cue.value then thats all that's needed
      if ((theCue.value != theCue.previousCue) && (theCue.value != theCue.nextCue)){
        return i;
      } else if (theCue.value == theCue.previousCue){
        let possible = getPreviousCue(i, newCues);
        if (possible == theCue.value){
          return i;
        }
      } else if (theCue.value == theCue.nextCue){
        let possible = getNextCue(i, newCues)
        if (possible == theCue.value){
          return i;
        }
      }
    }
  }
  return -1;
}

function codeSuitableForStart(theCode){
  if ((theCode == myTypes.chapter) || (theCode == myTypes.scene) || (theCode == myTypes.line)){
    return true;
  } else {
    return false;
  }
}

function codeForcesStop(theCode){
  if ((theCode == myTypes.sceneBlock) || (theCode == myTypes.wallaScripted) || (theCode == myTypes.wallaBlock) || (theCode == '')){
    return true;
  } else {
    return false;
  }
}

async function deleteAllFX(){
  //Deletes all rows that have Cues that end with/FX
  let details = await getFirstLastIndex()
  let isProtected = await unlockIfLocked();
  await Excel.run(async function(excel){ 
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let cueRange = scriptSheet.getRangeByIndexes(details.rowIndex, cueIndex, details.rowCount, 1);
    cueRange.load('values, rowIndex');
    await excel.sync();
    let cueValues = cueRange.values.map(x => x[0]);
    let deleteIndexes = [];
    let myIndex = -1;
    for (let i = 0; i < cueValues.length; i++){
      let myString = cueValues[i].toString();
      if (myString.endsWith("/FX")){
        myIndex += 1
        deleteIndexes[myIndex] = i + cueRange.rowIndex;
      }
    }
    console.log('deleteIndexes', deleteIndexes);
    if (deleteIndexes.length > 0){
      let myCount = 0;
      for (i = deleteIndexes.length - 1; i >= 0; i--){
        myCount += 1;
        console.log('Deleting: ', myCount, ' of ', deleteIndexes.length, ': ', deleteIndexes[i])
        await deleteThisRow(excel, deleteIndexes[i])
      }
    }
    
  })
  if (isProtected){
    await lockColumns();
  }
}

async function doSomeFormatting(excel, theIndexes, sheet, columnCount, fillClear, fillColour, fontColour){
  let lineRanges = [];
  let lineNoScriptLineRangeBefore = [];
  let lineNoScriptLineRangeAfter = [];
  let lineBorders = [];
  let myCount = 0;
  for (let i = 0; i < theIndexes.length; i++){
      lineRanges[i] = sheet.getRangeByIndexes(theIndexes[i], cueIndex, 1, columnCount);
      lineNoScriptLineRangeBefore[i] = sheet.getRangeByIndexes(theIndexes[i], cueIndex, 1, 5)
      lineNoScriptLineRangeAfter[i] = sheet.getRangeByIndexes(theIndexes[i], usCueIndex, 1, 2)
      if (fillClear){
       lineRanges[i].format.fill.clear();
      } else {
        lineRanges[i].format.fill.color = fillColour;
      }

      
      lineNoScriptLineRangeBefore[i].format.font.strikethrough = false;
      lineNoScriptLineRangeAfter[i].format.font.strikethrough = false;
      lineNoScriptLineRangeBefore[i].format.font.color = fontColour;
      lineNoScriptLineRangeAfter[i].format.font.color = fontColour;
      
      lineBorders[i] = lineRanges[i].format.borders

      doBorder(lineBorders[i], 'EdgeTop');
      doBorder(lineBorders[i], 'EdgeBottom');
      doBorder(lineBorders[i], 'EdgeLeft');
      doBorder(lineBorders[i], 'EdgeRight');
      doBorder(lineBorders[i], 'InsideHorizontal');
      doBorder(lineBorders[i], 'InsideVertical');

      myCount += 1;
      if (myCount >= 100){
        myCount = 0;
        await excel.sync();
        console.log('Row ' + (i + 1) + ' of ' + theIndexes.length);
      }
    }
    await excel.sync();
}


/*
chapter: 'Chapter',
  scene: 'Scene',
  line: 'Line',
  sceneBlock: 'Scene Block',
  wallaScripted: 'Walla Scripted',
  wallaBlock: 'Walla Block'
  */