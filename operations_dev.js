const firstDataRow = 3;
const lastDataRow = 9999;
const scriptSheetName = 'Script'
const columnsToLock = "A:T";

let sceneIndex, numberIndex;
let totalTakesIndex, ukTakesIndex, ukTakeNoIndex, ukDateIndex, ukStudioIndex, ukEngineerIndex, ukMarkUpIndex;
let usTakesIndex, usTakeNoIndex, usDateIndex, usStudioIndex, usEngineerIndex, usMarkUpIndex;
let wallaTakesIndex, wallaTakeNoIndex, wallaDateIndex, wallaStudioIndex, wallaEngineerIndex, wallaMarkUpIndex; 
let mySheetColumns;
let scriptSheet;

let sceneInput, lineNoInput, chapterInput

function auto_exec(){
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
  totalTakesIndex = findColumnIndex('Total Takes');
  
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

  chapterCalculationIndex = findColumnIndex('Chapter Calculation');

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
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    scriptSheet.protection.load('protected');
    let range = scriptSheet.getRange(columnsToLock);
    await excel.sync();
    
    console.log(scriptSheet.protection.protected);
    if (!scriptSheet.protection.protected){
      console.log("Not locked");
      range.format.protection.locked = true;
      scriptSheet.protection.protect({ selectionMode: "Normal", allowAutoFilter: true });
      await excel.sync();
      console.log("Now locked");
    } else {
      console.log("Locked");
    }
  })   
}

async function unlock(){
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    scriptSheet.protection.load('protected');
    await excel.sync();
    if (!scriptSheet.protection.protected){
      console.log("Already unlocked");
    } else {
      console.log("Currently locked");
      scriptSheet.protection.unprotect("")
      await excel.sync();
      console.log("Now not locked");
    }
  })
}

async function applyFilter(){
  /*Jade.listing:{"name":"Apply filter","description":"Applies empty filter to sheet"}*/
  await Excel.run(async function(excel){
    await unlock();
    const myRange = await getDataRange(excel);
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    scriptSheet.autoFilter.apply(myRange, 0, { criterion1: "*", filterOn: Excel.FilterOn.custom});
    scriptSheet.autoFilter.clearCriteria();
    await excel.sync();
    await lockColumns();
  })
}

async function removeFilter(){
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    scriptSheet.autoFilter.load('enabled')
    await excel.sync()
    if (scriptSheet.autoFilter.enabled){
      console.log("Autofilter enabled")
      scriptSheet.autoFilter.remove();
      await excel.sync();
    } else {
      console.log("Autofilter not enabled")
    }
    await lockColumns();
  })
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

    console.log("Found Index");
    console.log(myIndex);
    
    if (myIndex == -1){
      alert('Invalid Line Number');
    } else {
      scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      const myTarget = scriptSheet.getRangeByIndexes(myIndex + 2, startColumn, 1, 1);
      myTarget.select();
      await excel.sync();
    }
  })
}

async function findChapter(chapter){
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load("rowIndex");
    activeCell.load(("columnIndex"))
    await excel.sync()
    const startRow = activeCell.rowIndex;
    const startColumn = activeCell.columnIndex
    let range = await getChapterRange(excel);
    range.load("values");
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
      scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
      const myTarget = scriptSheet.getRangeByIndexes(myIndex + 2, startColumn, 1, 1);
      myTarget.select();
      await excel.sync();
    }
  })
}

async function firstScene(){
  await Excel.run(async function(excel){
    const minAndMax = await getSceneMaxAndMin();
    console.log("Min and Max");
    console.log(minAndMax);
    await findSceneNo(minAndMax.min);
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
  endRow.load("rowindex");
  await excel.sync();
  range = scriptSheet.getRangeByIndexes(2, sceneIndex, endRow.rowIndex, 1);
  await excel.sync();
  return range;
}

async function getLineRange(excel){
  scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
  const endRow = scriptSheet.getUsedRange().getLastRow();
  endRow.load("rowindex");
  await excel.sync();
  range = scriptSheet.getRangeByIndexes(2, numberIndex, endRow.rowIndex, 1);
  await excel.sync();
  return range;
}

async function getChapterRange(excel){
  scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
  const endRow = scriptSheet.getUsedRange().getLastRow();
  endRow.load("rowindex");
  await excel.sync();
  range = scriptSheet.getRangeByIndexes(2, chapterCalculationIndex, endRow.rowIndex, 1);
  await excel.sync();
  return range;
}

async function getDataRange(excel){
  scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
  const myLastRow = scriptSheet.getUsedRange().getLastRow();
  const myLastColumn = scriptSheet.getUsedRange().getLastColumn();
  myLastRow.load("rowindex");
  myLastColumn.load("columnindex")
  await excel.sync();
  
  const range = scriptSheet.getRangeByIndexes(1,0, myLastRow.rowIndex, myLastColumn.columnIndex + 1);
  await excel.sync();
  
  return range
}

async function getTargetSceneNumber(){
  const textValue = sceneInput.value;
  const sceneNumber = parseInt(textValue);
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
  if (isNaN(chapterNumber)){
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
    console.log(min.values[0][0]);
    console.log(max.values[0][0]);
    
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
    await unlock();
    console.log(studioRange);
    markupRange.values = [[markupText]]
    dateRange.values = [[dateInFormat()]];
    studioRange.values = [[studioText]];
    engineerRange.values = [[engineerText]];
    await excel.sync();
    await lockColumns();
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

async function theFormulas(){
  const sceneLineCountColumn = findColumnLetter("Scene Line Count") //B
  const sceneLineNumberRangeColumn = findColumnLetter("Scene Line Number Range"); //C
  const sceneNumberColumn = findColumnLetter("Scene Number"); //D
  const cueColumn = findColumnLetter("Cue") //F
  const numberColumn = findColumnLetter("Number"); //G
  const stageDirectionWallaDescriptionColumn = findColumnLetter("Stage Direction/ Walla description") //J
  const UKScriptColumn = findColumnLetter("UK script"); //K
  const ukNoOfTakesColumn = findColumnLetter("UK No of takes"); //T
  const ukTakeNoColumn = findColumnLetter("UK Take No"); //V
  console.log("uKTakeNoColumn");
  console.log(ukTakeNoColumn);
  const positionMinusColumn = findColumnLetter("Position -"); //BT
  const startLineColumn = findColumnLetter("Start Line"); //BU
  const positionEndSqaureBracketColumn = findColumnLetter("Position ]"); //BV
  const endLineColumn = findColumnLetter("End Line"); //BW
  const lineWordCountColumn = findColumnLetter("Line Word Count") //BY
  const sceneColumn = findColumnLetter("Scene"); //BZ
  const lineColumn = findColumnLetter("Line"); // CA
  const wordCountToThisLineColumn = findColumnLetter("Word count to this line"); //CB
  const sceneWordCountCalcColumn = findColumnLetter("Scene word count calc"); //CC
  const positionChapterColumn = findColumnLetter("Position Chapter"); //CD
  const chapterCalculationColumn = findColumnLetter("Chapter Calculation"); //CF
  const sceneBordersColumn = findColumnLetter("Scene Borders"); //CH
  const sceneLineCountCalculationColumn = findColumnLetter("Scene Line Count Calculation"); //CG
  const firstRow = "" + firstDataRow;
  const firstRestRow = "4";
  const lastRow = "" + lastDataRow;
  const columnFormulae = [
    {
      columnName: "Scene Word Count", //A
      formulaFirst: '=""',
      formulaRest: '=IF(' + cueColumn + firstRestRow + '="","",' + sceneWordCountCalcColumn + firstRestRow + ')'
    },
    {
      columnName: "Position -",
      formulaFirst: '=IF(' + sceneLineNumberRangeColumn + firstRow + '="",0,FIND("-",' + sceneLineNumberRangeColumn + firstRow + '))',
      formulaRest: '=IF(' + sceneLineNumberRangeColumn + firstRestRow + '="",0,FIND("-",' + sceneLineNumberRangeColumn + firstRestRow + '))'
    },
    {
      columnName: "Start Line",
      formulaFirst: 0,
      formulaRest: "=IF(" + positionMinusColumn + firstRestRow + "=0," + startLineColumn + firstRow + ",VALUE(MID(" + sceneLineNumberRangeColumn + firstRestRow + ",2," + positionMinusColumn + firstRestRow + "-2)))"
    },
    {
      columnName: "Position ]",
      formulaFirst: '=IF(' + sceneLineNumberRangeColumn + firstRow + '="",0,FIND("]",' + sceneLineNumberRangeColumn + firstRow + '))',
      formulaRest: '=IF(' + sceneLineNumberRangeColumn + firstRestRow + '="",0,FIND("]",' + sceneLineNumberRangeColumn + firstRestRow + '))'
    },
    {
      columnName: "End Line",
      formulaFirst: 0,
      formulaRest: "=IF(" + positionEndSqaureBracketColumn + firstRestRow + "=0," + endLineColumn + firstRow + ",VALUE(MID(" + sceneLineNumberRangeColumn + firstRestRow + "," + positionMinusColumn + firstRestRow + "+1," + positionEndSqaureBracketColumn + firstRestRow + "-" + positionMinusColumn + firstRestRow + "-1)))"
    },
    {
      columnName: "Valid Line Number",
      formulaFirst:  "=AND(" + numberColumn + firstRow + ">=" + startLineColumn + firstRow + ", " + numberColumn + firstRow + "<=" + endLineColumn + firstRow + ")",
      formulaRest: "=AND(" + numberColumn + firstRestRow + ">=" + startLineColumn + firstRestRow + ", " + numberColumn + firstRestRow + "<=" + endLineColumn + firstRestRow + ")"
    },
    {
      columnName: "Line Word Count", //BY
      formulaFirst:  0,
      formulaRest: '=IF(NOT(OR(' + ukTakeNoColumn + firstRestRow + '="",' + ukTakeNoColumn + firstRestRow + '=1)), 0, LEN(TRIM(' + UKScriptColumn + firstRestRow + ')) - LEN(SUBSTITUTE(' + UKScriptColumn + firstRestRow + ', " ", "")) + 1)'
    },
    {
      columnName: "Scene",
      formulaFirst:  1,
      formulaRest: '=IF(' + sceneNumberColumn + firstRestRow + '="",' +sceneColumn + firstRow + ',VALUE(' + sceneNumberColumn + firstRestRow + '))'
    },
    {
      columnName: "Line",
      formulaFirst:  0,
      formulaRest: "=" + numberColumn + firstRestRow + ""
    },
	  {
	    columnName: "Word count to this line",
      formulaFirst:  0,
      formulaRest: "=IF(" + sceneColumn + firstRestRow + "=" + sceneColumn + firstRow + "," + wordCountToThisLineColumn + firstRow + "+" + lineWordCountColumn + firstRestRow + "," + lineWordCountColumn + firstRestRow + ")"
  	},
	  {
	    columnName: "Scene word count calc",
      formulaFirst:  0,
      formulaRest: "=VLOOKUP(" + endLineColumn + firstRestRow + "," + "$" + lineColumn + "$" + firstRestRow + ":$" + wordCountToThisLineColumn + "$" + lastRow + ",2,FALSE)"
  	},
    {
      columnName: "Position Chapter", //CD
      formulaFirst: '=IF(' + stageDirectionWallaDescriptionColumn + firstRow + '="","",IF(ISERROR(FIND("Chapter",' + stageDirectionWallaDescriptionColumn + firstRow + ')),"",FIND("Chapter",' + stageDirectionWallaDescriptionColumn + firstRow + ')))',
      formulaRest: '=IF('+ stageDirectionWallaDescriptionColumn + firstRestRow + '="","",IF(ISERROR(FIND("Chapter",' + stageDirectionWallaDescriptionColumn + firstRestRow + ')),"",FIND("Chapter",' + stageDirectionWallaDescriptionColumn + firstRestRow + ')))'
    },
    {
      columnName: "Chapter Calculation", //CE
      formulaFirst: '=VALUE(IF(' + positionChapterColumn + firstRow + '="","",MID(' + stageDirectionWallaDescriptionColumn + firstRow + ',' + positionChapterColumn + firstRow + '+7,99)))',
      formulaRest: '=VALUE(IF(' + positionChapterColumn + firstRestRow + '="",' + chapterCalculationColumn + firstRow + ',MID(' + stageDirectionWallaDescriptionColumn + firstRestRow + ',' + positionChapterColumn + firstRestRow + '+7,99)))'
    },
    {
      columnName: "Chapter", //E
      formulaFirst: '=IF(' + cueColumn + firstRow + '="", "","Chapter " & TEXT(' + chapterCalculationColumn + firstRow + ', "0"))',
      formulaRest: '=IF(' + cueColumn + firstRestRow + '="", "","Chapter " & TEXT(' + chapterCalculationColumn + firstRestRow + ', "0"))'
    },
    
    {
      columnName: "Scene Borders", //CH
      formulaFirst: "Start",
      formulaRest: '=IF(' + cueColumn + firstRestRow + '="", "",IF(' + sceneBordersColumn + firstRow + '="Start","Original",IF(' + sceneColumn + firstRestRow + '<>' + sceneColumn + firstRow + ',"Original","Copy")))'
    },
    {
      columnName: "Scene Line Count Calculation", //CG
      formulaFirst: 0,
      formulaRest: '=' + endLineColumn + firstRestRow + '-' + startLineColumn + firstRestRow + '+1'
    },
    {
      columnName: "Scene Line Count", //B
      formulaFirst: 0,
      formulaRest: '=IF(' + cueColumn + firstRestRow + '="","",' + sceneLineCountCalculationColumn + firstRestRow + ')'
    }
  ]
  
  await unlock();
  await Excel.run(async function(excel){ 
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    for (let columnFormula of columnFormulae){
      const columnLetter = findColumnLetter(columnFormula.columnName);
      const myRange = columnLetter + firstRestRow + ":" + columnLetter + lastRow ;
      const myTopRow = columnLetter + firstRow;
      console.log(myRange + "  " + myTopRow);
      const range = scriptSheet.getRange(myRange);
      const topRowRange = scriptSheet.getRange(myTopRow);
      console.log(columnFormula.formulaRest + "   " + columnFormula.formulaFirst);
      range.formulas = columnFormula.formulaRest;
      topRowRange.formulas = columnFormula.formulaFirst;
      await excel.sync();
      console.log(range.formulas + "   " + topRowRange.formulas);
    }
  })
  await lockColumns();
}
async function insertRow(){
  let activeCell
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex');
    const dataRange = await getDataRange(excel);
    dataRange.load('address');
    await excel.sync();
    console.log(dataRange.address);
    console.log(activeCell.rowIndex);
    const myLastColumn = dataRange.getLastColumn();
    myLastColumn.load("columnindex")
    await excel.sync();
  
    const myRow = scriptSheet.getRangeByIndexes(activeCell.rowIndex,0, 1, myLastColumn.columnIndex+1);
    myRow.load('address');
    await excel.sync();
    console.log(myRow.address);
    await unlock();
    const newRow = myRow.insert("Down");
    newRow.load('address');
    myRow.load('address');
    await excel.sync();
    console.log(myRow.address);
    console.log(newRow.address);
    newRow.copyFrom(myRow, "All");
    await excel.sync();
    await correctFormulas(activeCell.rowIndex + 1);
    activeCell.select();
    await excel.sync();
  })
  return activeCell.rowIndex;
}

async function insertRowV2(currentRowIndex){
  await unlock();
  await Excel.run(async function(excel){
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const dataRange = await getDataRange(excel);
    const myLastColumn = dataRange.getLastColumn();
    myLastColumn.load("columnindex")
    await excel.sync();
    const myRow = scriptSheet.getRangeByIndexes(currentRowIndex, 0, 1, myLastColumn.columnIndex+1);
    const newRow = myRow.insert("Down");
    await excel.sync();
    newRow.copyFrom(myRow, "All");
    await excel.sync();
    await correctFormulas(currentRowIndex + 1);
  })
  await lockColumns();
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
    await unlock();
    myRow.delete("Up");
    myRow.load('address');
    await excel.sync();
    console.log(myRow.address);
    await correctFormulas(activeCell.rowIndex);
    await doTakesAndNumTakes(activeCell.rowIndex - 1, 'UK', false, false, false, false);
    selectCell.select();
    await excel.sync();
  })
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
      columnName: "Scene", //CA
      formulaRest: '=IF(' + sceneNumberColumn + firstRow + '="",' +sceneColumn + (firstRow - 1) + ',VALUE(' + sceneNumberColumn + firstRow + '))'
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
      columnName: "Scene Borders", //CH
      formulaRest: '=IF(' + sceneBordersColumn + (firstRow - 1) + '="Start","Original",IF(' + sceneColumn + firstRow + '<>' + sceneColumn + (firstRow - 1) + ',"Original","Copy"))'
    }
  ]
  
  await unlock();
  await Excel.run(async function(excel){ 
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    for (let columnFormula of columnFormulae){
      const columnLetter = findColumnLetter(columnFormula.columnName);
      const myRange = columnLetter + firstRow + ":" + columnLetter + (firstRow +1) ;
      console.log("Range to replace: " + myRange);
      const range = scriptSheet.getRange(myRange);
      console.log("Formula: " + columnFormula.formulaRest);
      range.formulas = columnFormula.formulaRest;
      await excel.sync();
      console.log("Formula after sync: " + range.formulas);
    }
  })
  await lockColumns();
}

async function insertTake(country, doAdditional, includeMarkUp, includeStudio, includeEngineer){
  const currentRowIndex = await insertRow();
  const doDate = true;
  console.log(currentRowIndex);
  await unlock();
  await doTakesAndNumTakes(currentRowIndex, country, doDate, doAdditional, includeMarkUp, includeStudio, includeEngineer);
  await lockColumns();
}

function zeroElement(value){
  return value[0];
}

async function addTakeDetails(country, doDate){
  await unlock()
  let myAction = radioButtonChoice();
  console.log('The action: ', myAction);

  await Excel.run(async function(excel){ 
    const activeCell = excel.workbook.getActiveCell();
    let selectCell = activeCell.getOffsetRange(1, 0);
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
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
      console.log('US Date Index', usDateIndex);
      markUpIndex = usMarkUpIndex;
      studioIndex = usStudioIndex;
      engineerIndex = usEngineerIndex;
      countryTakes = lineDetails.usTakes
      newLine = lineDetails.usTakes + 1;
    }else if (country == 'Walla'){
      takeNoIndex = wallaTakeNoIndex;
      dateRecordedIndex = wallaDateIndex;
      console.log('Walla Date Index', wallaDateIndex);
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
      await insertRowV2(currentRowIndex)
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

    await unlock();
    selectCell.select();
    await excel.sync();
  

    console.log("Line Details")
    console.log(lineDetails);
    doTheTidyUp(lineDetails)
        
  })
  await lockColumns();
}


async function findDetailsForThisLine(){
  await unlock();
  let result = {};
  
  await Excel.run(async function(excel){ 
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
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
  await unlock();
  await Excel.run(async function(excel){
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
        scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
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
              scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
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
              await unlock();
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
            scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
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
          scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
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
            await unlock();
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
    await unlock();
    await doTheTidyUp(lineDetails)
  })
    
  await lockColumns()
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

  })

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
  await unlock();
  await Excel.run(async function(excel){ 
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
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
  })
  await lockColumns();
}
async function hideRows(visibleType, country){
  let noOfTakesColumn;
  let takeNumberColumn;
  if (country == "UK"){
    noOfTakesColumn = findColumnLetter("UK No of takes");
    takeNumberColumn = findColumnLetter("UK Take No")
  }
  await unlock();
  await Excel.run(async function(excel){ 
    let myMessage = tag('takeMessage')
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let myRange = scriptSheet.getRange(noOfTakesColumn + firstDataRow + ":" + takeNumberColumn + lastDataRow);
    myRange.load('values')
    await excel.sync();
    console.log(myRange.values)
    console.log(myRange.values.length)
    console.log(myRange.values[0].length)

    //First unhide all
    let hideRange = scriptSheet.getRangeByIndexes(firstDataRow - 1, 0, lastDataRow - 2, 1);
    hideRange.load('address');
    hideRange.rowHidden = false;
    await excel.sync();
    console.log(hideRange.address);
    myMessage.innerText = "Showing all takes"

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
  await lockColumns();
}

async function showHideColumns(columnType){
  const sheetName = "Settings"
  const rangeName = "columnHide"
  let columnMessage = tag('columnMessage')
  let hideUnedited = tag('hideUnedited').checked;
  console.log('Hide Unedited', hideUnedited);
  await unlock();
  await Excel.run(async function(excel){ 
    let app = excel.workbook.application;
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
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
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
UK script	10	J	9
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

async function fillSceneNumber(){
  await Excel.run(async function(excel){ 
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let app = excel.workbook.application;
    app.suspendScreenUpdatingUntilNextSync();
    app.suspendApiCalculationUntilNextSync();
    const sceneNumberColumn = findColumnLetter('Scene Number');
    const sceneBordersColumn = findColumnLetter('Scene Borders');
    const sceneLineNumberRangeColumn = findColumnLetter('Scene Line Number Range')

    let borderRange = scriptSheet.getRange(sceneBordersColumn + firstDataRow + ":" +  sceneBordersColumn + lastDataRow);
    let sceneRange = scriptSheet.getRange(sceneNumberColumn + firstDataRow + ":" +  sceneNumberColumn + lastDataRow);
    let lineNoRange = scriptSheet.getRange(sceneLineNumberRangeColumn +firstDataRow + ':' + sceneLineNumberRangeColumn +lastDataRow);
    borderRange.load('values');
    sceneRange.load('values');
    sceneRange.load('numberFormat');
    lineNoRange.load('values')
    await excel.sync();
    console.log(lineNoRange.values)
    app.suspendScreenUpdatingUntilNextSync();
    app.suspendApiCalculationUntilNextSync();
    let borderValues = borderRange.values.map(x => x[0]);
    let sceneValues = sceneRange.values
    let sceneFormat = sceneRange.numberFormat;
    let lineNoValues = lineNoRange.values

    let currentValue = '';
    let currentFormat = '';
    let currentLineNo = '';
    for (let i = 0; i < borderValues.length; i++){
      if (borderValues[i] == 'Original'){
        currentValue = sceneValues[i][0];
        currentFormat = sceneFormat[i][0];
        currentLineNo = lineNoValues[i][0];
      } else if (borderValues[i] == 'Copy'){
        sceneValues[i][0] = currentValue;
        sceneFormat[i][0] = currentFormat;
        lineNoValues[i][0] = currentLineNo;
      } else if(borderValues[i] == ''){
        sceneValues[i][0] = '';
        lineNoValues[i][0] = '';
      }
    }

    sceneRange.values = sceneValues;
    sceneRange.numberFormat = sceneFormat;
    lineNoRange.values = lineNoValues;
    await excel.sync();

  }) 
}

async function setDefaultColumnWidths(){
  await Excel.run(async function(excel){ 
    scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
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
  })
}

async function setUpEvents(){
  sceneInput = tag('scene');
  lineNoInput = tag('lineNo')
  chapterInput = tag('chapter')
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
  if (admin.style.display === 'none'){
    admin.style.display = 'block';
  } else {
    admin.style.display = 'none';
  }
}
