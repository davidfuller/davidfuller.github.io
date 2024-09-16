async function auto_exec(){
  console.log('Operations loaded');
  console.log(jade_modules)
}
let mySheetColumns;
const firstDataRow = 3;
const lastDataRow = 9999;

async function getMySheetColumns(){
  console.log(mySheetColumns);
  return mySheetColumns
}

function findColumnIndex(name){
  return mySheetColumns.find((col) => col.name === name).index;
}

function findColumnLetter(name){
  return mySheetColumns.find((col) => col.name === name).column;
}


const columnsToLock = "A:Y";
const testRange = "A:B, D:E";

async function test(){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRanges(testRange);
    range.load('address');
    await excel.sync();
    console.log(range.address);
    console.log(range.areas.length);
  })
}
const myColumns = 
  [
    {
      columnName: "Scene",
      columnNo: 77
    },
    {
      columnName: "Line",
      columnNo: 78
    },
    {
      columnName: "UK Date Recorded",
      columnNo: 21
    },
    {
      columnName: "UK Studio",
      columnNo: 22
    },
    {
      columnName: "UK Engineer",
      columnNo: 23
    },
    {
      columnName: "US Date Recorded",
      columnNo: 26
    },
    {
      columnName: "US Studio",
      columnNo: 27
    },
    {
      columnName: "US Engineer",
      columnNo: 28
    },
    {
      columnName: "Walla Date Recorded",
      columnNo: 44
    },
    {
      columnName: "Walla Studio",
      columnNo: 45
    },
    {
      columnName: "Walla Engineer",
      columnNo: 46
    }
  ];

  /*
const sceneInput = tag("scene");
sceneInput.onkeydown = function(event){
  if(event.key === 'Enter'){
    alert(sceneInput.value)
  }
}
*/


async function lockColumns(){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange(columnsToLock);
    
    sheet.protection.load('protected');
    await excel.sync();
    
    console.log(sheet.protection.protected);
    if (!sheet.protection.protected){
      console.log("Not locked");
      range.format.protection.locked = true;
      sheet.protection.protect({ selectionMode: "Normal", allowAutoFilter: true });
      await excel.sync();
      console.log("Now locked");
    } else {
      console.log("Locked");
    }
  })   
}

async function unlock(){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    sheet.protection.load('protected');
    await excel.sync();
    if (!sheet.protection.protected){
      console.log("Already unlocked");
    } else {
      console.log("Currently locked");
      sheet.protection.unprotect("")
      await excel.sync();
      console.log("Now not locked");
    }
  })
}

async function applyFilter(){
  /*Jade.listing:{"name":"Apply filter","description":"Applies empty filter to sheet"}*/
  await Excel.run(async function(excel){
    await unlock();
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    const myRange = await getDataRange(excel);
    sheet.autoFilter.apply(myRange, 0, { criterion1: "*", filterOn: Excel.FilterOn.custom});
    sheet.autoFilter.clearCriteria();
    await excel.sync();
    await lockColumns();
  })
}

async function removeFilter(){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.load('enabled')
    await excel.sync()
    if (sheet.autoFilter.enabled){
      console.log("Autofilter enabled")
      sheet.autoFilter.remove();
      await excel.sync();
    } else {
      console.log("Autofilter not enabled")
    }
    await lockColumns();
  })
}

async function findScene(offset){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
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
      const myTarget = sheet.getRangeByIndexes(myIndex + 2, startColumn, 1, 1);
      await excel.sync();
      myTarget.select();
      await excel.sync();
    }
  })
}

async function findSceneNo(sceneNo){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
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
      const myTarget = sheet.getRangeByIndexes(myIndex + 2, startColumn, 1, 1);
      await excel.sync();
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
  const sceneColumn = findColumnIndex("Scene");
  console.log("Scene Colum");
  console.log(sceneColumn);
  const sheet = excel.workbook.worksheets.getActiveWorksheet();
  const endRow = sheet.getUsedRange().getLastRow();
  endRow.load("rowindex");
  await excel.sync();
  range = sheet.getRangeByIndexes(2, sceneColumn, endRow.rowIndex, 1);
  await excel.sync();
  return range;
}

async function getDataRange(excel){
  const sheet = excel.workbook.worksheets.getActiveWorksheet();
  const myLastRow = sheet.getUsedRange().getLastRow();
  const myLastColumn = sheet.getUsedRange().getLastColumn();
  myLastRow.load("rowindex");
  myLastColumn.load("columnindex")
  await excel.sync();
  
  const range = sheet.getRangeByIndexes(1,0, myLastRow.rowIndex, myLastColumn.columnIndex + 1);
  await excel.sync();
  
  return range
}

async function getTargetSceneNumber(){
  const textValue = tag("scene").value;
  const sceneNumber = parseInt(textValue);
  if (sceneNumber != NaN){
    console.log(sceneNumber);
    await findSceneNo(sceneNumber);
  }  else {
    alert("Please enter a number")
  }  
}

async function getSceneMaxAndMin(){
  let result = {};
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    const min = sheet.getRange("minScene");
    await excel.sync();
    min.load("values");
    await excel.sync();
    const max = sheet.getRange("maxScene");
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
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    const studioText = tag("studio-select").value;
    const engineerText = tag("engineer-select").value;
    let dateColumn;
    let studioColumn;
    let engineerColumn;
    if (country == 'UK'){
      dateColumn = findColumnIndex("UK Date Recorded");
      studioColumn = findColumnIndex("UK Studio");
      engineerColumn = findColumnIndex("UK Engineer");
    } else if ( country == 'US'){
      dateColumn = findColumnIndex("US Date Recorded");
      studioColumn = findColumnIndex("US Studio");
      engineerColumn = findColumnIndex("US Engineer");
    } else if ( country == 'Walla'){
      dateColumn = findColumnIndex("Walla Date Recorded");
      studioColumn = findColumnIndex("Walla Studio");
      engineerColumn = findColumnIndex("Walla Engineer");
    }
    
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load("rowIndex");
    await excel.sync();
    const myRow = activeCell.rowIndex;    
    console.log("Row Index");
    console.log(myRow);
    const dateRange = sheet.getRangeByIndexes(myRow, dateColumn, 1, 1);
    const studioRange = sheet.getRangeByIndexes(myRow, studioColumn, 1, 1);
    const engineerRange = sheet.getRangeByIndexes(myRow, engineerColumn, 1, 1);
    await excel.sync();
    await unlock();
    console.log(studioRange);
    dateRange.values = [[dateInFormat()]];
    studioRange.values = [[studioText]];
    engineerRange.values = [[engineerText]];
    await excel.sync();
    await lockColumns();
    engineerRange.select();
    await excel.sync();
    dateRange.select();
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
        result.push(temp);
      }
    }
    console.log(result);
    mySheetColumns = result;
    return result;
  })
}

async function theFormulas(){
  const sceneLineCountColumn = findColumnLetter("Scene Line Count") //B
  const sceneLineNumberRangeColumn = findColumnLetter("Scene Line Number Range"); //C
  const sceneNumberColumn = findColumnLetter("Scene Number"); //D
  const numberColumn = findColumnLetter("Number"); //F
  const UKScriptColumn = findColumnLetter("UK script"); //J
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
  const firstRow = "" + firstDataRow;
  const firstRestRow = "4";
  const lastRow = "" + lastDataRow;
  const columnFormulae = [
    {
      columnName: "Scene Word Count", //A
      formulaFirst: '=""',
      formulaRest: '=IF(' + sceneLineCountColumn + firstRestRow + '<>"",' + sceneWordCountCalcColumn + firstRestRow + ',"")'
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
  	}
  ]
  
  await unlock();
  await Excel.run(async function(excel){ 
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    for (let columnFormula of columnFormulae){
      const columnLetter = findColumnLetter(columnFormula.columnName);
      const myRange = columnLetter + firstRestRow + ":" + columnLetter + lastRow ;
      const myTopRow = columnLetter + firstRow;
      console.log(myRange + "  " + myTopRow);
      const range = sheet.getRange(myRange);
      const topRowRange = sheet.getRange(myTopRow);
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
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
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
  
    const myRow = sheet.getRangeByIndexes(activeCell.rowIndex,0, 1, myLastColumn.columnIndex+1);
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
async function deleteRow(){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
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
  const positionMinusColumn = findColumnLetter("Position -"); //BT
  const startLineColumn = findColumnLetter("Start Line"); //BU
  const positionEndSqaureBracketColumn = findColumnLetter("Position ]"); //BV
  const endLineColumn = findColumnLetter("End Line"); //BW
  const lineWordCountColumn = findColumnLetter("Line Word Count") //BY
  const sceneColumn = findColumnLetter("Scene"); //BZ
  const wordCountToThisLineColumn = findColumnLetter("Word count to this line"); //CB
  
  const columnFormulae = [
    {
      columnName: "Start Line", //BU
      formulaRest: "=IF(" + positionMinusColumn + firstRow + "=0," + startLineColumn + (firstRow - 1) + ",VALUE(MID(" + sceneLineNumberRangeColumn + firstRow + ",2," + positionMinusColumn + firstRow + "-2)))"
    },
    {
      columnName: "End Line", //BW
      formulaRest: "=IF(" + positionEndSqaureBracketColumn + firstRow + "=0," + endLineColumn + (firstRow - 1) + ",VALUE(MID(" + sceneLineNumberRangeColumn + firstRow + "," + positionMinusColumn + firstRow + "+1," + positionEndSqaureBracketColumn + firstRow + "-" + positionMinusColumn + firstRow + "-1)))"
    },
    {
      columnName: "Scene", //BZ
      formulaRest: '=IF(' + sceneNumberColumn + firstRow + '="",' +sceneColumn + (firstRow - 1) + ',VALUE(' + sceneNumberColumn + firstRow + '))'
    },
    {
	    columnName: "Word count to this line", //CB
      formulaRest: "=IF(" + sceneColumn + firstRow + "=" + sceneColumn + (firstRow - 1) + "," + wordCountToThisLineColumn + (firstRow -1) + "+" + lineWordCountColumn + firstRow + "," + lineWordCountColumn + firstRow + ")"
  	}
  ]
  
  await unlock();
  await Excel.run(async function(excel){ 
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    for (let columnFormula of columnFormulae){
      const columnLetter = findColumnLetter(columnFormula.columnName);
      const myRange = columnLetter + firstRow + ":" + columnLetter + (firstRow +1) ;
      console.log("Range to replace: " + myRange);
      const range = sheet.getRange(myRange);
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

async function findDetailsForThisLine(){
  await unlock();
  const totalTakesIndex = findColumnIndex('Total Takes');
  const ukTakesIndex = findColumnIndex('UK No of takes');
  const usTakesIndex = findColumnIndex('US No of takes');
  const wallaTakesIndex = findColumnIndex('Walla No Of takes');
  await Excel.run(async function(excel){ 
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    const activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex')
    await excel.sync();
    const currentRowIndex = activeCell.rowIndex
    let myIndecies = await getAllLinesWithThisNumber(excel, currentRowIndex);
    console.log("myIndecies");
    console.log(myIndecies);

    const totalTakesCell = sheet.getRangeByIndexes(myIndecies[0], totalTakesIndex, 1, 1);
    const ukTakesCell = sheet.getRangeByIndexes(myIndecies[0], ukTakesIndex, 1, 1);
    const usTakesCell = sheet.getRangeByIndexes(myIndecies[0], usTakesIndex, 1, 1);
    const wallaTakesCell = sheet.getRangeByIndexes(myIndecies[0], wallaTakesIndex, 1, 1);

    totalTakesCell.load('values');
    ukTakesCell.load('values');
    usTakesCell.load('values');
    wallaTakesCell.load('values');

    await excel.sync();
    let result = {};
    result.totalTakes = cleanTakes(totalTakesCell.values);
    result.ukTakes = cleanTakes(usTakesCell.values);
    result.usTakes = cleanTakes(usTakesCell.values);
    result.wallaTakes = cleanTakes(wallaTakesCell.values);

    console.log('Result');
    console.log(result);
    return result;

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


}
function cleanTakes(values){
  let temp = parseInt(values);
  if (temp != NaN){
      return temp;
    } else {
      return 0;
    }
}

async function getAllLinesWithThisNumber(excel, currentRowIndex){
  //returns an array of indexes
  const sheet = excel.workbook.worksheets.getActiveWorksheet();
  const numberIndex = findColumnIndex("Number");
  const numberColumn = findColumnLetter("Number");
  let currentNumberCell = sheet.getRangeByIndexes(currentRowIndex, numberIndex, 1, 1)
  currentNumberCell.load('values');
  let numberData = sheet.getRange(numberColumn + firstDataRow + ":" + numberColumn + lastDataRow);
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

async function doTakesAndNumTakes(currentRowIndex, country, doDate, doAdditional, includeMarkUp, includeStudio, includeEngineer){
  const numberColumn = findColumnLetter("Number");
  const numberIndex = findColumnIndex("Number")
  let noOfTakesIndex;
  if (country == "UK"){
    noOfTakesIndex = findColumnIndex("UK No of takes");
    dateRecordedIndex = findColumnIndex("UK Date Recorded");
    markUpIndex = findColumnIndex("UK Broadcast Assistant Markup");
    studioIndex = findColumnIndex("UK Studio");
    engineerIndex = findColumnIndex("UK Engineer")
  }
  await unlock();
  await Excel.run(async function(excel){ 
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    let currentNumberCell = sheet.getRangeByIndexes(currentRowIndex, numberIndex, 1, 1)
    currentNumberCell.load('values')
    let numberData = sheet.getRange(numberColumn + firstDataRow + ":" + numberColumn + lastDataRow);
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
      let numTakesRange = sheet.getRangeByIndexes(firstIndex, noOfTakesIndex, myIndecies.length, 2)
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
          let dateRange = sheet.getRangeByIndexes(rowIndex, dateRecordedIndex, 1, 1);
          let theDate = dateInFormat();
          dateRange.values = theDate;
        }
        if (!includeMarkUp){
          let markUpRange = sheet.getRangeByIndexes(rowIndex, markUpIndex, 1, 1);
          markUpRange.clear("Contents");
        }
        if (!includeStudio){
          console.log('Studio');
          let studioRange = sheet.getRangeByIndexes(rowIndex, studioIndex, 1, 1);
          studioRange.clear("Contents");
        }
        if(!includeEngineer){
          let engineerRange = sheet.getRangeByIndexes(rowIndex, engineerIndex, 1, 1);
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
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    let myRange = sheet.getRange(noOfTakesColumn + firstDataRow + ":" + takeNumberColumn + lastDataRow);
    myRange.load('values')
    await excel.sync();
    console.log(myRange.values)
    console.log(myRange.values.length)
    console.log(myRange.values[0].length)

    //First unhide all
    let hideRange = sheet.getRangeByIndexes(firstDataRow - 1, 0, lastDataRow - 2, 1);
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
            let hideRange = sheet.getRangeByIndexes(i + firstDataRow - 1, 0, 1, 1);
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
            let hideRange = sheet.getRangeByIndexes(i + firstDataRow - 1, 0, 1, 1);
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
  await unlock();
  await Excel.run(async function(excel){ 
    const settingsSheet = excel.workbook.worksheets.getItem(sheetName);
    const range = settingsSheet.getRange(rangeName);
    range.load('values');
    await excel.sync();
    console.log(range.values);
    let allIndex = range.values.findIndex(x => x[0] == 'All');
    console.log(allIndex);
    let unhideColumns = range.values[allIndex][1]
    console.log(unhideColumns);
    const dataSheet = excel.workbook.worksheets.getActiveWorksheet();
    const unhideColumnsRange = dataSheet.getRange(unhideColumns);
    unhideColumnsRange.columnHidden = false;
    await excel.sync();
    if (columnType == 'UK Script'){
      let ukIndex = range.values.findIndex(x => x[0] == 'UK Script');
      console.log(ukIndex);
      let hideUKColumns = range.values[ukIndex][2].split(",")
      console.log(hideUKColumns);
      for (let hide of hideUKColumns){
        let hideUKColumnsRange = dataSheet.getRange(hide);
        hideUKColumnsRange.load('address');
        await excel.sync();
        console.log(hideUKColumnsRange.address);
        hideUKColumnsRange.columnHidden = true;
        await excel.sync();  
      }
    }
    if (columnType == 'US Script'){
      let usIndex = range.values.findIndex(x => x[0] == 'US Script');
      console.log(usIndex);
      let hideUSColumns = range.values[usIndex][2].split(",")
      console.log(hideUSColumns);
      for (let hide of hideUSColumns){
        let hideUSColumnsRange = dataSheet.getRange(hide);
        hideUSColumnsRange.load('address');
        await excel.sync();
        console.log(hideUSColumnsRange.address);
        hideUSColumnsRange.columnHidden = true;
        await excel.sync();  
      }
    }
    if (columnType == 'Walla Script'){
      let wallaIndex = range.values.findIndex(x => x[0] == 'Walla Script');
      console.log(wallaIndex);
      let hideWallaColumns = range.values[wallaIndex][2].split(",")
      console.log(hideWallaColumns);
      for (let hide of hideWallaColumns){
        let hideWallaColumnsRange = dataSheet.getRange(hide);
        hideWallaColumnsRange.load('address');
        await excel.sync();
        console.log(hideWallaColumnsRange.address);
        hideWallaColumnsRange.columnHidden = true;
        await excel.sync();  
      }
    }
  })  
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

