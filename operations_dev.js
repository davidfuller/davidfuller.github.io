async function auto_exec(){
  console.log('Operations loaded');
  console.log(jade_modules)
}
let mySheetColumns;

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
  const columnNameFirst = "Position -";
  const columnNameLast = "Scene word count calc";
  const myFormulas = ['=IF(C4="",0,FIND("-",C4))' ];
  const columnFormulae = [
    {
      columnName: "Position -",
      formulaFirst: '=IF(C3="",0,FIND("-",C3))',
      formulaRest: '=IF(C4="",0,FIND("-",C4))'
    },
    {
      columnName: "Start Line",
      formulaFirst: 0,
      formulaRest: "=IF(BT4=0,BU3,VALUE(MID(C4,2,BT4-2)))"
    },
    {
      columnName: "Position ]",
      formulaFirst: '=IF(C3="",0,FIND("]",C3))',
      formulaRest: '=IF(C4="",0,FIND("]",C4))'
    },
    {
      columnName: "End Line",
      formulaFirst: 0,
      formulaRest: "=IF(BV4=0,BW3,VALUE(MID(C4,BT4+1,BV4-BT4-1)))"
    }
    
    
  ]
  const firstRow = "3";
  const firstRestRow = "4";
  const lastRow = "9999";
  await Excel.run(async function(excel){ 
    const sheet = excel.workbook.worksheets.getActiveWorksheet();
    for (let columnFormula of columnFormulae){
      const columnLetter = findColumnLetter(columnFormula.columnName);
      const myRange = columnLetter + firstRestRow + ":" + columnLetter + lastRow ;
      console.log(myRange);
      const range = sheet.getRange(myRange);
      range.formulas = columnFormula.formulaRest;
      await excel.sync();
      console.log(range.formulas);
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
