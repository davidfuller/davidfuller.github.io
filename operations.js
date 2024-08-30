function auto_exec(){
  console.log('Operations loaded');
  console.log(jade_modules)
}
const columnsToLock = "A:Y"
const myColumns = 
  [
    {
      columnName: "Scene",
      columnNo: 83
    },
    {
      columnName: "Line",
      columnNo: 84
    }
  ];

async function lockColumns(excel){
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
}

async function unlock(excel){
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
}

async function applyFilter(excel){
  /*Jade.listing:{"name":"Apply filter","description":"Applies empty filter to sheet"}*/
  await unlock(excel);
  const sheet = excel.workbook.worksheets.getActiveWorksheet();
  const myRange = await getDataRange(excel);
  sheet.autoFilter.apply(myRange, 0, { criterion1: "*", filterOn: Excel.FilterOn.custom});
  sheet.autoFilter.clearCriteria();
  await excel.sync();
  
  await lockColumns(excel);
}

async function removeFilter(excel){
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
  
  await lockColumns(excel);  
}

async function findNextScene(excel){
  
  const sceneColumn = myColumns.find(x => x.columnName == "Scene").columnNo;
  const sheet = excel.workbook.worksheets.getActiveWorksheet();
  const activeCell = excel.workbook.getActiveCell();
  activeCell.load("rowIndex");
  activeCell.load(("columnIndex"))
  const endRow = sheet.getUsedRange().getLastRow();
  endRow.load("rowindex");
  await excel.sync()
  const startRow = activeCell.rowIndex;
  const startColumn = activeCell.columnIndex
  const range = sheet.getRangeByIndexes(startRow, sceneColumn, endRow.rowIndex - startRow, 1);
  await excel.sync();
  
  range.load("values");
  await excel.sync();
  
  const currentValue = range.values[0][0];
  const myIndex = range.values.findIndex(a => a[0] == (currentValue + 1));
  console.log(myIndex + startRow);
  console.log(startColumn);
  if (myIndex == -1){
    alert('This is the final scene')
  } else {
    const myTarget = sheet.getRangeByIndexes(myIndex + startRow, startColumn, 1, 1);
    await excel.sync();
    myTarget.select();
    await excel.sync();
  }
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
