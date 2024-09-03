function auto_exec(){
  console.log('Operations loaded');
  console.log(jade_modules)
}
const columnsToLock = "A:Y"
const myColumns = 
  [
    {
      columnName: "Scene",
      columnNo: 77
    },
    {
      columnName: "Line",
      columnNo: 78
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
    
    const myIndex = range.values.findIndex(a => a[0] == (currentValue + offset));

    console.log("Found Index");
    console.log(myIndex);
    
    if (myIndex == -1){
      if (offset == 1){
        alert('This is the final scene')
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
      alert('This is the final scene')
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
  const sceneColumn = myColumns.find(x => x.columnName == "Scene").columnNo;
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
