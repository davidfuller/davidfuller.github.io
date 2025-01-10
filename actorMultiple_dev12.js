const forActorName = 'For Actors';
const multiActorTableName = 'faMultiActorTable';
const typeChoiceName = 'faChoice';
const textValueName = 'faTextSearch';
const listValueName = 'faCharacterChoice';
const allUsName = 'faSelect';
const multiActorColumns = [
  {name: 'No', column: 0},
  {name: 'Character', column: 1},
  {name: 'Type', column: 2},
  {name: 'All/US', column: 3},
  {name: 'Scene', column: 4},
]

async function auto_exec(){
  console.log('Actor Multiple');
}

async function addScript(){
  let characterColumn = getColumnNumber('Character');
  let sceneColumn = getColumnNumber('Scene');
  let columnCount = sceneColumn - characterColumn + 1; 
  let addRowNo = -1
  let actorDetails = await getActorDetails();
  let scenes = await jade_modules.scheduling.getSceneNumberActor();
  console.log('scenes', scenes);
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(forActorName);
    const range = sheet.getRange(multiActorTableName);
    range.load ('values, rowIndex, columnIndex');
    await excel.sync();
    for (let i = 0; i < range.values.length; i++){
      if (range.values[i][characterColumn] == ''){
        addRowNo = i;
        break;
      }
    }
    console.log('addRowNo', addRowNo)
    let resultRange = sheet.getRangeByIndexes(addRowNo + range.rowIndex, characterColumn + range.columnIndex, 1, columnCount);
    let resultArray = [[actorDetails.character, actorDetails.type, actorDetails.allUs, scenes.scenes.join(', ')]];
    resultRange.values = resultArray;
  })
}

function getColumnNumber(theName){
  let result = -1;
  for (let i = 0; i < multiActorColumns.length; i++){
    if (multiActorColumns[i].name == theName){
      result = multiActorColumns[i].column;
      break;
    }
  }
  return result
}

async function getActorDetails(){
  let details = {};
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(forActorName);
    const typeRange = sheet.getRange(typeChoiceName);
    const allUsRange = sheet.getRange(allUsName);
    typeRange.load('values');
    allUsRange.load('values');
    await excel.sync();
    details.type = typeRange.values[0][0];
    details.allUs = allUsRange.values[0][0];
    let characterRange;
    if (details.type == 'Text Search'){
      characterRange = sheet.getRange(textValueName);
    } else {
      characterRange = sheet.getRange(listValueName);
    }
    characterRange.load('values');
    await excel.sync();
    details.character = characterRange.values[0][0];
  })
  console.log('details', details);
  return details;
}

async function removeScript(){
  let tableRows = await tableRowsToClear();
  await clearRows(tableRows);
}

async function tableRowsToClear(){
  let details = [];
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(forActorName);
    const tableRange = sheet.getRange(multiActorTableName);
    tableRange.load('rowIndex, columnIndex, rowCount, columnCount');
    let selectedRanges = excel.workbook.getSelectedRanges();
    selectedRanges.load('address');
    await excel.sync();
    let myAddresses = selectedRanges.address.split(',');
    console.log('myAddresses', myAddresses)
    let myRanges = []
    for (let myAddress of myAddresses){
      console.log('address', myAddress)
      myRanges.push(sheet.getRange(myAddress));
    }
    for (let myRange of myRanges){
      myRange.load('rowIndex, columnIndex, rowCount');
    }
    await excel.sync();
    let testRanges = [];
    for (let myRange of myRanges){
      for (let i = 0; i < myRange.rowCount; i++){
        testRanges.push(sheet.getRangeByIndexes(myRange.rowIndex + i, myRange.columnIndex, 1, 1));
      }
    }

    for (let testRange of testRanges){
      testRange.load('rowIndex, columnIndex');
    }
    await excel.sync();
    
    let validRanges = []
    for (let myRange of testRanges){
      console.log('row', myRange.rowIndex, 'column', myRange.columnIndex);
      if ((myRange.rowIndex >= tableRange.rowIndex) && (myRange.rowIndex <= tableRange.rowIndex + tableRange.rowCount -1)){
        if ((myRange.columnIndex >= tableRange.columnIndex) && (myRange.columnIndex <= tableRange.columnIndex + tableRange.columnCount -1)){
          validRanges.push(myRange);
        }
      }
    }
    for (let validRange of validRanges){
      details.push(validRange.rowIndex - tableRange.rowIndex);
    }
  }) 
  console.log('details', details);
  return details;
}

async function clearRows(theRows){
  let characterColumn = getColumnNumber('Character');
  let sceneColumn = getColumnNumber('Scene');
  let columnCount = sceneColumn - characterColumn + 1; 
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(forActorName);
    const tableRange = sheet.getRange(multiActorTableName);
    tableRange.load('rowIndex, columnIndex, rowCount, columnCount');
    await excel.sync();
    let deleteRanges = [];
    for (let theRow of theRows){
      deleteRanges.push(sheet.getRangeByIndexes(theRow + tableRange.rowIndex, characterColumn + tableRange.columnIndex, 1, columnCount));
    }
    for (let deleteRange of deleteRanges){
      deleteRange.clear('Contents');
    }
  })
  await tidyTable();
}

async function tidyTable(){
  let characterColumn = getColumnNumber('Character');
  let sceneColumn = getColumnNumber('Scene');
  let columnCount = sceneColumn - characterColumn + 1; 
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(forActorName);
    const tableRange = sheet.getRange(multiActorTableName);
    tableRange.load('values, rowIndex, columnIndex, rowCount, columnCount');
    await excel.sync();
    let empty = [];
    for (let i = 0; i < tableRange.values.length; i++){
      if (tableRange.values[i][characterColumn] == ''){
        empty.push(i);
      }
    }
    console.log('empty', empty)
    for (let i = 0; i < tableRange.values.length; i++){
      if (empty.includes(i)){
        (console.log(i, 'is Empty'))
        for (j = i + 1; j <tableRange.values.length; j++){
          if (!empty.includes(j)){
            (console.log('j', j, 'is NOT Empty'))
            let newRange = sheet.getRangeByIndexes(i + tableRange.rowIndex, characterColumn + tableRange.columnIndex, 1, columnCount);
            let currentRange = sheet.getRangeByIndexes(j + tableRange.rowIndex, characterColumn + tableRange.columnIndex, 1, columnCount)
            newRange.copyFrom(currentRange, "values");
            await excel.sync();
            currentRange.clear("Contents");
            await excel.sync();
            break;
          }
        }
      }
    }
  })
}


