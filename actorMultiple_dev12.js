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
    let resultRange = sheet.getRangeByIndexes(addRowNo + range.rowIndex, characterColumn + range.columnIndex, 1, 3);
    let resultArray = [[actorDetails.character, actorDetails.type, actorDetails.allUs]];
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