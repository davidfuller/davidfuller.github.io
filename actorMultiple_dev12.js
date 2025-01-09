const forActorName = 'For Actors'
const multiActorTableName = 'faMultiActorTable'
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
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(forActorName);
    const range = sheet.getRange(multiActorTableName);
    range.load ('values, rowIndex');
    await excel.sync();
    for (let i = 0; i < range.values.length; i++){
      if (range.values[i][characterColumn] == ''){
        addRowNo = i;
        break;
      }
    }
    console.log('addRowNo', addRowNo)


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