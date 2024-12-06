function auto_exec(){
}
const usScriptName = 'US Script'
const usScriptColumns = {
  cue: 0, //A
  number: 1, //B
  character: 2, //C
  stageDirections: 3, //D
  ukScript: 4, //E
  ukScriptEdit: 5 ,//F
  usCue: 6, //G
  usScript: 7 //H
}

async function getUsCueIndexes(){
  let rowIndexes = [];
  await Excel.run(async function(excel){
    const usSheet = excel.workbook.worksheets.getItem(usScriptName);
    const usedRange = usSheet.getUsedRange();
    usedRange.load('rowIndex, rowCount');
    await excel.sync();
    const usCueRange = usSheet.getRangeByIndexes(usedRange.rowIndex, usScriptColumns.usCue, usedRange.rowCount, 1);
    usCueRange.load('values, rowIndex');
    await excel.sync();
    for (let i = 0; i < usCueRange.values.length; i++){
      if (usCueRange.values[i][0].trim() != ''){
        rowIndexes.push(i + usCueRange.rowIndex);
      }
    }
  })
  console.log('rowIndexes', rowIndexes);
  return rowIndexes;
}
