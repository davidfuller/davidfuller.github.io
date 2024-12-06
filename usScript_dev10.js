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

async function usScriptAdd(){
  const rowIndexes = await getUsCueIndexes();
  const details = await getUsScriptDetails(rowIndexes);
  await jade_modules.operations.findUsScriptCues(details)
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
        let index = i + usCueRange.rowIndex;
        if (index > 1) {
          rowIndexes.push(i + usCueRange.rowIndex);
        }
      }
    }
  })
  console.log('rowIndexes', rowIndexes);
  return rowIndexes;
}

async function getUsScriptDetails(rowIndexes){
  let details = [];
  await Excel.run(async function(excel){
    const usSheet = excel.workbook.worksheets.getItem(usScriptName);
    let cueRange = [];
    let characterRange = [];
    let ukScriptRange = [];
    let usCueRange = [];
    let usScriptRange = [];
    for (let i = 0; i < rowIndexes.length; i++){
      cueRange[i] = usSheet.getRangeByIndexes(rowIndexes[i], usScriptColumns.cue, 1, 1);
      cueRange[i].load('values');
      characterRange[i] = usSheet.getRangeByIndexes(rowIndexes[i], usScriptColumns.character, 1, 1);
      characterRange[i].load('values');
      ukScriptRange[i] = usSheet.getRangeByIndexes(rowIndexes[i], usScriptColumns.ukScript, 1, 1);
      ukScriptRange[i].load('values');
      usCueRange[i] = usSheet.getRangeByIndexes(rowIndexes[i], usScriptColumns.usCue, 1, 1);
      usCueRange[i].load('values');
      usScriptRange[i] = usSheet.getRangeByIndexes(rowIndexes[i], usScriptColumns.usScript, 1, 1);
      usScriptRange[i].load('values');
    }
    await excel.sync();
    for (let i = 0; i < cueRange.length; i++){
      //console.log('i', i, 'cue', cueRange[i].values, 'character', characterRange[i].values, 'ukScript', ukScriptRange[i].values, 'usCue', usCueRange[i].values, 'usScript', usScriptRange[i].values);
      details[i] = {
        cue: cueRange[i].values[0][0],
        character: characterRange.values[0][0],
        ukScript: ukScriptRange.values[0][0],
        usCue: usCueRange.values[0][0],
        usScript: usScriptRange.value[0][0]
      }
    }
  })
  console.log('details', details);
  return details
}
