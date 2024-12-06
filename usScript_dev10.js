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
  const usDetails = await getUsScriptDetails(rowIndexes);
  const ukDetails = await jade_modules.operations.findUsScriptCues(usDetails);
  const copyDetails = compareDetails(usDetails, ukDetails);
  if (copyDetails.length == rowIndexes.length){
    await jade_modules.operations.doTheCopy(copyDetails);
  } else {
    console.log('Incorrect number of copies', rowIndexes, copyDetails);
  }
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
        rowIndex: rowIndexes[i],
        cue: cueRange[i].values[0][0],
        character: characterRange[i].values[0][0],
        ukScript: ukScriptRange[i].values[0][0],
        usCue: usCueRange[i].values[0][0],
        usScript: usScriptRange[i].values[0][0]
      }
    }
  })
  console.log('details', details);
  return details
}

function compareDetails(usDetails, ukDetails){
  let copyDetails = []
  let compare = []
  for (let i = 0; i < usDetails.length; i++){
    let index = ukDetails.findIndex(x => x.cue == usDetails[i].cue)
    compare[i] = {}
    compare[i].character = (ukDetails[index].character.trim().toLowerCase() === usDetails[i].character.trim().toLowerCase());
    compare[i].ukScript = (ukDetails[index].ukScript.trim().toLowerCase() === usDetails[i].ukScript.trim().toLowerCase());
    console.log('i', i, 'usDetails', usDetails[i], 'index', index, 'compare', compare[i]);
    if (compare[i].character && compare[i].ukSccript){
      let details = {
        ukRowIndex: ukDetails.rowIndex,
        usRowIndex: usDetails.rowIndex,
        usCueColumnIndex: usScriptColumns.usCue,
        usScriptColumnIndex: usScriptColumns.usScript
      }
      copyDetails.push(details);
    }
  }
  return copyDetails;
}