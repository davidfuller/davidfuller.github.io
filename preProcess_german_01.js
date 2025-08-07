const lockedOriginalSheetName = 'Locked Original'
const germanProcessingSheetName = 'German Processing'
const originalLineAndTextName = 'loLineAndText'
const originalTextProcessingName = 'gpLineAndText'

const ukScriptSheetName = 'Script'
const cueColumnIndex = 5;
const numberColumnIndex = 6;
const characterColumnIndex = 7;
const ukScriptColumnIndex = 10;
const firstRowIndex = 3;
const lastRowCount = 10000;

async function doTheCopy(){
  await Excel.run(async function(excel){
    //get the sheets and ranges
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let processingTextRange = gpSheet.getRange(originalTextProcessingName);
    const origSheet = excel.workbook.worksheets.getItem(lockedOriginalSheetName);
    let origTextRange = origSheet.getRange(originalLineAndTextName);
    await excel.sync();

    //clear the processing range
    processingTextRange.clear('Contents')
    await excel.sync();
    //copy in the values
    processingTextRange.copyFrom(origTextRange, 'values');
    await excel.sync();
  })
}

async function getUKScript(){
  let lastRowIndex = await getUKData();
}

async function getUKData(){
  let ukData = {};
  ukData.cue = [];
  ukData.index = [];
  ukData.number = [];
  ukData.character = [];
  ukData.ukScript = [];
  await Excel.run(async function(excel){
    //get the sheets and ranges
    const ukScriptSheet = excel.workbook.worksheets.getItem(ukScriptSheetName);
    let cueRange = ukScriptSheet.getRangeByIndexes(firstRowIndex, cueColumnIndex, lastRowCount, 1);
    let numberRange = ukScriptSheet.getRangeByIndexes(firstRowIndex, numberColumnIndex, lastRowCount, 1);
    let characterRange = ukScriptSheet.getRangeByIndexes(firstRowIndex, characterColumnIndex, lastRowCount, 1);
    let ukScriptRange = ukScriptSheet.getRangeByIndexes(firstRowIndex, ukScriptColumnIndex, lastRowCount, 1);
    cueRange.load('rowIndex, values');
    numberRange.load('values');
    characterRange.load('values');
    ukScriptRange.load('values')
    await excel.sync();
    console.log('rowIndex: ', cueRange.rowIndex);
    let cueValues = cueRange.values.map(x => x[0]);
    let numberValues = numberRange.values.map(x => x[0]);
    let characterValues = characterRange.values.map(x => x[0]);
    let ukScriptValues = ukScriptRange.values.map(x => x[0]);
    console.log(cueValues)
    for (let i = 0; i < cueValues.length; i++){
      if (!isNaN(parseInt(cueValues[i]))){
        ukData.cue.push(parseInt(cueValues[i]))
        ukData.index.push(i);
        ukData.number.push(numberValues[i]);
        ukData.character.push(characterValues[i]);
        ukData.ukScript.push(ukScriptValues[i]);
      }
    }
    console.log('ukData: ', ukData);
  })
}