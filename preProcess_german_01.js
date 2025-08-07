const lockedOriginalSheetName = 'Locked Original'
const germanProcessingSheetName = 'German Processing'
const originalLineAndTextName = 'loLineAndText'
const originalTextProcessingName = 'gpLineAndText'

const ukScriptSheetName = 'Script'
const cueColumnIndex = 5;
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
  let ukData = {}
  ukData.cue = []
  await Excel.run(async function(excel){
    //get the sheets and ranges
    const ukScriptSheet = excel.workbook.worksheets.getItem(ukScriptSheetName);
    let cueRange = ukScriptSheet.getRangeByIndexes(firstRowIndex, cueColumnIndex, lastRowCount, 1);
    cueRange.load('rowIndex, values');
    await excel.sync();
    console.log('rowIndex: ', cueRange.rowIndex);
    let cueValues = cueRange.values.map(x => x[0])
    console.log(cueValues)
    for (let i = 0; i < cueValues.length; i++){
      if (!isNaN(parseInt(cueValues[i]))){
        ukData.cue.push(parseInt(cueValues[i]))
      }
    }
    console.log('ukData: ', ukData);
  })
}