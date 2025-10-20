const scriptSheetName = 'Script';
const cueColumnIndex = 5;
const startRow = 1;
const maxRowCount = 50000;

async function getUsedCueRange(){
  let cueRangeAddress = '';
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem('Script');
    let fullCueRange = scriptSheet.getRangeByIndexes(startRow, cueColumnIndex, maxRowCount, 1);
    let cueRange = fullCueRange.getUsedRange();
    cueRange.load('address');
    await excel.sync();
    cueRangeAddress = cueRange.address
 });
 console.log('Cue Range:', cueRangeAddress);
 return cueRangeAddress;
}

async function minMaxCueValues(){
  let cueRangeAddress = await getUsedCueRange();
  let minCueValue = 100000;
  let maxCueValue = 0;
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem('Script');
    let cueRange = scriptSheet.getRange(cueRangeAddress);
    cueRange.load('values')
    await excel.sync();
    let theValues = cueRange.values.map(x => x[0]);
    for (let i = 0; i < theValues.length; i++){
      let testValue = parseInt(theValues[i])
      if (!isNaN(testValue)){
        if (testValue > maxCueValue){
          maxCueValue = testValue;
        }
        if (testValue < minCueValue){
          minCueValue = testValue;
        }
      }
    }
  })
  let result = {min: minCueValue, max: maxCueValue}
  console.log('Result', result)
  return result
}