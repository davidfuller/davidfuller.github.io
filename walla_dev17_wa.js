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

async function sourceSheets(){
  sourceSheets = []
  await Excel.run(async function(excel){
    const worksheets = excel.workbook.worksheets;
    worksheets.load('items')
    await excel.sync();
    for (let i = 0; i < worksheets.items.length; i++){
      console.log(worksheets.items[i].name);
      if (worksheets.items[i].name.startsWith('Table')){
        sourceSheets.push(worksheets.items[i].name)
      }
    }
  })
  console.log('sourceSheets', sourceSheets)
  return sourceSheets
}
async function findCues(){
  let results = [];
  let minMax = await minMaxCueValues();
  let sourceSheetNames = await sourceSheets();
  for (let i = 0; i < sourceSheetNames.length; i++){
    await Excel.run(async function(excel){
      const thisSheet = excel.workbook.worksheets.getItem(sourceSheetNames[i]);
      let firstColumnRange = thisSheet.getRangeByIndexes(1, 1, 100, 1);
      firstColumnRange.load('values');
      await excel.sync();
      let theValues = firstColumnRange.values.map(x => x[0])
      for (let j = 0; j < theValues.length; j++){
        theNumber = parseInt(theValues[j]);
        if (!isNaN(theNumber)){
          if ((theNumber >= minMax.min) && (theNumber <= minMax.max)){
            let temp = {};
            temp.cue = theNumber;
            temp.sheetName = sourceSheetNames[i];
            results,push(temp)
          }  
        }
      }
    })
  }
  console.log('results', results);
  return results;
}