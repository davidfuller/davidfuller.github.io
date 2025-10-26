const scriptSheetName = 'Script';
const cueColumnIndex = 5;
const characterColumnIndex = 7;
const ukScriptColumnIndex = 10;
const germanScriptColumnIndex = 12;

const startRow = 1;
const maxRowCount = 50000;

async function getUsedCueRange(){
  let cueRangeAddress = '';
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
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
  let sourceSheetNames = []
  await Excel.run(async function(excel){
    const worksheets = excel.workbook.worksheets;
    worksheets.load('items')
    await excel.sync();
    for (let i = 0; i < worksheets.items.length; i++){
      console.log(worksheets.items[i].name);
      if (worksheets.items[i].name.startsWith('Table')){
        sourceSheetNames.push(worksheets.items[i].name)
      }
    }
  })
  console.log('sourceSheetNames', sourceSheetNames)
  return sourceSheetNames
}
async function findCues(){
  let results = [];
  let minMax = await minMaxCueValues();
  let sourceSheetNames = await sourceSheets();
  let contextText = '';
  const characterColumnIndex = 1;
  const scriptColumnIndex = 2;
  for (let i = 0; i < sourceSheetNames.length; i++){
    await Excel.run(async function(excel){
      const thisSheet = excel.workbook.worksheets.getItem(sourceSheetNames[i]);
      let firstColumnRange = thisSheet.getRangeByIndexes(0, 0, 100, 1);
      firstColumnRange.load('values, rowIndex');
      await excel.sync();
      let theValues = firstColumnRange.values.map(x => x[0])
      console.log(sourceSheetNames[i], 'theValues', theValues);
      for (let j = 0; j < theValues.length; j++){
        let tempContext = extractContext(theValues[j]);
        if (tempContext != ''){
          contextText = tempContext;
        }
        theNumber = parseInt(theValues[j]);
        if (!isNaN(theNumber)){
          if ((theNumber >= minMax.min) && (theNumber <= minMax.max)){
            let temp = {};
            temp.cue = theNumber;
            temp.sheetName = sourceSheetNames[i];
            temp.context = contextText;
            temp.rowIndex = firstColumnRange.rowIndex + j;
            let characterRange = thisSheet.getRangeByIndexes(temp.rowIndex, characterColumnIndex, 1, 1);
            let scriptRange = thisSheet.getRangeByIndexes(temp.rowIndex, scriptColumnIndex, 1, 1);
            characterRange.load('values');
            scriptRange.load('values');
            await excel.sync();
            temp.character = characterRange.values[0][0];
            temp.script = scriptRange.values[0][0];
            results.push(temp)
            contextText = '';
          }  
        }
      }
    })
  }
  results.sort((a, b) => a.cue - b.cue);
  console.log('results', results);
  return results;
}

function extractContext(text){
  console.log('text', text)
  let position = text.toString().toLowerCase().indexOf('context');
  let context = '';
  if (position != -1){
    context = text.substring(position + 8).trim();
  }
  return context;
}

async function findCueIndex(cue){
  const cueRangeAddress = await getUsedCueRange();
  let cueRowIndex = -1;
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let cueRange = scriptSheet.getRange(cueRangeAddress);
    cueRange.load('values, rowIndex');
    await excel.sync();
    let cueValues = cueRange.values.map(x => x[0]);
    let myIndex = cueValues.indexOf(cue);
    if (myIndex != -1){
      cueRowIndex = myIndex + cueRange.rowIndex;
    }
  })
  return cueRowIndex;
}

async function gatherData(){
  let results = await findCues();
  for (let result of results){
    result.rowIndex = await findCueIndex(result.cue);
    console.log('cue:', result.cue, 'rowIndex', result.rowIndex);
    if (result.rowIndex != -1){
      result.scriptData = await scriptData(result.rowIndex);
      console.log('scriptData', result.scriptData);
      result.wallaNextData = await wallaNextRows(result.rowIndex);
    }
  }
  console.log('results', results);
  return results;
}

async function scriptData(rowIndex){
  let data = {};
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let characterRange = scriptSheet.getRangeByIndexes(rowIndex, characterColumnIndex, 1, 1);
    characterRange.load('values');
    let ukScriptRange = scriptSheet.getRangeByIndexes(rowIndex, ukScriptColumnIndex, 1, 1);
    ukScriptRange.load('values');
    let germanScriptRange = scriptSheet.getRangeByIndexes(rowIndex, germanScriptColumnIndex, 1, 1);
    germanScriptRange.load('values');
    await excel.sync();
    data.character = characterRange.values[0][0];
    data.ukScript = ukScriptRange.values[0][0];
    data.germanScript = germanScriptRange.values[0][0];    
   })
   return data;
}

async function wallaNextRows(scriptRowIndex){
  let data = [];
  let rowIndex = scriptRowIndex + 1;
  let myCueValue = await cueValue(rowIndex)
  let cueString = myCueValue.toString().toLowerCase().trim();
  while (cueString.startsWith('w')){
    let temp = await scriptData(rowIndex);
    temp.rowIndex = rowIndex;
    data.push(temp);
    rowIndex += 1;
    myCueValue = await cueValue(rowIndex)
    cueString = myCueValue.toString().toLowerCase().trim();
  }
  console.log('next walla data', data) ;
  return data;
}

async function cueValue(rowIndex){
  let cueValue;
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let cueRange = scriptSheet.getRangeByIndexes(rowIndex, cueColumnIndex, 1, 1);
    cueRange.load('values');
    await excel.sync()
    cueValue = cueRange.values[0][0];
  })
  return cueValue;
}
  