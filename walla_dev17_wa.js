const scriptSheetName = 'Script';
const cueColumnIndex = 5;
const characterColumnIndex = 7;
const ukScriptColumnIndex = 10;
const germanScriptColumnIndex = 12;

const startRow = 1;
const maxRowCount = 50000;


const germanWallaColumns = {
  book: 0,
  cue: 1,
  character: 2,
  ukScript: 3,
  germanScript: 4,
  germanWallaMachineTranslation: 5,
  context: 6,
  numColumns: 7
};

const germanScriptedWallaName = 'German Scripted Walla';

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
            temp.bookNo = await bookNumber();
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
      await appendRow(result);
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
    characterRange.load('values, address');
    let ukScriptRange = scriptSheet.getRangeByIndexes(rowIndex, ukScriptColumnIndex, 1, 1);
    ukScriptRange.load('values, address');
    let germanScriptRange = scriptSheet.getRangeByIndexes(rowIndex, germanScriptColumnIndex, 1, 1);
    germanScriptRange.load('values, address');
    await excel.sync();
    data.character = {value: characterRange.values[0][0], address: characterRange.address};
    data.ukScript = {value: ukScriptRange.values[0][0], address: ukScriptRange.address};
    data.germanScript = {value: germanScriptRange.values[0][0], address: germanScriptRange};    
   })
   return data;
}

async function wallaNextRows(scriptRowIndex){
  let data = [];
  let rowIndex = scriptRowIndex + 1;
  let myCueValue = await cueValue(rowIndex)
  let cueString = myCueValue.value.toString().toLowerCase().trim();
  while (cueString.startsWith('w')){
    let temp = await scriptData(rowIndex);
    temp.rowIndex = rowIndex;
    temp.cue = myCueValue.value.toString().trim();
    temp.address = myCueValue.address;
    data.push(temp);
    rowIndex += 1;
    myCueValue = await cueValue(rowIndex)
    cueString = myCueValue.value.toString().toLowerCase().trim();
  }
  console.log('next walla data', data) ;
  return data;
}

async function cueValue(rowIndex){
  let myCueValue;
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let cueRange = scriptSheet.getRangeByIndexes(rowIndex, cueColumnIndex, 1, 1);
    cueRange.load('values, address');
    await excel.sync()
    myCueValue = {value: cueRange.values[0][0], address: cueRange.address};
  })
  return myCueValue;
}

async function bookNumber(){
  let bookNo = -1;
  let rangeName = 'CK3';
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let bookRange = scriptSheet.getRange(rangeName);
    bookRange.load('values');
    await excel.sync();
    let temp = bookRange.values[0][0]
    let bookValue = parseInt(temp.replace('Book ',''));
    if (!isNaN(bookValue)){
      bookNo = bookValue;
    }
  })
  return bookNo;
}

async function getGermanScriptedWallaUsedRange(){
  let used = {}
  await Excel.run(async function(excel){
    let wallaSheet = excel.workbook.worksheets.getItem(germanScriptedWallaName);
    let wallaRange = wallaSheet.getRangeByIndexes(startRow, germanWallaColumns.bookNo, maxRowCount, germanWallaColumns.numColumns);
    let usedWallaRange = wallaRange.getUsedRangeOrNullObject(true);
    await excel.sync();
    if (!usedWallaRange.isNullObject){
      usedWallaRange.load('address, rowIndex, rowCount');
      await excel.sync();
      console.log('Address', usedWallaRange.address, usedWallaRange.rowIndex, usedWallaRange.rowCount);
      used.address = usedWallaRange.address;
      used.rowIndex = usedWallaRange.rowIndex;
      used.rowCount = usedWallaRange.rowCount;
      used.empty = false;
    } else {
      console.log('Empty')
      used.empty = true;
    }
  })
  return used;
}

async function clearGermanScriptedWalla(){
  let used = await getGermanScriptedWallaUsedRange();
  if (!used.empty){
    await Excel.run(async function(excel){
      let wallaSheet = excel.workbook.worksheets.getItem(germanScriptedWallaName);
      let wallaRange = wallaSheet.getRange(used.address);
      wallaRange.clear('Contents');
    })
  }
}

async function appendRow(result){
  let used = await getGermanScriptedWallaUsedRange();
  let rowIndex = startRow
  if (!used.empty){
    rowIndex = used.rowIndex + used.rowCount
  }
  await Excel.run(async function(excel){
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let wallaSheet = excel.workbook.worksheets.getItem(germanScriptedWallaName);
    let bookRange = wallaSheet.getRangeByIndexes(rowIndex, germanWallaColumns.book, 1, 1);
    bookRange.values = [[result.bookNo]];
    let cueRange = wallaSheet.getRangeByIndexes(rowIndex, germanWallaColumns.cue, 1, 1);
    let sourceCueRange = scriptSheet.getRangeByIndexes(result.rowIndex, cueColumnIndex, 1, 1);
    copyValuesAndFormats(sourceCueRange, cueRange);
    bookRange.copyFrom(sourceCueRange, 'Formats');
    //cueRange.values =[[result.cue]];
    let characterRange = wallaSheet.getRangeByIndexes(rowIndex, germanWallaColumns.character, 1, 1);
    let sourceCharacterRange = scriptSheet.getRange(result.scriptData.character.address);
    copyValuesAndFormats(sourceCharacterRange, characterRange);
    //characterRange.values =[[result.character.value]];
    let ukScriptRange = wallaSheet.getRangeByIndexes(rowIndex, germanWallaColumns.ukScript, 1, 1);
    let sourceUkScriptRange = scriptSheet.getRange(result.scriptData.ukScript.address);
    copyValuesAndFormats(sourceUkScriptRange, ukScriptRange);
    //ukScriptRange.values =[[result.scriptData.ukScript.value]];
    let germanScriptRange = wallaSheet.getRangeByIndexes(rowIndex, germanWallaColumns.germanScript, 1, 1);
    germanScriptRange.values =[[result.scriptData.germanScript.value]];
    await excel.sync();
    for (let i = 0; i < result.wallaNextData.length; i++){
      rowIndex += 1;
      let bookRange = wallaSheet.getRangeByIndexes(rowIndex, germanWallaColumns.book, 1, 1);
      bookRange.values = [[result.bookNo]];
      let cueRange = wallaSheet.getRangeByIndexes(rowIndex, germanWallaColumns.cue, 1, 1);
      cueRange.values =[[result.wallaNextData[i].cue]];
      let characterRange = wallaSheet.getRangeByIndexes(rowIndex, germanWallaColumns.character, 1, 1);
      characterRange.values =[[result.wallaNextData[i].character.value]];
      let ukScriptRange = wallaSheet.getRangeByIndexes(rowIndex, germanWallaColumns.ukScript, 1, 1);
      ukScriptRange.values =[[result.wallaNextData[i].ukScript.value]];
    }
  }) 
}
      
function copyValuesAndFormats(sourceRange, destRange){
  destRange.copyFrom(sourceRange, 'Values');
  destRange.copyFrom(sourceRange, 'Formats');
}