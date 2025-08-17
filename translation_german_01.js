const germanProcessingSheetName = 'German Processing';
const translationCacheSheetName = 'Translation Cache';
const gpTranslationRangeName = 'gpTranslation';
const gpMachineTranslationRangeName = 'gpMachineTranslation'
const gpProcessedRangeName = 'gpProcessed';
const tcTranslationRangeName = 'tcTranslation';
const tcMachineTranslationRangeName = 'tcMachineTranslation'

async function copyValuesToCache(){
  showCopyCacheWait(true)
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let translationRange = gpSheet.getRange(gpTranslationRangeName);
    const tcSheet = excel.workbook.worksheets.getItem(translationCacheSheetName);
    let cacheRange = tcSheet.getRange(tcTranslationRangeName);
    await excel.sync()
    cacheRange.clear('Contents');
    cacheRange.copyFrom(translationRange, 'values');
    await excel.sync()
  })
  showCopyCacheWait(false)
}

async function fixMachineTranslationDisplay(){
  showFixWait(true);
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let translationRange = gpSheet.getRange(gpMachineTranslationRangeName);
    translationRange.copyFrom(translationRange, 'values')
    await excel.sync();
  })
  showFixWait(false);
}
function showFixWait(show){
  let ctl = tag('fix-machine-message')
  if (show){
    ctl.style.display = 'block';
  } else {
    ctl.style.display = 'none';
  }
}

function showCopyCacheWait(show){
  let ctl = tag('copy-cache-message')
  if (show){
    ctl.style.display = 'block';
  } else {
    ctl.style.display = 'none';
  }
}



async function getMachineTranslationFormula(){
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let translationRange = gpSheet.getRange(gpMachineTranslationRangeName);
    translationRange.load("formulas, rowIndex")
    await excel.sync();
    console.log('rowIndex', translationRange.rowIndex, 'formulas', translationRange.formulas);
  })
};

async function applyMachineTranslationFormula(rowIndex){
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let translationRange = gpSheet.getRange(gpMachineTranslationRangeName);
    translationRange.load("rowIndex, columnIndex")
    await excel.sync();
    let formulaRange = gpSheet.getRangeByIndexes(rowIndex, translationRange.columnIndex, 1,1)
    let newFormula = '=IF(G' + (rowIndex+1).toString() + ' <> 0,TRANSLATE(G' + (rowIndex+1).toString() + ',"de","en"),"")'
    console.log('New formula', newFormula);
    formulaRange.clear('Contents');
    await excel.sync();
    formulaRange.formulas = [[newFormula]];
    await excel.sync();
    formulaRange.load('values, address');
    await excel.sync();
    console.log('address', formulaRange.address, 'value', formulaRange.values);
  })
}

async function compareTranslationwithCache(doFormulae){
  let exceptions = [];
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let translationRange = gpSheet.getRange(gpTranslationRangeName);
    translationRange.load('values, rowIndex');
    const tcSheet = excel.workbook.worksheets.getItem(translationCacheSheetName);
    let cacheRange = tcSheet.getRange(tcTranslationRangeName);
    cacheRange.load('values, rowIndex');
    await excel.sync();

    let germanActual = translationRange.values.map(x => x[0]);
    let germanCache = cacheRange.values.map(x => x[0]);
    for (let i = 0; i < germanActual.length; i++){
      if(germanActual[i] != germanCache[i]){
        let temp = {index:i, actual: germanActual[i], cache: germanCache[i], rowIndex: i + translationRange.rowIndex}
        exceptions.push(temp);
      }
    }
  })
  console.log('exceptions', exceptions);
  
  if ((doFormulae) && (exceptions.length > 0)){
    await fillMachineFormula(exceptions[0].rowIndex)
    /**
    for (let i = 0; i < exceptions.length ; i++){
      await applyMachineTranslationFormula(exceptions[i].rowIndex);
    }
    */
  }
}

async function fillMachineFormula(startRowIndex){
  //Fills with formula from startRowIndex to bottom of GermanProcessed.
  let usedCount = await jade_modules.operations.getUsedRowCount(germanProcessingSheetName, gpProcessedRangeName);
  let lastRowIndex = usedCount.rowIndex + usedCount.rowCount - 1;
  let fillRowCount = lastRowIndex - startRowIndex + 1;
  console.log('usedCount', usedCount, 'lastRowIndex', lastRowIndex);
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let machineTranslationRange = gpSheet.getRange(gpMachineTranslationRangeName);
    machineTranslationRange.load('rowIndex, columnIndex');
    await excel.sync();
    
    let fillRange = gpSheet.getRangeByIndexes(startRowIndex, machineTranslationRange.columnIndex, fillRowCount, 1);
    fillRange.clear("Contents");
    await excel.sync();
    await applyMachineTranslationFormula(startRowIndex);
    let topCell = gpSheet.getRangeByIndexes(startRowIndex, machineTranslationRange.columnIndex, 1, 1);
    await excel.sync();
    topCell.autoFill(fillRange, 'FillDefault');
    await excel.sync();
  })
}

async function fillWithFormula(){
  let usedCount = await jade_modules.operations.getUsedRowCount(germanProcessingSheetName, gpProcessedRangeName);
  console.log('usedCount', usedCount);
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let machineTranslationRange = gpSheet.getRange(gpMachineTranslationRangeName);
    machineTranslationRange.load('rowIndex, columnIndex');
    machineTranslationRange.clear("Contents");
    await excel.sync();
    await applyMachineTranslationFormula(machineTranslationRange.rowIndex);
    let topCell = gpSheet.getRangeByIndexes(machineTranslationRange.rowIndex, machineTranslationRange.columnIndex, 1, 1);
    let formulaRange = gpSheet.getRangeByIndexes(machineTranslationRange.rowIndex, machineTranslationRange.columnIndex, usedCount.rowCount, 1);
    await excel.sync();
    topCell.autoFill(formulaRange, 'FillDefault');
    await excel.sync();
  })
}

async function machineTranslationValues(){
  let values;
  let rowIndex;
  let usedCount = await jade_modules.operations.getUsedRowCount(germanProcessingSheetName, gpProcessedRangeName);
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let machineTranslationRange = gpSheet.getRange(gpMachineTranslationRangeName);
    machineTranslationRange.load('rowIndex, columnIndex');
    await excel.sync();
    rowIndex = machineTranslationRange.rowIndex
    let valueRange = gpSheet.getRangeByIndexes(rowIndex, machineTranslationRange.columnIndex, usedCount.rowCount, 1);
    valueRange.load('values');
    await excel.sync()
    values = valueRange.values.map(x => x[0])
    console.log('Values', values);
  })
  return {values: values, rowIndex: rowIndex};
}

async function issueCells(doFormulae){
  let machineValues = await machineTranslationValues();
  console.log('machineValues', machineValues);
  let theIssues = []
  const issues = ['#CONNECT!', '#CALC!', '#BUSY']
  for (let i = 0; i < machineValues.values.length; i++){
    for (let words = 0; words < issues.length; words++){
      if(machineValues.values[i].includes(issues[words])){
        theIssues.push({index: i, value: machineValues.values[i], rowIndex: i + machineValues.rowIndex});
      }
    }
  }
  console.log('issues', theIssues);
  showIssuesMessage(theIssues.length.toString() + ' calculation issues');
  if ((doFormulae) && (theIssues.length > 0)){
    await fillMachineFormula(theIssues[0].rowIndex);
  }
  /**
  if (doFormulae){
    for(let i = 0; i < theIssues.length; i++){
      await applyMachineTranslationFormula(theIssues[i].rowIndex);
    }
  }
  **/
}

function showIssuesMessage(message){
  let ctrlMessage = tag('issues-message');
  ctrlMessage.innerText = message;
  
}



