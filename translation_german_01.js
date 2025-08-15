const germanProcessingSheetName = 'German Processing';
const translationCacheSheetName = 'Translation Cache';
const gpTranslationRangeName = 'gpTranslation';
const gpMachineTranslationRangeName = 'gpMachineTranslation'
const gpProcessedRangeName = 'gpProcessed';
const tcTranslationRangeName = 'tcTranslation';
const tcMachineTranslationRangeName = 'tcMachineTranslation'

async function copyValuesToCache(){
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
}

async function fixMachineTranslationDisplay(){
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let translationRange = gpSheet.getRange(gpMachineTranslationRangeName);
    translationRange.copyFrom(translationRange, 'values')
    await excel.sync();
  })
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
    formulaRange.formulas = [[newFormula]];
    await excel.sync();
    formulaRange.load('values, address');
    await excel.sync();
    console.log('address', formulaRange.address, 'value', formulaRange.values);
  })
}

async function compareTranslationwithCache(){
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
  for (let i = 0; i < exceptions.length ; i++){
    await applyMachineTranslationFormula(exceptions[i].rowIndex);
  }
  //exceptions.length
  //#CONNECT!
  //3832
}

async function fillWithFormula(){
  let usedCount = jade_modules.operations.getRowIndexLast(germanProcessingSheetName, gpProcessedRangeName);
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let machineTranslationRange = gpSheet.getRange(gpMachineTranslationRangeName);
    machineTranslationRange.load('rowIndex, columnIndex');
    machineTranslationRange.clear("Contents");
    await excel.sync();
  })
}



