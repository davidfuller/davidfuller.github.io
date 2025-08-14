const germanProcessingSheetName = 'German Processing';
const translationCacheSheetName = 'Translation Cache';
const gpTranslationRangeName = 'gpTranslation';
const gpMachineTranslationRangeName = 'gpMachineTranslation'
const tcTranslationRangeName = 'tcTranslation';

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
};

async function getMachineTranslationFormula(){
  await Excel.run(async function(excel) {
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let translationRange = gpSheet.getRange(gpMachineTranslationRangeName);
    translationRange.load("formulas, rowIndex")
    await excel.sync();
    console.log('rowIndex', translationRange.rowInde, 'formulas', translationRange.formulas);
  })
};



