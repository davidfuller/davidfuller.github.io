const lockedOriginalSheetName = 'Locked Original'
const germanProcessingSheetName = 'German Processing'
const originalLineAndTextName = 'loLineAndText'
const originalTextProcessingName = 'gpLineAndText'

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
    processingTextRange.copyFrom(origTextRange, 'values');
    await excel.sync();

  })
}