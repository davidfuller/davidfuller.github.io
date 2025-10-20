const scriptSheetName = 'Script';
const cueColumnIndex = 5;
const startRow = 1;
const maxRowCount = 50000;

async function getUsedCueRange(){
  let cueRangeAddress = '';
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem('Script');
    let fullCueRange = scriptSheet.getRangeByIndex(startRow, cueColumnIndex, maxRowCount, 1);
    let cueRange = fullCueRange.getUsedRange();
    cueRange.load('address');
    await excel.sync();
    cueRangeAddress = cueRange.address
 });
 console.log('Cue Range:', cueRangeAddress);
 return cueRangeAddress;
}