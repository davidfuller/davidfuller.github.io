const wallaSheetName = 'Walla Import';
const sourceTextRangeName = 'wiSource';

function auto_exec(){
  alert("I'm here")
}

async function parseSource(){
  await Excel.run(async (excel) => {
    let wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
    let sourceRange = wallaSheet.getRange(sourceTextRangeName);
    sourceRange.load('values')
    await excel.sync();
    console.log('sourceRange', sourceRange);
  })
}