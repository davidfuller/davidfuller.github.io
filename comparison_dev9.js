const pdfComparisonSheetName = 'PDF Comparison';
const sourceColumnIndex = 10;
const chaptersColumnIndex = 14;
const startRowIndex = 10;

async function getRowColumnDetails(){
  let details = {};
  await Excel.run(async function(excel){ 
    const pdfSheet = excel.workbook.worksheets.getItem(pdfComparisonSheetName);
    const usedRange = pdfSheet.getUsedRange();
    usedRange.load('rowIndex, rowCount, columnIndex, columnCount');
    await excel.sync();
    details = {
      rowIndex: usedRange.rowIndex,
      rowCount: usedRange.rowCount,
      columnIndex: usedRange.columnIndex,
      columnCount: usedRange.columnCount
    }
  })
  return details;
}

async function createChapters(){
  const details = await getRowColumnDetails();
  let chapters = [];
  let index = -1;
  textSoFar = '';
  await Excel.run(async function(excel){ 
    const pdfSheet = excel.workbook.worksheets.getItem(pdfComparisonSheetName);
    const rowCount = details.rowCount - details.rowIndex + 1 - startRowIndex;
    const sourceRange = pdfSheet.getRangeByIndexes(startRowIndex, sourceColumnIndex, rowCount, 1);
    sourceRange.load('rowIndex, values');
    await excel.sync();
    //console.log('sourceRange', sourceRange.values)
    sourceValues = sourceRange.values.map(x => x[0]);

    for (let i = 0; i < sourceValues.length; i++){
      //console.log('i', i, 'value', sourceValues[i]);
      let text = sourceValues[i].trim();
      if (text != ''){
        //Does the string include 'chapter'
        if (text.toLowerCase().includes('chapter')){
          //Finish last chapter and start new one.
          if (textSoFar != ''){
            index += 1;
            chapters[index] = textSoFar;
          }
          textSoFar = text;
        } else {
          //append to textSoFar
          textSoFar += text
        }
      }
    }
    index += 1;
    chapters[index] = textSoFar;
    chapterValues = chapters.map(x => [x]);
    console.log('chapterValues', chapterValues);
    console.log(startRowIndex, chaptersColumnIndex, chapterValues.length, 1);
    let chapterRange = pdfSheet.getRangeByIndexes(startRowIndex, chaptersColumnIndex, chapterValues.length, 1);
    chapterRange.clear('Values');
    chapterRange.values = chapterValues;
  })
  
  
}