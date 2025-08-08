const lockedOriginalSheetName = 'Locked Original';
const germanProcessingSheetName = 'German Processing';
const originalLineAndTextName = 'loLineAndText';
const originalTextProcessingName = 'gpLineAndText';
const ukScriptRangeName = 'gpLineCharacterAndUKScript';
const processedRangeName = 'gpProcessed';
const originalRangeName = 'gpOriginal';

const ukScriptSheetName = 'Script'
const cueColumnIndex = 5;
const numberColumnIndex = 6;
const characterColumnIndex = 7;
const ukScriptColumnIndex = 10;
const firstRowIndex = 3;
const lastRowCount = 10000;

//offsets wityhin the UK Script Range in German processing
const cueOffset = 0;
const numberOffset = 1;
const characterOffset = 2;
const ukScriptOffset = 3;

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
    //copy in the values
    processingTextRange.copyFrom(origTextRange, 'values');
    await excel.sync();
  })
}

async function getUKScript(){
  //get the data from the uk script sheet
  let ukData = await getUKData();
  //get the row and column indexes for the script part
  let rowIndex
  let columnIndex
  let rowCount
  await Excel.run(async function(excel){
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let ukScriptRange = gpSheet.getRange(ukScriptRangeName);
    ukScriptRange.load('rowIndex, columnIndex, rowCount')
    await excel.sync()
    rowIndex = ukScriptRange.rowIndex
    columnIndex = ukScriptRange.columnIndex
    rowCount = ukScriptRange.rowCount
  })
  
  //Fill in the cue
  await fillRangeByIndexes(germanProcessingSheetName, rowIndex, columnIndex + cueOffset, rowCount, ukData.cue, true);
  //Fill in the number
  await fillRangeByIndexes(germanProcessingSheetName, rowIndex, columnIndex + numberOffset, rowCount, ukData.number, true);
  //Fill in the character
  await fillRangeByIndexes(germanProcessingSheetName, rowIndex, columnIndex + characterOffset, rowCount, ukData.character, true);
  //Fill in the number
  await fillRangeByIndexes(germanProcessingSheetName, rowIndex, columnIndex + ukScriptOffset, rowCount, ukData.ukScript, true);

}

async function getUKData(){
  //The data is for the full length UK Script Sheet, but only the script. No scene, walla etc.
  /*
    The data is:
      index       the rowIndex of the data
      cue         the value in the cue column
      number      the value in the number column
      character   the name of the character (or narrator) *Incudes narrator(cut)
      ukScript    The text of the script.
  */
  let ukData = {};
  ukData.cue = [];
  ukData.index = [];
  ukData.number = [];
  ukData.character = [];
  ukData.ukScript = [];
  await Excel.run(async function(excel){
    //get the sheets and ranges
    const ukScriptSheet = excel.workbook.worksheets.getItem(ukScriptSheetName);
    let cueRange = ukScriptSheet.getRangeByIndexes(firstRowIndex, cueColumnIndex, lastRowCount, 1);
    let numberRange = ukScriptSheet.getRangeByIndexes(firstRowIndex, numberColumnIndex, lastRowCount, 1);
    let characterRange = ukScriptSheet.getRangeByIndexes(firstRowIndex, characterColumnIndex, lastRowCount, 1);
    let ukScriptRange = ukScriptSheet.getRangeByIndexes(firstRowIndex, ukScriptColumnIndex, lastRowCount, 1);
    cueRange.load('rowIndex, values');
    numberRange.load('values');
    characterRange.load('values');
    ukScriptRange.load('values')
    await excel.sync();
    console.log('rowIndex: ', cueRange.rowIndex);
    let cueValues = cueRange.values.map(x => x[0]);
    let numberValues = numberRange.values.map(x => x[0]);
    let characterValues = characterRange.values.map(x => x[0]);
    let ukScriptValues = ukScriptRange.values.map(x => x[0]);
    console.log(cueValues)
    for (let i = 0; i < cueValues.length; i++){
      if (!isNaN(parseInt(cueValues[i]))){
        if (!((characterValues[i].trim() == '') && (ukScriptValues[i] == ''))){
          ukData.cue.push(parseInt(cueValues[i]))
          ukData.index.push(i);
          ukData.number.push(numberValues[i]);
          ukData.character.push(characterValues[i]);
          ukData.ukScript.push(ukScriptValues[i]);
        }
      }
    }
    console.log('ukData: ', ukData);
  })
  return ukData;
}

async function fillRangeByIndexes(sheetName, rowIndex, columnIndex, rowCount, dataArray, doClear){
 await Excel.run(async function(excel){
  const mySheet = excel.workbook.worksheets.getItem(sheetName);
  const myRange = mySheet.getRangeByIndexes(rowIndex, columnIndex, rowCount, 1);
  myRange.load("rowIndex, columnIndex");
  if (doClear){
    myRange.clear("Contents")
  }
  await excel.sync();

  const destRange = mySheet.getRangeByIndexes(myRange.rowIndex, myRange.columnIndex, dataArray.length, 1)
  destRange.load('address');
  await excel.sync();
  console.log('address:', destRange.address);
  let temp = []
  for (let i = 0; i < dataArray.length; i++){
    temp[i] = [];
    temp[i][0] = dataArray[i]; 
  }
  console.log(temp)
  destRange.values = temp;
  await excel.sync();
 }) 
}

async function findThisBlock(){
  //Gets the cell of the active row in 'German Processed' column
  //Finds that block in the 'German Original' column
   await Excel.run(async function(excel){
    const activeCell = excel.workbook.getActiveCell();
    const gpProcessSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let processedRange = gpProcessSheet.getRange(processedRangeName);
    let originalRange = gpProcessSheet.getRange(originalRangeName);

    activeCell.load('rowIndex');
    processedRange.load('columnIndex');
    originalRange.load('rowIndex, columnIndex, values')

    await excel.sync();

    //Get the text from that row in processed.
    let searchTextRange = gpProcessSheet.getRangeByIndexes(activeCell.rowIndex, processedRange.columnIndex, 1, 1);
    searchTextRange.load('values');
    await excel.sync();
    let searchText = searchTextRange.values[0][0];
    console.log('Search Text', searchText)
   })

}
