const lockedOriginalSheetName = 'Locked Original';
const germanProcessingSheetName = 'German Processing';
const originalLineAndTextName = 'loLineAndText';
const originalLineName = 'loLineNos';
const originalTextName = 'loText';
const joinsRangeName = 'loJoins';
const originalTextProcessingName = 'gpLineAndText';
const ukScriptRangeName = 'gpLineCharacterAndUKScript';
const processedRangeName = 'gpProcessed';
const processedLineNoRangeName = 'gpLineNo';
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

const textInputProcessAddress = "process-address";
const textInputSourceRow = "source-row";
const textAreaOriginalText = "original-text";
const textAreaReplaceText = "replace-text";

async function doTheCopy() {
  let joinsIndexes = await findJoins();
  let joinedValues;
  await Excel.run(async function(excel) {
    //get the sheets and ranges
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let processingTextRange = gpSheet.getRange(originalTextProcessingName);
    let processingLineRange = gpSheet.getRange(processedLineNoRangeName);
    
    const origSheet = excel.workbook.worksheets.getItem(lockedOriginalSheetName);
    let origLineNoRange = origSheet.getRange(originalLineName);
    let originalTextRange = origSheet.getRange(originalTextName);
    originalTextRange.load('rowIndex, values');
    await excel.sync();

    //clear the processing range
    processingTextRange.clear('Contents')
    //copy the line numbers
    processingLineRange.copyFrom(origLineNoRange, 'values');
    await excel.sync();

    //Now do the joined tet
    let textValues = originalTextRange.values.map(x => x[0]);
    joinedValues = createJoinedText(textValues, joinsIndexes, originalTextRange.rowIndex);
  })
  console.log('Joined Values in doTheCopy', joinedValues)
  await jade_modules.operations.fillRange(germanProcessingSheetName, originalRangeName, joinedValues, true);
}

async function getUKScript() {
  //get the data from the uk script sheet
  let ukData = await getUKData();
  //get the row and column indexes for the script part
  let rowIndex
  let columnIndex
  let rowCount
  await Excel.run(async function(excel) {
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

async function getUKData() {
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
  await Excel.run(async function(excel) {
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
    for (let i = 0; i < cueValues.length; i++) {
      if (!isNaN(parseInt(cueValues[i]))) {
        if (!((characterValues[i].trim() == '') && (ukScriptValues[i] == ''))) {
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

async function fillRangeByIndexes(sheetName, rowIndex, columnIndex, rowCount, dataArray, doClear) {
  await Excel.run(async function(excel) {
    const mySheet = excel.workbook.worksheets.getItem(sheetName);
    const myRange = mySheet.getRangeByIndexes(rowIndex, columnIndex, rowCount, 1);
    myRange.load("rowIndex, columnIndex");
    if (doClear) {
      myRange.clear("Contents")
    }
    await excel.sync();

    const destRange = mySheet.getRangeByIndexes(myRange.rowIndex, myRange.columnIndex, dataArray.length, 1)
    destRange.load('address');
    await excel.sync();
    console.log('address:', destRange.address);
    let temp = []
    for (let i = 0; i < dataArray.length; i++) {
      temp[i] = [];
      temp[i][0] = dataArray[i];
    }
    console.log(temp)
    destRange.values = temp;
    await excel.sync();
  })
}

async function findThisBlock() {
  //Gets the cell of the active row in 'German Processed' column
  //Finds that block in the 'German Original' column
  await Excel.run(async function(excel) {
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
    searchTextRange.load('values, address');
    await excel.sync();
    let searchText = (searchTextRange.values[0][0]).toLowerCase();
    console.log('Search Text', searchText)
    putInTextArea(textInputProcessAddress, searchTextRange.address)
    putInTextArea(textAreaOriginalText, searchText);

    originalTexts = originalRange.values.map((x => x[0]));
    let foundRowIndex = 0;
    for (let i = 0; i < originalTexts.length; i++) {
      if (originalTexts[i].toLowerCase().includes(searchText)) {
        foundRowIndex = i + originalRange.rowIndex;
        putInTextArea(textAreaOriginalText, originalTexts[i]);
        putInTextArea(textInputSourceRow, foundRowIndex);
        break;
      }
    }
    console.log('Found Row Index', foundRowIndex);

    if (foundRowIndex > 0) {
      let rangeToSelect = gpProcessSheet.getRangeByIndexes(foundRowIndex, originalRange.columnIndex, 1, 1)
      rangeToSelect.select();
      await excel.sync();
    }
  })
}

function putInTextArea(textAreaID, text) {
  let textArea = tag(textAreaID);
  textArea.value = text;
}
async function returnToProcessedCell() {
  let textArea = tag(textInputProcessAddress);
  let cellAddress = textArea.value;
  console.log('cellAddress', cellAddress);
  await Excel.run(async function(excel) {
    const gpProcessSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let toProcessedRange = gpProcessSheet.getRange(cellAddress);
    toProcessedRange.select();
    await excel.sync();
  })
}

async function findJoins(){
  let indexes = [];
  await Excel.run(async function(excel) {
    const origSheet = excel.workbook.worksheets.getItem(lockedOriginalSheetName);
    let joinsRange = origSheet.getRange(joinsRangeName);
    joinsRange.load('rowIndex, values');
    await excel.sync();
    let joinsText = joinsRange.values.map(x => x[0]);
    for (let i = 0; i < joinsText.length; i++){
      if (joinsText[i].toLowerCase() == 'join'){
        indexes.push(joinsRange.rowIndex + i)
      }
    }
  })
  console.log('Joins Row Indexes', indexes)
  return indexes;
}
function createJoinedText(textValues, joinIndexes, textRowIndex){
  //returns array with the relevant text joined
  let joinedText = [];
  let previousAJoin = false;
  let nextIndex = 0;
  let doThis = true;
  for (let i = 0; i < textValues.length; i++){
    thisRowIndex = i + textRowIndex;
    if (previousAJoin){
      if (thisRowIndex < nextIndex){
        //Not there yet do nothing
        doThis = false
      } else {
        doThis = true
      }
    }
    if (doThis){
      //Previous item was not a join
      console.log('joinIndexes', joinIndexes);
      if (joinIndexes.includes(thisRowIndex)){
        //Do a join
        //Test to see how many joins....
        let tempRowIndex = thisRowIndex + 1;
        let done = true;
        let lastOffset = 1;
        do {
          if (joinIndexes.includes(tempRowIndex)){
            tempRowIndex += 1;
            lastOffset += 1;
            done = false;
          } else {
            done = true;
          }
        } while (!done);
        let tempResult = ''
        for (let offset = 0; offset <= lastOffset; offset++){
          if (textValues?.[i + offset]){
            tempResult = tempResult + ' ' + textValues[i + offset]
          }
        }
        joinedText.push(tempResult.trim());
        previousAJoin = true;
        nextIndex = tempRowIndex + 1
      } else {
        //put in joinedText
        joinedText.push(textValues[i]);
        previousAJoin = false;
      }
    }
  }
  console.log('joinedText', joinedText);
  return joinedText
}
async function findInLockedOriginal() {
  //Gets the cell of the active row in 'German Processed' column
  //Finds that block in the 'German Text' column of the Locked Origial
  await Excel.run(async function(excel) {
    const activeCell = excel.workbook.getActiveCell();
    const gpProcessSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    let processedRange = gpProcessSheet.getRange(processedRangeName);
    
    activeCell.load('rowIndex');
    processedRange.load('columnIndex');
    await excel.sync();

    //Get the text from that row in processed.
    let searchTextRange = gpProcessSheet.getRangeByIndexes(activeCell.rowIndex, processedRange.columnIndex, 1, 1);
    searchTextRange.load('values, address');
    await excel.sync();
    
    let searchText = (searchTextRange.values[0][0]).toLowerCase();
    console.log('Search Text', searchText)

    const lockedOriginalSheet =excel.workbook.worksheets.getItem(lockedOriginalSheetName);
    let originalTextRange = lockedOriginalSheet.getRange(originalTextName);
    originalTextRange.load('rowIndex, values, columnIndex')
    await excel.sync();

    let originalText = originalTextRange.values.map(x => x[0]);
    for (i = 0; i < originalText.length; i++){
      if (originalText[i].toLowerCase == searchText){
        let selectedRowIndex = i + originalTextRange.rowIndex;
        let selectRange = lockedOriginalSheet.getRangeByIndexes(selectedRowIndex, originalTextRange.columnIndex, 1, 1);
        selectRange.select();
        await excel.sync();
      }
    }
  })
}
