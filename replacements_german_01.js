async function auto_exec() {}

const replacementsSheetName = 'Replacements';
const replacementsTableRangeName = 'reTable';

const textInputSourceRow = "source-row";
const textAreaOriginalText = "original-text";
const textAreaReplaceText = "replace-text";

const openSpeechChar = '»';
const closeSpeechChar = '«';
const openSingleQuoteChar = '›';
const closeSingleQuoteChar = '‹';
const eolChar = '|eol|'

const loadMessageLabelName = 'load-message';

const missingTextString = '[MISSING TEXT]'
const missingTextSubstringLength = 40;

async function appendData(rowIndex, searchText, replaceText) {
  //Finds the first empty row and adds the data
  await Excel.run(async function(excel) {
    // get the table range
    const reSheet = excel.workbook.worksheets.getItem(replacementsSheetName);
    let tableRange = reSheet.getRange(replacementsTableRangeName);
    tableRange.load('rowIndex, columnIndex, values');
    await excel.sync();
    //find the first empty row
    let emptyRowIndex = -1;
    for (let i = 0; i < tableRange.values.length; i++) {
      if (tableRange.values[i][0].toString().trim() == '') {
        emptyRowIndex = i + tableRange.rowIndex;
        break;
      }
    }
    console.log('emptyRowIndex', emptyRowIndex)
    let targetData = reSheet.getRangeByIndexes(emptyRowIndex, tableRange.columnIndex, 1, 3);
    targetData.load('values, address');
    await excel.sync();
    console.log('targetData', targetData.values);
    console.log('address', targetData.address)
    let temp = [
      [rowIndex, searchText, replaceText]
    ]
    targetData.values = temp;
    await excel.sync();
  })
}

async function addToReplacements() {
  // Gets the details from the html and appends to the table  
  const sourceRowTextInput = tag(textInputSourceRow);
  const searchTextArea = tag(textAreaOriginalText);
  const replaceTextArea = tag(textAreaReplaceText);

  let rowIndex = sourceRowTextInput.value;
  let searchText = searchTextArea.value;
  let replaceText = replaceTextArea.value;

  await appendData(rowIndex, searchText, replaceText)

}


async function doTheReplacements() {
  jade_modules.preprocess.showMessage(loadMessageLabelName, 'Doing replacements');
  //Takes each row of table
  //If has valid rowIndex then do a replacement
  let rowIndex;
  let searchText;
  let replaceText;
  await Excel.run(async function(excel) {
    // get the table range
    const reSheet = excel.workbook.worksheets.getItem(replacementsSheetName);
    let tableRange = reSheet.getRange(replacementsTableRangeName);
    tableRange.load('values');
    await excel.sync();

    for (let i = 0; i < tableRange.values.length; i++) {
      let testRowIndex = tableRange.values[i][0];
      if (!isNaN(parseInt(testRowIndex))) {
        rowIndex = parseInt(testRowIndex);
        searchText = tableRange.values[i][1];
        replaceText = tableRange.values[i][2];
        await jade_modules.operations.doAReplacement(rowIndex, searchText, replaceText);
      }
    }
  })
  jade_modules.preprocess.hideMessage(loadMessageLabelName);
}

async function replacementsAndProcess(){
  await doTheReplacements();
  await jade_modules.operations.processGerman();
}

function copySearchReplacingDoubleQuotes(){
  const searchTextArea = tag(textAreaOriginalText);
  const replaceTextArea = tag(textAreaReplaceText);
  
  let searchText = searchTextArea.value;
  let replaceText = searchText.replace(openSpeechChar, openSingleQuoteChar).replace(closeSpeechChar, closeSingleQuoteChar).trim();
  
  replaceTextArea.value = replaceText;
}

function isolateQuotedBit() {
  const searchTextArea = tag(textAreaOriginalText);  
  
  let searchText = searchTextArea.value;
  let openLocations = jade_modules.operations.locations(openSpeechChar, searchText);
  let closeLocations = jade_modules.operations.locations(closeSpeechChar, searchText);
  let result = [];
  console.log('searchText', searchText);
  console.log('openLocations', openLocations);
  console.log('closeLocations', closeLocations);
  
  if ((openLocations.length > 0) && (openLocations.length == closeLocations.length)){
    for (let i = 0; i < openLocations.length; i++) {
      result[i] = searchText.substring(openLocations[i], closeLocations[i] + 1);
      console.log(i, result[i])
    }
  }
  
  searchTextArea.value = result[0];
  
}

function createMissingSearchAndReplace(){
  // Takes last missingTextSubstringLength (40) chars of searchText and makes it searchText
  // Makes replaceTex equal searchText plus [MISSING TEXT]
  const searchTextArea = tag(textAreaOriginalText);
  const replaceTextArea = tag(textAreaReplaceText);
  
  let searchText = searchTextArea.value;
  let newSearchText = searchText.substr(-missingTextSubstringLength);
  searchTextArea.value = newSearchText;
  console.log(newSearchText.trim().slice(-1));
  if (newSearchText.trim().slice(-1) == closeSpeechChar){
    console.loh("I'm here");
    replaceTextArea.value = newSearchText + missingTextString;
  } else {
    replaceTextArea.value = newSearchText + eolChar +  missingTextString;
  }
}

function insertEol(){
  const searchTextArea = tag(textAreaOriginalText);
  let insertCharPosition = searchTextArea.selectionEnd
  console.log(insertCharPosition);
  let theText = searchTextArea.value;
  let before = theText.substring(0, insertCharPosition);
  let after = theText.substring(insertCharPosition);
  console.log('text', theText, 'Before', before, 'after', after );
  const replaceTextArea = tag(textAreaReplaceText);
  replaceTextArea.value = before + eolChar + after;
}
