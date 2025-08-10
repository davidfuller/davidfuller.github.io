async function auto_exec() {}

const replacementsSheetName = 'Replacements';
const replacementsTableRangeName = 'reTable';

const textInputSourceRow = "source-row";
const textAreaOriginalText = "original-text";
const textAreaReplaceText = "replace-text";

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
}

async function replacementsAndProcess(){
  await doTheReplacements();
  await jade_modules.operations.processGerman();
  
}