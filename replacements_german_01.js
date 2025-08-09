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
    console.log('table range', tableRange.values);
    for (let i = 0; i < tableRange.values[0].length; i++) {
      if (tableRange.values[0][i].trim() == '') {
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
    let temp = [[rowIndex, searchText, replaceText]]
    targetData.values = temp;
    await excel.sync();
  })
}

async function addToReplacements() {
  console.log('My start')
  const sourceRowTextInput = tag(textInputSourceRow);
  const searchTextArea = tag(textAreaOriginalText);
  const replaceTextArea = tag(textAreaReplaceText);
  
  let rowIndex = sourceRowTextInput.value;
  let searchText = searchTextArea.value;
  let replaceText = replaceTextArea.value;
  
  await appendData(rowIndex, searchText, replaceText)

}