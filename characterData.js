const linkedDataSheetName = 'Linked_Data';
const characterSheetName = 'Characters';
const settingsSheetName = 'Settings';
const codeVersion = '1.0';
function auto_exec(){
  console.log('Hello');
}

async function makeTheFullList(){
  await Excel.run(async function(excel){ 
    let linkedDataSheet = excel.workbook.worksheets.getItem(linkedDataSheetName);
    let resultRange = linkedDataSheet.getRange('ldAllResults');
    resultRange.clear("Contents");
    resultRange.load('rowIndex, columnIndex');
    await excel.sync();
    let startRow = resultRange.rowIndex
    for (let i = 1; i<= 7; i++){
      let rangeName = 'ldSheet' + i;
      let thisRange = linkedDataSheet.getRange(rangeName);
      thisRange.load('values')
      await excel.sync();
      let myValues = thisRange.values.map(x => x[0]);
      let filteredValues = myValues.filter((x) => x != 0)
      let filteredRangedValues = []
      for (let j = 0; j < filteredValues.length; j++){
        filteredRangedValues[j] = [filteredValues[j]];
      }
      console.log(i, myValues, filteredValues, filteredRangedValues);
      //let myIndecies = myData.map((x, i) => [x, i]).filter(([x, i]) => x == targetValue).map(([x, i]) => i + firstDataRow - 1);
      let tempRange = linkedDataSheet.getRangeByIndexes(startRow, resultRange.columnIndex, filteredValues.length, 1);
      tempRange.values = filteredRangedValues;
      await excel.sync();
      startRow = startRow + filteredValues.length
    }
    resultRange.removeDuplicates([0], false);
    await excel.sync();
    const sortFields = [
      {
        key: 0,
        ascending: true
      }
    ]
    resultRange.sort.apply(sortFields);
    await excel.sync();
    console.log('The full list is made');
  })
}
async function whichBooks(){
  await Excel.run(async function(excel){ 
    let linkedDataSheet = excel.workbook.worksheets.getItem(linkedDataSheetName);
    let characterSheet = excel.workbook.worksheets.getItem(characterSheetName); 
    let waitMessageRange = characterSheet.getRange('chMessage');
    waitMessageRange.values = [['Please wait...']]
    let waitMessage = tag('wait-message');
    waitMessage.style.display = 'block';
    await excel.sync();
    let results = [];
    let resultIndex = -1;
    for (let i = 1; i<= 7; i++){
      let rangeName = 'ldIsInBook0' + i;
      let thisRange = linkedDataSheet.getRange(rangeName);
      thisRange.load('values')
      await excel.sync();
      if (thisRange.values[0][0]){
        resultIndex += 1;
        results[resultIndex] = i;
      }
    }
    resultValue = results.join(', ');

    let booksRange = characterSheet.getRange('chBooks');
    booksRange.values = [[resultValue]];
    waitMessageRange.values = [['']];
    waitMessage.style.display = 'none';
  })
}
async function registerExcelEvents(){
  await Excel.run(async (excel) => {
    let characterSheet = excel.workbook.worksheets.getItem(characterSheetName); 
    characterSheet.onChanged.add(handleChange);
    await excel.sync();
    console.log("Event handler successfully registered for onChanged event for four sheets.");
  })
}

async function handleChange(event) {
  await Excel.run(async (excel) => {
      await excel.sync();        
      if ((event.address == 'C10') && event.source == 'Local'){
        await whichBooks();
      }
  })
}

async function showMain(){
  let main = tag(main-page);
  main.style.display = 'block';
  let wait = tag('start-wait');
  wait.style.display = 'none';
  await Excel.run(async (excel) => {
    let settingsSheet = excel.workbook.worksheets.getItem(settingsSheetName);
    let dateRange = settingsSheet.getRange('seData');
    dateRange.load('text');
    let versionRange = settingsSheet.getRange('seVersion');
    versionRange.load('values');
    await excel.sync();
    let versionString = 'Version ' + versionRange.values + ' Code: ' + codeVersion + ' Released: ' + dateRange.text;
    let versionInfo = tag('sheet-version')
    versionInfo.innerText = versionString;
  })
}
