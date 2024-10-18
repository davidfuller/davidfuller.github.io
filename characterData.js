const linkedDataSheetName = 'Linked_Data';
const characterSheetName = 'Characters';
const settingsSheetName = 'Settings';
const codeVersion = '1.0';
function auto_exec(){
  console.log('Hello');
}

async function makeTheFullList(){
  let waitMessage = tag('admin-wait-message');
  waitMessage.style.display = 'block';
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
  waitMessage.style.display = 'none';
}
async function whichBooks(){
  await Excel.run(async function(excel){ 
    let linkedDataSheet = excel.workbook.worksheets.getItem(linkedDataSheetName);
    let characterSheet = excel.workbook.worksheets.getItem(characterSheetName); 
    let waitMessageRange = characterSheet.getRange('chMessage');
    waitMessageRange.values = [['Please wait...']]
    let waitMessage = tag('wait-message');
    waitMessage.style.display = 'block';
    let booksRange = characterSheet.getRange('chBooks');
    booksRange.values = [['']];
    let numRange = characterSheet.getRange('chNumBooks');
    numRange.values = [['']];
    let characterNameRange = characterSheet.getRange('chCharacterName');
    characterNameRange.load('values')
    await excel.sync();
    let characterName = characterNameRange.values[0][0]
    if (characterName != ''){
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
      booksRange.values = [[resultValue]];
      numRange.values = [[results.length]];
    }
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
  let mainPage = tag('main-page');
  mainPage.style.display = 'block';
  let waitPage = tag('start-wait');
  waitPage.style.display = 'none';
  await Excel.run(async (excel) => {
    let settingsSheet = excel.workbook.worksheets.getItem(settingsSheetName);
    let dateRange = settingsSheet.getRange('seDate');
    dateRange.load('text');
    await excel.sync();
    let versionRange = settingsSheet.getRange('seVersion');
    versionRange.load('values');
    await excel.sync();
    let versionString = 'Version ' + versionRange.values + ' Code: ' + codeVersion + ' Released: ' + dateRange.text;
    let versionInfo = tag('sheet-version')
    versionInfo.innerText = versionString;
  })
}

async function refreshLinks(){
  let waitMessage = tag('admin-wait-message');
  waitMessage.style.display = 'block';
  await Excel.run(async (excel) => {
    let theLinks = excel.workbook.linkedWorkbooks
    theLinks.load('workbookLinksRefreshMode', 'items');
    await excel.sync();
    console.log(theLinks.workbookLinksRefreshMode, theLinks.items, theLinks.items[0].id);
    theLinks.refreshAll();
  })
  waitMessage.style.display = 'none';
}

function showAdmin(){
  let admin = tag('admin')
  if (admin.style.display === 'block'){
    admin.style.display = 'none';
  } else {
    admin.style.display = 'block';
  }
}

async function textSearch(){
  await Excel.run(async function(excel){ 
    let linkedDataSheet = excel.workbook.worksheets.getItem(linkedDataSheetName);
    let characterSheet = excel.workbook.worksheets.getItem(characterSheetName); 
    let waitMessageRange = characterSheet.getRange('chMessage');
    waitMessageRange.values = [['Please wait...']]
    let waitMessage = tag('wait-message');
    waitMessage.style.display = 'block';

    let textSearchRange = characterSheet.getRange('chTextSearch');
    textSearchRange.load('values');
    await excel.sync();

    let searchText = textSearchRange.values[0][0]
    let theTable = characterSheet.getRange('chTable');
    theTable.clear('Contents');

    if (searchText != ''){
      let results = [];
      let resultIndex = -1;
      for (let i = 1; i<= 7; i++){
        let rangeName = 'ldSheet' + i;
        let thisRange = linkedDataSheet.getRange(rangeName);
        thisRange.load('values')
        await excel.sync();
        let myValues = thisRange.values.map(x => x[0]);
        let filteredValues = myValues.filter((x) => x != 0)
        for (let j = 0; j < filteredValues.length; j++){
          if (filteredValues[j].toLowerCase().includes(searchText.toLowerCase())){
            let theIndex = doesCharacterAlreadyExist(results, filteredValues[j]);
            if (theIndex != -1){
              let booksArray = results[theIndex].books;
              booksArray.push(i);
              results[theIndex].books = booksArray;
            } else {
              resultIndex += 1;
              results[resultIndex] = {character: filteredValues[j], books: [i] }
            }
          }
        }
      }
      console.log('Results: ', results)
      let displayResult = [];
      for (let i = 0; i < results.length; i++){
        displayResult[i] = [results[i].character, results[i].books.join(', '), results.books.length]
      }
      console.log('Display Result', displayResult);
      theTable.values = displayResult;
      await excel.sync();
    }
    waitMessageRange.values = [['']];
    waitMessage.style.display = 'none';
  })
}

function doesCharacterAlreadyExist(resultArray, newCharacter){
  for (let i = 0; i < resultArray.length; i++){
    if (resultArray[i].character == newCharacter){
      return i;
    }
  }
  return -1

}
