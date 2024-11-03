const linkedDataSheetName = 'Linked_Data';
const characterSheetName = 'Characters';
const sceneSheetName = 'Scenes';
const settingsSheetName = 'Settings';
const codeVersion = '2.1';
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
    let characterSheet = excel.workbook.worksheets.getItem(characterSheetName); 
    let waitMessageRange = characterSheet.getRange('chMessage');
    waitMessageRange.values = [['Please wait...']]
    let waitMessage = tag('wait-message');
    waitMessage.style.display = 'block';
    let characterNameRange = characterSheet.getRange('chCharacterName');
    characterNameRange.load('values')
    await excel.sync();
    let characterName = characterNameRange.values[0][0]
    if (characterName != ''){
      let results = await findCharacter(characterName, true)
      if (results[0].valid){
       await display(results);
      }
    }
    waitMessageRange.values = [['']];
    waitMessage.style.display = 'none';
  })
} 

function numBooks(theBooks){
  theBooks = '' + theBooks;
  let numBooks;
  if (theBooks.includes(',')){
    numBooks = theBooks.split(',').length;
  } else {
    if (isNaN(parseInt(theBooks))){
      numBooks = 0;
    } else {
      numBooks = 1;
    }
  }
  return numBooks;
}
async function whichBooksOld(){
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
      if ((event.address == 'B7') && event.source == 'Local'){
        await textSearch();
      }
      if ((event.address == 'B9') && event.source == 'Local'){
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
    let characterSheet = excel.workbook.worksheets.getItem(characterSheetName); 
    let waitMessageRange = characterSheet.getRange('chMessage');
    waitMessageRange.values = [['Please wait...']]
    let waitMessage = tag('wait-message');
    waitMessage.style.display = 'block';

    let textSearchRange = characterSheet.getRange('chTextSearch');
    textSearchRange.load('values');
    await excel.sync();

    let searchText = textSearchRange.values[0][0]
    if (searchText != ''){
      let results = await findCharacter(searchText, false)
      console.log('Results: ', results)
      await display(results);
    }
    waitMessageRange.values = [['']];
    waitMessage.style.display = 'none';
  })
}

async function textSearchOld(){
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
    theTable.load('rowIndex, columnIndex, columnCount');
    await excel.sync();
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
        displayResult[i] = [results[i].character, results[i].books.join(', '), results[i].books.length]
      }
      console.log('Display Result', displayResult);
      let displayRange = characterSheet.getRangeByIndexes(theTable.rowIndex, theTable.columnIndex, displayResult.length, theTable.columnCount);
      displayRange.values = displayResult;
      await excel.sync();
      const sortFields = [
        {
          key: 0,
          ascending: true
        }
      ]
      theTable.sort.apply(sortFields);
      let numItems = characterSheet.getRange('chItems');
      numItems.values = displayResult.length
      
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

async function gatherData(){
  //This takes the data from each of the books and creates total data
  const resultName = 'ldTotal';
  const numBooks = 7;
  const bookNameBase = 'ldSheet';

  await Excel.run(async function(excel){
    let linkedDataSheet = excel.workbook.worksheets.getItem(linkedDataSheetName);
    //get results range and clear it
    let resultRange = linkedDataSheet.getRange(resultName);
    resultRange.load('address');
    resultRange.clear("Contents")
    await excel.sync();
    console.log('address: ', resultRange.address);
    
    for (let i = 0; i < numBooks; i++){
      let newRows = [];
      let bookName = bookNameBase + (i + 1);
      // Get the book details
      console.log('bookName', bookName);
      let bookRange = linkedDataSheet.getRange(bookName);
      bookRange.load('text, address, rowCount');
      await excel.sync();
      console.log ('Book: ', i, 'rowCount:', bookRange.rowCount, 'data: ', bookRange.text, 'length', bookRange.text.length);
      resultRange.load('values, rowIndex, rowCount, columnIndex, columnCount')
      await excel.sync();
      console.log ('result rowCount', resultRange.rowCount, 'values: ', resultRange.values);
      //let currentNames = resultRange.values.map(x => x[0]).filter((x) => {x != '' })
      currentNames = [];
      for (let i = 0; i < resultRange.values.length; i++){
        if (resultRange.values[i][0] != ''){
          currentNames.push(resultRange.values[i]);
        }
      }
      console.log('currentNames: ', currentNames);
      for (let item = 0; item < bookRange.text.length; item++){
        let thisCharacter = bookRange.text[item][0];
        if (thisCharacter != '0'){
          let found = false;
          for (let charIndex = 0; charIndex < currentNames.length; charIndex++){
            if (currentNames[charIndex][0] == thisCharacter){
              //Do something with charIndex
              console.log('Already exists: ', thisCharacter, ' before ', currentNames[charIndex] )
              currentNames[charIndex][1] = currentNames[charIndex][1] + ', ' + (i + 1);
              currentNames[charIndex][2] = currentNames[charIndex][2] + parseInt(bookRange.text[item][1]); 
              currentNames[charIndex][3] = currentNames[charIndex][3] + parseInt(bookRange.text[item][2]); 
              currentNames[charIndex][4] = currentNames[charIndex][4] + ', ' + bookRange.text[item][3]; 
              found = true;
              console.log('Already exists: ', thisCharacter, ' after ', currentNames[charIndex] )
            } 
          }
          if (!found){
            let newElement = [thisCharacter, '' + (i + 1), bookRange.text[item][1], bookRange.text[item][2], bookRange.text[item][3]];
            //console.log('Item:', item, 'New element: ', newElement)
            newRows.push(newElement);
          }
        }
      }
      //Add currentNames back in.
      console.log(resultRange.rowIndex, resultRange.columnIndex, currentNames.length, resultRange.columnCount);
      console.log('currentNames', currentNames);
      let tempRange;
      if (currentNames.length > 0){
        tempRange = linkedDataSheet.getRangeByIndexes(resultRange.rowIndex, resultRange.columnIndex, currentNames.length, resultRange.columnCount);
        tempRange.values = currentNames;
        await excel.sync();
      }
      
      
      //Now do the new rows
      let startRowIndex = resultRange.rowIndex + currentNames.length;
      console.log(startRowIndex, resultRange.columnIndex, newRows.length, resultRange.columnCount);
      console.log('New rows', newRows);
      if (newRows.length > 0){
        tempRange = linkedDataSheet.getRangeByIndexes(startRowIndex, resultRange.columnIndex, newRows.length, resultRange.columnCount);
        tempRange.values = newRows;
        await excel.sync();
      }
    }
    const sortFields = [
      {
        key: 0,
        ascending: true
      }
    ]
    resultRange.sort.apply(sortFields);
    await excel.sync();

    let wordCountAllBooksRange = linkedDataSheet.getRange('ldWordCountAllBooks')
    wordCountAllBooksRange.load('rowIndex, columnIndex, columnCount');
    wordCountAllBooksRange.clear("Contents")
    await excel.sync();
    let startRowIndex = wordCountAllBooksRange.rowIndex;
    //now do the scene word count
    for (let i = 0; i < numBooks; i++){
      let bookRange = linkedDataSheet.getRange('ldWordCount' + (i + 1));
      bookRange.load('values')
      await excel.sync();
      let result = [];
      let index = - 1;
      for (let i = 0; i < bookRange.values.length; i++){
        if (!((bookRange.values[i][0] == 0) || (bookRange.values[i][1] == 0))){
          index += 1;
          result[index] = bookRange.values[i];
        }
      }
      console.log('Book: ', (i+1), 'Word Count: ', result);
      let tempRange = linkedDataSheet.getRangeByIndexes(startRowIndex, wordCountAllBooksRange.columnIndex, result.length, wordCountAllBooksRange.columnCount);
      tempRange.values = result;
      await excel.sync;
      startRowIndex = startRowIndex + result.length;
    }
  })
}

async function findCharacter(characterName, exact){
  const resultName = 'ldTotal';
  let results = [];
  await Excel.run(async function(excel){
    let linkedDataSheet = excel.workbook.worksheets.getItem(linkedDataSheetName);
    let resultRange = linkedDataSheet.getRange(resultName)
    resultRange.load('values');
    await excel.sync();
    
    characterNames = [];
    for (let i = 0; i < resultRange.values.length; i++){
      if (resultRange.values[i][0] != ''){
        let temp = {
          index: i,
          character: resultRange.values[i][0]
        }
        characterNames.push(temp);
      }
    }
    if (exact){
      let index = characterNames.findIndex(x => (x.character == characterName))
      let result = {};
      if (index == -1){
        result = {
          valid: false
        }
      } else {
        let item = characterNames[index].index;
        result = {
          valid: true,
          character: resultRange.values[item][0],
          books: resultRange.values[item][1],
          sceneWords: resultRange.values[item][2],
          lineWords: resultRange.values[item][3],
          scenes: resultRange.values[item][4]
        }
      }
      results = [result];
    } else {
      for (let i = 0; i < characterNames.length; i++){
        if (characterNames[i].character.toLowerCase().includes(characterName.toLowerCase())){
          let item = characterNames[i].index;
          result = {
            valid: true,
            character: resultRange.values[item][0],
            books: resultRange.values[item][1],
            sceneWords: resultRange.values[item][2],
            lineWords: resultRange.values[item][3],
            scenes: resultRange.values[item][4]
          }
          results.push(result)
        }
      }
    }
    console.log('Results: ', results)  
  })
  return results;
}

async function display(results){
  await Excel.run(async function(excel){
    let displayResult = [];
    let allBooks = [];
    for (let i = 0; i < results.length; i++){
      let sceneDisplay;
      if (results[i].scenes == 0){
        sceneDisplay = '';
      } else {
        sceneDisplay = results[i].scenes
      }
      displayResult[i] = [results[i].character, results[i].books, numBooks(results[i].books), sceneDisplay];
      let tempArray = ('' + results[i].books).split(', ');
      allBooks = allBooks.concat(tempArray);
    }
    console.log('Display Result', displayResult);
    let characterSheet = excel.workbook.worksheets.getItem(characterSheetName); 
    let theTable = characterSheet.getRange('chTable');
    theTable.clear('Contents');
    theTable.load('rowIndex, columnIndex, columnCount');
    await excel.sync();
    
    let displayRange = characterSheet.getRangeByIndexes(theTable.rowIndex, theTable.columnIndex, displayResult.length, theTable.columnCount);
    displayRange.values = displayResult;
    await excel.sync();
    const sortFields = [
      {
        key: 0,
        ascending: true
      }
    ]
    theTable.sort.apply(sortFields);
    let numItems = characterSheet.getRange('chItems');
    numItems.values = displayResult.length

    let totalLinesWords = 0;
    let totalSceneWords = 0;
    let scenesUsed = [];
    for (let i = 0; i < results.length; i++){
      // lets create a list of scenes used by this result
      if (results[i].scenes != 0){
        //Split it as a string
        let theScenes = ('' + results[i].scenes).split(', ');
        scenesUsed = scenesUsed.concat(theScenes);
        totalLinesWords += results[i].lineWords;
      }
    }
    //remove duplicates
    let uniqueScenes = [...new Set(scenesUsed)];
    const linkedDataSheet = excel.workbook.worksheets.getItem(linkedDataSheetName);
    let sceneWordCount = linkedDataSheet.getRange('ldWordCountAllBooks');
    sceneWordCount.load('values');
    await excel.sync();

    let wordCountData = sceneWordCount.values
    let justScenes = wordCountData.map(x => x[0])

    for(let i = 0; i < uniqueScenes.length; i++){
      let myIndex = justScenes.indexOf(parseInt(uniqueScenes[i]))
      if (myIndex != -1){
        totalSceneWords = totalSceneWords + wordCountData[myIndex][1];
      }
    }

    let linesUsedRange = characterSheet.getRange('chLinesUsed');
    let fullScenesRange = characterSheet.getRange('chFullScene');

    linesUsedRange.values = [[totalLinesWords]];
    fullScenesRange.values = [[totalSceneWords]];

    await excel.sync();
    
    //now do the total books
    let uniqueBooks = [...new Set(allBooks)];
    let sortedBooks = uniqueBooks.sort().join(', ');
    let sortedBooksRange = characterSheet.getRange('chAllBooks');
    sortedBooksRange.values = [[sortedBooks]];
    await excel.sync();
    
  })
}

async function doSearch(){
  await Excel.run(async function(excel){
    const characterSheet = excel.workbook.worksheets.getItem(characterSheetName);
    let choiceRange = characterSheet.getRange('chSearchType');
    choiceRange.load('values');
    await excel.sync();

    if (choiceRange.values[0][0] == 'Text Search'){
      await textSearch();
    } else {
      await whichBooks();
    }
  })
}

async function createSceneList(){
  const characterIndex = 0;
  const scenesIndex = 4;
  let characterData = []
  let myIndex = - 1;
  let whichBooks = getCheckedBooks();
  await Excel.run(async function(excel){
    const linkedDataSheet = excel.workbook.worksheets.getItem(linkedDataSheetName);
    let dataRange = linkedDataSheet.getRange('ldTotal');
    let bookRange = [];
    dataRange.load('values');
    for (let i = 1; i <= 7; i++){
      console.log('ldBook0' + i + 'SceneRange')
      bookRange[i] = linkedDataSheet.getRange('ldBook0' + i + 'SceneRange');
      bookRange[i].load('values');
      await excel.sync();
    }
    await excel.sync();
    const dataValues = dataRange.values
    console.log('Data Values', dataValues);
    for (let i = 0; i < dataValues.length; i++){
      if (dataValues[i][characterIndex].trim() != ''){
        myIndex += 1
        characterData[myIndex] = {
          name: dataValues[i][characterIndex],
          scenes: getSceneArray(dataValues[i][scenesIndex])
        }
      }
    }
    console.log('characterData', characterData);
    let sceneData = []
    for (let i = 0; i < characterData.length; i++){
      for (let j = 0; j < characterData[i].scenes.length; j++){
        let sceneNo = characterData[i].scenes[j];
        if((!isNaN(sceneNo)) && (sceneNo > 0)){
          if (Array.isArray(sceneData[sceneNo])){
            console.log(sceneNo, ' is array')
            sceneData[sceneNo] = sceneData[sceneNo].concat([characterData[i].name]);
          } else {
            console.log(sceneNo, ' is NOT array')
            sceneData[sceneNo] = [characterData[i].name];
          }
        }
      }
    }
    console.log('sceneData', sceneData);
    let bookScenes = [];
    for (let i = 1; i <= 7; i++){
      bookScenes[i] = {
        min: bookRange[i].values[0][0],
        max: bookRange[i].values[1][0]
      }
    }
    console.log('Book Scenes', bookScenes);

    let resultData = [];
    let resultIndex = -1;
    for (let i = 1; i<= characterData.length; i++){
      let bookNo = bookFromScene(i, bookScenes);
      if (whichBooks.includes(bookNo)){
        resultIndex += 1;
        resultData[resultIndex]= [i, bookNo, sceneData[i].join(' | ') , sceneData[i].length];
      }
    }

    let sceneSheet = excel.workbook.worksheets.getItem('Scenes');
    let sceneTableRange = sceneSheet.getRange('scTable');
    sceneTableRange.load('rowIndex, columnIndex, columnCount');
    sceneTableRange.clear('Contents');
    await excel.sync();
    let tempRange = sceneSheet.getRangeByIndexes(sceneTableRange.rowIndex, sceneTableRange.columnIndex, resultData.length, sceneTableRange.columnCount);
    tempRange.values = resultData;
    await excel.sync();
  });
  console.log('Which books', whichBooks);
}

function bookFromScene(sceneNo, bookScenes){
  for (let i = 1; i <= bookScenes.length; i++){
    if ((sceneNo >= bookScenes[i].min) && (sceneNo <= bookScenes[i].max)){
      return i;
    }
  }
  return 0;
}
function getSceneArray(sceneString){
  let result = []
  if ((sceneString == '')||(sceneString == 0)){
    return result;
  }
  let scenes = sceneString.toString().split(', ');
  result = scenes.map(x => parseInt(x));
  return result;
}
function getCheckedBooks(){
  let result = [];
  for (let i = 1; i <= 7; i++){
    let chk = tag('book-' + i)
    if (chk.checked){
      result.push(i);
    }
  }
  return result
}

async function showScenePage(){
  const sceneBackgroundColour = '#ffafaf';
  const sceneColour = '#640000';
  let mainPage = tag("main-page");
  mainPage.style.display = 'none';
  let scenePage = tag("scene-page");
  scenePage.style.display = 'block'
  let characterSummary = tag('character-summary');
  characterSummary.style.backgroundColor = sceneBackgroundColour;
  characterSummary.style.color = sceneColour;
  await Excel.run(async function(excel){
    const sceneSheet = excel.workbook.worksheets.getItem(sceneSheetName);
    sceneSheet.activate();
  })
}
async function gotoMain(){
  const sceneBackgroundColour = '#ffffb3';
  const sceneColour = '#424200';
  let mainPage = tag("main-page");
  mainPage.style.display = 'block';
  let scenePage = tag("scene-page");
  scenePage.style.display = 'none'
  let characterSummary = tag('character-summary');
  characterSummary.style.backgroundColor = sceneBackgroundColour;
  characterSummary.style.color = sceneColour;
  await Excel.run(async function(excel){
    const characterSheet = excel.workbook.worksheets.getItem(characterSheetName);
    characterSheet.activate();
  })
}
