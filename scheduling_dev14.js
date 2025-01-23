let characterlistSheet, forDirectorSheet, forActorSheet, forSchedulingSheet;
const characterListName = 'Character List';
const scriptSheetName = 'Script'
const forDirectorName = 'For Directors';
const forActorName = 'For Actors'
const forSchedulingName = 'For Scheduling'
const locationSheetName = 'Locations'
const forDirectorTableName = 'fdTable';
const forActorsTableName = "faTable";
const forSchedulingTableName = 'fsTable'
const locationTableName = 'loTable'
const numItemsDirectorsName = 'fdItems';
const numItemsActorsName = 'faItems';
const numItemsSchedulingName = 'fsItems';
const numItemsLocationName = 'loItems'
const actorScriptName = 'Actor Script';
const actorScriptBookName = 'asBook';
const actorScriptCharacterName = 'asCharacter';
const actorScriptCharcaterHeadingName = 'asCharacterHeading'
const actorScriptTableName = 'asTable'

let myFormats = {
  purple: '#f3d1f0',
  green: '#daf2d0',
  lightGrey: '#a6a6a6',
  orange: '#f7c7ac'
}

let choiceType ={
  list: 'List Search',
  text: 'Text Search'
}

let scriptChoice = {
  all: 'all',
  selection: 'selection',
  none: 'none'
}

function auto_exec(){
}
async function loadReduceAndSortCharacters(){
  await Excel.run(async function(excel){ 
    characterlistSheet = excel.workbook.worksheets.getItem(characterListName);
    let characters = await jade_modules.operations.getCharacters(scriptSheetName, null);
    console.log('the characters', characters);
    let characterRange = characterlistSheet.getRange('clCharacters');
    characterRange.clear("Contents");
    characterRange.load('values');
    await excel.sync();
    console.log(characterRange.values);
    characterRange.values = characters;
    await excel.sync();
    characterRange.removeDuplicates([0], false);
    await excel.sync();
    
    const sortFields = [
      {
        key: 0,
        ascending: true
      }
    ]
    characterRange.sort.apply(sortFields);
    await excel.sync();
    
  })  
}

async function getDirectorInfo(){
  await Excel.run(async function(excel){
    let waitLabel = tag('director-wait');
    waitLabel.style.display = 'block';
    forDirectorSheet = excel.workbook.worksheets.getItem(forDirectorName);
    const waitCell = forDirectorSheet.getRange('fdMessage');
    waitCell.values = 'Please wait...';
    await excel.sync();
    let characterChoiceRange = forDirectorSheet.getRange('fdCharacterChoice');
    characterChoiceRange.load('values');
    await excel.sync();
    let characterName = characterChoiceRange.values[0][0];
    console.log('Character ',characterName);
    let character = {name: characterName, type: choiceType.list}
    let myData = await jade_modules.operations.getDirectorDataV2(character);
    console.log('Scheduling myData', myData);
    let dataRange = forDirectorSheet.getRange(forDirectorTableName);
    let numItems = forDirectorSheet.getRange(numItemsDirectorsName);
    dataRange.clear("Contents");
    dataRange.load('rowCount');
    await excel.sync();
    let dataArray = [];
    for (i = 0; i < dataRange.rowCount; i++){
      let thisRow = new Array(5).fill("");
      if (i < myData.length){
        thisRow = [myData[i].sceneNumber, myData[i].lineNumber, myData[i].ukNumTakes, myData[i].ukTakeNum, myData[i].ukDateRecorded];
      }
      dataArray.push(thisRow);
    }
    console.log('dataArray', dataArray, 'rowCount', dataRange.rowCount, 'dataLength', myData.length);
    dataRange.values = dataArray;
    numItems.values = myData.length;
    await excel.sync();
    waitLabel.style.display = 'none';
    waitCell.values = '';
    await excel.sync();
  }) 
}

async function getActorInformation(){
  await Excel.run(async function(excel){
    let waitLabel = tag('actor-wait');
    waitLabel.style.display = 'block';
    let forActorSheet = excel.workbook.worksheets.getItem(forActorName);
    const waitCell = forActorSheet.getRange('faMessage');
    waitCell.values = 'Please wait...';
    
    let usOnly = await isUsOnly();
    let character = await getActor(forActorName);
    console.log('Character ',character.name, character.type);
    let myData = await jade_modules.operations.getDirectorDataV2(character);
    let myLocation = await jade_modules.operations.getLocations();
    console.log('Scheduling myData', myData);
    console.log('Locations', myLocation);
    
    let tempData = [];
    if (usOnly){
      for (let i = 0; i < myData.length; i++){
        if (myData[i].usCue.trim() != ''){
          tempData.push(myData[i])
        }
      }
      myData = tempData;
    }
    
    
    let dataRange = forActorSheet.getRange(forActorsTableName);
    let numItems = forActorSheet.getRange(numItemsActorsName);
    dataRange.clear("Contents");
    dataRange.load('rowCount');
    dataRange.load('rowIndex');
    dataRange.load('columnIndex');

    await excel.sync();
    
    let dataArray = [];
    // dataArray for putting in range [character, scene, lines, location]
    // lines are comma seperated for when both character and scene are the same
    // e.g ["RON", "123", "456, 789", "Hogwarts"]
    console.log('Start of loops', dataArray)
    for (i = 0; i < myData.length; i++){
      //Looping through the director data (myData)
      if (i < myData.length){
        //Does this character & scene already exist in data array
        let myIndex = dataArray.findIndex(x => (x[1] == myData[i].sceneNumber) && (x[0] == myData[i].character));
        //returns -1 if not found
        console.log('myIndex', myIndex)
        //Find the location of the scene.
        let theLocation = myLocation.find(x => x.sceneNumber == myData[i].sceneNumber)
        console.log('location', theLocation);
        if (myIndex == -1){
          // We need to add a new row. 
          console.log(i, "New Row")
          if (theLocation == null){
            thisRow = [myData[i].character, myData[i].sceneNumber, myData[i].lineNumber, ""];
          } else {
            thisRow = [myData[i].character, myData[i].sceneNumber, myData[i].lineNumber, theLocation.location];
          }
          console.log('thisRow', thisRow)
          let newIndex = dataArray.length;
          console.log('newIndex', newIndex);
          dataArray[newIndex] = thisRow;
          console.log('dataArray[newIndex]', dataArray[newIndex]);
        } else {
          //Having found a row we should add the line
          //However it is possible the same line number is present on multiple consecutive rows
          //This should not be added. Test for this
          if ((i > 0) && (myData[i - 1].lineNumber != myData[i].lineNumber)){
            console.log('Array before:', dataArray[myIndex]);
            //Add the lineNumber to array element 2 concatanated with a ", "
            dataArray[myIndex][2] = dataArray[myIndex][2] + ", " + myData[i].lineNumber;
            console.log("Found Index",  myIndex, "dataArray", dataArray[myIndex]);
          }
        }
        console.log("i", i, "dataArray", dataArray);
      } else {
        let thisRow = new Array(4).fill("");
        dataArray.push(thisRow);
        console.log("Empty i", i, "dataArray", dataArray);
      }
    }
    console.log('dataArray', dataArray, 'rowCount', dataRange.rowCount, 'dataLength', myData.length, 'dataArray.length', dataArray.length);
    if (dataArray.length > 0){
      let displayRange = forActorSheet.getRangeByIndexes(dataRange.rowIndex, dataRange.columnIndex, dataArray.length, 4);
      displayRange.values = dataArray;
    }
    numItems.values = dataArray.length;    
    await excel.sync();
    waitLabel.style.display = 'none';
    waitCell.values = '';
    await excel.sync();
  })
  await displayScenes();
}

async function getLocationInfo(){
  await Excel.run(async function(excel){
    let waitLabel = tag('location-wait');
    waitLabel.style.display = 'block';
    let locationSheet = excel.workbook.worksheets.getItem(locationSheetName);
    const waitCell = locationSheet.getRange('loMessage');
    waitCell.values = 'Please wait...';
    await excel.sync();
    
    let locationChoiceRange = locationSheet.getRange('loLocationChoice');
    locationChoiceRange.load('values');
    await excel.sync();
    let locationName = locationChoiceRange.values[0][0];
    console.log('Location Text ',locationName);
    let myData = await jade_modules.operations.getLocationData(locationName);
    console.log('Scheduling myData', myData);
    
    let dataRange = locationSheet.getRange(locationTableName);
    let numItems = locationSheet.getRange(numItemsLocationName);
    dataRange.clear("Contents");
    dataRange.load('rowCount');
    dataRange.load('rowIndex');
    dataRange.load('columnIndex');
    dataRange.load('columnCount')

    await excel.sync();
    
    let dataArray = [];
    console.log('Start of loops', dataArray)
    let sceneArray = myData.map(x => x.sceneNumber);
    let locationArray = myData.map(x => x.location);
    let lineArray = myData.map(x => x.lineNumber);
    console.log('Scene Number', sceneArray, 'location', locationArray, 'lineNo', lineArray);

    let characterData = await jade_modules.operations.gatherActorsforScene(sceneArray);
    console.log('Character Data', characterData);

    let result = []
    for (let i = 0; i < sceneArray.length; i++){
      result[i] = [sceneArray[i], locationArray[i], lineArray[i], characterData[i].characters.join(', ')]
    }

    console.log('Result', result);
    numItems.values = [[result.length]];

    if (result.length > 0){
      let tempRange = locationSheet.getRangeByIndexes(dataRange.rowIndex, dataRange.columnIndex, result.length, dataRange.columnCount);
      tempRange.values = result;
    }
    waitCell.values = '';
    waitLabel.style.display = 'none';
  })  
}

async function searchCharacter(){
  await getActorInformation();
}

async function getForSchedulingInfo(){
  await Excel.run(async function(excel){
    let waitLabel = tag('scheduling-wait');
    waitLabel.style.display = 'block';
    forSchedulingSheet = excel.workbook.worksheets.getItem(forSchedulingName);
    const waitCell = forSchedulingSheet.getRange('fsMessage');
    waitCell.values = 'Please wait...';
    await excel.sync();
    
    let character = await getActor(forSchedulingName);
    let myData = await jade_modules.operations.getDirectorDataV2(character);
    console.log('Scheduling myData', myData);
    
    let dataArray = [];
    let totalSceneWordCount = 0;
    let totalLineWordCount = 0;
    let sceneArray = [];
    let arrayIndex = -1;
    for (let i = 0; i < myData.length; i++){
      let newRow;
      if (myData[i].sceneWordCount == ''){
        myData[i].sceneWordCount = 0;
      }
      if (i == 0){
        newRow = {
          sceneNumber: myData[i].sceneNumber,
          sceneWordCount: myData[i].sceneWordCount,
          characterWordCount: myData[i].lineWordCount,
          characters: [myData[i].character]
        }
        dataArray.push(newRow);
        arrayIndex += 1;
        sceneArray[arrayIndex] = [];
        sceneArray[arrayIndex][0] = myData[i].sceneNumber;
        totalSceneWordCount += myData[i].sceneWordCount;
        totalLineWordCount += myData[i].lineWordCount;
      } else {
        if (myData[i - 1].lineNumber != myData[i].lineNumber){
          let myIndex = dataArray.findIndex(x => x.sceneNumber == myData[i].sceneNumber);
          if (myIndex == -1){
            newRow = {
              sceneNumber: myData[i].sceneNumber,
              sceneWordCount: myData[i].sceneWordCount,
              characterWordCount: myData[i].lineWordCount,
              characters: [myData[i].character]
            }
            dataArray.push(newRow);
            totalSceneWordCount += myData[i].sceneWordCount;
            console.log(i, 'totalscene', totalSceneWordCount, 'sceneWordCount', myData[i].sceneWordCount, 'sceneNo', myData[i].sceneNumber);
            totalLineWordCount += myData[i].lineWordCount;
          } else {
            dataArray[myIndex].characterWordCount = dataArray[myIndex].characterWordCount + myData[i].lineWordCount;
            totalLineWordCount += myData[i].lineWordCount;
            if (!(dataArray[myIndex].characters.some(x => x.toLowerCase() == myData[i].character.toLowerCase()))){
              dataArray[myIndex].characters.push(myData[i].character);
            }
          }
        }
      } 
    }
    
    for(let i = 0; i < dataArray.length; i++){
      sceneArray[i] = [];
      sceneArray[i][0] = dataArray[i].characters.join('|');
      sceneArray[i][1] = dataArray[i].sceneNumber;
    }
    console.log('dataArray', dataArray, 'totalScene', totalSceneWordCount, 'totalLine', totalLineWordCount, 'sceneNumbers', sceneArray);
    
    
    let dataRange = forSchedulingSheet.getRange(forSchedulingTableName);
    let numItems = forSchedulingSheet.getRange(numItemsSchedulingName);
    let linesUsedRange = forSchedulingSheet.getRange('fsLinesUsed')
    let fullScenesRange = forSchedulingSheet.getRange('fsFullScene')
    dataRange.clear("Contents");
    dataRange.load('rowCount');
    dataRange.load('rowIndex');
    dataRange.load('columnIndex');
    dataRange.load('columnCount');
    await excel.sync();

    if (sceneArray.length > 0){
      let displayRange = forSchedulingSheet.getRangeByIndexes(dataRange.rowIndex, dataRange.columnIndex, sceneArray.length, dataRange.columnCount);
      displayRange.values = sceneArray;
    }
    numItems.values = sceneArray.length;
    linesUsedRange.values = totalLineWordCount;
    fullScenesRange.values = totalSceneWordCount;
    await excel.sync();
    waitLabel.style.display = 'none';
    waitCell.values = '';
    await excel.sync();

  })
}
async function directorGoToLine(){
  await Excel.run(async function(excel){
    const forDirectorSheet = excel.workbook.worksheets.getItem(forDirectorName);
    const lineIndex = 2;
    let activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex');
    await excel.sync(); 
    let rowIndex = activeCell.rowIndex;
    if (rowIndex >= 10){
      let lineNumberCell = forDirectorSheet.getRangeByIndexes(rowIndex, lineIndex, 1, 1);
      lineNumberCell.load('values');
      await excel.sync(); 
      console.log('lineNumber', lineNumberCell.values);
      let lineNumber = parseInt(lineNumberCell.values[0][0])
      console.log('lineNumber', lineNumber);
      if (!isNaN(lineNumber)){
        await jade_modules.operations.findLineNo(lineNumber);
        activeCell = excel.workbook.getActiveCell();
        activeCell.load('rowIndex');
        await excel.sync(); 
        let rowIndex = activeCell.rowIndex;
        const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
        let columnIndex = await jade_modules.operations.findColumnIndex('Number');
        let tempRange = scriptSheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);
        tempRange.select();
        await excel.sync();
        await jade_modules.operations.showMainPage();
      } else {
        alert('Not a line number');
      }
    } else {
      alert('Must be in a line with valid line number');
    }
  })
}
async function actorGoToLine(){
  await Excel.run(async function(excel){
    const forActorSheet = excel.workbook.worksheets.getItem(forActorName);
    const lineIndex = 3;
    let activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex');
    await excel.sync(); 
    let rowIndex = activeCell.rowIndex;
    if (rowIndex >= 10){
      let lineNumberCell = forActorSheet.getRangeByIndexes(rowIndex, lineIndex, 1, 1);
      lineNumberCell.load('values');
      await excel.sync(); 
      console.log('lineNumber', lineNumberCell.values);
      let lineNumber = parseInt(lineNumberCell.values[0][0])
      console.log('lineNumber', lineNumber);
      if (!isNaN(lineNumber)){
        let rowIndex = await jade_modules.operations.findLineNo(lineNumber);
        const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
        let columnIndex = await jade_modules.operations.findColumnIndex('Number');
        let tempRange = scriptSheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);
        tempRange.select();
        await excel.sync();
        await jade_modules.operations.showMainPage();
      } else {
        alert('Not a line number');
      }
    } else {
      alert('Must be in a line with valid line number');
    }
  })
}

async function schedulingGoToLine(){
  await Excel.run(async function(excel){
    const forSchedulingSheet = excel.workbook.worksheets.getItem(forSchedulingName);
    const sceneIndex = 2;
    let activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex');
    await excel.sync(); 
    let rowIndex = activeCell.rowIndex;
    if (rowIndex >= 10){
      let sceneNumberCell = forSchedulingSheet.getRangeByIndexes(rowIndex, sceneIndex, 1, 1);
      sceneNumberCell.load('values');
      await excel.sync(); 
      console.log('sceneNumber', sceneNumberCell.values);
      let sceneNumber = parseInt(sceneNumberCell.values[0][0])
      console.log('sceneNumber', sceneNumber);
      if (!isNaN(sceneNumber)){
        await jade_modules.operations.findSceneNo(sceneNumber);
        activeCell = excel.workbook.getActiveCell();
        activeCell.load('rowIndex');
        await excel.sync(); 
        let rowIndex = activeCell.rowIndex;
        const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
        let columnIndex = await jade_modules.operations.findColumnIndex('Scene Number');
        let tempRange = scriptSheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);
        tempRange.select();
        await excel.sync();
        await jade_modules.operations.showMainPage();
      } else {
        alert('Not a scene number');
      }
    } else {
      alert('Must be in a line with valid scene number');
    }
  })
}

async function locationGoToLine(){
  await Excel.run(async function(excel){
    const locationSheet = excel.workbook.worksheets.getItem(locationSheetName);
    const sceneIndex = 1;
    let activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex');
    await excel.sync(); 
    let rowIndex = activeCell.rowIndex;
    if (rowIndex >= 10){
      let sceneNumberCell = locationSheet.getRangeByIndexes(rowIndex, sceneIndex, 1, 1);
      sceneNumberCell.load('values');
      await excel.sync(); 
      console.log('sceneNumber', sceneNumberCell.values);
      let sceneNumber = parseInt(sceneNumberCell.values[0][0])
      console.log('sceneNumber', sceneNumber);
      if (!isNaN(sceneNumber)){
        await jade_modules.operations.findSceneNo(sceneNumber);
        activeCell = excel.workbook.getActiveCell();
        activeCell.load('rowIndex');
        await excel.sync(); 
        let rowIndex = activeCell.rowIndex;
        const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
        let columnIndex = await jade_modules.operations.findColumnIndex('Scene Number');
        let tempRange = scriptSheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);
        tempRange.select();
        await excel.sync();
        await jade_modules.operations.showMainPage();
      } else {
        alert('Not a scene number');
      }
    } else {
      alert('Must be in a line with valid scene number');
    }
  })
}

async function isUsOnly(){
  let usOnly
  await Excel.run(async function(excel){
    let forActorSheet = excel.workbook.worksheets.getItem(forActorName);
    let selectRange = forActorSheet.getRange('faSelect');
    selectRange.load('values');
    await excel.sync();
    usOnly = selectRange.values[0][0] != 'All';
  })
  return usOnly;
}

async function createScript(sheetName = 'Actor Script', isMultiScript = false, multiScriptDetails = null, scenesDone = 0, totalNumScenes = 0){
  let startTime = new Date().getTime();
  let actorWait = tag('script-wait');
  actorWait.style.display = 'block';
  let isAllNaN = true;
  let sceneNumbers, usOnly
  let message = tag('multi-message');
  let baseMessage = message.innerText + '. ';

  if (isMultiScript){
    sceneNumbers = multiScriptDetails.scenes;
    usOnly = multiScriptDetails.allUs == 'US Only';
  } else {
    sceneNumbers = await getSceneNumberActor();
    usOnly = await isUsOnly();
  }
  
  if (sceneNumbers.scenes.length > 0){
    let book = await jade_modules.operations.getBook();
    let character = {};
    if (isMultiScript){
      character.name = multiScriptDetails.character;
      character.type = multiScriptDetails.type;
      sheetName = multiScriptDetails.sheetName
    } else {
      character = await getActor(forActorName);
    }
    await clearActorScriptBody(sheetName);    
    await topOfFirstPage(book, character, usOnly, sheetName);  
  
    let theRowIndex = 1;
    let rowIndexes;
    let sceneBlockRows = await jade_modules.operations.getSceneBlockRows();
    for (let i = 0; i < sceneNumbers.scenes.length; i++){
      actorWait.innerText = 'Please wait... Doing scene: ' + sceneNumbers.scenes[i] + ' (' + (i + 1) + ' of ' + sceneNumbers.scenes.length + ')';
      message.innerText = baseMessage + 'Scene: ' + (i + 1 + scenesDone) + ' of ' + totalNumScenes;
      let sceneNumber = sceneNumbers.scenes[i]
      if (!isNaN(sceneNumber)){
        isAllNaN = false;
        let indexes = await jade_modules.operations.getRowIndeciesForScene(sceneNumber, usOnly);
        console.log('Indexes: ', indexes);
        if (indexes.length > 0){
          //let sceneBlockText = await jade_modules.operations.getSceneBlockNear(indexes[0]);
          let sceneBlockText = await jade_modules.operations.getSceneBlockText(sceneNumber, sceneBlockRows)
          let doPageBreak = i > 0;
          let rowDetails = await putDataInActorScriptSheet(sceneBlockText, theRowIndex, doPageBreak, sheetName);
          if (rowDetails.sceneBlockRowIndexes.length == 0){
            console.log('Missing scene block (from rowDetails): ', sceneNumber)
          }
          //give 1 row of scpace between sceneblock and script
          theRowIndex = rowDetails.nextRowIndex + 1;
          rowIndexes = await jade_modules.operations.getActorScriptRanges(indexes, theRowIndex, usOnly, sheetName);
          await formatActorScript(sheetName, rowDetails.sceneBlockRowIndexes, rowIndexes, character.name, usOnly);
          theRowIndex = rowIndexes[rowIndexes.length - 1].startRow + rowIndexes[rowIndexes.length - 1].rowCount + 1;
        } else {
          console.log('Missing scene block: ', sceneNumber)
        }
      }
    }
    await makeChaptersBlackFont(sheetName, theRowIndex);
    if (isAllNaN){
      alert('Please select a scene')
    } else {
      if (!isMultiScript){
        await jade_modules.operations.showActorScript()
      }
    }
  } else {
    alert('Please select a scene')
  }
  actorWait.innerText = 'Please wait...'
  actorWait.style.display = 'none';
  let endTime = new Date().getTime();
  console.log('Complete Create Script taken:', (endTime - startTime) / 1000);
  return sceneNumbers.scenes.length + scenesDone;
}

async function displayScenes(){
  console.log('In displayScenes');
  let theScenes = await getSceneNumberActor();
  console.log('display', theScenes.display);
  let scenesDisplay = tag('actor-scene-display');
  scenesDisplay.innerText = theScenes.display;  
}

function setDefaultIfNotSet(){
  if (getActorScriptChoice() == scriptChoice.none){
    const radioAll = tag('radAllScenes');
    radioAll.checked = true; 
  }
}

function getActorScriptChoice(){
  const radioAll = tag('radAllScenes');
  const radioHighlighted = tag('radHighlighted');
  
  if (radioAll.checked){
    return scriptChoice.all;
  } else if (radioHighlighted.checked) {
    return scriptChoice.selection;
  } else {
    return scriptChoice.none;
  }
}

async function getSceneNumberActor(){
  let sceneNumbers = {};
  let theChoice = getActorScriptChoice();
  
  await Excel.run(async function(excel){
    const forActorSheet = excel.workbook.worksheets.getItem(forActorName);
    const sceneColumnIndex = 2;
    let myAddresses = [];
    let display = ''
    if (theChoice == scriptChoice.all){
      let tableRange = forActorSheet.getRange('faTable');
      tableRange.load('address');
      await excel.sync();
      myAddresses[0] = tableRange.address
      display = 'All'
    } else if (theChoice == scriptChoice.selection){
      let selectedRanges = excel.workbook.getSelectedRanges();
      selectedRanges.load('address');
      await excel.sync();
      myAddresses = selectedRanges.address.split(',')  
    } else {
      return null;
    }
        
    let tempRange = []
    for (let i = 0; i < myAddresses.length; i++){
      tempRange[i] = forActorSheet.getRange(myAddresses[i]);
      tempRange[i].load('rowIndex, rowCount')
    }
    await excel.sync();
    let sceneRange = [];
    let sceneIndex = -1;
    for (let i = 0; i < myAddresses.length; i++){
      for (let row = 0; row < tempRange[i].rowCount; row++){
        let thisRow = row + tempRange[i].rowIndex;
        if (thisRow > 9){
          sceneIndex += 1;
          sceneRange[sceneIndex] = forActorSheet.getRangeByIndexes(thisRow, sceneColumnIndex, 1, 1)
          sceneRange[sceneIndex].load('values');
        }
      }
    }
    await excel.sync();
    let result = []
    let resultIndex = -1;
    for (let i = 0; i < sceneRange.length; i++){
      let thisNumber = parseInt(sceneRange[i].values[0][0]);
      if (!isNaN(thisNumber)){
        if (!result.includes(thisNumber)){
          resultIndex += 1;
          result[resultIndex] = thisNumber;
        }
      }
    }
    if (theChoice == scriptChoice.selection){
      if (result.length == 0){
        display = 'Nothing selected'
      } else {
        display = result.join(', ');
      }
    } else if (theChoice == scriptChoice.all){
      if (result.length > 10){
        display = 'All (' + result.slice(0, 3).join(', ') + ' ... ' + result.slice(-2).join(', ') + ')';
      } else {
        display = 'All (' + result.join(', ') + ')';
      }
      
    }
    console.log('The results', result);
    sceneNumbers = {scenes: result, display: display}
  }).catch(e => console.log('My error', e));
  return sceneNumbers;
}


async function topOfFirstPage(book, character, usOnly, sheetName){
  await Excel.run(async function(excel){
    const actorScriptSheet = excel.workbook.worksheets.getItem(sheetName);
    let bookRange = actorScriptSheet.getRange(actorScriptBookName);
    bookRange.values = book;
    let headingRange = actorScriptSheet.getRange(actorScriptCharcaterHeadingName);
    if (character.type == choiceType.list){
      headingRange.values = [['Character: ']]
    } else {
      headingRange.values = [['Character: (Text Search)']]
    }
    let characterRange = actorScriptSheet.getRange(actorScriptCharacterName);
    let numColumns = 2;
    if (usOnly){
      characterRange.values = character.name.charAt(0).toUpperCase() + character.name.slice(1) + ' (US Script)';
      numColumns = 3;
    } else {
      characterRange.values = character.name.charAt(0).toUpperCase() + character.name.slice(1);
    }
    characterRange.unmerge()
    characterRange.load('rowIndex, columnIndex')
    
    await excel.sync();
    let mergeRange = actorScriptSheet.getRangeByIndexes(characterRange.rowIndex, characterRange.columnIndex, 1, numColumns);
    mergeRange.merge(true);
  })
}
async function clearActorScriptBody(sheetName){
  await Excel.run(async function(excel){
    const actorScriptSheet = excel.workbook.worksheets.getItem(sheetName);
    actorScriptSheet.horizontalPageBreaks.removePageBreaks();
  
    let tableRange = actorScriptSheet.getRange(actorScriptTableName);
    tableRange.clear("Contents");
    tableRange.clear("Formats");
  })
}
async function putDataInActorScriptSheet(sceneBlock, startRowIndex, doPageBreak, sheetName){
  let rowDetails = {};
  await Excel.run(async function(excel){
    const actorScriptSheet = excel.workbook.worksheets.getItem(sheetName);
    let sceneBlockColumnIndex = 0;
    console.log(startRowIndex, sceneBlockColumnIndex, sceneBlock.length, 1)
    console.log(sceneBlock);
    let sceneBlockIndexes = [];
    if (sceneBlock.length > 0){
      let temp = [];
      for (let i = 0; i < sceneBlock.length; i++){
        temp[i] = [sceneBlock[i]];
      }
      console.log('temp:', temp);
      let range = actorScriptSheet.getRangeByIndexes(startRowIndex, sceneBlockColumnIndex, sceneBlock.length, 1);
      if (doPageBreak){
        actorScriptSheet.horizontalPageBreaks.add(range)
      };
      range.values = temp;
      await excel.sync();
      
      for (let i = 0; i < sceneBlock.length; i++){
        sceneBlockIndexes[i] = startRowIndex + i;
      }
    }
    let nextRowIndex = startRowIndex + sceneBlock.length
    rowDetails = { nextRowIndex: nextRowIndex, sceneBlockRowIndexes: sceneBlockIndexes}
  })
  return rowDetails;
}

async function getActor(sheetName){
  let character = {};
  let choiceRangeName, characterListRangeName, characterTextRangeName;
  console.log('sheetName', sheetName, forActorName, forSchedulingName);
  if (sheetName == forActorName){
    choiceRangeName = 'faChoice';
    characterListRangeName = 'faCharacterChoice';
    characterTextRangeName = 'faTextSearch';
  } else if (sheetName == forSchedulingName){
    choiceRangeName = 'fsChoice'
    characterListRangeName = 'fsCharacterChoice';
    characterTextRangeName = 'fsTextSearch';
  }
  console.log('range names: ', choiceRangeName);
  await Excel.run(async function(excel){
    let theSheet = excel.workbook.worksheets.getItem(sheetName);
    let choiceRange = theSheet.getRange(choiceRangeName);
    choiceRange.load('values')
    await excel.sync();
    let characterChoiceRange, myChoiceType;
    console.log(choiceRange.values[0][0]);
    if (choiceRange.values[0][0] == choiceType.list){
      characterChoiceRange = theSheet.getRange(characterListRangeName);
      myChoiceType = choiceType.list;
    } else if (choiceRange.values[0][0] == choiceType.text){
      characterChoiceRange = theSheet.getRange(characterTextRangeName);
      myChoiceType = choiceType.text;
    }
    characterChoiceRange.load('values');
    await excel.sync();
    character.name = characterChoiceRange.values[0][0];
    character.type = myChoiceType;
    console.log('Character ',character);
  })
  return character;
}

async function showActorScript(){
  await Excel.run(async function(excel){
    let actorScriptSheet = excel.workbook.worksheets.getItem(actorScriptName);
    actorScriptSheet.activate();
  })
}

async function formatActorScript(sheetName, sceneBlockRowIndexes, scriptRowIndexes, character, usOnly){
  await removeBorders(sheetName);
  await formatSceneBlocks(sheetName, sceneBlockRowIndexes, usOnly);
  await formatHeading(sheetName);
  for (let i = 0; i < scriptRowIndexes.length; i++){
    await cueColumnFontColour(sheetName, scriptRowIndexes[i]);
    await clearScriptFill(sheetName,scriptRowIndexes[i]);
    //await clearStrikethrough(sheetName, scriptRowIndexes[i]);
    await highlightCharacters(sheetName, character, scriptRowIndexes[i]);
  }
}

async function removeBorders(sheetName){
  await Excel.run(async function(excel){
    let theSheet = excel.workbook.worksheets.getItem(sheetName);
    let usedRange = theSheet.getUsedRange();
    let theBorders = usedRange.format.borders;
    theBorders.getItem("EdgeTop").style = "None";
    theBorders.getItem("EdgeBottom").style = "None";
    theBorders.getItem("EdgeLeft").style = "None";
    theBorders.getItem("EdgeRight").style = "None";
    theBorders.getItem("InsideVertical").style = "None";
    theBorders.getItem("InsideHorizontal").style = "None";
    theBorders.getItem("DiagonalDown").style = "None";
    theBorders.getItem("DiagonalUp").style = "None";
  })
}

async function formatSceneBlocks(sheetName, rowIndexes, usOnly){
  let firstColumn = 0;
  let columnCount = 4;
  if (usOnly){
    columnCount = 5;
  }

  for (let i = 0; i < rowIndexes.length; i++){
    await mergeTheRow(sheetName, rowIndexes[i], 1, firstColumn, columnCount);
  }

  await Excel.run(async function(excel){
    let theSheet = excel.workbook.worksheets.getItem(sheetName);
    for (let i = 0; i < rowIndexes.length; i++){
      let theRange = theSheet.getRangeByIndexes(rowIndexes[i], firstColumn, 1, columnCount);
      theRange.format.font.name = 'Courier New';
      theRange.format.font.size = 12;
      theRange.format.font.bold = true;
      theRange.format.fill.color = myFormats.purple;
      theRange.format.horizontalAlignment = 'Center';
      theRange.format.verticalAlignment = 'Top';
      await jade_modules.operations.mergedRowAutoHeight(excel, theSheet, theRange);
    }
  })
  
}
async function mergeTheRow(sheetName, rowIndex, rowCount, firstColumnIndex, columnCount){
  await Excel.run(async function(excel){
    let theSheet = excel.workbook.worksheets.getItem(sheetName);
    let myMergeRange = theSheet.getRangeByIndexes(rowIndex, firstColumnIndex, rowCount, columnCount);
    myMergeRange.load('address');
    let mergedAreas = myMergeRange.getMergedAreasOrNullObject();
    mergedAreas.load("cellCount");
    await excel.sync();
    if (!(mergedAreas.cellCount == (rowCount * columnCount))){
      myMergeRange.merge(true);
    }
  })
}
async function formatHeading(sheetName){
  let rowIndex = 0
  let firstColumn = 0;
  let columnCount = 3;
  await Excel.run(async function(excel){
    let theSheet = excel.workbook.worksheets.getItem(sheetName);
    let theRange = theSheet.getRangeByIndexes(rowIndex, firstColumn, 1, columnCount);
    theRange.format.font.name = 'Courier New';
    theRange.format.font.size = 14;
    theRange.format.font.bold = true;
    theRange.format.fill.color = myFormats.green;
    theRange.format.horizontalAlignment = 'Left';
    theRange.format.verticalAlignment = 'Top';
  })
}

async function cueColumnFontColour(sheetName, rowDetails){
  await Excel.run(async function(excel){
    let cueColumnIndex = 0;
    let theSheet = excel.workbook.worksheets.getItem(sheetName);
    let theRange = theSheet.getRangeByIndexes(rowDetails.startRow, cueColumnIndex, rowDetails.rowCount, 1);
    theRange.format.font.color = myFormats.lightGrey;
  })
}

async function makeChaptersBlackFont(sheetName, lastRowIndex){
  await Excel.run(async function(excel){
    let stageDirectionColumnIndex = 2;
    let firstRowIndex = 5
    let rowCount = lastRowIndex - firstRowIndex + 1;
    let theSheet = excel.workbook.worksheets.getItem(sheetName);
    let myRange = theSheet.getRangeByIndexes(firstRowIndex, stageDirectionColumnIndex, rowCount, 1);
    myRange.load('values, rowIndex');
    await excel.sync();
    let values = myRange.values.map(x => x[0]);
    console.log('Values:', values, 'RowIndex', myRange.rowIndex, 'lastRowIndex', lastRowIndex)
    for (let i = 0; i < values.length; i++){
      if (values[i].toLowerCase().includes('chapter') && values[i].toLowerCase().includes('book')){
        let tempRowIndex = i + myRange.rowIndex;
        console.log('Found: ', i, 'rowIndex: ', tempRowIndex)
        let tempRange = theSheet.getRangeByIndexes(tempRowIndex, stageDirectionColumnIndex, 1, 1);
        tempRange.format.font.color = '#000000';
        await excel.sync();
      }
    }
  })
}

async function clearScriptFill(sheetName, rowDetails){
  await Excel.run(async function(excel){
    let cueColumnIndex = 0;
    let columnCount = 4;
    let theSheet = excel.workbook.worksheets.getItem(sheetName);
    let theRange = theSheet.getRangeByIndexes(rowDetails.startRow, cueColumnIndex, rowDetails.rowCount, columnCount);
    theRange.format.fill.clear();
  })
}

async function clearStrikethrough(sheetName, rowDetails){
  await Excel.run(async function(excel){
    let cueColumnIndex = 0;
    let columnCount = 4;
    let theSheet = excel.workbook.worksheets.getItem(sheetName);
    let theRange = theSheet.getRangeByIndexes(rowDetails.startRow, cueColumnIndex, rowDetails.rowCount, columnCount);
    theRange.format.font.strikethrough = false;
  })
}

async function highlightCharacters(sheetName, character, rowDetails){
  let characterColumnIndex = 1;
  await Excel.run(async (excel) => {
    let theSheet = excel.workbook.worksheets.getItem(sheetName);
    let theRange = theSheet.getRangeByIndexes(rowDetails.startRow, characterColumnIndex, rowDetails.rowCount, 1);
    let conditionalFormat = theRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
    
    conditionalFormat.textComparison.format.fill.color = myFormats.orange;
    conditionalFormat.textComparison.rule = {
      operator: Excel.ConditionalTextOperator.contains,
      text: character
    };
    
    await excel.sync();
  });
}

async function processCharacterListForWordAndScene(){
  let myWait = tag('character-wait');
  myWait.style.display = 'block'
  await Excel.run(async function(excel){
    const characterListSheet = excel.workbook.worksheets.getItem(characterListName);
    let characterRange = characterListSheet.getRange('clCharacters');
    let detailsRange = characterListSheet.getRange('clTable')
    characterRange.load('values, rowIndex');
    detailsRange.load('columnIndex, columnCount')
    detailsRange.clear("Contents");
    await excel.sync();
    let myCharacters = characterRange.values.map(x => x[0]);
    console.log('Characters: ', myCharacters, 'rowIndex: ', characterRange.rowIndex )
    for (let i = 0; i < myCharacters.length; i ++){
      if (myCharacters[i] != ''){
        let character = {name: myCharacters[i], type: choiceType.list};
        let details = await getWordCountForCharacter(character);
        let tempRange = characterListSheet.getRangeByIndexes(i + characterRange.rowIndex, detailsRange.columnIndex, 1, detailsRange.columnCount);
        tempRange.values = [[details.sceneWordCount, details.lineWordCount, details.scenes]];
      }
    }
  })
  myWait.style.display = 'none'
}

async function getWordCountForCharacter(characterName){
  let myData = await jade_modules.operations.getDirectorDataV2(characterName);
  console.log('Scheduling myData', myData);
    
  let dataArray = [];
  let totalSceneWordCount = 0;
  let totalLineWordCount = 0;
  let sceneArray = [];
  let arrayIndex = -1;
  for (let i = 0; i < myData.length; i++){
    let newRow;
    if (myData[i].sceneWordCount == ''){
      myData[i].sceneWordCount = 0;
    }
    if (i == 0){
      newRow = {
        sceneNumber: myData[i].sceneNumber,
        sceneWordCount: myData[i].sceneWordCount,
        characterWordCount: myData[i].lineWordCount
      }
      dataArray.push(newRow);
      arrayIndex += 1;
      sceneArray[arrayIndex] = [];
      sceneArray[arrayIndex][0] = myData[i].sceneNumber;
      totalSceneWordCount += myData[i].sceneWordCount;
      totalLineWordCount += myData[i].lineWordCount;
    } else {
      if (myData[i - 1].lineNumber != myData[i].lineNumber){
        let myIndex = dataArray.findIndex(x => x.sceneNumber == myData[i].sceneNumber);
        if (myIndex == -1){
          newRow = {
            sceneNumber: myData[i].sceneNumber,
            sceneWordCount: myData[i].sceneWordCount,
            characterWordCount: myData[i].lineWordCount
          }
          dataArray.push(newRow);
          arrayIndex += 1;
          sceneArray[arrayIndex] = [];
          sceneArray[arrayIndex][0] = myData[i].sceneNumber;
          totalSceneWordCount += myData[i].sceneWordCount;
          console.log(i, 'totalscene', totalSceneWordCount, 'sceneWordCount', myData[i].sceneWordCount, 'sceneNo', myData[i].sceneNumber);
          totalLineWordCount += myData[i].lineWordCount;
        } else {
          dataArray[myIndex].characterWordCount = dataArray[myIndex].characterWordCount + myData[i].lineWordCount;
          totalLineWordCount += myData[i].lineWordCount;
        }
      }
    } 
  }

  return {
    sceneWordCount: totalSceneWordCount,
    lineWordCount: totalLineWordCount,
    scenes: sceneArray.map(x => x[0]).join(', ')
  }
}

async function createSceneWordCountData(){
  let myWait = tag('character-wait');
  myWait.style.display = 'block';
  await Excel.run(async function(excel){
    console.log('Starting');
    const characterListSheet = excel.workbook.worksheets.getItem(characterListName);
    let sceneWordCountRange = characterListSheet.getRange('clSceneWordCount');
    sceneWordCountRange.clear("Contents") ;
    sceneWordCountRange.load('rowIndex, columnIndex, columnCount')
    await excel.sync();
    
    let countDetails = await jade_modules.operations.getSceneWordCount();
    let display = []
    for (let i = 0; i < countDetails.length; i++){
      display[i] = [ countDetails[i].scene, countDetails[i].wordCount];
    }
    console.log(sceneWordCountRange.rowIndex, sceneWordCountRange.columnIndex, display.length, sceneWordCountRange.columnCount)
    let displayRange = characterListSheet.getRangeByIndexes(sceneWordCountRange.rowIndex, sceneWordCountRange.columnIndex, display.length, sceneWordCountRange.columnCount);
    displayRange.values = display;
  })
  myWait.style.display = 'none';
}
