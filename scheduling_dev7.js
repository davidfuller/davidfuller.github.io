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

function auto_exec(){
}
async function loadReduceAndSortCharacters(){
  await Excel.run(async function(excel){ 
    characterlistSheet = excel.workbook.worksheets.getItem(characterListName);
    let characters = await jade_modules.operations.getCharacters();
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
    let myData = await jade_modules.operations.getDirectorData(characterName);
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

async function getActorInfo(){
  await getActorInformation(true);
}

async function getActorInformation(asDropdown){
  await Excel.run(async function(excel){
    let waitLabel = tag('actor-wait');
    waitLabel.style.display = 'block';
    let forActorSheet = excel.workbook.worksheets.getItem(forActorName);
    const waitCell = forActorSheet.getRange('faMessage');
    waitCell.values = 'Please wait...';
    await excel.sync();
    
    let characterName = await getActor(asDropdown);
    console.log('Character ',characterName);
    let myData = await jade_modules.operations.getDirectorData(characterName, asDropdown);
    let myLocation = await jade_modules.operations.getLocations();
    console.log('Scheduling myData', myData);
    console.log('Locations', myLocation);
    
    let dataRange = forActorSheet.getRange(forActorsTableName);
    let numItems = forActorSheet.getRange(numItemsActorsName);
    dataRange.clear("Contents");
    dataRange.load('rowCount');
    dataRange.load('rowIndex');
    dataRange.load('columnIndex');

    await excel.sync();
    
    let dataArray = [];
    console.log('Start of loops', dataArray)
    for (i = 0; i < myData.length; i++){
      if (i < myData.length){
        let myIndex = dataArray.findIndex(x => x[1] == myData[i].sceneNumber)
        console.log('myIndex', myIndex)
        let theLocation = myLocation.find(x => x.sceneNumber == myData[i].sceneNumber)
        console.log('location', theLocation);
        if (myIndex == -1){
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
          if ((i > 0) && (myData[i - 1].lineNumber != myData[i].lineNumber)){
            console.log('Array before:', dataArray[myIndex]);
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

async function getActorText(){
  await getActorInformation(false);
}

async function getForSchedulingInfo(){
  await Excel.run(async function(excel){
    let waitLabel = tag('scheduling-wait');
    waitLabel.style.display = 'block';
    forSchedulingSheet = excel.workbook.worksheets.getItem(forSchedulingName);
    const waitCell = forSchedulingSheet.getRange('fsMessage');
    waitCell.values = 'Please wait...';
    await excel.sync();
    
    let characterChoiceRange = forSchedulingSheet.getRange('fsCharacterChoice');
    characterChoiceRange.load('values');
    await excel.sync();
    let characterName = characterChoiceRange.values[0][0];
    console.log('Character ',characterName);
    let myData = await jade_modules.operations.getDirectorData(characterName);
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
    console.log('dataArray', dataArray, 'totalScene', totalSceneWordCount, 'totalLine', totalLineWordCount, 'sceneNumbers', sceneArray);
    let dataRange = forSchedulingSheet.getRange(forSchedulingTableName);
    let numItems = forSchedulingSheet.getRange(numItemsSchedulingName);
    let linesUsedRange = forSchedulingSheet.getRange('fsLinesUsed')
    let fullScenesRange = forSchedulingSheet.getRange('fsFullScene')
    dataRange.clear("Contents");
    dataRange.load('rowCount');
    dataRange.load('rowIndex');
    dataRange.load('columnIndex');
    await excel.sync();

    if (sceneArray.length > 0){
      let displayRange = forSchedulingSheet.getRangeByIndexes(dataRange.rowIndex, dataRange.columnIndex, sceneArray.length, 1);
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

async function createScript(){
  let actorWait = tag('actor-wait');
  actorWait.style.display = 'block';
  let sceneNumber = await getSceneNumberActor();
  if (!isNaN(sceneNumber)){
    let indexes = await jade_modules.operations.getRowIndeciesForScene(sceneNumber);
    console.log('Indexes: ', indexes);
    let book = await jade_modules.operations.getBook();
    let sceneBlockText = await jade_modules.operations.getSceneBlockNear(indexes[0]);
    let characterName = await getActor();
    let rowDetails = await putDataInActorScriptSheet(book, characterName, sceneBlockText);
    //give 1 row of scpace between sceneblock and script
    rowDetails[0].nextRowIndex += 1;
    let rowIndexes = await jade_modules.operations.getActorScriptRanges(indexes, rowDetails[0].nextRowIndex);
    await formatActorScript(actorScriptName, rowDetails[0].sceneBlockRowIndexes, rowIndexes, characterName);
    await showActorScript();
  } else {
    alert('Please select a scene')
  }
  actorWait.style.display = 'none';
}

async function getSceneNumberActor(){
  let sceneNumber;
  await Excel.run(async function(excel){
    const forActorSheet = excel.workbook.worksheets.getItem(forActorName);
    const sceneIndex = 1;
    let activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex');
    await excel.sync(); 
    let rowIndex = activeCell.rowIndex;
    if (rowIndex >= 10){
      let sceneCell = forActorSheet.getRangeByIndexes(rowIndex, sceneIndex, 1, 1);
      sceneCell.load('values');
      await excel.sync(); 
      console.log('scene', sceneCell.values);
      sceneNumber = parseInt(sceneCell.values[0][0])
      console.log('sceneNumber', sceneNumber);
    }
  })
  return sceneNumber;
}

async function putDataInActorScriptSheet(book, character, sceneBlock){
  let rowDetails = [];
  await Excel.run(async function(excel){
    const actorScriptSheet = excel.workbook.worksheets.getItem(actorScriptName);
    let bookRange = actorScriptSheet.getRange(actorScriptBookName);
    bookRange.values = book;
    let headingRange = actorScriptSheet.getRange(actorScriptCharcaterHeadingName);
    headingRange.values = [['Character: ']]
    let characterRange = actorScriptSheet.getRange(actorScriptCharacterName);
    characterRange.values = character;
    characterRange.unmerge()
    characterRange.load('rowIndex, columnIndex')
    await excel.sync();
    let mergeRange = actorScriptSheet.getRangeByIndexes(characterRange.rowIndex, characterRange.columnIndex, 1, 2);
    mergeRange.merge(true);
    let tableRange = actorScriptSheet.getRange(actorScriptTableName);
    tableRange.clear("Contents");
    tableRange.clear("Formats");
    let startRowIndex = 1
    let sceneBlockColumnIndex = 0;
    console.log(startRowIndex,sceneBlockColumnIndex, sceneBlock.length, 1)
    console.log(sceneBlock);
    let temp = [];
    for (let i = 0; i < sceneBlock.length; i++){
      temp[i] = [sceneBlock[i]];
    }
    console.log('temp:', temp);
    let range = actorScriptSheet.getRangeByIndexes(startRowIndex, sceneBlockColumnIndex, sceneBlock.length, 1);
    range.values = temp;
    await excel.sync();
    
    let sceneBlockIndexes = [];
    for (let i = 0; i < sceneBlock.length; i++){
      sceneBlockIndexes[i] = startRowIndex + i;
    }
    let nextRowIndex = startRowIndex + sceneBlock.length
    rowDetails.push(
      {
        nextRowIndex: nextRowIndex,
        sceneBlockRowIndexes: sceneBlockIndexes
      }
    )
  })
  return rowDetails;
}

async function getActor(asDropDown){
  let characterName;
  await Excel.run(async function(excel){
    let forActorSheet = excel.workbook.worksheets.getItem(forActorName);
    let characterChoiceRange
    if (asDropDown){
      characterChoiceRange = forActorSheet.getRange('faCharacterChoice');
    } else {
      characterChoiceRange = forActorSheet.getRange('faTextSearch');
    }
    characterChoiceRange.load('values');
    await excel.sync();
    characterName = characterChoiceRange.values[0][0];
    console.log('Character ',characterName);
  })
  return characterName;
}

async function showActorScript(){
  await Excel.run(async function(excel){
    let actorScriptSheet = excel.workbook.worksheets.getItem(actorScriptName);
    actorScriptSheet.activate();
  })
}

async function formatActorScript(sheetName, sceneBlockRowIndexes, scriptRowIndexes, character){
  await removeBorders(sheetName);
  await formatSceneBlocks(sheetName, sceneBlockRowIndexes);
  await formatHeading(sheetName);
  for (let i = 0; i < scriptRowIndexes.length; i++){
    await cueColumnFontColour(sheetName, scriptRowIndexes[i]);
    await clearScriptFill(sheetName,scriptRowIndexes[i]);
    await clearStrikethrough(sheetName, scriptRowIndexes[i]);
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

async function formatSceneBlocks(sheetName, rowIndexes){
  let firstColumn = 0;
  let columnCount = 4;

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