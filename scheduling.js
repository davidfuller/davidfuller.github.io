let characterlistSheet, forDirectorSheet, forActorSheet, forSchedulingSheet;
const characterListName = 'Character List';
const scriptSheetName = 'Script'
const forDirectorName = 'For Directors';
const forActorName = 'For Actors'
const forSchedulingName = 'For Scheduling'
const forDirectorTableName = 'fdTable';
const forActorsTableName = "faTable";
const forSchedulingTableName = 'fsTable'
const numItemsDirectorsName = 'fdItems';
const numItemsActorsName = 'faItems';
const numItemsSchedulingName = 'fsItems';

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
  await Excel.run(async function(excel){
    let waitLabel = tag('actor-wait');
    waitLabel.style.display = 'block';
    forActorSheet = excel.workbook.worksheets.getItem(forActorName);
    const waitCell = forActorSheet.getRange('faMessage');
    waitCell.values = 'Please wait...';
    await excel.sync();
    
    let characterChoiceRange = forActorSheet.getRange('faCharacterChoice');
    characterChoiceRange.load('values');
    await excel.sync();
    let characterName = characterChoiceRange.values[0][0];
    console.log('Character ',characterName);
    let myData = await jade_modules.operations.getDirectorData(characterName);
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
        let myIndex = dataArray.findIndex(x => x[0] == myData[i].sceneNumber)
        console.log('myIndex', myIndex)
        let theLocation = myLocation.find(x => x.sceneNumber == myData[i].sceneNumber)
        console.log('location', theLocation);
        if (myIndex == -1){
          console.log(i, "New Row")
          if (theLocation == null){
            thisRow = [myData[i].sceneNumber, myData[i].lineNumber, ""];
          } else {
            thisRow = [myData[i].sceneNumber, myData[i].lineNumber, theLocation.location];
          }
          console.log('thisRow', thisRow)
          let newIndex = dataArray.length;
          console.log('newIndex', newIndex);
          dataArray[newIndex] = thisRow;
          console.log('dataArray[newIndex]', dataArray[newIndex]);
        } else {
          if ((i > 0) && (myData[i - 1].lineNumber != myData[i].lineNumber)){
            console.log('Array before:', dataArray[myIndex]);
            dataArray[myIndex][1] = dataArray[myIndex][1] + ", " + myData[i].lineNumber;
            console.log("Found Index",  myIndex, "dataArray", dataArray[myIndex]);
          }
        }
        console.log("i", i, "dataArray", dataArray);
      } else {
        let thisRow = new Array(3).fill("");
        dataArray.push(thisRow);
        console.log("Empty i", i, "dataArray", dataArray);
      }
    }
    console.log('dataArray', dataArray, 'rowCount', dataRange.rowCount, 'dataLength', myData.length, 'dataArray.length', dataArray.length);
    let displayRange = forActorSheet.getRangeByIndexes(dataRange.rowIndex, dataRange.columnIndex, dataArray.length, 3);
    displayRange.values = dataArray;
    numItems.values = dataArray.length;    
    await excel.sync();
    waitLabel.style.display = 'none';
    waitCell.values = '';
    await excel.sync();
  })  
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

    let displayRange = forSchedulingSheet.getRangeByIndexes(dataRange.rowIndex, dataRange.columnIndex, sceneArray.length, 1);
    displayRange.load('address');
    await excel.sync()
    console.log(displayRange.address);
    displayRange.values = sceneArray;
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
        await jade_modules.operations.showMainPage();
        const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
        let columnIndex = await jade_modules.operations.findColumnIndex('Number');
        let tempRange = scriptSheet.getRangeByIndexes(10, columnIndex, 1, 1);
        tempRange.select();
        await excel.sync();
        await jade_modules.operations.findLineNo(lineNumber);
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
    const lineIndex = 2;
    let activeCell = excel.workbook.getActiveCell();
    activeCell.load('rowIndex');
    await excel.sync(); 
    let rowIndex = activeCell.rowIndex;
    if (rowIndex >= 10){
      let lineNumberCell = forActorSheet.getRangeByIndexes(rowIndex, lineIndex, 1, 1);
      lineNumberCell.load('values');
      await excel.sync(); 
      console.log('lineNumber', lineNumberCell.values);
      let lineNumber = parseInt(lineNumberCell.values[0][0].split(',')[0])
      console.log('lineNumber', lineNumber);
      if (!isNaN(lineNumber)){
        await jade_modules.operations.showMainPage();
        const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
        let columnIndex = await jade_modules.operations.findColumnIndex('Number');
        let tempRange = scriptSheet.getRangeByIndexes(10, columnIndex, 1, 1);
        tempRange.select();
        await excel.sync();
        await jade_modules.operations.findLineNo(lineNumber);
      } else {
        alert('Not a line number');
      }
    } else {
      alert('Must be in a line with valid line number');
    }
  })
}