let characterlistSheet, forDirectorSheet, forActorSheet;
const characterListName = 'Character List';
const forDirectorName = 'For Directors';
const forActorName = 'For Actors'
const forDirectorTableName = 'fdTable';
const forActorsTableName = "faTable";
const numItemsDirectorsName = 'fdItems';
const numItemsActorsName = 'faTable';

function auto_exec(){
}
async function loadReduceAndSortCharacters(){
  await Excel.run(async function(excel){ 
    characterlistSheet = excel.workbook.worksheets.getItem(characterListName);
    let characters = await jade_modules.operations.getCharacters();
    console.log(characters);
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
    forDirectorSheet = excel.workbook.worksheets.getItem(forDirectorName);
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
  }) 
}
async function getActorInfo(){
  await Excel.run(async function(excel){
    forActorSheet = excel.workbook.worksheets.getItem(forActorName);
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
    await excel.sync();
    let dataArray = [];
    for (i = 0; i < dataRange.rowCount; i++){
      let thisRow = new Array(3).fill("");
      let myIndex = dataArray.findIndex(x => x[0] == myData[i].sceneNumber)
      let theLocation = myLocation.find(x => x.sceneNumber == myData[i].sceneNumber)
      if (myIndex == -1){
        console.log(i, "New Row")
        thisRow = [myData[i].sceneNumber, myData[i].lineNumber, theLocation.location]
        dataArray.push(thisRow);
      } else {
        dataArray[myIndex][2] = dataArray[myIndex][2] + ", " + theLocation.location
        console.log("Found Index",  myIndex, "dataArray", dataArray[myIndex]);
      }
      console.log("Index", i, "dataArray", dataArray);
    }
    console.log('dataArray', dataArray, 'rowCount', dataRange.rowCount, 'dataLength', myData.length);
    dataRange.values = dataArray;
    numItems.values = dataArray.length;
    await excel.sync();
  })


  
}