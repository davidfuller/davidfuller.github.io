function auto_exec(){
}

const characterListSheetName = 'Character List';
const characterRangeName = 'clCharacters'
const logSheetName = 'log';
const logRangeName = 'lgTable';
const settingsSheetName = 'Settings';
const versionRangeName = 'seVersion';
const dateRangeName = 'seDate';
const forActorSheetName = 'For Actors'

async function doTheFullTest(){
  let messages = [];
  messages.push(addMessage('Start of test'));

  //lock the sheet
  if (await jade_modules.operations.isLocked()){
    messages.push(addMessage('Sheet is locked'));
  } else {
    messages.push(addMessage('Sheet is unlocked, locking now'));
    await jade_modules.operations.lockColumns();
    messages.push(addMessage  ('Sheet is locked'));
  }

  //unfilter the sheet
  if (await jade_modules.operations.isFiltered()){
    messages.push(addMessage('Sheet is filtered, un-filtering now'));
    await jade_modules.operations.removeFilter();
    messages.push(addMessage('Sheet un-filtered'));
  } else {
    messages.push(addMessage('Sheet is un-filtered'));
  }

  //unhide character list
  messages.push(addMessage('Unhiding character List Sheet'));
  await unHide(characterListSheetName);

  //check the list
  messages.push(addMessage('Checking Characters'));
  let issues = await checkCharacters()

  if (issues != -1){
    messages.push(addMessage('Character issue before word count at:' + issues));
  } else {
    //character word count
    messages.push(addMessage('Doing character word count'))
    await jade_modules.scheduling.processCharacterListForWordAndScene();
    
    messages.push(addMessage('Doing scene word count'))
    await jade_modules.scheduling.createSceneWordCountData()

    messages.push(addMessage('Checking Characters after counts'));
    issues = await checkCharacters()
    if (issues != -1){
      messages.push(addMessage('Character issue after word count at:' + issues));
    } else {
      messages.push(addMessage('No character issues after word count'));
    }
  }

  //hide character list
  messages.push(addMessage('Hiding character List Sheet'));
  await hide(characterListSheetName);

  //update settings
  messages.push(addMessage('Updating settings in sheet'));
  await upDateSettings();

  //checkWalla
  messages.push(addMessage('Checking Walla Cues'));
  let wallaDetails = await jade_modules.operations.getWallaCues();
  console.log('wallaDetails', wallaDetails);
  messages.push(addMessage((wallaDetails.wallaIssues + wallaDetails.cueColumnIssues) + ' Walla cues issues'));

  //check cue numbers
  messages.push(addMessage('Checking for duplicate cue numbers'))
  let duplicateIndexes = jade_modules.operations.findDuplicateLineNumbers();
  if (duplicateIndexes.length > 0){
    messages.push(addMessage(duplicateIndexes.length + ' duplicate cue numbers at indexes: ' + duplicateIndexes.join(", ")));
  } else {
    messages.push(addMessage('No duplicate cue numbers'));
  }

  await moveMessages();

  await insertMessages(1, messages);
  console.log('messages', messages)
}


function addMessage(message){
  let result = {}
  result.time = new Date();
  result.message = message;
  return result;
}

async function unHide(sheetName){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(sheetName);
    sheet.visibility = 'Visible';
    sheet.activate();
  });
}

async function hide(sheetName){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(sheetName);
    sheet.visibility = 'Hidden';
  });
}

async function checkCharacters(){
  let issue = -1;
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(characterListSheetName);
    const range = sheet.getRange(characterRangeName);
    range.load('values');
    await excel.sync();
    
    let foundSpace = false
    for (let i = 0; i < range.values.length; i++){
      let thisValue = range.values[i][0].toString();
      if (!foundSpace){
        if (thisValue.trim() == ''){
          foundSpace = true;
          console.log('Found space at', i);
        }
      } else if (thisValue.trim() != ''){
        issue = i
        console.log('Issue at', i);
        break;
      }
    }
  });
  return issue;

}

async function insertMessages(columnNo, messages){
  await Excel.run(async function(excel){
    //getColumnRange
    const sheet = excel.workbook.worksheets.getItem(logSheetName);
    const range = sheet.getRange(logRangeName);
    range.load('rowIndex, rowCount, columnIndex');
    await excel.sync();

    let column = range.columnIndex + 2 * (columnNo - 1);
    const targetRange = sheet.getRangeByIndexes(range.rowIndex, column, range.rowCount, 2);
    targetRange.load('address');
    await excel.sync();

    console.log('address:', targetRange.address);

    //clearColumn
    targetRange.clear('Contents')

    //getTargetRange
    const targetValueRange = sheet.getRangeByIndexes(range.rowIndex, column, messages.length, 2)
    myValues = []
    for (let i = 0; i < messages.length; i++){
      console.log(i, messages[i]);
      myValues[i] = [jsDateToExcelDate(messages[i].time), messages[i].message];
    }
    //insert Data
    console.log('myValues', myValues);
    targetValueRange.values = myValues;
    sheet.activate();
    await excel.sync();
  })
}

function excelDateToJSDate(excelDate){
  //takes a number and return javascript Date object
  return new Date(Math.round((excelDate - 25569) * 86400 * 1000));
}

function jsDateToExcelDate(jsDate){
  //takes javascript a Date object to an excel number
  let returnDateTime = 25569.0 + ((jsDate.getTime()-(jsDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
  return returnDateTime
}

function createSpreadsheetDate(){
  // If before 16:59 use today. After 17.00 uses tomorrow
  //If Saturday or Sunday, use Monday
  const currentTime = new Date();
  let hour = currentTime.getHours();
  let newDate;
  if (hour < 16){
    newDate = currentTime;
  } else {
    newDate = addDays(currentTime, 1);
  }
  console.log('newDate first', newDate);
  let day = newDate.getDay();
  if (day == 6){
    //Saturday
    newDate = addDays(newDate, 2);// Monday
  }
  if (day == 0){
    //Sunday
    newDate = addDays(newDate, 1);// Monday
  }
  console.log('newDate second', newDate);
  return newDate;
}

function addDays(date, days) {
  const newDate = new Date(date);
  newDate.setDate(date.getDate() + days);
  return newDate;
}

async function upDateSettings(){
  await Excel.run(async function(excel){
    //getColumnRange
    const sheet = excel.workbook.worksheets.getItem(settingsSheetName);
    sheet.activate();
    const versionRange = sheet.getRange(versionRangeName);
    const dateRange = sheet.getRange(dateRangeName);

    versionRange.load('values');
    await excel.sync();

    let oldVersion = versionRange.values[0][0]
    let digits = oldVersion.split('.');
    console.log('oldVersion', oldVersion, 'digits', digits)
    if (digits.length == 3){
      if (!isNaN(parseInt(digits[2]))){
        let newDigit = parseInt(digits[2]) + 1;
        if (newDigit < 10) {
          newDigit = '0' + newDigit
        }
        let newDigits = digits[0] + '.' + digits[1] + '.' + newDigit;
        console.log('newDigits', newDigits);
        versionRange.values = [[newDigits]];
        await excel.sync();
      }
    }
    let newDate = createSpreadsheetDate();
    dateRange.values = [[jsDateToExcelDate(newDate)]];
  })  
}
async function moveMessages(){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(logSheetName);
    const range = sheet.getRange(logRangeName);
    range.load('rowIndex, rowCount, columnIndex');
    await excel.sync();
    for (let columnNo = 10; columnNo > 1; columnNo--){
      //getColumnRange
      let columnTarget = range.columnIndex + 2 * (columnNo - 1);
      const targetRange = sheet.getRangeByIndexes(range.rowIndex, columnTarget, range.rowCount, 2);
      const sourceRange = sheet.getRangeByIndexes(range.rowIndex, columnTarget - 2, range.rowCount, 2);
      targetRange.load('address');
      sourceRange.load('address');
      await excel.sync();
      console.log('target address:', targetRange.address, 'source address', sourceRange.address);
      targetRange.copyFrom(sourceRange, "values");
    }
  })
}

async function checkForActorConditionalFormatting(){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(forActorSheetName);
    let characterTextSearchRange = sheet.getRange('faTextSearch');
    characterTextSearchRange.load('conditionalFormats');
    await excel.sync();
    console.log('conditional formats', characterTextSearchRange.conditionalFormats.toJSON());
    let items = characterTextSearchRange.conditionalFormats.items
    await excel.sync()
    console.log('items', items);
  })
}