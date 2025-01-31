function auto_exec(){
}

const characterListSheetName = 'Character List';
const characterRangeName = 'clCharacters'

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
      console.log(i, thisValue)
      if (!foundSpace){
        if (thisValue.trim() == ''){
          foundSpace = true;
          console.log('Found space at', i);
        } else if (thisValue.trim() != ''){
          issue = i
          console.log('Issue at', i);
        }
      }
    }
  });
  return issue;

}