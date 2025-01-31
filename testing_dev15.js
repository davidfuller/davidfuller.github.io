function auto_exec(){
}

const characterListSheetName = 'Character List';

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

  //character word count

  messages.push(addMessage('Doing character word count'))
  await jade_modules.scheduling.processCharacterListForWordAndScene();

  //hide character list
  messages.push(addMessage('Hiding character List Sheet'));
  await hide(characterListSheetName);
  console.log(messages)
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