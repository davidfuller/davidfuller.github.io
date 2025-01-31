function auto_exec(){
}

async function doTheFullTest(){
  let messages = [];
  messages.push(addMessage('Start of test'));

  //lock the sheet
  if (await jade_modules.operations.isLocked()){
    messages.push(addMessage('Sheet is locked'));
  } else {
    messages.push(addMessage('Sheet is unlocked, locking now'));
    await jade_modules.operations.lockColumns();
    messages.push(addMessage('Sheet is locked'));
  }

  console.log(messages)
}


function addMessage(message){
  let result = {}
  result.time = new Date();
  result.message = message;
  return result;
}