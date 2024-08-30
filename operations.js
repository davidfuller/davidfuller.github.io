function auto_exec(){
  console.log('Operations loaded')
}
const columnsToLock = "A:Y"

async function lockColumns(excel){
  const sheet = excel.workbook.worksheets.getActiveWorksheet();
  var range = sheet.getRange(columnsToLock);
  
  sheet.protection.load('protected');
  await excel.sync();
  
  console.log(sheet.protection.protected);
  if (!sheet.protection.protected){
    console.log("Not locked");
    range.format.protection.locked = true;
    sheet.protection.protect({ selectionMode: "Normal", allowAutoFilter: true });
    await excel.sync();
    console.log("Now locked");
  } else {
    console.log("Locked");
  }   
}

async function unlock(excel){
  const sheet = excel.workbook.worksheets.getActiveWorksheet();
  sheet.protection.load('protected');
  await excel.sync();
  if (!sheet.protection.protected){
    console.log("Already unlocked");
  } else {
    console.log("Currently locked");
    sheet.protection.unprotect("")
    await excel.sync();
    console.log("Now not locked");
  }
}
