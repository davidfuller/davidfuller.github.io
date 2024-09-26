let characterlistSheet;
const characterListName = 'Character List'

function auto_exec(){
}
async function loadReduceAndSortCharacters(){
  await Excel.run(async function(excel){ 
    characterlistSheet = excel.workbook.worksheets.getItem(characterListName);
    let characters = await jade_modules.operations.getCharacters();
    console.log(characters);
    let characterRange = characterlistSheet.getRange('clCharacters');
    characterRange.clear("Contents")
    await excel.sync();
    characterRange.values = characters;
    await excel.sync();
  })  
}
