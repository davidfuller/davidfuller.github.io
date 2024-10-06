const wallaSheetName = 'Walla Import';
const sourceTextRangeName = 'wiSource';
const namedCharacters = 'Named Characters - For reaction sounds and walla';

function auto_exec(){
}

async function parseSource(){
  await Excel.run(async (excel) => {
    let wallaSheet = excel.workbook.worksheets.getItem(wallaSheetName);
    let sourceRange = wallaSheet.getRange(sourceTextRangeName);
    sourceRange.load('values')
    await excel.sync();
    let mySourceText = sourceRange.values[0][0];
    let theLines = mySourceText.split('\n');
    for (let i = 1; i < theLines.length; i++){
      splitLine(theLines[i]);
    }
  })
}

function splitLine(theLine){
  //first split with '-'
  theSections = theLine.split('-');
  theCharacter = theSections[0].trim();
  console.log(theCharacter)
  thePosition = theSections[1].trim()
  wholeScene = thePosition.toLowerCase().indexOf('whole scene')
  firstLine = thePosition.toLowerCase().indexOf('line')
  console.log(thePosition, parseInt(thePosition), wholeScene, firstLine);

}