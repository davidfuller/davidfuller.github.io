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
  let theSections = theLine.split('-');
  let theCharacter = theSections[0].trim();
  console.log(theCharacter)
  let thePosition = theSections[1].trim()
  let wholeScene = thePosition.toLowerCase().indexOf('whole scene')
  let firstLine = thePosition.toLowerCase().indexOf('line')
  let lineNo;
  if (firstLine != -1){
    lineNo = parseInt(thePosition.substring(firstLine + 4));
  }
  let theRestPosition = theLine.toLowerCase().indexOf(thePosition.toLowerCase());
  let theRest;
  if (theRestPosition != -1){
    theRest = theLine.substring(theRestPosition);
  }
  console.log(thePosition, parseInt(thePosition), wholeScene, firstLine, lineNo, theRestPosition, theRest);

}