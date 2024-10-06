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
    let theResults = [];
    for (let i = 1; i < theLines.length; i++){
      theResults[i - 1] = splitLine(theLines[i]);
    }
    console.log('theResults', theResults)
  })
}

/* where to put the data
    Walla cue Number - auto calculated starting at W00001

    Walla line range - either 'whole scene' or the best guess of line range
    Type of Walla - top line
    Walla characters - character
    Walla description - the description
    Num of characters - count the characters
    walla original - All
*/
function splitLine(theLine){
  //first split with '-'
  let theSections = theLine.split('-');
  let theCharacter = theSections[0].trim();
  let individualCharacters = theCharacter.split(',')
  
  let thePosition = theSections[1].trim()
  let wholeScene = thePosition.toLowerCase().indexOf('whole scene')
  let firstLine = thePosition.toLowerCase().indexOf('line')
  let lineNo = -1;
  if (firstLine != -1){
    lineNo = parseInt(thePosition.substring(firstLine + 4));
  }
  let theRestPosition = theLine.toLowerCase().indexOf(thePosition.toLowerCase());
  let theRest = '';
  if (theRestPosition != -1){
    theRest = theLine.substring(theRestPosition);
  }
  let lastBit = theSections[theSections.length - 1];
  

  let theDescription;
  let lastBitPosition;
  let lineRange;
  if (isNaN(parseInt(lastBit))){
    theDescription = lastBit.trim();
    lastBitPosition = theLine.toLowerCase().indexOf(lastBit.toLowerCase());
    lineRange = theLine.substring(theRestPosition, lastBitPosition - 2).trim() ;
  } else {
    theDescription = 'N/A'
    lineRange = theRest;
  }
  
  result = {
    all: theLine,
    character: theCharacter,
    wholeScene: (wholeScene != -1),
    line: lineNo,
    rest: theRest,
    description: theDescription,
    lineRange:  lineRange,
    numCharacters: individualCharacters.length
  }
  return result;

}