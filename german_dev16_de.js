function auto_exec(){}

const germanProcessingSheetName = 'German Processing';
const scriptSheetName = 'Script';
const settingsSheetName = 'Settings';

const scriptRangeNames = [
  { name: 'scUKCue',
    range: 'F3:F30000',
    heading: '',
    formula: ''
  },
  { name: 'scUKNumber',
    range: 'G3:G30000',
    heading: '',
    formula: ''
  },
  { name: 'scUKCharacter',
    range: 'H3:H30000',
    heading: '',
    formula: ''
  },
  { name: 'scUKScript',
    range: 'K3:K30000',
    heading: '',
    formula: ''
  },
  { name: 'scUKCueWorking',
    range: 'DA3:DA30000',
    heading: 'UK Cue (Working)',
    formula: '=IF(ISNUMBER(F3),F3,"")'
  },
  { name: 'scUKNumberWorking',
    range: 'DB3:DB30000',
    heading: 'UK No (Working)',
    formula: '=IF(ISNUMBER(G3),G3,"")'
  },
  { name: 'scUKCharacterWorking',
    range: 'DC3:DC30000',
    heading: 'UK Character (Working)',
    formula: '=IF(H3=0,"",H3)'
  },
  { name: 'scUKScriptWorking',
    range: 'DD3:DD30000',
    heading: 'UK Script (Working)',
    formula: '=IF(K3=0,"",K3)'
  },
  { name: 'scGermanProcessed',
    range: 'DE3:DE30000',
    heading: 'German',
    formula: ''
  },
  { name: 'scGermanComments',
    range: 'DF3:DF30000',
    heading: 'Comments',
    formula: ''
  },
  { name: 'scUKCheck',
    range: 'DG3:DG30000',
    heading: 'UK Check',
    formula: ''
  },
  { name: 'scGermanScript',
    range: 'M3:M30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanComment',
    range: 'N3:N30000',
    heading: '',
    formula: ''
  },
  { name: 'scTotalTakes',
    range: 'U3:U30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanTakes',
    range: 'V3:V30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanTakeNo',
    range: 'W3:W30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanMarkup',
    range: 'X3:X30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanDate',
    range: 'Y3:Y30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanStudio',
    range: 'Z3:Z30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanEngineer',
    range: 'AA3:AA30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanRetake',
    range: 'AB3:AB30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanRemove',
    range: 'AC3:AC30000',
    heading: '',
    formula: ''
  },
  { name: 'scStageDirections',
    range: 'J3:J30000',
    heading: '',
    formula: ''
  },
  { name: 'scEnglishStageDirections',
    range: 'DJ3:DJ30000',
    heading: 'English Directions',
    formula: '=IF(J3=0,"",J3)'
  },
  { name: 'scGermanStageDirections',
    range: 'DK3:DK30000',
    heading: 'German Directions',
    formula: '=IF(DJ3=0,"",TRANSLATE(DJ3,"en","de"))'
  },
  { name: 'scEnglishStageDirectionsCopy',
    range: 'DL3:DL30000',
    heading: 'English Directions Copy',
    formula: ''
  },
  { name: 'scGermanStageDirectionsCopy',
    range: 'DM3:DM30000',
    heading: 'German Directions Copy',
    formula: ''
  }
]

const scriptRangeNamesAfterMove = [
  { name: 'scGermanCharacter',
    range: 'O3:O30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanDirection',
    range: 'P3:P30000',
    heading: '',
    formula: ''
  },
  { name: 'scGermanPresentCharacters',
    range: 'Q3:Q30000',
    heading: '',
    formula: ''
  }
]

async function createScriptNames(afterMove){
  let numAdded = 0;
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    
    //get the names in the workbook
    let theNames = excel.workbook.names.load();
    await excel.sync();
    let currentNames = theNames.items.map(x => x.name);
    console.log('currentNames', currentNames)

    let myNames;
    if (afterMove){
      myNames = scriptRangeNamesAfterMove;
    } else {
      myNames = scriptRangeNames;
    }

    for (let i = 0; i < myNames.length;i++){
      console.log(i, myNames[i].name);
      if (!currentNames.includes(myNames[i].name)){
        // It doesn't currently exist... add it
        let newRange = scriptSheet.getRange(myNames[i].range);
        excel.workbook.names.add(myNames[i].name, newRange);
        await excel.sync();
        numAdded += 1;
      }
    }
  })
  console.log('numAdded', numAdded)
}


async function setUpNewColumns(){
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    for (let i = 0; i < scriptRangeNames.length;i++){
      if(scriptRangeNames[i].heading != ''){
        console.log(i, 'Doing heading', scriptRangeNames[i].heading)
        let tempRange = scriptSheet.getRange(scriptRangeNames[i].name)
        tempRange.load('rowIndex, columnIndex');
        await excel.sync();
        let headerRange = scriptSheet.getRangeByIndexes(tempRange.rowIndex - 1, tempRange.columnIndex, 1, 1);
        headerRange.values = [[scriptRangeNames[i].heading]];
        await excel.sync();
        if (scriptRangeNames[i].formula != ''){
          console.log(i, 'Doing formula', scriptRangeNames[i].formula)
          let topCell = scriptSheet.getRangeByIndexes(tempRange.rowIndex, tempRange.columnIndex, 1, 1);
          topCell.formulas = [[scriptRangeNames[i].formula]];
          await excel.sync();
          topCell.autoFill(tempRange, 'FillDefault');
          await excel.sync();
        }
      }
    }
  })
}

async function processTheGerman(){
  let ukCueDetails = await getRangeDetails(scriptSheetName, 'scUKCueWorking');
  let ukNumberDetails = await getRangeDetails(scriptSheetName, 'scUKNumberWorking');
  let ukCharacterDetails = await getRangeDetails(scriptSheetName, 'scUKCharacterWorking');
  let ukScriptDetails = await getRangeDetails(scriptSheetName, 'scUKScriptWorking');
  console.log('ukCueDetails', ukCueDetails);
  console.log('ukNumberDetails', ukNumberDetails);
  console.log('ukCharacterDetails', ukCharacterDetails);
  console.log('ukScriptDetails', ukScriptDetails);

  let gpCueDetails = await getRangeDetails(germanProcessingSheetName, 'gpUKCue');
  let gpNumberDetails = await getRangeDetails(germanProcessingSheetName, 'gpUKLine');
  let gpCharacterDetails = await getRangeDetails(germanProcessingSheetName, 'gpUKCharacter');
  let gpScriptDetails = await getRangeDetails(germanProcessingSheetName, 'gpUKScript');

  let scGermanProcessed = await getRangeDetails(scriptSheetName, 'scGermanProcessed');
  await clearRangeContents(scriptSheetName, 'scGermanProcessed');
  
  let scGermanComments = await getRangeDetails(scriptSheetName, 'scGermanComments');
  await clearRangeContents(scriptSheetName, 'scGermanComments');
  
  let scUKCheck = await getRangeDetails(scriptSheetName, 'scUKCheck');
  await clearRangeContents(scriptSheetName, 'scUKCheck');

  let gpGermanProcessed = await getRangeDetails(germanProcessingSheetName, 'gpProcessed');
  let gpGermanComments = await getRangeDetails(germanProcessingSheetName, 'gpComments');

  console.log('gpCueDetails', gpCueDetails);
  console.log('gpNumberDetails', gpNumberDetails);
  console.log('gpCharacterDetails', gpCharacterDetails);
  console.log('gpScriptDetails', gpScriptDetails);

  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    const gpSheet = excel.workbook.worksheets.getItem(germanProcessingSheetName);
    await excel.sync();
    for (let i = 0; i < ukCharacterDetails.values.length; i++){
    //for (let i = 0; i < 2000; i++){
      germanMessage(true, 'Doing ' + i + ' of ' + ukCharacterDetails.values.length);
      let j = 0;
      let maxJ = 50000
      while ((!sameCueDetails(ukCueDetails, i, gpCueDetails,j)) & (j < maxJ)){
        j++
      }
  
      if (j < maxJ){
        if (ukNumberDetails.values[i] == gpNumberDetails.values[j]){
          //console.log(i,j, 'Number match');
          if (ukCharacterDetails.values[i] == gpCharacterDetails.values[j]){
            //console.log(i,j, 'Character match');
            if (scriptEqual(ukScriptDetails.values[i], gpScriptDetails.values[j])){
              //console.log(i,j, 'Script Match');
              let gpRowIndex = j + gpCueDetails.rowIndex;
              let scRowIndex = i + ukCueDetails.rowIndex
              let germanProcessRange = gpSheet.getRangeByIndexes(gpRowIndex, gpGermanProcessed.columnIndex, 1, 1);
              let germanCommentRange = gpSheet.getRangeByIndexes(gpRowIndex, gpGermanComments.columnIndex, 1, 1);
              let ukCheckRange = gpSheet.getRangeByIndexes(gpRowIndex, gpScriptDetails.columnIndex, 1, 1);
              germanProcessRange.load('values');
              germanCommentRange.load('values');
              ukCheckRange.load('values');
              await excel.sync();
              //console.log('German', germanProcessRange.values[0][0]);
              //console.log('Comment', germanCommentRange.values[0][0]);
              //console.log('UK Check', ukCheckRange.values[0][0]);
              
              let scGermanProcessedRange = scriptSheet.getRangeByIndexes(scRowIndex, scGermanProcessed.columnIndex, 1, 1);
              let scGermanCommentsRange = scriptSheet.getRangeByIndexes(scRowIndex, scGermanComments.columnIndex, 1, 1);
              let scUkCheckRange = scriptSheet.getRangeByIndexes(scRowIndex, scUKCheck.columnIndex, 1, 1);
              scGermanProcessedRange.values = [[germanProcessRange.values[0][0]]];
              scGermanCommentsRange.values =[[germanCommentRange.values[0][0]]];
              scUkCheckRange.values = [[ukCheckRange.values[0][0]]];
              await excel.sync(); 
            } else {
              console.log(i,j, '=====================> Failed on Script Match');
            }
          } else {
            console.log(i,j, '=====================> Failed on Character Match');
          }
        } else {
          console.log(i,j, '=====================> Failed on Number Match');        
        }
      }
      //console.log(i,j, ukCueDetails.values[i], gpCueDetails.values[j]);
    }
    germanMessage(false, '');
  })

}

function scriptEqual(uk, ge){
  if (uk == ge){
    return true;
  } else if (uk.startsWith("'")){
    if (uk.substring(1) == ge){
      return true;
    }
  } else if (ge.startsWith("'")){
    if (uk == ge.substring(1)){
      return true;
    }
  } else {
    return false;
  }
}

async function getRangeDetails(sheetName, rangeName){
  console.log('sheetName', sheetName, 'rangeName', rangeName)
  let details = {}
  await Excel.run(async function(excel){
    const theSheet = excel.workbook.worksheets.getItem(sheetName);
    let theRange = theSheet.getRange(rangeName);
    theRange.load('rowIndex, columnIndex, rowCount, columnCount, values');
    await excel.sync();
    let theValues = theRange.values.map(x => x[0]);
    details.rowIndex = theRange.rowIndex;
    details.columnIndex = theRange.columnIndex;
    details.rowCount = theRange.rowCount;
    details.columnCount = theRange.columnCount;
    details.values = theValues;
  })
  return details;
}

function sameCueDetails(uk, i,  gp, j){
  if (uk.values[i] != ''){
    return uk.values[i] == gp.values[j]
  } else {
    return false;
  }
  
}
async function clearRangeContents(sheetName, rangeName){
  console.log('Clear range sheetName', sheetName, 'rangeName', rangeName)
  await Excel.run(async function(excel){
    const theSheet = excel.workbook.worksheets.getItem(sheetName);
    let theRange = theSheet.getRange(rangeName);
    theRange.select();
    theRange.clear('Contents');
    await excel.sync();
  })
}

function germanMessage(show, message){
  let lblGerman = tag('german-wait')
  if (show){
    lblGerman.style.display = 'block';
  } else {
    lblGerman.style.display = 'none';
  }
  lblGerman.innerText = message;
}

const columnSwap =[
  {us: 'US Cue', german: 'German Script', width: 50},
  {us: 'US Script', german: 'German Comment', width: 30}
]

const takeColumnSwap =[
  {uk: 'UK No of takes', german: 'German No of takes', width: 12},
  {uk: 'UK Take No', german: 'German Take No', width: 12},
  {uk: 'UK Broadcast Assistant Markup', german: 'German Broadcast Assistant Markup', width: 30},
  {uk: 'UK Date Recorded', german: 'German Date Recorded', width: 12},
  {uk: 'UK Studio', german: 'German Studio', width: 20},
  {uk: 'UK Engineer', german: 'German Engineer', width: 20},
  {uk: 'UK Retake Required', german: 'German Retake Required', width: 12},
  {uk: 'UK Remove from Edit', german: 'German Remove from Edit', width: 12}
]

async function changeUStoGermanColumns(){
  await Excel.run(async function(excel){
    const settingsSheet = excel.workbook.worksheets.getItem(settingsSheetName);
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let columnDataRange = settingsSheet.getRange('columnData');
    columnDataRange.load('rowIndex, columnIndex, values');
    await excel.sync();

    let columnNames = columnDataRange.values.map(x => x[0]);
    let columnIndexes = columnDataRange.values.map(x => x[3]);

    for (let index = 0; index < columnSwap.length; index++){
      for (let i = 0; i < columnNames.length; i++){
        if(columnSwap[index].us == columnNames[i]){
          console.log('Found', columnSwap[index].us, 'rowIndex', i + columnDataRange.rowIndex);
          console.log('columnIndex', columnIndexes[i])
          columnHeaderRange = scriptSheet.getRangeByIndexes(1, columnIndexes[i], 1, 1)
          await excel.sync();
          columnHeaderRange.values = [[columnSwap[index].german]];
          await excel.sync();

          let theNameRange = settingsSheet.getRangeByIndexes(i + columnDataRange.rowIndex, columnDataRange.columnIndex, 1, 1);
          theNameRange.values = [[columnSwap[index].german]];
          await excel.sync();

          let theWidthRange = settingsSheet.getRangeByIndexes(i + columnDataRange.rowIndex, columnDataRange.columnIndex + 4, 1, 1);
          theWidthRange.values = [[columnSwap[index].width]];
          await excel.sync();
        }
      }
    }
  })
}

async function isCellMerged(sheetName, rowIndex, columnIndex){
  let result;
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(sheetName);
    let myRange = sheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);
    myRange.select();
    await excel.sync();
    let tempSelect = excel.workbook.getSelectedRange()
    tempSelect.load('address, cellCount');
    await excel.sync();
    console.log('Selected cell', tempSelect.address, tempSelect.cellCount);
    result = tempSelect.cellCount > 1;
  })
  return result;
}

async function copyToMainScript(){
  await Excel.run(async function(excel){
    let scriptSourceDetails = await getRangeDetails(scriptSheetName, 'scGermanProcessed');
    let scriptDestinationDetails = await getRangeDetails(scriptSheetName, 'scGermanScript');
    //await clearRangeContents(scriptSheetName, 'scGermanScript');
    let commentSourceDetails = await getRangeDetails(scriptSheetName, 'scGermanComments');
    let commentDestinationDetails = await getRangeDetails(scriptSheetName, 'scGermanComment');
    //await clearRangeContents(scriptSheetName, 'scGermanComment');
    let scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    //for (let i = 0; i < scriptSourceDetails.values.length; i++){
    for (let i = 0; i < 100; i++){
      germanMessage(true, 'Doing: ' + i + ' of ' + scriptSourceDetails.values.length);
      let theValue = scriptSourceDetails.values[i].trim()
      let theComment = commentSourceDetails.values[i].trim()
      let theRowIndex = i + scriptSourceDetails.rowIndex;
      if (!await isCellMerged(scriptSheetName, theRowIndex, scriptDestinationDetails.columnIndex)){
        let destRange = scriptSheet.getRangeByIndexes(theRowIndex, scriptDestinationDetails.columnIndex, 1, 1);
        if (theValue != ''){
          destRange.values = [[theValue]];
          destRange.select()
          if (!await isCellMerged(scriptSheetName, theRowIndex, commentDestinationDetails.columnIndex)){
            let destCommentRange = scriptSheet.getRangeByIndexes(theRowIndex, commentDestinationDetails.columnIndex, 1, 1);
            if (theComment != ''){
              destCommentRange.values = [[theComment]];
              console.log('Comment', theComment);
              await excel.sync();
            }
            if (theComment.toLowerCase() == 'ok'){
              console.log('clear');
              destCommentRange.clear('Contents');
            }
          }
          await excel.sync();
        } else {
          destRange.clear('Contents');
          await excel.sync();
        }
      }
    }
  })
}

async function changeUKtoGermanTakes(){
  await Excel.run(async function(excel){
    const settingsSheet = excel.workbook.worksheets.getItem(settingsSheetName);
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let columnDataRange = settingsSheet.getRange('columnData');
    columnDataRange.load('rowIndex, columnIndex, values');
    await excel.sync();

    let columnNames = columnDataRange.values.map(x => x[0]);
    let columnIndexes = columnDataRange.values.map(x => x[3]);

    for (let index = 0; index < takeColumnSwap.length; index++){
      for (let i = 0; i < columnNames.length; i++){
        if(takeColumnSwap[index].uk == columnNames[i]){
          console.log('Found', takeColumnSwap[index].us, 'rowIndex', i + columnDataRange.rowIndex);
          console.log('columnIndex', columnIndexes[i])
          columnHeaderRange = scriptSheet.getRangeByIndexes(1, columnIndexes[i], 1, 1)
          await excel.sync();
          columnHeaderRange.values = [[takeColumnSwap[index].german]];
          await excel.sync();

          let theNameRange = settingsSheet.getRangeByIndexes(i + columnDataRange.rowIndex, columnDataRange.columnIndex, 1, 1);
          theNameRange.values = [[takeColumnSwap[index].german]];
          await excel.sync();

          let theWidthRange = settingsSheet.getRangeByIndexes(i + columnDataRange.rowIndex, columnDataRange.columnIndex + 4, 1, 1);
          theWidthRange.values = [[takeColumnSwap[index].width]];
          await excel.sync();
        }
      }
    }
  })
}

async function clearTakeRanges(){
  await clearRangeContents(scriptSheetName, 'scTotalTakes');
  await clearRangeContents(scriptSheetName, 'scGermanTakes');

  await clearRangeContents(scriptSheetName, 'scGermanTakeNo');
  await clearRangeContents(scriptSheetName, 'scGermanMarkup');

  await clearRangeContents(scriptSheetName, 'scGermanDate');
  await clearRangeContents(scriptSheetName, 'scGermanStudio');

  await clearRangeContents(scriptSheetName, 'scGermanEngineer');
  await clearRangeContents(scriptSheetName, 'scGermanRetake');
  await clearRangeContents(scriptSheetName, 'scGermanRemove');
}

async function copyDirections(){
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let sourceUKRange = scriptSheet.getRange('scEnglishStageDirections');
    let destUKRange = scriptSheet.getRange('scEnglishStageDirectionsCopy');
    destUKRange.copyFrom(sourceUKRange, "values");
    await excel.sync();
    let sourceGermanRange = scriptSheet.getRange('scGermanStageDirections');
    let destGermanRange = scriptSheet.getRange('scGermanStageDirectionsCopy');
    destGermanRange.copyFrom(sourceGermanRange, "values");
    await excel.sync();
  })
}

async function copyTranslatedDirection(){
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let sourceGermanRange = scriptSheet.getRange('scGermanStageDirectionsCopy');
    let destGermanRange = scriptSheet.getRange('scGermanDirection');
    destGermanRange.copyFrom(sourceGermanRange, "values");
    await excel.sync();
  })
}



