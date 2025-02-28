function auto_exec(){
}

const characterListSheetName = 'Character List';
const characterRangeName = 'clCharacters'
const logSheetName = 'log';
const logRangeName = 'lgTable';
const settingsSheetName = 'Settings';
const versionRangeName = 'seVersion';
const dateRangeName = 'seDate';
const forActorSheetName = 'For Actors';
const forSchedulingSheetName = 'For Scheduling';

const inputFont = 
  {
    "autoIndent": false,
    "columnWidth": 264.75,
    "horizontalAlignment": "General",
    "indentLevel": 0,
    "readingOrder": "Context",
    "rowHeight": 15.75,
    "shrinkToFit": false,
    "textOrientation": 0,
    "useStandardHeight": false,
    "useStandardWidth": false,
    "verticalAlignment": "Bottom",
    "wrapText": false,
    "borders": [
        {
            "color": "#51154A",
            "sideIndex": "EdgeTop",
            "style": "Continuous",
            "tintAndShade": 0,
            "weight": "Thin"
        },
        {
            "color": "#51154A",
            "sideIndex": "EdgeBottom",
            "style": "Continuous",
            "tintAndShade": 0,
            "weight": "Thin"
        },
        {
            "color": "#51154A",
            "sideIndex": "EdgeLeft",
            "style": "Continuous",
            "tintAndShade": 0,
            "weight": "Thin"
        },
        {
            "color": "#51154A",
            "sideIndex": "EdgeRight",
            "style": "Continuous",
            "tintAndShade": 0,
            "weight": "Thin"
        },
        {
            "color": "#000000",
            "sideIndex": "InsideVertical",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#000000",
            "sideIndex": "InsideHorizontal",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#000000",
            "sideIndex": "DiagonalDown",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#000000",
            "sideIndex": "DiagonalUp",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        }
    ],
    "fill": {
        "color": "#FFFFFF",
        "pattern": null,
        "patternColor": "#FFFFFF",
        "patternTintAndShade": null,
        "tintAndShade": null
    },
    "font": {
        "bold": false,
        "color": "#000000",
        "italic": false,
        "name": "Aptos Narrow",
        "size": 11,
        "strikethrough": false,
        "subscript": false,
        "tintAndShade": 0,
        "superscript": false,
        "underline": "None"
    }
  }
const labelFontText = 
  {
    "autoIndent": false,
    "columnWidth": 110.25,
    "horizontalAlignment": "Right",
    "indentLevel": 0,
    "readingOrder": "Context",
    "rowHeight": 15.75,
    "shrinkToFit": false,
    "textOrientation": 0,
    "useStandardHeight": false,
    "useStandardWidth": false,
    "verticalAlignment": "Top",
    "wrapText": false,
    "borders": [
        {
            "color": "#FBE2D5",
            "sideIndex": "EdgeTop",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#FBE2D5",
            "sideIndex": "EdgeBottom",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#FBE2D5",
            "sideIndex": "EdgeLeft",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#51154A",
            "sideIndex": "EdgeRight",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "blue",
            "sideIndex": "InsideVertical",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "blue",
            "sideIndex": "InsideHorizontal",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "green",
            "sideIndex": "DiagonalDown",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "green",
            "sideIndex": "DiagonalUp",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        }
    ],
    "fill": {
        "color": "#FBE2D5",
        "pattern": null,
        "patternColor": "#FBE2D5",
        "patternTintAndShade": null,
        "tintAndShade": null //0.799981688894314
    },
    "font": {
        "bold": true,
        "color": "#51154A",
        "italic": false,
        "name": "Aptos Display",
        "size": 12,
        "strikethrough": false,
        "subscript": false,
        "superscript": false,
        "tintAndShade": 0,
        "underline": "None"
    }
  }
const labelFontList = 
  {
    "autoIndent": false,
    "columnWidth": 110.25,
    "horizontalAlignment": "Right",
    "indentLevel": 0,
    "readingOrder": "Context",
    "rowHeight": 15.75,
    "shrinkToFit": false,
    "textOrientation": 0,
    "useStandardHeight": false,
    "useStandardWidth": false,
    "verticalAlignment": "Top",
    "wrapText": false,
    "borders": [
        {
            "color": "#FBE2D5",
            "sideIndex": "EdgeTop",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#FBE2D5",
            "sideIndex": "EdgeBottom",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#51154A",
            "sideIndex": "EdgeLeft",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#51154A",
            "sideIndex": "EdgeRight",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "blue",
            "sideIndex": "InsideVertical",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "blue",
            "sideIndex": "InsideHorizontal",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "green",
            "sideIndex": "DiagonalDown",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "green",
            "sideIndex": "DiagonalUp",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        }
    ],
    "fill": {
        "color": "#FBE2D5",
        "pattern": null,
        "patternColor": "#FBE2D5",
        "patternTintAndShade": null,
        "tintAndShade": null //0.799981688894314
    },
    "font": {
        "bold": true,
        "color": "#51154A",
        "italic": false,
        "name": "Aptos Display",
        "size": 12,
        "strikethrough": false,
        "subscript": false,
        "superscript": false,
        "tintAndShade": 0,
        "underline": "None"
    }
  }
const myConditionalFormats = [
  {
    name: 'faTextSearch',
    mainFontStyle: inputFont,
    rule: '=$D$8 = "List Search"',
    doFillColor: true,
    fillColor: "#FBE2D5",
    doFontColor: true,
    fontColor: "#FBE2D5",
    doBorders: true,
    borders: [
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeTop",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeBottom",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeLeft",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeRight",
          "style": "Continuous"
      }
    ]
  },
  {
    name: 'faCharacterChoice',
    mainFontStyle: inputFont,
    rule: '=$D$8 = "Text Search"',
    doFillColor: true,
    fillColor: "#FBE2D5",
    doFontColor: true,
    fontColor: "#FBE2D5",
    doBorders: true,
    borders: [
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeTop",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeBottom",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeLeft",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeRight",
          "style": "Continuous"
      }
    ]
  },
  {
    name: 'faCharacterChoiceLabel',
    mainFontStyle: labelFontList,
    rule: '=$D$8 = "Text Search"',
    doFillColor: false,
    fillColor: "#FBE2D5",
    doFontColor: true,
    fontColor: "#FBE2D5",
    doBorders: false,
    borders: [
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeTop",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeBottom",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeLeft",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeRight",
          "style": "Continuous"
      }
    ]
  },
  {
    name: 'faTextChoiceLabel',
    mainFontStyle: labelFontText,
    rule: '=$D$8 = "List Search"',
    doFillColor: false,
    fillColor: "#FBE2D5",
    doFontColor: true,
    fontColor: "#FBE2D5",
    doBorders: false,
    borders: [
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeTop",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeBottom",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeLeft",
          "style": "Continuous"
      },
      {
          "color": "#FBE2D5",
          "sideIndex": "EdgeRight",
          "style": "Continuous"
      }
    ]
  }
]

const schedulingLabelFontText = 
{
  "autoIndent": false,
  "columnWidth": 121.5,
  "horizontalAlignment": "Right",
  "indentLevel": 0,
  "readingOrder": "Context",
  "rowHeight": 15.75,
  "shrinkToFit": false,
  "textOrientation": 0,
  "useStandardHeight": false,
  "useStandardWidth": false,
  "verticalAlignment": "Top",
  "wrapText": false,
  "borders": [
        {
            "color": "#DAF2D0",
            "sideIndex": "EdgeTop",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#DAF2D0",
            "sideIndex": "EdgeBottom",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#DAF2D0",
            "sideIndex": "EdgeLeft",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "#DAF2D0",
            "sideIndex": "EdgeRight",
            "style": "Continuous",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "blue",
            "sideIndex": "InsideVertical",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "blue",
            "sideIndex": "InsideHorizontal",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "green",
            "sideIndex": "DiagonalDown",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        },
        {
            "color": "green",
            "sideIndex": "DiagonalUp",
            "style": "None",
            "tintAndShade": null,
            "weight": "Thin"
        }
    ],
  "fill": {
      "color": "#DAF2D0",
      "pattern": null,
      "patternColor": "#DAF2D0",
      "patternTintAndShade": null,
      "tintAndShade": null
  },
  "font": {
      "bold": true,
      "color": "#51154A",
      "italic": false,
      "name": "Aptos Display",
      "size": 12,
      "strikethrough": false,
      "subscript": false,
      "superscript": false,
      "tintAndShade": 0,
      "underline": "None"
  }
}

const schedulingFontText =
{
  "autoIndent": false,
  "columnWidth": 263.25,
  "horizontalAlignment": "Left",
  "indentLevel": 0,
  "readingOrder": "Context",
  "rowHeight": 15.75,
  "shrinkToFit": false,
  "textOrientation": 0,
  "useStandardHeight": false,
  "useStandardWidth": false,
  "verticalAlignment": "Top",
  "wrapText": true,
  "borders": [
      {
          "color": "#51154A",
          "sideIndex": "EdgeTop",
          "style": "Continuous",
          "tintAndShade": 0,
          "weight": "Thin"
      },
      {
          "color": "#51154A",
          "sideIndex": "EdgeBottom",
          "style": "Continuous",
          "tintAndShade": 0,
          "weight": "Thin"
      },
      {
          "color": "#51154A",
          "sideIndex": "EdgeLeft",
          "style": "Continuous",
          "tintAndShade": 0,
          "weight": "Thin"
      },
      {
          "color": "#51154A",
          "sideIndex": "EdgeRight",
          "style": "Continuous",
          "tintAndShade": -0,
          "weight": "Thin"
      },
      {
          "color": "#000000",
          "sideIndex": "InsideVertical",
          "style": "None",
          "tintAndShade": null,
          "weight": "Thin"
      },
      {
          "color": "#000000",
          "sideIndex": "InsideHorizontal",
          "style": "None",
          "tintAndShade": null,
          "weight": "Thin"
      },
      {
          "color": "#000000",
          "sideIndex": "DiagonalDown",
          "style": "None",
          "tintAndShade": null,
          "weight": "Thin"
      },
      {
          "color": "#000000",
          "sideIndex": "DiagonalUp",
          "style": "None",
          "tintAndShade": null,
          "weight": "Thin"
      }
  ],
  "fill": {
      "color": "#FFFFFF",
      "pattern": null,
      "patternColor": "#FFFFFF",
      "patternTintAndShade": null,
      "tintAndShade": null
  },
  "font": {
      "bold": false,
      "color": "#000000",
      "italic": false,
      "name": "Aptos Narrow",
      "size": 11,
      "strikethrough": false,
      "subscript": false,
      "superscript": false,
      "tintAndShade": 0,
      "underline": "None"
  }
}
   
const mySchedulingConditionalFormats = [
  {
    name: 'fsTextSearchLabel',
    mainFontStyle: schedulingLabelFontText,
    rule: '=$D$8 = "List Search"',
    doFillColor: true,
    fillColor: "#DAF2D0",
    doFontColor: true,
    fontColor: "#DAF2D0",
    doBorders: true,
    borders: [
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeTop",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeBottom",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeLeft",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeRight",
          "style": "Continuous"
      }
    ]
  },
  {
    name: 'fsListSearchLabel',
    mainFontStyle: schedulingLabelFontText,
    rule: '=$D$8 = "Text Search"',
    doFillColor: true,
    fillColor: "#DAF2D0",
    doFontColor: true,
    fontColor: "#DAF2D0",
    doBorders: true,
    borders: [
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeTop",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeBottom",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeLeft",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeRight",
          "style": "Continuous"
      }
    ]
  },
  {
    name: 'fsTextSearch',
    mainFontStyle: schedulingFontText,
    rule: '=$D$8 = "List Search"',
    doFillColor: true,
    fillColor: "#DAF2D0",
    doFontColor: true,
    fontColor: "#DAF2D0",
    doBorders: true,
    borders: [
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeTop",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeBottom",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeLeft",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeRight",
          "style": "Continuous"
      }
     ]
  },
  {
    name: 'fsCharacterChoice',
    mainFontStyle: schedulingFontText,
    rule: '=$D$8 = "Text Search"',
    doFillColor: true,
    fillColor: "#DAF2D0",
    doFontColor: true,
    fontColor: "#DAF2D0",
    doBorders: true,
    borders: [
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeTop",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeBottom",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeLeft",
          "style": "Continuous"
      },
      {
          "color": "#DAF2D0",
          "sideIndex": "EdgeRight",
          "style": "Continuous"
      }
     ]
  }
]

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
  //console.log('wallaDetails', wallaDetails);
  messages.push(addMessage((wallaDetails.wallaIssues + wallaDetails.cueColumnIssues) + ' Walla cues issues'));

  //check cue numbers
  messages.push(addMessage('Checking for duplicate cue numbers'))
  let duplicateIndexes = await jade_modules.operations.findDuplicateLineNumbers();
  if (duplicateIndexes.length > 0){
    messages.push(addMessage(duplicateIndexes.length + ' duplicate cue numbers at indexes: ' + duplicateIndexes.join(", ")));
  } else {
    messages.push(addMessage('No duplicate cue numbers'));
  }

  //conditional formatting actors
  messages.push(addMessage('Conditional formatting For Actors'))
  await checkForActorConditionalFormatting();

  //conditional formatting scheduling
  messages.push(addMessage('Conditional formatting For Scheduling'))
  await checkForSchedulingConditionalFormatting();

  await moveMessages();

  await insertMessages(1, messages);
  //console.log('messages', messages)
}


function addMessage(message){
  displayMessage(message);
  let result = {}
  result.time = new Date();
  result.message = message;
  console.log(result);
  return result;
}

function displayMessage(message){
  let testMessage = tag('test-message');
  testMessage.style.display = 'block';
  testMessage.innerText += message + '\n';
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
          //console.log('Found space at', i);
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

    //console.log('address:', targetRange.address);

    //clearColumn
    targetRange.clear('Contents')

    //getTargetRange
    const targetValueRange = sheet.getRangeByIndexes(range.rowIndex, column, messages.length, 2)
    myValues = []
    for (let i = 0; i < messages.length; i++){
      //console.log(i, messages[i]);
      myValues[i] = [jsDateToExcelDate(messages[i].time), messages[i].message];
    }
    //insert Data
    //console.log('myValues', myValues);
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
  //console.log('newDate first', newDate);
  let day = newDate.getDay();
  if (day == 6){
    //Saturday
    newDate = addDays(newDate, 2);// Monday
  }
  if (day == 0){
    //Sunday
    newDate = addDays(newDate, 1);// Monday
  }
  //console.log('newDate second', newDate);
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
    //console.log('oldVersion', oldVersion, 'digits', digits)
    if (digits.length == 3){
      if (!isNaN(parseInt(digits[2]))){
        let newDigit = parseInt(digits[2]) + 1;
        if (newDigit < 10) {
          newDigit = '0' + newDigit
        }
        let newDigits = digits[0] + '.' + digits[1] + '.' + newDigit;
        //console.log('newDigits', newDigits);
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
      //console.log('target address:', targetRange.address, 'source address', sourceRange.address);
      targetRange.copyFrom(sourceRange, "values");
    }
  })
}

async function checkForActorConditionalFormatting(){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(forActorSheetName);
    let characterTextSearchRange = sheet.getRange('faCharacterChoiceLabel');
    characterTextSearchRange.load('conditionalFormats, format/*, format/font, format/fill, format/borders');
    await excel.sync();
    //console.log('conditional formats', characterTextSearchRange.conditionalFormats.toJSON());
    //console.log('format', characterTextSearchRange.format.toJSON());
    //console.log('format/font', characterTextSearchRange.format.font.toJSON());
    //console.log('format/fill', characterTextSearchRange.format.fill.toJSON());
    //console.log('format/borders', characterTextSearchRange.format.borders.toJSON());
    
    await excel.sync();
    for (let myFormat of myConditionalFormats){
      //console.log('Doing cell', myFormat.name);
      //console.log('mainFont', myFormat.mainFontStyle);
      let range = sheet.getRange(myFormat.name);
      //fill
      range = doTheMainFont(range, myFormat.mainFontStyle);
      await excel.sync();

      range.conditionalFormats.clearAll();
      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      conditionalFormat.custom.rule.formula = myFormat.rule;
      if (myFormat.doFontColor){
        conditionalFormat.custom.format.font.color = myFormat.fontColor;
      }
      if (myFormat.doFillColor){
        conditionalFormat.custom.format.fill.color = myFormat.fillColor;
      }
      if (myFormat.doBorders){
        let myBorders = conditionalFormat.custom.format.borders;
        myBorders.load('count, items');
        await excel.sync();
        for (let border of myFormat.borders){
          let myEdge = myBorders.getItem(border.sideIndex);
          myEdge.load('sideIndex, color, style');
          await excel.sync();
          myEdge.color = border.color;
          myEdge.style = border.style;
          await excel.sync();
          //console.log('myEdge After', myEdge.toJSON());
        }
      }
    }
  })
}
async function checkForSchedulingConditionalFormatting(){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(forSchedulingSheetName);
    await getFontDetails(forSchedulingSheetName, 'fsTextSearchLabel');
    await getFontDetails(forSchedulingSheetName, 'fsTextSearch');
    for (let myFormat of mySchedulingConditionalFormats){
      //console.log('Doing cell', myFormat.name);
      //console.log('mainFont', myFormat.mainFontStyle);
      let range = sheet.getRange(myFormat.name);
      //fill
      range = doTheMainFont(range, myFormat.mainFontStyle);
      await excel.sync();
      
      range.conditionalFormats.clearAll();
      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      conditionalFormat.custom.rule.formula = myFormat.rule;
      if (myFormat.doFontColor){
        conditionalFormat.custom.format.font.color = myFormat.fontColor;
      }
      if (myFormat.doFillColor){
        conditionalFormat.custom.format.fill.color = myFormat.fillColor;
      }
      if (myFormat.doBorders){
        let myBorders = conditionalFormat.custom.format.borders;
        myBorders.load('count, items');
        await excel.sync();
        for (let border of myFormat.borders){
          let myEdge = myBorders.getItem(border.sideIndex);
          myEdge.load('sideIndex, color, style');
          await excel.sync();
          myEdge.color = border.color;
          myEdge.style = border.style;
          await excel.sync();
          //console.log('myEdge After', myEdge.toJSON());
        }
      }
    }
  })
}

function doTheMainFont(range, style){
  //console.log('style', style)
  range.format.fill.color = style.fill.color;
  
  range.format.fill.pattern = style.fill.pattern;
  range.format.fill.patternColor = style.fill.patternColor;
  range.format.fill.patternTintAndShade = style.fill.patternTintAndShade;
  range.format.fill.tintAndShade = style.fill.tintAndShade;
  
  range.format.autoIndent = style.autoIndent;
  range.format.columnWidth = style.columnWidth;
  range.format.horizontalAlignment = style.horizontalAlignment;
  range.format.indentLevel = style.indentLevel;
  range.format.readingOrder = style.readingOrder;
  
  range.format.rowHeight = style.rowHeight;
  range.format.shrinkToFit = style.shrinkToFit;
  range.format.useStandardHeight = style.useStandardHeight;
  range.format.textOrientation = style.textOrientation;
  range.format.useStandardWidth = style.useStandardWidth;
    
  range.format.verticalAlignment = style.verticalAlignment;
  range.format.wrapText = style.wrapText;
  
  range.format.font.bold = style.font.bold;
  range.format.font.color = style.font.color;
  range.format.font.italic = style.font.italic;
  range.format.font.name = style.font.name;
  range.format.font.size = style.font.size;
  range.format.font.strikethrough = style.font.strikethrough;
  range.format.font.subscript = style.font.subscript;
  
  range.format.font.tintAndShade = style.font.tintAndShade;
  range.format.font.superscript = style.font.superscript;
  range.format.font.underline = style.font.underline;
  
  for (let border of style.borders){
    let myBorder = range.format.borders.getItem(border.sideIndex);
    myBorder.color = border.color;
    myBorder.style = border.style;
    myBorder.tintAndShade = border.tintAndShade;
    myBorder.weight = border.weight;
    
  }
  
  return range;
}

async function getFontDetails(sheetName, rangeName){
  await Excel.run(async function(excel){
    const sheet = excel.workbook.worksheets.getItem(sheetName);
    let testRange = sheet.getRange(rangeName);
    testRange.load('conditionalFormats/*, conditionalFormats/custom/rule, conditionalFormats/custom/format/fill, conditionalFormats/custom/format/font, conditionalFormats/custom/format/borders, format/*, format/font, format/fill, format/borders');
    await excel.sync();
    //console.log('Range:', rangeName)
    //console.log('format', testRange.format.toJSON());
    //console.log('conditional formats', testRange.conditionalFormats.toJSON());
  })
   
}

function showAdminForActor(){
  let admin = tag('admin-actor')
  if (admin.style.display === 'block'){
    admin.style.display = 'none';
  } else {
    admin.style.display = 'block';
  }
}

function showAdminForScheduling(){
  let admin = tag('admin-scheduling')
  if (admin.style.display === 'block'){
    admin.style.display = 'none';
  } else {
    admin.style.display = 'block';
  }
}