function auto_exec(){}

const germanProcessingSheetName = 'German Processing';
const scriptSheetName = 'Script';

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
  }

]

async function createScriptNames(){
  let numAdded = 0;
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    
    //get the names in the workbook
    let theNames = excel.workbook.names.load();
    await excel.sync();
    let currentNames = theNames.items.map(x => x.name);
    console.log('currentNames', currentNames)

    for (let i = 0; i < scriptRangeNames.length;i++){
      if (!currentNames.includes(scriptRangeNames[i].name)){
        // It doesn't currently exist... add it
        let newRange = scriptSheet.getRange(scriptRangeNames[i].range);
        excel.workbook.names.add(scriptRangeNames[i].name, newRange);
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
        let tempRange = scriptSheet.getRange(scriptRangeNames[i].name)
        tempRange.load('rowIndex, columnIndex');
        await excel.sync();
        let headerRange = scriptSheet.getRangeByIndexes(tempRange.rowIndex - 1, tempRange.columnIndex, 1, 1);
        headerRange.values = [[scriptRangeNames[i].heading]];
        await excel.sync();
        if (scriptRangeNames[i].formula != ''){
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



