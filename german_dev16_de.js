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
    formula: '=F3'
  },
  { name: 'scUKNumberWorking',
    range: 'DB3:DB30000',
    heading: 'UK No (Working)',
    formula: '=G3'
  },
  { name: 'scUKCharacterWorking',
    range: 'DC3:DC30000',
    heading: 'UK Character (Working)',
    formula: '=H3'
  },
  { name: 'scUKScriptWorking',
    range: 'DD3:DD30000',
    heading: 'UK Script (Working)',
    formula: '=K3'
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
        let topCell = scriptSheet.getRangeByIndexes(tempRange.rowIndex, tempRange.columnIndex, 1, 1);
        topCell.formulas = [['=F3']];
        await excel.sync();
      }
    }
  })
}
//topCell.autoFill(fillRange, 'FillDefault');


