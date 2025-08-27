function auto_exec(){}

const germanProcessingSheetName = 'German Processing';
const scriptSheetName = 'Script';

const scriptRangeNames = [
  { name: 'scUKCue',
    range: 'F3:F30000'
  },
  { name: 'scUKNumber',
    range: 'G3:G30000'
  },
  { name: 'scUKCharacter',
    range: 'H3:H30000'
  },
  { name: 'scUKScript',
    range: 'K3:K30000'
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



