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
  console.log('createSciptNames')
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    let theNames = excel.workbook.names.load();
    await excel.sync();
    console.log(theNames.items.length);
    for (let i = 0; i < theNames.items.length; i++){
      console.log(i, theNames.items[i].name);
    }
    await excel.sync();
    console.log('After excel.sync()');
    console.log(scriptRangeNames);
    console.log(scriptSheet.names);
    for (let i = 0; i < scriptRangeNames.length;i++){
      let tempRange = excel.workbook.names.getItemOrNullObject(scriptRangeNames[i].name).getRangeOrNullObject();
      tempRange.load('address');
      await excel.sync();
      console.log(i, 'tempRange', tempRange.address);
      
      if (tempRange.isNullObject){
        // add the name;
        let newRange = scriptSheet.getRange(scriptRangeNames[i].range);
        excel.workbook.names.add(scriptRangeNames[i].name, newRange);
        await excel.sync();
      }
    }
  })
}



