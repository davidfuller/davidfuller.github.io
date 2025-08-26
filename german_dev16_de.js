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

]

async function createScriptNames(){
  console.log('createSciptNames')
  await Excel.run(async function(excel){
    const scriptSheet = excel.workbook.worksheets.getItem(scriptSheetName);
    await excel.sync();
    console.log('After excel.sync()');
    console.log(scriptRangeNames);
    for (let i = 0; i < scriptRangeNames.length;i++){
      let tempRange = excel.workbook.names.getItemOrNullObject(scriptRangeNames[i].name);
      await excel.sync();
      if (tempRange.isNullObject{
        // add the name;
        let newRange = scriptSheet.getRange(scriptRangeNames[i].range);
        scriptSheet.names.add(scriptRangeNames[i].name, newRange);
        await excel.sync();
      }
    }
  })
}



