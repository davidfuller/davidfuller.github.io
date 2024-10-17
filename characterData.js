const linkedDataSheetName = 'Linked_Data'


async function makeTheFullList(){
  await Excel.run(async function(excel){ 
    let linkedDataSheet = excel.workbook.worksheets.getItem(linkedDataSheetName);
    for (let i = 1; i<= 7; i++){
      let rangeName = 'ldSheet' + i;
      let thisRange = linkedDataSheet.getRange(rangeName);
      thisRange.load('values')
      await excel.sync();
      let myValues = thisRange.values.map(x => x[0]);
      let filteredValues = myValues.filter((x) => x != 0)
      console.log(i, myValues, filteredValues);
      //let myIndecies = myData.map((x, i) => [x, i]).filter(([x, i]) => x == targetValue).map(([x, i]) => i + firstDataRow - 1);
    }
  })
}
