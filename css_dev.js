function auto_exec(){
}
async function mainCSS(){
  const css =`
body {
  background-color: #ffc901;
  color: #d34c01;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
}
button {
  margin-left: 30px;
  margin-bottom: 5px;
  margin-top: 5px;
  padding-left: 15px;
  padding-right: 15px;
  height: 36px;
  border: none;
  width: 150px;
  border-radius: 7px;
  background-color: #d34c01;
  color: #fef3df;
  font-size: 16px;
  cursor: pointer;
}
h1 {
  margin-left: 30px;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
}
h2 {
  margin-left: 30px;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
  font-size: 12px;
  font-weight: normal;
}
label {
  margin-left: 30px;
  font-size: 16px;
}
  
#firstButtons, #filterButtons, #dateStudioEngineer, #showTakes, #showColumns {
  margin-left: 30px;
  margin-top: 10px;
  width: 370px;
  padding: 10px;
  border-width: 1px;
  border-radius: 5px;
  border-color: #d34c01;
  border-style: solid;
}

#jump {
  margin-left: 30px;
  margin-top: 10px;
  width: 370px;
  padding: 10px;
  border-width: 1px;
  border-radius: 5px;
  border-color: #d34c01;
  border-style: solid;
}

#btnGo {
  width: fit-content;
  height: fit-content;
  margin-left: 5px;
}
#btnFillUK {
  margin-left: 30px;
  margin-bottom: 5px;
  margin-top: 5px;
  padding-left: 15px;
  padding-right: 15px;
  height: 24px;
  border: none;
  width: fit-content;
  border-radius: 4px;
  background-color: #d34c01;
  color: #fef3df;
  font-size: 12px;
  cursor: pointer;
}

#btnFillUS, #btnFillWalla{
  margin-left: 5px;
  margin-bottom: 5px;
  margin-top: 5px;
  padding-left: 15px;
  padding-right: 15px;
  height: 24px;
  border: none;
  width: fit-content;
  border-radius: 4px;
  background-color: #d34c01;
  color: #fef3df;
  font-size: 12px;
  cursor: pointer;
}
#scene {
  width: 50px;
  background-color: #ffc901;
  color: #d34c01;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
  border-width: 1px;
  border-radius: 5px;
  border-color: #d34c01;
  border-style: solid;
}

#chkAboveDetails {
  background-color: #ffc901;
  color: #d34c01;
  border-width: 1px;
  border-radius: 5px;
  border-color: #d34c01;
  border-style: solid;
}

#lblAboveDetails {
  background-color: #ffc901;
  color: #d34c01;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
  margin-left: 5px;
  margin-bottom: 5px;
  margin-top: 5px;
  padding-left: 5px;
  padding-right: 5px;
  height: 24px;
  border: none;
  width: fit-content;
  font-size: 12px;
  cursor: pointer;
}

#studio, #engineer, #studio-select, #engineer-select {
  width: 150px;
  background-color: #ffc901;
  color: #d34c01;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
  border-width: 1px;
  border-radius: 5px;
  border-color: #d34c01;
  border-style: solid;
}
.row {
  display: flex;
}

.column {
  padding: 10px;
}

#column1 {
  flex: 30%
}
#column2 {
  flex: 60%
}

#takeMessage, #columnMessage {
  font-size: 14px;
  margin-left: 10px;
}

#btnUnhideAll, #btnShowLast, #btnShowFirst, #btnShowAll, #btnShowUK, #btnShowUS, #btnShowWalla, #btnAddTakeUK, #btnRemoveTakeUK {
  margin-left: 10px;
  margin-bottom: 5px;
  margin-top: 5px;
  padding-left: 15px;
  padding-right: 15px;
  height: 24px;
  border: none;
  width: fit-content;
  border-radius: 4px;
  background-color: #d34c01;
  color: #fef3df;
  font-size: 12px;
  cursor: pointer;
}
                    `;
  Jade.set_css(css);
  console.log('CSS done')
}