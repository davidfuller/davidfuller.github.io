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
  
#firstButtons, #filterButtons, #dateStudioEngineer, #showTakes {
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
  width: fit-content;
  height: fit-content;
  margin-left: 30px;
}

#btnFillUS, #btnFillWalla{
  width: fit-content;
  height: fit-content;
  margin-left: 5px;

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

#takeMessage {
  font-size: 14px;
  margin-left: 10px;
}

#btnUnhideAll, #btnShowLast, #btnShowFirst {
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

}