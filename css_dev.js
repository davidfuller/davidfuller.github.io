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

/*
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
*/
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
/* Start of checkbox stuff */
/* Customize the label (the container) */
.container {
  display: block;
  position: relative;
  padding-left: 15px;
  margin-bottom: 12px;
  cursor: pointer;
  font-size: 12px;
  -webkit-user-select: none;
  -moz-user-select: none;
  -ms-user-select: none;
  user-select: none;
}

/* Hide the browser's default checkbox */
.container input {
  position: absolute;
  opacity: 0;
  cursor: pointer;
  height: 0;
  width: 0;
}

/* Create a custom checkbox */
.checkmark {
  position: absolute;
  top: 0;
  left: 0;
  height: 12px;
  width: 12px;
  background-color: #fef3df;
}

/* On mouse-over, add a grey background color */
.container:hover input ~ .checkmark {
  background-color: #fbd284;
}

/* When the checkbox is checked, add a blue background */
.container input:checked ~ .checkmark {
  background-color: #d34c01;
}

/* Create the checkmark/indicator (hidden when not checked) */
.checkmark:after {
  content: "";
  position: absolute;
  display: none;
}

/* Show the checkmark when checked */
.container input:checked ~ .checkmark:after {
  display: block;
}

/* Style the checkmark/indicator */
.container .checkmark:after {
  left: 0px;
  top: 0px;
  width: 5px;
  height: 5px;
  border: solid white;
  border-width: 0 3px 3px 0;
  -webkit-transform: rotate(45deg);
  -ms-transform: rotate(45deg);
  transform: rotate(45deg);
}
                    `;
  Jade.set_css(css);
  console.log('CSS done')
}