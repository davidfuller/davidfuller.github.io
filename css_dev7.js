function auto_exec(){
}
async function mainCSS(){
  const css =`
body {
  background-color: #d8dfe5;
  color: #46656F;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
}
button {
  margin-left: 5px;
  margin-bottom: 5px;
  margin-top: 5px;
  padding-left: 5px;
  padding-right: 5px;
  height: 24px;
  border: none;
  width: fit-content;
  border-radius: 5px;
  background-color: #46656F;
  color: #fef3df;
  font-size: 12px;
  cursor: pointer;
}
h1 {
  margin-left: 15px;
  margin-bottom: 5px;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
  font-size: 18px;
}
h2 {
  margin-left: 15px;
  margin-bottom: 10px;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
  font-size: 14px;
  font-weight: normal;
}
label {
  margin-left: 5px;
  font-size: 12px;
}
/* Section formatting */
#firstButtons, #filterButtons, #dateStudioEngineer, #showTakes, #showColumns, #jump, #admin, #forDirectorButtons, #forActorsButtons, #forSchedulingButtons, #wallaImportButtons, #locationButtons {
  margin-left: 15px;
  margin-top: 10px;
  width: 370px;
  padding-left: 10px;
  padding-top: 0px;
  border-width: 1px;
  border-radius: 5px;
  border-color: #46656F;
  border-style: solid;
}

#admin {
  display: none;
}

#btnGo, #btnGoLine, #btnGoChapter {
  width: fit-content;
  margin-left: 5px;
}

#btnFillUK, #btnFillUS, #btnFillWalla{
  margin-left: 10px;
  margin-bottom: 5px;
  margin-top: 5px;
  padding-left: 10px;
  padding-right: 10px;
  border: none;
  width: fit-content;
  border-radius: 4px;
  background-color: #46656F;
  color: #fef3df;
  font-size: 12px;
  cursor: pointer;
}
#scene, #lineNo, #chapter, #add-chapter {
  width: 50px;
  background-color: #d8dfe5;
  color: #46656F;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
  border-width: 1px;
  border-radius: 5px;
  border-color: #46656F;
  border-style: solid;
}

/*
#chkAboveDetails {
  background-color: #d8dfe5;
  color: #46656F;
  border-width: 1px;
  border-radius: 5px;
  border-color: #46656F;
  border-style: solid;
}

#lblAboveDetails {
  background-color: #d8dfe5;
  color: #46656F;
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
#studio, #engineer, #studio-select, #engineer-select, #chapter-scene-select, #markup, #walla-scene {
  width: 150px;
  background-color: #d8dfe5;
  color: #46656F;
  font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
  border-width: 1px;
  border-radius: 5px;
  border-color: #46656F;
  border-style: solid;
}

#markup {
  width: 200px;
}

.row {
  display: flex;
}

.column {
  padding: 10px;
}

.column-jump {
  padding-left: 10px;
}

#column1 {
  flex: 20%
}
#column2 {
  flex: 60%
}

#column-add-one {
  flex: 65%;
  padding: 0px;
}

#column-add-two {
  flex: 35%;
  padding: 7px;
}

#column-jump-one {
  flex: 5%
}
#column-jump-two {
  flex: 60%
}
#takeMessage, #columnMessage, #lockMessage {
  font-size: 14px;
  margin-left: 10px;
}

.section-label {
  font-size: 14px;
  font-weight: bold;
}

#btnUnhideAll, #btnShowLast, #btnShowFirst, #btnShowAll, #btnShowUK, #btnShowUS, #btnShowWalla {
  margin-left: 10px;
  margin-bottom: 5px;
  margin-top: 5px;
  padding-left: 10px;
  padding-right: 10px;
  border: none;
  width: fit-content;
  border-radius: 4px;
  background-color: #46656F;
  color: #fef3df;
  font-size: 12px;
  cursor: pointer;
}

 #btnAddTakeUK, #btnRemoveTakeUK, #btnAddTakeUS, #btnRemoveTakeUS, #btnAddTakeWalla, #btnRemoveTakeWalla {
  margin-left: 4px;
  margin-bottom: 5px;
  margin-top: 5px;
  padding-left: 10px;
  padding-right: 10px;
  border: none;
  width: 130px;
  border-radius: 4px;
  background-color: #46656F;
  color: #fef3df;
  font-size: 12px;
  cursor: pointer;
}
/* Start of checkbox stuff */
/* Customize the label (the container) */
.container {
  display: block;
  position: relative;
  padding-left: 10px;
  margin-left: 10px;
  margin-bottom: 0px;
  cursor: pointer;
  font-size: 12px;
  /*left: 216px;
  top: -26px;*/
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
  top: 2px;
  left: -4px;
  height: 8px;
  width: 8px;
  background-color: #d8dfe5;
  border: #46656F ;
  border-radius: 50%;
  border-width: 2px;
  border-style: solid;
}

/* On mouse-over, add a grey background color */
.container:hover input ~ .checkmark {
  background-color: #46656F7f;
}

/* When the checkbox is checked, add a blue background */
.container input:checked ~ .checkmark {
  background-color: #46656F;
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
  left: 2px;
  top: 2px;
  width: 4px;
  height: 4px;
  border-radius: 50%;
  background: white;
}
/* *********CHECKBOX ******** */
/* Customize the label (the container) */
.container2 {
  display: block;
  position: relative;
  padding-left: 10px;
  margin-left: 10px;
  margin-bottom: 0px;
  cursor: pointer;
  font-size: 12px;
  -webkit-user-select: none;
  -moz-user-select: none;
  -ms-user-select: none;
  user-select: none;
}

/* Hide the browser's default checkbox */
.container2 input {
  position: absolute;
  opacity: 0;
  cursor: pointer;
  height: 0;
  width: 0;
}

/* Create a custom checkbox */
.checkmark2 {
  position: absolute;
  top: 2px;
  left: -4px;
  height: 8px;
  width: 8px;
  background-color: #d8dfe5;
  border: #46656F ;
  border-radius: 50%;
  border-width: 2px;
  border-style: solid;
}

/* On mouse-over, add a grey background color */
.container2:hover input ~ .checkmark2 {
  background-color: #46656F7f;
}

/* When the checkbox is checked, add a blue background */
.container2 input:checked ~ .checkmark2 {
  background-color: #46656F;
}

/* Create the checkmark/indicator (hidden when not checked) */
.checkmark2 {
  content: "";
  position: absolute;
  display: none;
}

/* Show the checkmark when checked */
.container2 input:checked ~ .checkmark2:after {
  display: block;
}

/* Style the checkmark/indicator */
.container2 .checkmark2:after {
  left: 2px;
  top: 2px;
  width: 4px;
  height: 4px;
  border-radius: 50%;
  background: white;
  -webkit-transform: rotate(45deg);
  -ms-transform: rotate(45deg);
  transform: rotate(45deg);
}
/* Customize the label (the container) */
.container3 {
  display: block;
  position: relative;
  padding-left: 18px;
  margin-left: 10px;
  margin-bottom: 10px;
  cursor: pointer;
  font-size: 12px;
  -webkit-user-select: none;
  -moz-user-select: none;
  -ms-user-select: none;
  user-select: none;
}

/* Hide the browser's default checkbox */
.container3 input {
  position: absolute;
  opacity: 0;
  cursor: pointer;
  height: 0;
  width: 0;
}

/* Create a custom checkbox */
.checkmark3 {
  position: absolute;
  top: 0px;
  left: 1px;
  height: 10px;
  width: 10px;
  background-color: #d8dfe5;
  border: #46656F ;
  border-radius: 4px;
  border-width: 2px;
  border-style: solid;
}

/* On mouse-over, add a grey background color */
.container3:hover input ~ .checkmark3 {
  background-color: #46656F7f;
}

/* When the checkbox is checked, add a blue background */
.container3 input:checked ~ .checkmark3 {
  background-color: #46656F;
}

/* Create the checkmark/indicator (hidden when not checked) */
.checkmark3:after {
  content: "";
  position: absolute;
  display: none;
}

/* Show the checkmark when checked */
.container3 input:checked ~ .checkmark3:after {
  display: block;
}

/* Style the checkmark/indicator */
.container3 .checkmark3:after {
  left: 3px;
  top: 0px;
  width: 2px;
  height: 6px;
  border: solid white;
  border-width: 0 3px 3px 0;
  -webkit-transform: rotate(45deg);
  -ms-transform: rotate(45deg);
  transform: rotate(45deg);
}

#min-and-max, #min-and-max-lineNo, #min-and-max-chapter, #add-min-and-max-chapter, .jump-label {
  font-size: 13px;
}

#jump-label-line-no, #jump-label-chapter, #jump-label-scene, #add-scene-label-chapter {
  position: relative;
  font-size: 13px;
}
#jump-label-line-no {
  top: 14px;
}
#jump-label-chapter {
  top: 24px;
}
#jump-label-scene {
  top: 3px;
}

#show-hide {
  margin-left: 15px;
  font-size: 10px;
  padding-left: 5px;
  cursor: pointer;
}

#main-page {
  display: none;
}

#start-wait {
  display: block;
}

#for-director-page, #for-actor-page, #for-scheduling-page, #walla-import-page, #location-page {
  display: none;
}

#for-director-label, #for-actor-label, #for-scheduling-label, #walla-import-label, #location-label {
  font-size: 22px;
  font-weight: bold;
  margin-left: 15px
}

#btnMainPage, #btnForDirectorPage, #forActorsButtons {
  margin-left: 15px;
}

#btnGetSchedulingInfo, #btnGetActorInfo, #btnGetDirectorInfo, #btnActorGoToLine, #btnDirectorGoToLine, #btnSchedulingGoToLine, #btnAddSceneBlock, #btnGetLocationInfo, #btnLocationGoTo, #btnGetActorText, #btnChangeFills {
  height: auto;
  padding: 10px;
}

#director-wait, #actor-wait, #scheduling-wait, #formula-wait, #scene-wait, #scene-add-wait, #load-message{
  display: none;
  font-size: 15px;
  margin-left: 15px;
}
#fillButton {
  display: none;
}
#Script-Controller {
  height: 100vh;
  padding: 5px;
  background-color: #d8dfe5;
  color: #46656F;
}
#sheet-version {
  font-size: 12px;
}
.actor-button {
  background-color: #592509;
  color: #fbe2d5; 
  vertical-align: top;
}

#forActorsButtons {
  border-color: #592509;
}

#forDirectorButtons{
  border-color: #06394d;
}
#forSchedulingButtons {
  border-color: #1d3a10;
}

#wallaImportButtons {
  border-color: #481343;
}

#locationButtons {
  border-color: #0d3714;
}

.actor-nav{
  background-color: #592509;
  color: #fbe2d5; 
  margin-left: 10px;
  margin-right: 0px;
}

.director-button {
  color: #caedfb;
  background-color: #06394d; 
  vertical-align: top;
}

.director-nav{
  background-color: #06394d;
  color: #caedfb; 
  margin-left: 10px;
  margin-right: 0px;
}

.scheduling-button {
  background-color: #1d3a10;
  color: #daf2d0; 
  vertical-align: top;
}
.scheduling-nav{
  background-color: #1d3a10;
  color: #daf2d0; 
  margin-left: 10px;
  margin-right: 0px;
}
.walla-button {
  background-color: #481343;
  color: #f2ceef; 
  vertical-align: top;
}
.walla-nav{
  background-color: #481343;
  color: #f2ceef; 
  margin-left: 10px;
  margin-right: 0px;
}
.location-button {
  background-color: #0d3714;
  color: #c1f0c8; 
  vertical-align: top;
}
.location-nav{
  background-color: #0d3714;
  color: #c1f0c8; 
  margin-left: 10px;
  margin-right: 0px;
}

#btnActorScript {
  display: block;
}
`;
  Jade.set_css(css);
  console.log('CSS done')
}
