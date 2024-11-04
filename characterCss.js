function auto_exec(){
}
async function mainCSS(){
  const css =`
body {
  background-color: #ffffb3;
  color: #424200;
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
  background-color: #424200;
  color: #ffffb3;
  font-size: 12px;
  cursor: pointer;
}
h1 {
  margin-left: 15px;
  margin-bottom: 5px;
  font-size: 18px;
}
h2 {
  margin-left: 15px;
  margin-bottom: 10px;
  font-size: 14px;
  font-weight: normal;
}

#admin, #main-navigation, #main-controls, #scene-navigation, #scene-controls {
  margin-left: 15px;
  margin-top: 10px;
  width: 370px;
  padding-left: 10px;
  padding-top: 10px;
  padding-bottom: 10px;
  border-width: 1px;
  border-radius: 5px;
  border-color: #424200;
  border-style: solid;
  display: block;
}
#scene-page, #scene-navigation, #scene-controls {
  border-color: #640000;
}

#main-page, #admin, #scene-page {
  display: none;
}

#wait-message, #admin-wait-message, #scene-wait-message {
  display: none;
  font-size: 12px;
  margin-left: 5px;
}

#show-hide {
  margin-left: 5px;
  font-size: 12px;
  padding-left: 5px;
  cursor: pointer;
}

.section-label {
  margin-left: 5px;
  font-size: 14px;
  font-weight: bold;
}
#character-summary {
  height: 100vh !important;
  padding: 5px;
  background-color: #ffffb3;
  color: #424200;
}
#btnMain, #btnCreateSceneList, #btnSceneScene, #btnSelectAll, #btnSelectNone {
  background-color: #640000;
  color: #ffafaf;
}
/* Customize the label (the container) */
.container-books {
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
.container-books input {
  position: absolute;
  opacity: 0;
  cursor: pointer;
  height: 0;
  width: 0;
}

/* Create a custom checkbox */
.checkmark-books {
  position: absolute;
  top: 0px;
  left: 1px;
  height: 10px;
  width: 10px;
  background-color: #ffafaf;
  border: #640000;
  border-radius: 4px;
  border-width: 2px;
  border-style: solid;
}

/* On mouse-over, add a grey background color */
.container-books:hover input ~ .checkmark-books {
  background-color: #640000;
}

/* When the checkbox is checked, add a blue background */
.container-books input:checked ~ .checkmark-books {
  background-color: #640000;
}

/* Create the checkmark/indicator (hidden when not checked) */
.checkmark-books:after {
  content: "";
  position: absolute;
  display: none;
}

/* Show the checkmark when checked */
.container-books input:checked ~ .checkmark-books:after {
  display: block;
}

/* Style the checkmark/indicator */
.container-books .checkmark-books:after {
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

  `;
  Jade.set_css(css);
  console.log('CSS done')
}