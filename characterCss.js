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

#main-page, #admin, #scene-page {
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
  display: none;
}
#scene-page {
  border-color: #640000;
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
  height: 100vh;
  padding: 5px;
  background-color: #ffffb3;
  color: #424200;
}
#btnMain, #btnCreateSceneList {
  background-color: #640000;
  color: #ffafaf;
}
  `;
  Jade.set_css(css);
  console.log('CSS done')
}