function auto_exec(){
}
async function mainCSS(){
  const css =`
body {
  background-color: #ffffb3;
  color: #424200;
  font-family: 'Aptos Narrow', 'Arial Narrow', 'Arial', 'Helvetica', 'sans-serif';
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
  font-family: 'Aptos Narrow', 'Arial Narrow', 'Arial', 'Helvetica', 'sans-serif';
  font-size: 18px;
}
h2 {
  margin-left: 15px;
  margin-bottom: 10px;
  font-family: 'Aptos Narrow', 'Arial Narrow', 'Arial', 'Helvetica', 'sans-serif';
  font-size: 14px;
  font-weight: normal;
}
  `;
  Jade.set_css(css);
  console.log('CSS done')
}