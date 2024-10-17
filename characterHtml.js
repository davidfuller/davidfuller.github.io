function auto_exec(){
}
async function mainHTML(){
  const html = `
<h1>MuVi2 Character Summary</h1>
<h2 id="sheet-version"></h2>
<div id='start-wait'>
  <h1>Please wait...</h1>
</div>
<div id="main-page">
  <button id='btnPrepare' onclick="jade_modules.characterdata.makeTheFullList()">Make the list</button>
  <button id="btnCalculate" onclick="jade_modules.characterdata.whichBooks()">Calculate which books</button>
  <label id="wait-message">Please wait...</label>
</div>
   `;

await Jade.open_canvas("character-summary", html, true);
console.log('Canvas open');
await jade_modules.characterdata.registerExcelEvents();
await jade_modules.characterdata.makeTheFullList();
await jade_modules.characterdata.showMain();
console.log("I'm here data loaded.");
}