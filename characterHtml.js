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
  <button id='btnPrepare' onclick="jade_modules.characterData.makeTheFullList()">Make the list</button>
</div>
   `;

await Jade.open_canvas("character-summary", html, true);
console.log('Canvas open');

console.log("I'm here data loaded.");
}