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
  <button id="btnScene" onclick="jade_modules.characterdata.showScenePage()">Characters in Scenes</button><br/>
  <button id="btnSearch" onclick="jade_modules.characterdata.doSearch()">Search</button>
  <label id="wait-message">Please wait...</label><br/>
  <a id='show-hide' onclick="jade_modules.characterdata.showAdmin()">Show/hide admin</a>
</div>
<div id="admin">
  <label class="section-label">Admin</label><br/>
  <button id='btnPrepare' onclick="jade_modules.characterdata.gatherData()">Make the list</button><br/>
  <button id="btnRefresh" onclick="jade_modules.characterdata.refreshLinks()">Refresh links</button>
  <label id="admin-wait-message">Please wait...</label>
</div>
<div id="scene-page">
  <button id="btnMain" onclick="jade_modules.characterdata.gotoMain()">Which character in which book</button><br/>
  <button id="btnCreateSceneList" onclick="jade_modules.characterdata.createSceneList()">Create Scene List</button>
  <label id="scene-wait-message">Please wait...</label>
</div>
   `;

await Jade.open_canvas("character-summary", html, true);
console.log('Canvas open');
await jade_modules.characterdata.registerExcelEvents();
await jade_modules.characterdata.refreshLinks();
await jade_modules.characterdata.gatherData ();
await jade_modules.characterdata.showMain();
await jade_modules.characterdata.gotoMain();
console.log("I'm here data loaded.");
}