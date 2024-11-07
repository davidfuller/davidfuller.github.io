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
  <div id="main-navigation">
    <button id="btnMainMain" onclick="jade_modules.characterdata.gotoMain()">Which character in which book</button>
    <button id="btnScene" onclick="jade_modules.characterdata.showScenePage()">Characters in Scenes</button>
    <button id="btnAllCharacters" onclick="jade_modules.characterdata.showAllCharacters()">All Characters</button><br/>
  </div>
  <div id="main-controls">
    <button id="btnSearch" onclick="jade_modules.characterdata.doSearch()">Search</button>
    <label id="wait-message">Please wait...</label><br/>
    <a id='show-hide' onclick="jade_modules.characterdata.showAdmin()">Show/hide admin</a>
  </div>
</div>
<div id="admin">
  <label class="section-label">Admin</label><br/>
  <button id='btnPrepare' onclick="jade_modules.characterdata.gatherData()">Make the list</button>
  <button id="btnRefresh" onclick="jade_modules.characterdata.refreshLinks()">Refresh links</button>
  <button id="btnCreateSceneListAdmin" onclick="jade_modules.characterdata.createSceneList()">Create Scene List</button>
  <label id="admin-wait-message">Please wait...</label>
</div>
<div id="scene-page">
  <div id="scene-navigation">
    <button id="btnMain" onclick="jade_modules.characterdata.gotoMain()">Which character in which book</button>
    <button id="btnSceneScene" onclick="jade_modules.characterdata.showScenePage()">Characters in Scenes</button>
    <button id="btnAllCharactersScene" onclick="jade_modules.characterdata.showAllCharacters()">All Characters</button><br/>
  </div>
  <div id="scene-controls">
    <button id="btnCreateSceneList" onclick="jade_modules.characterdata.createSceneList()">Create Scene List</button><br/>
    <button id="btnSelectAll" onclick="jade_modules.characterdata.selectBooks(true)">Select all books</button>
    <button id="btnSelectNone" onclick="jade_modules.characterdata.selectBooks(false)">Select no books</button>
    <label id="scene-wait-message">Please wait...</label>
    <div id="book-scheckboxes">
      <label class="container-books">Book 1
        <input type="checkbox" id='book-1' checked="checked">
        <span class="checkmark-books"></span>
      </label>
      <label class="container-books">Book 2
        <input type="checkbox" id='book-2' checked="checked">
        <span class="checkmark-books"></span>
      </label>
      <label class="container-books">Book 3
        <input type="checkbox" id='book-3' checked="checked">
        <span class="checkmark-books"></span>
      </label>
      <label class="container-books">Book 4
        <input type="checkbox" id='book-4' checked="checked">
        <span class="checkmark-books"></span>
      </label>
      <label class="container-books">Book 5
        <input type="checkbox" id='book-5' checked="checked">
        <span class="checkmark-books"></span>
      </label>
      <label class="container-books">Book 6
        <input type="checkbox" id='book-6' checked="checked">
        <span class="checkmark-books"></span>
      </label>
      <label class="container-books">Book 7
        <input type="checkbox" id='book-7' checked="checked">
        <span class="checkmark-books"></span>
      </label>
    </div>
  </div>
</div>
<div id="all-characters-page">
  <div id="all-character-navigation">
    <button id="btnMainAllChar" onclick="jade_modules.characterdata.gotoMain()">Which character in which book</button>
    <button id="btnSceneAllChar" onclick="jade_modules.characterdata.showScenePage()">Characters in Scenes</button>
    <button id="btnAllCharactersAllChar" onclick="jade_modules.characterdata.showAllCharacters()">All Characters</button><br/>
  </div>
  <div id="all-character-controls">
    <button id='btnRefreshNames' onclick="jade_modules.characterdata.refreshNames()">Refresh Names</button>
    <label id="all-character-wait-message">Please wait...</label>
  </div>
</div>
   `;

await Jade.open_canvas("character-summary", html, true);
console.log('Canvas open');
await jade_modules.characterdata.registerExcelEvents();
await jade_modules.characterdata.refreshLinks();
await jade_modules.characterdata.gatherData ();
await jade_modules.characterdata.showMain();
await jade_modules.characterdata.gotoMain();
await jade_modules.characterdata.createSceneList();
console.log("I'm here data loaded.");
}