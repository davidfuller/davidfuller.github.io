function auto_exec(){
}
async function mainHTML(){
  const html = `
<h1>MuVi2 Script Controller</h1>
<h2 id="sheet-version"></h2>
<div id='start-wait'>
  <h1>Please wait...</h1>
</div>
<div id="main-page">
  <button id="btnForDirectorPage" onclick="jade_modules.operations.showForDirector()">For Director</button>
  <button id="btnForActorPage" onclick="jade_modules.operations.showForActorsPage()">For Actors</button>
  <button id="btnForSchedulingPage" onclick="jade_modules.operations.showForSchedulingPage()">For Scheduling</button>
  <button id="btnWallaImport" onclick="jade_modules.operations.showWallaImportPage()">Walla Import</button>
  <button id="btnLocationPage" onclick="jade_modules.operations.showLocation()">Location</button><br/>
  <a id='show-hide' onclick="jade_modules.operations.showAdmin()">Show/hide admin</a>
  <div id="admin">
    <label class="section-label">Admin</label><br/>
    <button id='btnFormula' onclick="jade_modules.operations.theFormulas()">Formula</button>
    <label id="formula-wait">Please wait...</label>
    <button id='btnSceneCalc' onclick="jade_modules.operations.fillSceneNumber()">Scene Number</button>
    <label id="scene-wait">Please wait...</label>
    <button id='btnDefaultColumn' onclick="jade_modules.operations.setDefaultColumnWidths()">Default Columns Widths</button>
    <button id='btnLoadCharacters' onclick="jade_modules.scheduling.loadReduceAndSortCharacters()">Load characters</button>
    <button id="btnAddHandler" onclick="jade_modules.operations.registerExcelEvents()">Register events</button>
    <button id="btnClearWallal" onclick="jade_modules.operations.clearWalla()">Clear Walla</button>
    <button id="btnWallaCues" onclick="jade_modules.operations.calculateWallaCues()">Create Walla Cues</button>
    <button id="btnCalculateType" onclick="jade_modules.operations.createTypeCodes()">Create Type codes</button><br/>
    <button id="btnDeleteSceneWalla" onclick="jade_modules.operations.deleteAllSceneAndWallaBlocks()">Delete all Scene and Walla blocks</button>
    <div class="row">
      <div class="column" id="column-add-one">
        <button id="btnAddSceneBlock" onclick="jade_modules.operations.addSceneBlock()">Add scene block</button>
        <label for="chapter-scene-select">Chapter/Scene</label><br/>
        <label id='scene-add-wait'>Please wait...</label><br/>
        <button id="btnAddWallaBlockNamed" onclick="jade_modules.operations.getSceneWallaInformation(1)">Add walla block (Named)</button>
        <button id="btnAddWallaBlockUnnamed" onclick="jade_modules.operations.getSceneWallaInformation(2)">Add walla block (Un-named)</button>
        <button id="btnAddWallaBlockGeneral" onclick="jade_modules.operations.getSceneWallaInformation(3)">Add walla block (General)</button>
        <label for="walla-scene">Scene</label><br/>
      </div>
      <div class="column" id="column-add-two">
        <select id="chapter-scene-select"><option value="">Please select</option></select>
        <button id='btnGoChapterScene' onclick="jade_modules.operations.goSceneChapter()">Go</button><br/>
        <button id="btnRefreshList" onclick="jade_modules.operations.fillChapterAndScene()">Refresh List</button>
        <input type="text" id="walla-scene" name="walla-scene">
        <button id='btnGoWallaScene' onclick="jade_modules.operations.goWallaScene()">Go</button>
      </div>
    </div>
  </div>
  <div id="firstButtons">
    <label class="section-label">Lock / Unlock</label><br/>
    <button id='btnLock' onclick="jade_modules.operations.lockColumns()">Lock sheet</button>
    <button id='btnUnlock' onclick="jade_modules.operations.unlock()">Unlock sheet</button><br/>
  </div>
  <div id="filterButtons" >
    <label class="section-label">Filtering</label><br/>
    <button id='btnFilter' onclick="jade_modules.operations.applyFilter()">Apply Filter</button>
    <button id='btnRemoveFilter' onclick="jade_modules.operations.removeFilter()">Remove Filter</button><br/>
  </div>
  <div id="showTakes">
    <label class="section-label">Take selection: </label><span id="takeMessage">Showing all takes</span><br/>
    <button id='btnUnhideAll' onclick="jade_modules.operations.hideRows('all', 'UK')">All Takes</button>
    <button id='btnShowFirst' onclick="jade_modules.operations.hideRows('first', 'UK')">First Takes</button>
    <button id='btnShowLast' onclick="jade_modules.operations.hideRows('last', 'UK')">Last Takes</button>
  </div>
  <div id="showColumns">
    <label class="section-label">Column selection:</label><span id="columnMessage">Showing all columns</span><br/>
    <button id='btnShowAll' onclick="jade_modules.operations.showHideColumns('all')">All Columns</button>
    <button id='btnShowUK' onclick="jade_modules.operations.showHideColumns('UK Script')">UK Script</button>
    <button id='btnShowUS' onclick="jade_modules.operations.showHideColumns('US Script')">US Script</button>
    <button id='btnShowWalla' onclick="jade_modules.operations.showHideColumns('Walla Script')">Walla Script</button>
    <label class="container3">Always hide UK Script without dialog tags
      <input type="checkbox" id='hideUnedited' checked="checked">
      <span class="checkmark3"></span>
    </label>
  </div>
  <div id="jump">
    <label class="section-label">Jump to scene</label><br/>
    <button id='btnFirst' onclick="jade_modules.operations.firstScene()">First scene</button>
    <button id='btnPrev' onclick="jade_modules.operations.findScene(-1)">Prev scene</button>
    <button id='btnNext' onclick="jade_modules.operations.findScene(1)">Next scene</button>
    <button id='btnLast' onclick="jade_modules.operations.lastScene()">Last scene</button><br/>
    <div class="row">
      <div class="column-jump" id="column-jump-one">
        <label id='jump-label-scene' for="scene">Scene No</label><br/>
        <label id='jump-label-line-no' for="lineNo">Cue/line no</label><br/>
        <label id='jump-label-chapter' for="chapter">Chapter</label><br/>
      </div>
      <div class="column-jump" id="column-jump-two">
        <input type="text" id="scene" name="scene">
        <button id='btnGo' onclick="jade_modules.operations.getTargetSceneNumber()">Go</button>
        <span id='min-and-max'></span><br/>
        <input type="text" id="lineNo" name="lineNo">
        <button id='btnGoLine' onclick="jade_modules.operations.getTargetLineNo()">Go</button>
        <span id='min-and-max-lineNo'></span><br/>
        <input type="text" id="chapter" name="chapter">
        <button id='btnGoChapter' onclick="jade_modules.operations.getTargetChapter()">Go</button>
        <span id='min-and-max-chapter'></span><br/>
      </div>
    </div>
  </div>
  <div id="dateStudioEngineer">
    <label class="section-label">Add / remove takes</label><br/>
    <button id="btnAddTakeUK" onclick="jade_modules.operations.addTakeDetails('UK', true)">Add Take UK</button>
    <button id="btnRemoveTakeUK" onclick="jade_modules.operations.removeTake('UK')">Remove Take UK</button><br/>
    <button id="btnAddTakeUS" onclick="jade_modules.operations.addTakeDetails('US', true)">Add Take US</button>
    <button id="btnRemoveTakeUS" onclick="jade_modules.operations.removeTake('US')">Remove Take US</button><br/>
    <button id="btnAddTakeWalla" onclick="jade_modules.operations.addTakeDetails('Walla', true)">Add Take Walla</button>
    <button id="btnRemoveTakeWalla" onclick="jade_modules.operations.removeTake('Walla')">Remove Take Walla</button><br/>
    <label class="container">Just date
      <input type='radio' id='radJustDate' checked="checked" name='radio'>
      <span class="checkmark"></span>
    </label>
    <label class="container">Details from row above
      <input type='radio' id='radAboveDetails' name='radio'>
      <span class="checkmark"></span>
    </label>
    <label class="container">Details from the input below
      <input type='radio' id='radBelowDetails' name='radio'>
      <span class="checkmark"></span>
    </label>
    <div class="row">
      <div class="column" id="column-one">
        <label for="markup">Markup</label>
        <label for="studio-select">Studio</label><br/>
        <label for="engineer-select">Engineer</label>
      </div>
      <div class="column" id="column-two">
        <input type="text" id="markup" name="markup">
        <select id="studio-select"><option value="">Please select</option></select><br/>
        <select id="engineer-select"><option value="">Please select</option></select>
      </div>
    </div>
    <div id="fillButton">
      <button id='btnFillUK' onclick="jade_modules.operations.fill('UK')">Fill UK</button>
      <button id='btnFillUS' onclick="jade_modules.operations.fill('US')">Fill US</button>
      <button id='btnFillWalla' onclick="jade_modules.operations.fill('Walla')">Fill Walla</button><br/>
    </div>
  </div>
</div>
<div id="for-director-page">
  <label id='for-director-label'>For Director Commands</label><br/>
  <button id="btnMainPage" onclick="jade_modules.operations.showMainPage()">Main Page</button>
  <button id="btnForActorPage" onclick="jade_modules.operations.showForActorsPage()">For Actors</button>
  <button id="btnForSchedulingPage" onclick="jade_modules.operations.showForSchedulingPage()">For Scheduling</button><br/>
  <div id="forDirectorButtons">
    <button id='btnGetDirectorInfo' onclick="jade_modules.scheduling.getDirectorInfo()">Calculate for<br/>director information<br/>for selected character</button>
    <button id="btnDirectorGoToLine" onclick="jade_modules.scheduling.directorGoToLine()">Go to this line in script</button>
    <label id="director-wait">Please wait...</label>
  </div>
</div>
<div id="for-actor-page">
  <label id='for-actor-label'>For Actor Commands</label><br/>
  <button id="btnMainPage" onclick="jade_modules.operations.showMainPage()">Main Page</button>
  <button id="btnForDirectorPage" onclick="jade_modules.operations.showForDirector()">For Director</button>
  <button id="btnForSchedulingPage" onclick="jade_modules.operations.showForSchedulingPage()">For Scheduling</button>
  <button id="btnWallaImport" onclick="jade_modules.operations.showWallaImportPage()">Walla Import</button><br/>
  <div id="forActorsButtons">
    <button id='btnGetActorInfo' onclick="jade_modules.scheduling.getActorInfo()">Calculate for<br/>actor information<br/>for selected character</button>
    <button id="btnActorGoToLine" onclick="jade_modules.scheduling.actorGoToLine()">Go to this line in script<br/>(First line if multiple)</button>
    <label id='actor-wait'>Please wait...</label>
  </div>
</div>
<div id="for-scheduling-page">
  <label id='for-scheduling-label'>For Scheduling Commands</label><br/>
  <button id="btnMainPage" onclick="jade_modules.operations.showMainPage()">Main Page</button>
  <button id="btnForDirectorPage" onclick="jade_modules.operations.showForDirector()">For Director</button>
  <button id="btnForActorPage" onclick="jade_modules.operations.showForActorsPage()">For Actors</button><br/>
  <div id="forSchedulingButtons">
    <button id='btnGetSchedulingInfo' onclick="jade_modules.scheduling.getForSchedulingInfo()">Calculate for<br>scheduling information<br>for selected character</button>
    <button id="btnSchedulingGoToLine" onclick="jade_modules.scheduling.schedulingGoToLine()">Goto first line of<br/>the selected scene</button>
    <label id='scheduling-wait'>Please wait...</label>
  </div>
</div>
<div id="walla-import-page">
  <label id='walla-import-label'>Walla Import</label><br/>
  <button id="btnMainPage" onclick="jade_modules.operations.showMainPage()">Main Page</button>
  <button id="btnForDirectorPage" onclick="jade_modules.operations.showForDirector()">For Director</button>
  <button id="btnForActorPage" onclick="jade_modules.operations.showForActorsPage()">For Actors</button><br/>
  <div id="wallaImportButtons">
    <button id='btnParseSource' onclick="jade_modules.wallaimport.parseSource()">Parse source text</button>
    <button id='btnLoadScript' onclick="jade_modules.wallaimport.loadIntoScriptSheet()">Load into script sheet</button>
    <label id='load-message'>This item is already present</label>
  </div>
</div>
<div id="location-page">
  <label id='location-label'>Locations</label><br/>
  <button id="btnMainPage" onclick="jade_modules.operations.showMainPage()">Main Page</button>
  <button id="btnForDirectorPage" onclick="jade_modules.operations.showForDirector()">For Director</button>
  <button id="btnForActorPage" onclick="jade_modules.operations.showForActorsPage()">For Actors</button><br/>
  <div id="locationButtons">
    <button id='btnGetLocationInfo' onclick="jade_modules.scheduling.getLocationInfo()">Get info for<br/>location</button>
    <button id='btnLocationGoTo' onclick="jade_modules.scheduling.locationGoToLine()">Goto first line of<br/>of selected scene</button>
    <label id='location-wait'>Please wait...</label>
  </div>
</div>
  `;

  await Jade.open_canvas("Script-Controller", html, true);
  console.log('Canvas open');
  await jade_modules.operations.getDataFromSheet('Settings','studioChoice','studio-select');
  await jade_modules.operations.getDataFromSheet('Settings','engineerChoice','engineer-select');
  await jade_modules.operations.getColumnData('Settings', 'columnData');
  await jade_modules.operations.initialiseVariables();
  await jade_modules.operations.displayMinAndMax();
  await jade_modules.scheduling.loadReduceAndSortCharacters();
  await jade_modules.operations.setDefaultColumnWidths();
  await jade_modules.operations.showHideColumns('all');
  await jade_modules.operations.setUpEvents();
  await jade_modules.operations.registerExcelEvents();
  await jade_modules.operations.fillChapterAndScene();
  await jade_modules.operations.showMain();

  console.log("I'm here data loaded. Dev5");
}

 