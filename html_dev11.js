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
  <div id="nav-buttons">
    <button id="btnForDirectorPageMain" class="nav-button" onclick="jade_modules.operations.showForDirector()">For Director</button>
    <button id="btnForActorPage" class="nav-button" onclick="jade_modules.operations.showForActorsPage()">For Actors</button>
    <button id="btnForSchedulingPage" class="nav-button" onclick="jade_modules.operations.showForSchedulingPage()">For Scheduling</button>
    <button id="btnWallaImport" class="nav-button" onclick="jade_modules.operations.showWallaImportPage()">Walla Import</button>
    <button id="btnLocationPage" class="nav-button" onclick="jade_modules.operations.showLocation()">Location</button><br/>
    <a id='show-hide' onclick="jade_modules.operations.showAdmin()">Show/hide admin</a>
    <a id='show-hide-comparison' onclick="jade_modules.operations.showComparison()">Show/hide comparison</a>
    <a id='show-hide-us-script' onclick="jade_modules.operations.showUsScript()">Show/hide US Script</a>
  </div>
  <div id="admin">
    <label class="section-label">Admin</label><br/>
    <button id='btnFormula' onclick="jade_modules.operations.theFormulas()">Formula</button>
    <label id="formula-wait">Please wait...</label>
    <button id='btnSceneCalc' onclick="jade_modules.operations.fillSceneNumber()">Scene Number</button>
    <label id="scene-wait">Please wait...</label>
    <button id='btnFindDuplicates' onclick="jade_modules.operations.findDuplicateLineNumbers()">Find Duplicates</button>
    <button id='btnDefaultColumn' onclick="jade_modules.operations.setDefaultColumnWidths()">Default Columns Widths</button>
    <button id='btnLoadCharacters' onclick="jade_modules.scheduling.loadReduceAndSortCharacters()">Load characters</button>
    <button id="btnAddHandler" onclick="jade_modules.operations.registerExcelEvents()">Register events</button>
    <button id="btnClearWallal" onclick="jade_modules.operations.clearWalla()">Clear Walla</button>
    <button id="btnWallaCues" onclick="jade_modules.operations.calculateWallaCues()">Create Walla Cues</button>
    <button id="btnCalculateType" onclick="jade_modules.operations.createTypeCodes()">Create Type codes</button><br/>
    <button id="btnDeleteSceneWalla" onclick="jade_modules.operations.deleteAllSceneAndWallaBlocks()">Delete all Scene and Walla blocks</button><br/>
    <button id="btnChangeFills" onclick="jade_modules.operations.fillColorLinesAndScriptedWalla()">Change fill for<br/>lines and scripted walla</button>
    <label id="character-wait">Please wait...</label>
    <button id="btnProcessCharacterListForWordAndScene" onclick="jade_modules.scheduling.processCharacterListForWordAndScene()">Character Word Count</button>
    <button id="btnAddScneWordCount" onclick="jade_modules.scheduling.createSceneWordCountData()">Add Scene Word Count<br/>to Character List</button>
    <button id="btnCheckScenes" onclick="jade_modules.operations.checkAllTheSceneBreaks()">Check all the<br/>scene breaks</button>
    <button id="btnCheckCopyScenes" onclick="jade_modules.operations.copyTextV2(false, true)">Check the new text V2</button>
    <button id="btnCheckCopyScenesNoComparison" onclick="jade_modules.operations.copyTextV2(false, false)">Check the new text V2<br/>No comparison report</button>
    <button id="btnCopyScenes" onclick="jade_modules.operations.copyTextV2(true, false)">Copy the new text V2</button>
    <button id="btnDeleteFX" onclick="jade_modules.operations.deleteAllFX()">Delete all FX</button>
    <button id="btnReconcileLocation" onclick="jade_modules.operations.reconcileLocations(false)">Reconcile locations<br/>no correction</button>
    <button id="btnReconcileLocationCorrection" onclick="jade_modules.operations.reconcileLocations(true)">Reconcile locations<br/>WITH correction</button>
    <button id="btnAutoSceneBlock" onclick="jade_modules.operations.autoSceneBlockCreation()">Auto scene block</button>
    <button id="btnLoadNewSheetCharacters" onclick="jade_modules.operations.loadNewSheetCharacters()">Load new sheet characters</button>
    <button id="btnNewCharacters" onclick="jade_modules.operations.newCharacters()">Missing characters</button>
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
        <button id='btnGoChapterScene' onclick="jade_modules.operations.goSceneChapter()">Go</button>
        <button id="btnRefreshList" onclick="jade_modules.operations.fillChapterAndScene()">Refresh List</button>
        <br/>
        <input type="text" id="walla-scene" name="walla-scene">
        <button id='btnGoWallaScene' onclick="jade_modules.operations.goWallaScene()">Go</button>
      </div>
    </div>
  </div>
  <div id="comparison">
    <button id="btnCreateChapters" onclick="jade_modules.comparison.createChapters()">Create Chapters</button>
    <button id="btnReadDecision" onclick="jade_modules.comparison.createResult()">Create Results</button>
    <button id="btnChapterAndResult" onclick="jade_modules.comparison.createChaptersAndResults()">Chapter + Result</button><br/>
    <label for="chapter-compare-select">Chapter</label><select id="chapter-compare-select"><option value="">Please select chapter</option></select><br/>
    <button id="btnArchiveComparison" onclick="jade_modules.comparison.copySheets()">Archive Comparison Sheets</button>
    <button id="btnClearDecisionAndResult" onclick="jade_modules.comparison.clearDecisionAndResult()">Clear Decision And Result</button>
    <button id="btnCreateScriptLinks" onclick="jade_modules.comparison.getLinksToTextFromChapter()">Create Script Links</button>
    <button id="btnSelectResultLowestTrue" onclick="jade_modules.comparison.selectResultLowestTrue()">Lowest True</button>
    <button id="btnActiveCellTextArea" onclick="jade_modules.comparison.putSelectedCellInTextArea()">Active cell search text</button>
    <textarea id="search-text" cols="40" rows="8"></textarea>
    <button id="btnCorrectTextLf" class="compButton" onclick="jade_modules.comparison.correctTextReplaceLF(true)">Replace Line-Feed</button>
    <button id="btnCorrectTextSpaceQuote" class="compButton" onclick="jade_modules.comparison.correctTextSpaceQuotes(true)">Correct Space Quotes</button>
    <button id="btnFindDecision"  class="compButton"onclick="jade_modules.comparison.findSearchTextInDecision()">Find text in Decision</button>
    <button id="btnFindInPDF" class="compButton" onclick="jade_modules.comparison.findInPDF()">Find in PDF Comparison</button><br/>
    <button id="btnFindNext" class="compButton" onclick="jade_modules.comparison.fixNextIssue()">Fix Next Issue</button>
    <button id="btnComparisonLoop"class="compButton" onclick="jade_modules.comparison.comparisonLoop()">Auto Run Comparison</button><br/>
    <button id="btnAutoChapter" class="compButton" onclick="jade_modules.comparison.autoSelectChapter()">Auto Chapter</button>
    <button id="btnDoKeeps" class="compButton" onclick="jade_modules.comparison.doKeepsAndManuals()">Do keeps</button><br/>
    <button id="btnDoKeepsResults" class="compButton" onclick="jade_modules.comparison.doKeepsFollowedByResults()">Keeps/Results/Autorun</button>
    <button id="btnCellJoin" class="compButton" onclick="jade_modules.comparison.dealWithCellJoin()">Cell Join</button><br/>
    <button id="btnCellMerge" class="compButton" onclick="jade_modules.comparison.mergeCells()">Cell Merge & Autorun</button>
    <button id="btnFindMergeAutoRun" class="compButton" onclick="jade_modules.comparison.findMergeAutorun()">Find & Merge & Autorun</button>
  </div>
  <div id="us-script">
    <button id="btnAddUsScript" onclick="jade_modules.usscript.usScriptAdd()">Add US Script</button>
    <button id="btnClearUsCueScript" onclick="jade_modules.operations.clearUsCueAndScript()">Clear US Cue and Script</button>
  </div>
  <div id="firstButtons">
    <label class="section-label" onclick="jade_modules.operations.showLock()">Lock / Unlock</label><span id="lockMessage"></span><a id='show-hide-lock' onclick="jade_modules.operations.showLock()">Show/hide lock</a>
    <div id='lock-buttons'>
      <button id='btnLock' onclick="jade_modules.operations.lockColumns()">Lock sheet</button>
      <button id='btnUnlock' onclick="jade_modules.operations.unlock()">Unlock sheet</button><br/>
    </div>
  </div>
  <div id="filterButtons" >
    <label class="section-label" onclick="jade_modules.operations.showFilter()">Filtering</label><a id='show-hide-filter' onclick="jade_modules.operations.showFilter()">Show/hide filtering</a>
    <div id='filter-buttons'>
      <button id='btnFilter' onclick="jade_modules.operations.applyFilter()">Apply Filter</button>
      <button id='btnRemoveFilter' onclick="jade_modules.operations.removeFilter()">Remove Filter</button><br/>
    </div>
  </div>
  <div id="jump">
    <label class="section-label" onclick="jade_modules.operations.showJump()">Jump to scene</label><a id='show-hide-jump' onclick="jade_modules.operations.showJump()">Show/hide jumping</a>
    <div id="jump-buttons">
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
  </div>
  <div id="showColumns">
    <label class="section-label" onclick="jade_modules.operations.showColumns()">Column selection:</label><span id="columnMessage">Showing all columns</span><a id='show-hide-columns' onclick="jade_modules.operations.showColumns()">Show/hide columns</a>
    <div id="column-buttons">
      <button id='btnDefaultColumn' onclick="jade_modules.operations.setDefaultColumnWidths()">Default Columns Widths</button><br/>
      <button id='btnShowAll' onclick="jade_modules.operations.showHideColumns('all')">All Columns</button>
      <button id='btnShowUK' onclick="jade_modules.operations.showHideColumns('UK Script')">UK Script</button>
      <button id='btnShowUS' onclick="jade_modules.operations.showHideColumns('US Script')">US Script</button>
      <button id='btnShowWalla' onclick="jade_modules.operations.showHideColumns('Walla Script')">Walla Script</button>
      <button id='btnShowScriptSupervisor' onclick="jade_modules.operations.showHideColumns('Script Supervisor')">Script Supervisor</button>
      <label class="container3">Always hide UK Script without dialog tags
        <input type="checkbox" id='hideUnedited' checked="checked">
        <span class="checkmark3"></span>
      </label>
    </div>
  </div>
  <div id="showTakes">
    <label class="section-label" onclick="jade_modules.operations.showTakes()">Take selection: </label><span id="takeMessage">Showing all takes</span><a id='show-hide-takes' onclick="jade_modules.operations.showTakes()">Show/hide takes</a>
    <div id="takes-buttons">
      <button id='btnUnhideAll' onclick="jade_modules.operations.hideRows('all', 'UK')">All Takes</button>
      <button id='btnShowFirst' onclick="jade_modules.operations.hideRows('first', 'UK')">First Takes</button>
      <button id='btnShowLast' onclick="jade_modules.operations.hideRows('last', 'UK')">Last Takes</button><br/>
      <button id='btnShowTakesOnly' onclick="jade_modules.operations.hideRows('takeOnly', 'UK')">Takes Only</button>
      <button id='btnShowFirstOnly' onclick="jade_modules.operations.hideRows('firstOnly', 'UK')">First Takes Only</button>
      <button id='btnShowLastOnly' onclick="jade_modules.operations.hideRows('lastOnly', 'UK')">Last Takes Only</button><br/>
      <button id='btnColouredTakes' onclick="jade_modules.operations.doColourTakes()">Coloured Takes</button><br/>
    </div>
  </div>
  <div id="dateStudioEngineer">
    <label class="section-label" onclick="jade_modules.operations.showAddRemove()">Add / remove takes</label><a id='show-hide-add-remove' onclick="jade_modules.operations.showAddRemove()">Show/hide add/remove takes</a>
    <div id="add-remove-buttons">
      <label id="take-wait">Please wait...</label>
      <label id="take-message"></label>
      <div id="message-box">
        <div id="message-prompt">Are you sure want to clear these takes</div>
        <div id="buttons-yes-no">
          <button id="message-yes" onclick="jade_modules.operations.messageYes()">Yes</button>
          <button id="message-no" onclick="jade_modules.operations.messageNo()">No</button>
        </div>
      </div>
      <div id="hidden-old-takes">
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
      </div>
      <div class="row">
        <div class="column" id="column-one">
          <label for="character-select">Character</label><br/>
          <br/>
          <label for="takes-select">No. takes</label><br/>
          <label for="studio-select">Studio</label><br/>
          <label for="engineer-select">Engineer</label>
          <label for="markup">Markup</label>
        </div>
        <div class="column" id="column-two">
          <select id="character-select"><option value="">Please select</option></select>
          <button id="btnFilterCharacter" onclick="jade_modules.operations.filterCharacter()">Filter</button>
          <button id="btnClearFilterCharacter" onclick="jade_modules.operations.removeFilter()">Clear</button><br/>
          <label class="container4">Show scene blocks
            <input type="checkbox" id='show-scene-blocks' checked="checked">
            <span class="checkmark4"></span>
          </label>
          <select id="takes-select"><option value=""></option></select><br/>
          <select id="studio-select"><option value="">Please select</option></select><br/>
          <select id="engineer-select"><option value="">Please select</option></select>
          <input type="text" id="markup" name="markup"><button id="btnDefaultMarkup" onclick="jade_modules.operations.addDefaultMarkUp()">Default</button><br/>
          <div id="add-remove-markup">
            <button id="btnAddMarkup" onclick="jade_modules.operations.addMarkUpToSelected(false)">Add</button>
            <button id="btnReplaceMarkup" onclick="jade_modules.operations.addMarkUpToSelected(true)">Replace</button><br/>
            <label id="markup-message">The Add/Replace buttons allow you to add or replace text to the selected markup fields</label>
          </div>
        </div>
      </div>
      <div>
        <button id="btnApplyTakeUk" class="take-button" onclick="jade_modules.operations.applyTakeDetails('UK')">Apply Take UK</button>
        <button id="btnApplyTakeUs" class="take-button" onclick="jade_modules.operations.applyTakeDetails('US')">Apply Take US</button>
        <button id="btnApplyTakeWalla" class="take-button" onclick="jade_modules.operations.applyTakeDetails('Walla')">Apply Take Walla</button><br/>
        <button id="btnClearTakeUk" class="take-button" onclick="jade_modules.operations.clearTakeDetails('UK')">Clear Take UK</button>
        <button id="btnClearTakeUs" class="take-button" onclick="jade_modules.operations.clearTakeDetails('US')">Clear Take US</button>
        <button id="btnClearTakeWalla" class="take-button" onclick="jade_modules.operations.clearTakeDetails('Walla')">Clear Take Walla</button>
      </div>
          
      <div id="fillButton">
        <button id='btnFillUK' onclick="jade_modules.operations.fill('UK')">Fill UK</button>
        <button id='btnFillUS' onclick="jade_modules.operations.fill('US')">Fill US</button>
        <button id='btnFillWalla' onclick="jade_modules.operations.fill('Walla')">Fill Walla</button><br/>
      </div>
    </div>
  </div>
</div>
<div id="for-director-page">
  <label id='for-director-label'>For Director Commands</label><br/>
  <div id="director-nav-buttons">
    <button id="btnMainPage" class="director-nav" onclick="jade_modules.operations.showMainPage()">Main Page</button>
    <button id="btnForActorPage" class="director-nav" onclick="jade_modules.operations.showForActorsPage()">For Actors</button>
    <button id="btnForSchedulingPage" class="director-nav" onclick="jade_modules.operations.showForSchedulingPage()">For Scheduling</button>
    <button id="btnWallaImport" class="director-nav" onclick="jade_modules.operations.showWallaImportPage()">Walla Import</button>
    <button id="btnLocationPage" class="director-nav" onclick="jade_modules.operations.showLocation()">Location</button><br/>
  </div>
  <div id="forDirectorButtons">
    <button id='btnGetDirectorInfo' class='director-button' onclick="jade_modules.scheduling.getDirectorInfo()">Calculate for<br/>director information<br/>for selected character</button>
    <button id="btnDirectorGoToLine" class='director-button' onclick="jade_modules.scheduling.directorGoToLine()">Go to this line in script</button>
    <label id="director-wait">Please wait...</label>
  </div>
</div>
<div id="for-actor-page">
  <label id='for-actor-label'>For Actor Commands</label><br/>
  <div id="actor-nav-buttons">
    <button id="btnMainPage" class="actor-nav" onclick="jade_modules.operations.showMainPage()">Main Page</button>
    <button id="btnForDirectorPageActor" class="actor-nav" onclick="jade_modules.operations.showForDirector()">For Director</button>
    <button id="btnForSchedulingPage" class="actor-nav" onclick="jade_modules.operations.showForSchedulingPage()">For Scheduling</button>
    <button id="btnWallaImport" class="actor-nav" onclick="jade_modules.operations.showWallaImportPage()">Walla Import</button>
    <button id="btnLocationPage" class="actor-nav" onclick="jade_modules.operations.showLocation()">Location</button><br/>
  </div>
  <div id="forActorsButtons">
    <button id='btnGetActorInfo' class='actor-button' onclick="jade_modules.scheduling.searchCharacter()">Search Character</button>
    <button id="btnActorGoToLine" class='actor-button' onclick="jade_modules.scheduling.actorGoToLine()">Go to this line in script<br/>(First line if multiple)</button>
    <label id='actor-wait'>Please wait...</label>
  </div>
  <div id="scriptButtons">
    <div class="row-script">
      <div class="column-script" id="column-one-script">
        <button id="btnActorScript" class='actor-button' onclick="jade_modules.scheduling.createScript()">Create Actor Script</button>
        <label class="container-actor">All scenes
          <input type='radio' id='radAllScenes' name='radio' onclick="jade_modules.scheduling.displayScenes()">
          <span class="checkmark-actor"></span>
        </label>
        <label class="container-actor">Highlighted scenes
          <input type='radio' id='radHighlighted' checked="checked" name='radio' onclick="jade_modules.scheduling.displayScenes()">
          <span class="checkmark-actor"></span>
        </label>
      </div>
      <div class="column-script" id="column-two-script">
        <label id="actor-scene-heading">Scenes</label><br/>
        <span id="actor-scene-display"></span>
      </div>
    </div>
    <label id='script-wait'>Please wait...</label>
  </div>
</div>
<div id="for-scheduling-page">
  <label id='for-scheduling-label'>For Scheduling Commands</label><br/>
  <div id="scheduling-nav-buttons">
    <button id="btnMainPage" class="scheduling-nav" onclick="jade_modules.operations.showMainPage()">Main Page</button>
    <button id="btnForDirectorPageScheduling" class="scheduling-nav" onclick="jade_modules.operations.showForDirector()">For Director</button>
    <button id="btnForActorPage" class="scheduling-nav" onclick="jade_modules.operations.showForActorsPage()">For Actors</button>
    <button id="btnWallaImport" class="scheduling-nav" onclick="jade_modules.operations.showWallaImportPage()">Walla Import</button>
    <button id="btnLocationPage" class="scheduling-nav" onclick="jade_modules.operations.showLocation()">Location</button><br/>
  </div>
  <div id="forSchedulingButtons">
    <button id='btnGetSchedulingInfo' class="scheduling-button" onclick="jade_modules.scheduling.getForSchedulingInfo()">Calculate for<br>scheduling information<br>for selected character</button>
    <button id="btnSchedulingGoToLine" class="scheduling-button" onclick="jade_modules.scheduling.schedulingGoToLine()">Goto first line of<br/>the selected scene</button>
    <label id='scheduling-wait'>Please wait...</label>
  </div>
</div>
<div id="walla-import-page">
  <label id='walla-import-label'>Walla Import</label><br/>
  <div id="walla-nav-buttons">
    <button id="btnMainPage" class="walla-nav" onclick="jade_modules.operations.showMainPage()">Main Page</button>
    <button id="btnForDirectorPageWalla" class="walla-nav" onclick="jade_modules.operations.showForDirector()">For Director</button>
    <button id="btnForActorPage" class="walla-nav" onclick="jade_modules.operations.showForActorsPage()">For Actors</button>
    <button id="btnForSchedulingPage" class="walla-nav" onclick="jade_modules.operations.showForSchedulingPage()">For Scheduling</button>
    <button id="btnLocationPage" class="walla-nav" onclick="jade_modules.operations.showLocation()">Location</button><br/>
  </div>
  <div id="wallaImportButtons">
    <button id='btnParseSource' class="walla-button" onclick="jade_modules.wallaimport.parseSource()">Parse source text</button>
    <button id='btnLoadScript' class="walla-button" onclick="jade_modules.wallaimport.loadIntoScriptSheet()">Load into script sheet</button>
    <label id='load-message'>This item is already present</label>
  </div>
</div>
<div id="location-page">
  <label id='location-label'>Locations</label><br/>
  <div id="location-nav-buttons">
    <button id="btnMainPage" class="location-nav" onclick="jade_modules.operations.showMainPage()">Main Page</button>
    <button id="btnForDirectorPageLocation" class="location-nav" onclick="jade_modules.operations.showForDirector()">For Director</button>
    <button id="btnForActorPage" class="location-nav" onclick="jade_modules.operations.showForActorsPage()">For Actors</button>
    <button id="btnForSchedulingPage" class="location-nav" onclick="jade_modules.operations.showForSchedulingPage()">For Scheduling</button>
    <button id="btnWallaImport" class="location-nav" onclick="jade_modules.operations.showWallaImportPage()">Walla Import</button><br/>
  </div>
  <div id="locationButtons">
    <button id='btnGetLocationInfo' class="location-button" onclick="jade_modules.scheduling.getLocationInfo()">Get info for<br/>location</button>
    <button id='btnLocationGoTo' class="location-button" onclick="jade_modules.scheduling.locationGoToLine()">Goto first line of<br/>of selected scene</button>
    <label id='location-wait'>Please wait...</label>
  </div>
</div>
<div id="script-page">
  <label id='script-label'>Script</label><br/>
  <div id="script-nav-buttons">
    <button id="btnMainPageScript" class="script-nav" onclick="jade_modules.operations.showMainPage()">Main Page</button>
    <button id="btnForDirectorPageScript" class="script-nav" onclick="jade_modules.operations.showForDirector()">For Director</button>
    <button id="btnForActorPageScript" class="script-nav" onclick="jade_modules.operations.showForActorsPage()">For Actors</button>
    <button id="btnForSchedulingPageScript" class="script-nav" onclick="jade_modules.operations.showForSchedulingPage()">For Scheduling</button>
    <button id="btnWallaImportScript" class="script-nav" onclick="jade_modules.operations.showWallaImportPage()">Walla Import</button>
    <button id="btnLocationPageScript" class="script-nav" onclick="jade_modules.operations.showLocation()">Location</button><br/>
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
  await jade_modules.operations.hiddenRows();
  await jade_modules.operations.fillChapterAndScene();
  await jade_modules.operations.showMain();
  await jade_modules.operations.gatherTakeInformation(false);
  await jade_modules.comparison.fillChapter();
  jade_modules.operations.setDefaultRadioButton();
  await jade_modules.comparison.autoSelectChapter();
  await jade_modules.operations.fillCharacterAndTakesDropdowns();

  console.log("I'm here data loaded. Dev10");
}

 