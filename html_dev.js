function auto_exec(){
}
async function mainHTML(){
  const html = `
<h1>MuVi2 Script Controller</h1>
<h2>Version Beta 2. September 25, 2024</h2>
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
  <label class="container3">Always hide UK Script Unedited
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
<a id='show-hide' onclick="jade_modules.operations.showAdmin()">Show/hide admin</a>
<div id="admin">
  <label class="section-label">Admin</label><br/>
  <button id='btnFormula' onclick="jade_modules.operations.theFormulas()">Formula</button>
  <button id='btnSceneCalc' onclick="jade_modules.operations.fillSceneNumber()">Scene Number</button>
  <button id='btnDefaultColumn' onclick="jade_modules.operations.setDefaultColumnWidths()">Default Columns Widths</button>
</div>
  `;

  await Jade.open_canvas("Script Controller", html, true);
  console.log('Canvas open');
  await jade_modules.operations.getDataFromSheet('Settings','studioChoice','studio-select');
  await jade_modules.operations.getDataFromSheet('Settings','engineerChoice','engineer-select');
  await jade_modules.operations.getColumnData('Settings', 'columnData');
  await jade_modules.operations.initialiseVariables();
  await jade_modules.operations.displayMinAndMax();
  await jade_modules.operations.setDefaultColumnWidths();
  await jade_modules.operations.showHideColumns('all');
  await jade_modules.operations.setUpEvents();

  console.log("I'm here data loaded");
}

 