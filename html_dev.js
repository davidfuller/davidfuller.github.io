function auto_exec(){
}
async function mainHTML(){
  const html = `
<h1>MuVi2 Script Controller</h1>
<h2>Development edition</h2>
<div id="firstButtons">
  <button id='btnLock' onclick="jade_modules.operations.lockColumns()">Lock sheet</button>
  <button id='btnUnlock' onclick="jade_modules.operations.unlock()">Unlock sheet</button><br/>
  <button id='btnFirst' onclick="jade_modules.operations.firstScene()">First scene</button>
  <button id='btnLast' onclick="jade_modules.operations.lastScene()">Last scene</button><br/>
  <button id='btnPrev' onclick="jade_modules.operations.findScene(-1)">Prev scene</button>
  <button id='btnNext' onclick="jade_modules.operations.findScene(1)">Next scene</button><br/>
  <button id='btnFormula' onclick="jade_modules.operations.theFormulas()">Formula</button>
</div>
<div id="showTakes">
  <div id="takeMessage">Showing all takes</div>
  <button id='btnUnhideAll' onclick="jade_modules.operations.hideRows('all', 'UK')">All Takes</button>
  <button id='btnShowFirst' onclick="jade_modules.operations.hideRows('first', 'UK')">First Takes</button>
  <button id='btnShowLast' onclick="jade_modules.operations.hideRows('last', 'UK')">Last Takes</button>
</div>
<div id="showColumns">
  <div id="columnMessage">Showing all columns</div>
  <button id='btnShowAll' onclick="jade_modules.operations.showHideColumns('all')">All Columns</button>
  <button id='btnShowUK' onclick="jade_modules.operations.showHideColumns('UK Script')">UK Script</button>
  <button id='btnShowUS' onclick="jade_modules.operations.showHideColumns('US Script')">US Script</button>
  <button id='btnShowWalla' onclick="jade_modules.operations.showHideColumns('Walla Script')">Walla Script</button>
</div>
<div id="filterButtons" >
  <button id='btnFilter' onclick="jade_modules.operations.applyFilter()">Apply Filter</button>
  <button id='btnRemoveFilter' onclick="jade_modules.operations.removeFilter()">Remove Filter</button><br/>
</div>
<div id="jump">
  <label for="scene">Jump to scene</label>
  <input type="text" id="scene" name="scene">
  <button id='btnGo' onclick="jade_modules.operations.getTargetSceneNumber()">Go</button><br/>
</div>
<div id="dateStudioEngineer">
  <button id="btnAddTakeUK" onclick="jade_modules.operations.addTakeDetails('UK', true, false, false, false)">Add Take UK</button>
  <button id="btnRemoveTakeUK" onclick="jade_modules.operations.removeTake('UK')">Remove Take UK</button>
  label class="container">Just date
    <input type='radio' id='radJustDate' checked="checked" name='radio'>
    <span class="checkmark"></span>
  </label>
  <label class="container">Details from above
    <input type='radio' id='radAboveDetails' name='radio'>
    <span class="checkmark"></span>
  </label>
  <label class="container">Details from below
    <input type='radio' id='radBelowDetails' name='radio'>
    <span class="checkmark"></span>
  </label>
  <button id="btnTest" onclick="jade_modules.operations.checkboxChecked('radJustDate')">Test</button>
  <div class="row">
    <div class="column" id="column-one">
      <label for="studio-select">Studio</label><br/>
      <label for="engineer-select">Engineer</label>
    </div>
    <div class="column" id="column-two">
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
  `;

  await Jade.open_canvas("Script Controller", html, true);
  console.log('Canvas open');
  await jade_modules.operations.getDataFromSheet('Settings','studioChoice','studio-select');
  await jade_modules.operations.getDataFromSheet('Settings','engineerChoice','engineer-select');
  await jade_modules.operations.getColumnData('Settings', 'columnData');
  await jade_modules.operations.initialiseVariables();
  console.log("I'm here data loaded");
}

 