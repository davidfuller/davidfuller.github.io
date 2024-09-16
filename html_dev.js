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
  <button id='btnInsertRow' onclick="jade_modules.operations.insertTake('UK', true, false, false, false)">Insert Take</button>
  <button id='btnDeleteRow' onclick="jade_modules.operations.deleteRow()">Delete</button><br/>
  <button id='btnFormula' onclick="jade_modules.operations.theFormulas()">Formula</button>
</div>
<div id="showTakes"></div>
  <div id="takeMessage">Showing all takes</div>
  <button id='btnTest' onclick="jade_modules.operations.hideRows('last', 'UK')">Show Last Take</button><br/>
  <button id='btnUnhideAll' onclick="jade_modules.operations.hideRows('all', 'UK')">Show All Takes</button>
  <button id='btnShowFirst' onclick="jade_modules.operations.hideRows('first', 'UK')">Show First Take</button><br/>
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
  await jade_modules.operations.getDataFromSheet('Settings','studioChoice','studio-select');
  await jade_modules.operations.getDataFromSheet('Settings','engineerChoice','engineer-select');
  await jade_modules.operations.getColumnData('Settings', 'columnData');
  console.log("I'm here data loaded");
}

 