async function mainHTML(){
  const html = `<h1>MuVi2 Script Controller</h1>
                <div id="firstButtons">
                  <button id='btnLock' onclick="jade_modules.operations.lockColumns()">Lock sheet</button>
                  <button id='btnUnlock' onclick="jade_modules.operations.unlock()">Unlock sheet</button><br/>
                  <button id='btnFirst' onclick="jade_modules.operations.firstScene()">First scene</button>
                  <button id='btnLast' onclick="jade_modules.operations.lastScene()">Last scene</button><br/>
                  <button id='btnPrev' onclick="jade_modules.operations.findScene(-1)">Prev scene</button>
                  <button id='btnNext' onclick="jade_modules.operations.findScene(1)">Next scene</button><br/>
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
                <div id="dateStudioEngineer"></div>
                  <label for="studio">Studio</label>
                  <input type="text" id="studio" name="studio">
                  <label for="engineer">Engineer</label>
                  <input type="text" id="engineer" name="engineer">
                  <button id='btnFillUK' onclick="jade_modules.operations.fillUK()">Fill UK</button><br/>
                </div>
                `;

  Jade.open_canvas("Script Controller", html, true);
}

