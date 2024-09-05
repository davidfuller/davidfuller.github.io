async function mainHTML(){
  const html = `<h1>MuVi2 Script Controller</h1>
                <h2>Development edition</h2>
                <div id="firstButtons">
                  <button id='btnLock' onclick="jade_modules.operations.lockColumns()">Lock sheet</button>
                  <button id='btnUnlock' onclick="jade_modules.operations.unlock()">Unlock sheet</button><br/>
                  <button id='btnFirst' onclick="jade_modules.operations.firstScene()">First scene</button>
                  <button id='btnLast' onclick="jade_modules.operations.lastScene()">Last scene</button><br/>
                  <button id='btnPrev' onclick="jade_modules.operations.findScene(-1)">Prev scene</button>
                  <button id='btnNext' onclick="jade_modules.operations.findScene(1)">Next scene</button><br/>
                  <button id='btnTest' onclick="jade_modules.operations.getDataFromRange('Settings','studioChoice')">Test</button><br/>
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
                  <label for="studio">Studio</label>
                  <input type="text" id="studio" name="studio"><br/>
                  <label for="engineer">Engineer</label>
                  <input type="text" id="engineer" name="engineer"><br/>
                  <button id='btnFillUK' onclick="jade_modules.operations.fill('UK')">Fill UK</button>
                  <button id='btnFillUS' onclick="jade_modules.operations.fill('US')">Fill US</button>
                  <button id='btnFillWalla' onclick="jade_modules.operations.fill('Walla')">Fill Walla</button><br/>
                </div>
                <select>
                  <option value="Audible Studio One">Audible Studio One</option>
                  <option value="Audible Studio Two">Audible Studio Two</option>
                  <option value="Audible Studio Three">Audible Studio Three</option>
                  <option value="Audible Studio Four">Audible Studio Four</option>
                </select>
                `;

  Jade.open_canvas("Script Controller", html, true);
}

