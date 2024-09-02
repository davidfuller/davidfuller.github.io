async function mainHTML(){
  const html = `<h1>MuVi2 Script Controller</h1>
                <button id='btnLock' onclick="jade_modules.operations.lockColumns()">Lock sheet</button>
                <button id='btnUnlock' onclick="jade_modules.operations.unlock()">Unlock sheet</button>
                <button id='btnNext' onclick="jade_modules.operations.findScene(1)">Next scene</button>
                <button id='btnPrev' onclick="jade_modules.operations.findScene(-1)">Previous scene</button><br/>
                <label for="scene">Jump to scene</label>
                <input type="text" id="scene" name="scene">
                <button id='btnJump' onclick="jade_modules.operations.getTargetSceneNumber()">Go</button><br/>`;

  Jade.open_canvas("Script Controller", html, true);
}

