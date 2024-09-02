async function mainHTML(){
  const html = `<h1>MuVi2 Script Controller</h1>
                <button id='btnLock' onclick="jade_modules.operations.lockColumns()">Lock sheet</button>
                <button id='btnLock' onclick="jade_modules.operations.unlock()">Unlock sheet</button>`;
  Jade.open_canvas("Script Controller", html, true);
}

