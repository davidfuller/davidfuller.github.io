function auto_exec(){
}
async function mainHTML(){
  const html = `
<h1>MuVi2 German Scripted Walla Controller</h1>
<h2 id="sheet-version"></h2>
<div id='start-wait'>
  <h1>Please wait...</h1>
</div>
<div id="main-page">
  <div id="nav-buttons">
    <a id='show-hide' onclick="jade_modules.operations.showAdmin()">Show/hide admin</a>
  </div>
  <div id="admin">
    <label class="section-label">Admin</label><br/>
    <button id='btnCueValues' onclick="jade_modules.walla.minMaxCueValues()">Test</button>
    <button id='btnSourceSheets' onclick="jade_modules.walla.sourceSheets()">Source Sheets</button>
    <button id='btnFindCues' onclick="jade_modules.walla.findCues()">Find Cues</button>
  </div>
</div>
  `;

  await Jade.open_canvas("Script-Controller", html, true);
  console.log('Canvas open');
  await jade_modules.operations.showMain();
  console.log("I'm here data loaded. Dev 17");
}

 