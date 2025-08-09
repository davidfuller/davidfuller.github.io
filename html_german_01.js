function auto_exec(){
}
async function mainHTML(){
  const html = `
<h2 id="sheet-version"></h2>
<div id='start-wait'>
  <h1>Please wait...</h1>
</div>
<div id="main-page">
  <h1>Hello Fran and Tom</h1>
  <div id="admin">
    <label class="section-label">Admin</label><br/>
    <button id='btnLoadScript' onclick="jade_modules.preprocess.doTheCopy()">Load German Original</button>
    <button id='btnProcessGerman' onclick="jade_modules.operations.processGerman()">Process German</button>
    <button id='btnLoadUKScript' onclick="jade_modules.preprocess.getUKScript()">Load UK Script</button>
    <button id='btnFindInOriginal' onclick="jade_modules.preprocess.findThisBlock()">Find this text in Original</button>
    <button id='btnReturnProcessed' onclick="jade_modules.preprocess.returnToProcessedCell()">Return to processed cell</button><br/>
    <input type="text" id="process-address" name="process-address">
    <input type="text" id="source-row" name="source-row">
    <label id="search-lable">Search</label><br/>
    <textarea id="original-text" cols="40" rows="8"></textarea><br/>
    <label id="replace-label">Replace</label><br/>
    <textarea id="replace-text" cols="40" rows="8"></textarea>
  </div>  
</div>`;
  await Jade.open_canvas("Script-Controller", html, true);
  console.log('Canvas open');
  await jade_modules.operations.showMain();
}