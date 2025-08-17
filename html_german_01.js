function auto_exec(){
}
async function mainHTML(){
  const html = `
<h2 id="sheet-version"></h2>
<div id='start-wait'>
  <h1>Please wait...</h1>
</div>
<div id="main-page">
  <h1>German Processing</h1>
  <div id="processing">
    <label class="section-label" onclick="jade_modules.operations.showProcessing()">Processing</label><a id='show-hide-processing' onclick="jade_modules.operations.showProcessing()">Show/hide processing</a>
    <div id="processing-group">
      <button id='btnLoadReplaceProcess' onclick="jade_modules.preprocess.loadReplaceProcess()">Load/Replace/Process</button> <br/>
      <label id="load-message"></label>
      <button id='btnLoadLoadOriginal' onclick="jade_modules.preprocess.loadOriginal()">Load German Block</button> <br/>
      <button id='btnFindInLockedOriginal' onclick="jade_modules.preprocess.findInLockedOriginal()">Find In Locked Original</button><br/>
      <label id="search-label">Search</label><br/>
      <button id='btnCopySearchToReplace' onclick="jade_modules.replacements.copySearchToReplace()()">Copy (Manual)</button>
      <button id='btnReplaceAllDoubleQuotesWithSingles' onclick="jade_modules.replacements.replaceAllDoubleQuotesWithSingles()">Replace Double Quote With Single</button>
      <button id='btnIsolateQuotedBit' onclick="jade_modules.replacements.isolateQuotedBit()">Isolate Quote Bit</button>
      <button id='btnCreateMissingSearchAndReplace' onclick="jade_modules.replacements.createMissingSearchAndReplace()">Create Missing Text</button>
      <button id='btnInsertEol' onclick="jade_modules.replacements.insertEol()">Insert End of Line</button>
      <button id='btnInsertExtra' onclick="jade_modules.replacements.insertExtra()">Insert [EXTRA:]</button>
      <textarea id="original-text" cols="40" rows="8"></textarea><br/>
      <label id="replace-label">Replace</label><br/>
      <button id='btnAddExtraMissingText' onclick="jade_modules.replacements.addExtraMissingText()">Add extra MISSING TEXT</button>
      <button id='btnIsertMoreEol' onclick="jade_modules.replacements.insertMoreEol()()">Add extra eol</button><br/>
      <button id="btnDoubleReplace" onclick="jade_modules.replacements.replaceDoubleQuotesInSelection()">Replace Double Quotes In Selection</button>
      <textarea id="replace-text" cols="40" rows="8"></textarea><br/>
      <button id='btnAddToReplacementsProcess' onclick="jade_modules.replacements.addToReplacementsProcess()">Add to Replacements then Process</button>
      <button id='btnAddToReplacements' onclick="jade_modules.replacements.addToReplacements()">Add to Replacements</button>
      <br/>
    </div>
  </div>
  <div id="admin">
    <label class="section-label" onclick="jade_modules.operations.showAdmin()">Admin</label><a id='show-hide-admin' onclick="jade_modules.operations.showAdmin()">Show/hide admin</a><br/>
    <div id="admin-group">
      <button id='btnApplyTransactionFormulaToAll' onclick="jade_modules.translation.fillWithFormula()">Formula All</button><br/>
      <button id='btnMachineTranslationValues' onclick="jade_modules.translation.machineTranslationValues()">Values</button><br/>
      <button id='btnMachineTranslationIssueCells' onclick="jade_modules.translation.issueCells(false)">Calculation Issues</button><br/>
      <label id="issues-message"></label><br/>
      <button id='btnMachineTranslationIssueCellsWithFormula' onclick="jade_modules.translation.issueCells(true)">Issues Add Formulas</button><br/>
      <button id='btnCopyValuesToCache' onclick="jade_modules.translation.copyValuesToCache()">Cache Machine Translations</button><br/>
      <button id='btnGetTranslationFormulas' onclick="jade_modules.translation.getMachineTranslationFormula()">Get Machine Translation Formula</button><br/>
      <button id='btnFixMachineTranslationDisplay' onclick="jade_modules.translation.fixMachineTranslationDisplay()">Fix Machine Translation</button><label id="fix-machine-message">Please wait...</label><br/>
      <button id='btnApplyMachineTranslationFormula' onclick="jade_modules.translation.applyMachineTranslationFormula(13)">Apply formula</button><br/>
      <button id='btnCompareWithCache' onclick="jade_modules.translation.compareTranslationwithCache(false)">Compare with Cache</button><br/>
      <button id='btnCompareWithCacheFormulae' onclick="jade_modules.translation.compareTranslationwithCache(true)">Compare with Cache and Add Formulae</button><br/>
      <button id='btnLoadScript' onclick="jade_modules.preprocess.doTheCopy()">Load German Original</button><br/>
      <button id='btnProcessGerman' onclick="jade_modules.operations.processGerman()">Process German</button>
      <button id='btnLoadUKScript' onclick="jade_modules.preprocess.getUKScript()">Load UK Script</button><br/>
      <button id='btnFindInOriginal' onclick="jade_modules.preprocess.findThisBlock(true, true)">Find this text in Original</button>
      <button id='btnReturnProcessed' onclick="jade_modules.preprocess.returnToProcessedCell()">Return to processed cell</button><br/>
      <button id='btnFindInLockedOriginal' onclick="jade_modules.preprocess.findInLockedOriginal()">Find In Locked Original</button>
      <button id='btnDoTheReplacements' onclick="jade_modules.replacements.doTheReplacements()">Do The Replacements</button>
      <button id='btnReplaceProcess' onclick="jade_modules.replacements.replacementsAndProcess()">Replace/Process</button>
      <br/>
      <input type="text" id="process-address" name="process-address">
      <input type="text" id="source-row" name="source-row">
    </div>
  </div>    
  <div id="jump">
    <label class="section-label" onclick="jade_modules.operations.showJump()">Jump to...</label><a id='show-hide-jump' onclick="jade_modules.operations.showJump()">Show/hide jumping</a>
    <div id="jump-buttons">
      <button id='btnFirst' onclick="jade_modules.operations.firstScene()">First chapter</button>
      <button id='btnPrev' onclick="jade_modules.operations.findScene(-1)">Prev chapter</button>
      <button id='btnNext' onclick="jade_modules.operations.findScene(1)">Next chapter</button>
      <button id='btnLast' onclick="jade_modules.operations.lastScene()">Last chapter</button><br/>
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
</div>`;
  await Jade.open_canvas("Script-Controller", html, true);
  console.log('Canvas open');
  await jade_modules.operations.showMain();
}