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
</div>`;
  await Jade.open_canvas("Script-Controller", html, true);
  console.log('Canvas open');
  await jade_modules.operations.showMain();
}