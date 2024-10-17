async function auto_exec(){
  console.log("The very beginning");
  await Jade.load_js("https://davidfuller.github.io/characterData.js", "characterData");
  console.log('After character data');
  await Jade.load_js("https://davidfuller.github.io/characterHtml.js", "html");
  console.log('After html');
  await Jade.load_js("https://davidfuller.github.io/characterCss.js", "css");
  console.log('After css');
  await jade_modules.css.mainCSS();
  await jade_modules.html.mainHTML();
}