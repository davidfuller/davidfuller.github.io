function mainHTML(){
  const html = [];

  html.push(`<h1>MuVi2 Script Controller</h1>`);
  html.push(`<button id='btnLock' onclick="jade_modules.operations.lockColumns()">Lock sheet</button>`);
  html.push(``);
  html.push(``);
  html.push(``);
  html.push(``);

  html.join(" ");
  Jade.open_canvas("Script Controller", html, true);

}

