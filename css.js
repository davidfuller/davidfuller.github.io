async function mainCSS(){
  const css =`body {
              background-color: #ffc901;
              color: #d34c01;
              font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
            }
            button {
              margin-left: 30px;
              margin-bottom: 5px;
              margin-top: 5px;
              padding-left: 15px;
              padding-right: 15px;
              height: 40px;
              border: none;
              border-radius: 7px;
              background-color: #d34c01;
              color: #fef3df;
              font-size: 18px;
              cursor: pointer;
            }
            h1 {
              margin-left: 30px;
              font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
            }`;
  Jade.set_css(css);

}