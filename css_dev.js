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
                      height: 36px;
                      border: none;
                      width: 150px;
                      border-radius: 7px;
                      background-color: #d34c01;
                      color: #fef3df;
                      font-size: 16px;
                      cursor: pointer;
                    }
                    h1, h2{
                      margin-left: 30px;
                      font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
                    }
                    label {
                      margin-left: 30px;
                      font-size: 16px;
                    }
                      
                    #firstButtons, #filterButtons, #dateStudioEngineer {
                      margin-left: 30px;
                      margin-top: 10px;
                      width: 370px;
                      padding: 10px;
                      border-width: 1px;
                      border-radius: 5px;
                      border-color: #d34c01;
                      border-style: solid;
                    }

                     #jump {
                      margin-left: 30px;
                      margin-top: 10px;
                      width: 370px;
                      padding: 10px;
                      border-width: 1px;
                      border-radius: 5px;
                      border-color: #d34c01;
                      border-style: solid;
                    }

                    #btnGo {
                      width: fit-content;
                      height: fit-content;
                      margin-left: 5px;
                    }
                    #btnFillUK {
                      width: fit-content;
                      height: fit-content;
                      margin-left: 30px;
                    }
                    #btnFillUS, #btnFillWalla {
                      width: fit-content;
                      height: fit-content;
                      margin-left: 5px;
                      
                    }
                    #scene {
                      width: 50px;
                      background-color: #ffc901;
                      color: #d34c01;
                      font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
                      border-width: 1px;
                      border-radius: 5px;
                      border-color: #d34c01;
                      border-style: solid;
                    }

                    #studio, #engineer, #studio-select, #engineer-select {
                      width: 150px;
                      background-color: #ffc901;
                      color: #d34c01;
                      font-family: 'Aptos Narrow', 'Arial Narrow'Arial, Helvetica, sans-serif;
                      border-width: 1px;
                      border-radius: 5px;
                      border-color: #d34c01;
                      border-style: solid;
                    }
                    .float-container {
                          border: 3px solid #fff;
                          padding: 20px;
                      }
                      
                      .float-child {
                          width: 50%;
                          float: left;
                          padding: 20px;
                          border: 2px solid red;
                      }  
                    
                    `;
  Jade.set_css(css);

}