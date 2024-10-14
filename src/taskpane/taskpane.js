/*  SETUP VAR AND EVENT LISTENER  */
var pbxVersion = null;
var git_values;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("home-btn1").addEventListener("click", showHome);
    document.getElementById("home-btn2").addEventListener("click", showHome);
    document.getElementById("version-sel-btn20").addEventListener("click", showFunctions);
    document.getElementById("version-sel-btn18").addEventListener("click", showFunctions);
    document.getElementById("version-sel-btntools").addEventListener("click", showTools);
    document.getElementById("chrome").addEventListener("click", showHome);
    document.getElementById("f-gen-pages").addEventListener("click", genPagesBtn);
  }
});

/* GITHUB VARIABLES */



/* TASKPANE DOMS */ 
function showHome(){
  document.getElementById("functions-container").style.display ="none";
  document.getElementById("tools-container").style.display ="none";
  document.getElementById("version-sel-container").style.display ="block"
}

function showFunctions(e){
  pbxVersion = "v18";
  if(e.srcElement.id=="version-sel-btn20"){
    pbxVersion = "v20";
  }
  document.getElementById("version-sel-container").style.display ="none";
  document.getElementById("functions-container").style.display ="grid";
  document.getElementById("tools-container").style.display ="none";
  document.getElementById("version-info").innerText = pbxVersion;
}

function showTools(e){
  document.getElementById("version-sel-container").style.display ="none";
  document.getElementById("functions-container").style.display ="none";
  document.getElementById("tools-container").style.display ="grid";
}



/* EXCEL PART */ 
export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

function genPagesBtn(){
  switch (pbxVersion) {
    case "v20": {
      genPages(git_values[pbxVersion]);
    }
    break;
  
    default:
    break;
  }
}

export async function genPages(vesion) {
  try {
    await Excel.run(async (context) => {
      //CREATE PAGE 
      const names = context.workbook.worksheets.load("items/name");
      let exitsPages = new Array();
      await context.sync();
      names.items.forEach((pageName) => {
        exitsPages.push(pageName.name);
      })
      vesion.pages.forEach((page, i) => {
        if (exitsPages.includes(page.name) == false) {
          context.workbook.worksheets.add(page.name);
        }
      })
    });

    //LOAD DATA
    await Excel.run(async (context) => {
        vesion.pages.forEach(page => {
          const columns = numberToColumn(page.data.length-1);
          const selPage = context.workbook.worksheets.getItem(page.name);
          const range = selPage.getRange(`A1:${columns}1`);
          range.load("values");
          range.values = [page.data];
          const range2 = selPage.getRange(`A:${columns}`);
          range2.numberFormat = "@";
      });
    });

  }catch (error) {
    console.log(error);
  }
}



/* LIBS */ 
function numberToColumn(number) {
  let a = "";
  let result = Math.trunc(number/26);
  if(result > 0){
    a += String.fromCharCode(65+result-1)
    a += String.fromCharCode(65+number%26);
  }else{
    a += String.fromCharCode(65+number%26); 
  }
  return a;
}

async function gitVal() {
  fetch('https://raw.githubusercontent.com/ZSirag/3CX/main/settings.json')
    .then(res => res.json())
    .then(json => {
    git_values = json;
    /*let selElm = document.getElementById("excel-phone");
    for (let i = 0; i < git_values.phones.length; i++) {
      const elm = document.createElement("option");
      elm.innerHTML = git_values.phones[i].name;
      elm.value = i;
      selElm.appendChild(elm);
    }*/
  })
}

gitVal();