/*  SETUP VAR AND EVENT LISTENER  */
var pbxVersion = null;
var importConfigData = null;
var import3CXBackupData = null; 
var git_values, environment, fileHandle;

const fs = new FileReader();
const xml = new DOMParser();

const importJsonOption = {
  types: [
    {
      description: 'Json Settings',
      accept: {
        'file/*': ['.json']
      }
    },
  ],
  excludeAcceptAllOption: true,
  multiple: false
};

const importXmlOption = {
  types: [
    {
      description: '3CX Backup',
      accept: {
        'file/*': ['.xml']
      }
    },
  ],
  excludeAcceptAllOption: true,
  multiple: false
};


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    //DOMS 
    document.getElementById("home-btn1").addEventListener("click", showHome);
    document.getElementById("home-btn2").addEventListener("click", showHome);
    document.getElementById("version-sel-btn20").addEventListener("click", showFunctions);
    document.getElementById("version-sel-btn18").addEventListener("click", showFunctions);
    document.getElementById("version-sel-btntools").addEventListener("click", showTools);
    document.getElementById("chrome").addEventListener("click", showHome);

    //EXT FUNCTIONS
    document.getElementById("f-gen-pages").addEventListener("click", genPagesBtn);
    document.getElementById("f-import-config").addEventListener("click", readConfigFile);

    //TOOL FUNCTIONS
    document.getElementById("t-3cxbackup").addEventListener("click", read3CXBackup);


    //ENV SETUP
    environment = "excel";
    if (window!=window.top){
      environment = "web";
    }
  }
});

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

async function readConfigFile(){
  readFile(importJsonOption);
}

async function read3CXBackup() {
  readFile(importXmlOption);
}

async function readFile(fileType){
  try {
    if(environment == "web"){
        const fileElem = document.createElement("input");
        fileElem.type = "file";
        fileElem.accept = fileType.types[0].accept["file/*"][0];
        fileElem.click();
        fileElem.addEventListener("change", (event)=> {
          const data = event.target.files[0];
          const excelReader = new FileReader();
          excelReader.onload = function (evnt) {
            console.log(evnt.target.result)
            onFileRead(evnt.target.result, fileType.types[0].accept["file/*"][0]);
          }
        excelReader.readAsText(data); 
      })
    }else{
        [fileHandle] = await window.showOpenFilePicker(fileType);
        const file = await fileHandle.getFile();
        onFileRead(await file.text(), fileType.types[0].accept["file/*"][0]);
      }
    }
   catch (error) {
    console.log(error);
  }
}

async function onFileRead(data, type) {
  if(type == ".json"){
    importConfigData = JSON.parse(data);
    updateSettings(importConfigData);
  }else{
    import3CXBackupData = xml.parseFromString(data, "text/xml");
  }
}

async function updateSettings(data) {
  createOptionsFromJSON("department", data.department ,"Name");
  createOptionsFromJSON("trunk", data.trunk ,"TrunkName");
  document.getElementById("fqdn").value = data.fqdn;
}


function createOptions(parentID, data){
  const parent = document.getElementById(parentID);
  for (let i = 0; i < data.length; i++) {
    let opt = document.createElement("option");
    opt.innerText = data[i].text;
    opt.value = data[i].value;
    parent.appendChild(opt);
  }
}

function createOptionsFromJSON(parentID, data, key){
  const parent = document.getElementById(parentID);
  for (let i = 0; i < data.length; i++) {
    console.log(data);
    let opt = document.createElement("option");
    opt.innerText = data[i][key];
    opt.value = i;
    parent.appendChild(opt);
  }
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
  genPages(git_values[pbxVersion]);
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
        if (exitsPages.incluwdes(page.name) == false) {
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

/* GITHUB VARIABLES */

async function gitVal() {
  fetch('https://raw.githubusercontent.com/ZSirag/3CX/main/settings.json')
    .then(res => res.json())
    .then(json => {
    git_values = json;
    let selElm = document.getElementById("phone-brand");
    for (let i = 0; i < git_values.phones.length; i++) {
      const elm = document.createElement("option");
      elm.innerHTML = git_values.phones[i].name;
      elm.value = i;
      selElm.appendChild(elm);
    }
  })
}

gitVal();