/*  SETUP VAR AND EVENT LISTENER  */
var importConfigData = null;
var import3CXBackupData = null; 
var git_values, environment, fileHandle;

const fs = new FileReader();
const xml = new DOMParser();
const regexRows = /[^:0-9]/g;

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
    console.log(info);

    //DOMS 
    document.getElementById("home-btn1").addEventListener("click", showHome);
    document.getElementById("home-btn2").addEventListener("click", showHome);
    document.getElementById("version-sel-btn20").addEventListener("click", showFunctions);
    document.getElementById("version-sel-btntools").addEventListener("click", showTools);
    document.getElementById("chrome").addEventListener("click", showHome);

    //EXT FUNCTIONS
    document.getElementById("f-gen-pages").addEventListener("click", genPagesBtn);
    document.getElementById("f-import-config").addEventListener("click", readConfigFile);
    document.getElementById("department-btn").addEventListener("click", addDepartment);
    document.getElementById("phone-brand-btn").addEventListener("click", addPhone);

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
  document.getElementById("version-sel-container").style.display ="none";
  document.getElementById("functions-container").style.display ="grid";
  document.getElementById("tools-container").style.display ="none";
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
  parent.value = 0;
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

export async function addDepartment() {
  try {
    await Excel.run(async(context)=> {
      console.log("TETS");
      const extPage = context.workbook.worksheets.getItem("Extentions");
      const departmentRange = await advSelect(context, {columns: {start: "g", end: "g"}});
      const departmentData = extPage.getRange(departmentRange);
      departmentData.values = "Sel. Department";
      const tempDepartments = [];
      for (let i = 0; i < importConfigData.department.length; i++) {
        tempDepartments.push(importConfigData.department[i].Name);
      } 
      console.log(tempDepartments);
      departmentData.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: `${tempDepartments.join(",")}`
        }
      }
    });
  } catch (error) {
    console.log(error);
  }
}


export async function addPhone(){
  try {
    await Excel.run(async (context) => {
      const extPage = context.workbook.worksheets.getItem("Extentions");
      const brand = document.getElementById("phone-brand").value;      

      const phoneRange = await advSelect(context, {columns: {start: "h", end: "h"}});
      const sbcRange = await advSelect(context, {columns: {start: "k", end: "k"}});
      const langRange = await advSelect(context, {columns: {start: "i", end: "i"}});

      const phoneRangeData = extPage.getRange(phoneRange);
      const sbcRangeData = extPage.getRange(sbcRange);
      const langRangeData = extPage.getRange(langRange);
      
      
      phoneRangeData.values = "Sel. Device";
      sbcRangeData.values = "Sel. SBC";
      langRangeData.values = "Sel. Lang";


      phoneRangeData.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: `${git_values.phones[brand].models.join(",")}`
        }
      }
      langRangeData.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: `${git_values.phones[brand].langs.join(",")}`
        }
      }
      let tmpSbc = [];
      for (let i = 0; i < importConfigData.sbc.length; i++) {
        const sbc = importConfigData.sbc[i].DisplayName;
        tmpSbc.push(sbc)
      }
      sbcRangeData.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: `${tmpSbc.join(",")}`
        }
      }
    })
  }catch(error){
    console.log(error)
  }
}



// ADV SELECT
async function advSelect(context, settings){
  //LOAD SELECTED RANGE
  let tmpRange = context.workbook.getSelectedRange();
  tmpRange.load("address");
  await context.sync();
  let range = ((tmpRange.address).split("!"))[1];

  // AVOID 1ST ROW SELECTION
  let rows = range.split(":");
  for(let i = 0; i < rows.length; i++){
    rows[i] = rows[i].replace(regexRows, "");
    if(rows[i] == 1){
      rows[i] = 2;
    }
  }
 
  //COMBINE SELECTED ROWS WT PASSED COLUMN
  if(rows.length == 1 && settings.columns.start == settings.columns.end){
    return `${settings.columns.start}${rows[0]}`;
  }
  if(rows.length == 1 && settings.columns.start != settings.columns.end){
    return `${settings.columns.start}${rows[0]}:${settings.columns.end}${rows[0]}`;;
  }
  return `${settings.columns.start}${rows[0]}:${settings.columns.end}${rows[1]}`;;
}

function genPagesBtn(){
  genPages(git_values["v20"]);
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