/*  SETUP VAR AND EVENT LISTENER  */
var importConfigData = null;
var import3CXBackupData = null;
var git_values, environment, fileHandle;

const fs = new FileReader();
const xml = new DOMParser();
const regexRows = /[^:0-9]/g;
const backupPhoneBooks = {data: []};

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
    document.getElementById("f-gen-exts").addEventListener("click", genExt);
    document.getElementById("f-export-exts").addEventListener("click", exportExts);
    document.getElementById("f-gen-pbook").addEventListener("click", genPbook);
    document.getElementById("f-export-pbook").addEventListener("click", exportPhonebook);

    //TOOL FUNCTIONS
    document.getElementById("t-3cxbackup").addEventListener("click", read3CXBackup);
    document.getElementById("load-bk-pbooks").addEventListener("click", loadPhonebook);
    document.getElementById("export-bk-download").addEventListener("click", exportPhonebookBackup);

    //ENV SETUP
    environment = "excel";
    if (window != window.top) {
      environment = "web";
    }
  }
});

/* TASKPANE DOMS */
function showHome() {
  document.getElementById("functions-container").style.display = "none";
  document.getElementById("tools-container").style.display = "none";
  document.getElementById("version-sel-container").style.display = "block"
}

function showFunctions(e) {
  document.getElementById("version-sel-container").style.display = "none";
  document.getElementById("functions-container").style.display = "grid";
  document.getElementById("tools-container").style.display = "none";
}

function showTools(e) {
  document.getElementById("version-sel-container").style.display = "none";
  document.getElementById("functions-container").style.display = "none";
  document.getElementById("tools-container").style.display = "grid";
}

async function readConfigFile() {
  readFile(importJsonOption);
}

async function read3CXBackup() {
  readFile(importXmlOption);
}

async function readFile(fileType) {
  try {
    if (environment == "web") {
      const fileElem = document.createElement("input");
      fileElem.type = "file";
      fileElem.accept = fileType.types[0].accept["file/*"][0];
      fileElem.click();
      fileElem.addEventListener("change", (event) => {
        const data = event.target.files[0];
        const excelReader = new FileReader();
        excelReader.onload = function (evnt) {
          console.log(evnt.target.result)
          onFileRead(evnt.target.result, fileType.types[0].accept["file/*"][0]);
        }
        excelReader.readAsText(data);
      })
    } else {
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
  if (type == ".json") {
    importConfigData = JSON.parse(data);
    updateSettings(importConfigData);
  } else {
    import3CXBackupData = xml.parseFromString(data, "text/xml");
    updatePhonebook(import3CXBackupData);
  }
}

async function updatePhonebook(data) {
  const tenant = data.querySelector("Tenants");
  const publicPhonebook = tenant.querySelector("PhoneBookEntries");
  if(publicPhonebook.hasChildNodes()){
    backupPhoneBooks.data.push({name: "Company phonebook", data: populatePB([publicPhonebook])});
  }
  const exts = data.documentElement.querySelectorAll("Extension");
  for(let i = 0; i < exts.length; i++){
    const tempExt = exts[i];
    const tempPhonebook = tempExt.getElementsByTagName("PhoneBookEntries");
    if(tempPhonebook[0].hasChildNodes()){
      const fileName = `${(tempExt.getElementsByTagName("Number"))[0].innerHTML} - ${(tempExt.getElementsByTagName("FirstName"))[0].innerHTML}`;  
      backupPhoneBooks.data.push({name: fileName, data: populatePB(tempPhonebook)});
    }
  }
  createOptionsFromJSON("t-phonebook-list", backupPhoneBooks.data, "name");
  console.log(backupPhoneBooks);  
}

function populatePB(rawPBook) {
  console.log(rawPBook);
  const pbEntry = rawPBook[0].getElementsByTagName("PhoneBookEntry");
  const dataOut = [];
  for(let j = 0; j < pbEntry.length; j++){
    let data = new Array(18).fill("");
    pbEntry[j].childNodes.forEach((item, index) => {
      switch (item.tagName) {
        case "FirstName": data[0] = item.innerHTML;
        break;
        case "LastName": data[1] = item.innerHTML;
        break;
        case "CompanyName": data[2] = item.innerHTML;
        break;
        case "PhoneNumber": data[3] = item.innerHTML;
        break;
        default: {
            if(item.tagName != undefined){
              data[index+4] = item.innerHTML
            }
          };
        break;
      }
    });

    data = filterArray(data, git_values["v20"].template.backupImportFilter);
    dataOut.push(data);
  }
  return dataOut;
}

async function updateSettings(data) {
  createOptionsFromJSON("department", data.department, "Name");
  createOptionsFromJSON("trunk", data.trunk, "TrunkName");
  document.getElementById("fqdn").value = data.fqdn;
}

function createOptionsFromJSON(parentID, data, key) {
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


export async function loadPhonebook() {
  try{
    await Excel.run(async (context) => {
      const selElm = document.getElementById("t-phonebook-list").value;
      if(selElm != "null"){
        const dataOut = [];
        for (let i = 0; i < backupPhoneBooks.data[selElm].data.length; i++) {
          const template = git_values["v20"].template.pbook.slice();
          dataOut.push(combineArray(backupPhoneBooks.data[selElm].data[i], template, git_values["v20"].template.backupImportOffset));
        }
        const pbookOutPage = context.workbook.worksheets.getItem("Out Phonebook");
        const outPbookPageRange = pbookOutPage.getRange(`A2:N${dataOut.length+1}`);
        outPbookPageRange.values = dataOut;
      }
    });
  } catch(error){
    console.log(error);
  }
}

export async function addDepartment() {
  try {
    await Excel.run(async (context) => {
      console.log("TETS");
      const extPage = context.workbook.worksheets.getItem("Extentions");
      const departmentRange = await advSelect(context, { columns: { start: "g", end: "g" } });
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

export async function addPhone() {
  try {
    await Excel.run(async (context) => {
      const extPage = context.workbook.worksheets.getItem("Extentions");
      const brand = document.getElementById("phone-brand").value;

      const phoneRange = await advSelect(context, { columns: { start: "h", end: "h" } });
      const sbcRange = await advSelect(context, { columns: { start: "k", end: "k" } });
      const langRange = await advSelect(context, { columns: { start: "i", end: "i" } });

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
  } catch (error) {
    console.log(error)
  }
}

export async function genPbook() {
  try {
    await Excel.run(async (context) => {
      const pbookPage = context.workbook.worksheets.getItem("Phonebook");
      const pbookRange = await advSelect(context, { columns: { start: "a", end: "h" }});
      const pbookRowData = pbookPage.getRange(pbookRange);
      pbookRowData.load("values");
      await context.sync();
      const dataOut = new Array(pbookRowData.values.length);
      for (let i = 0; i < dataOut.length; i++) {
        let template = git_values["v20"].template.pbook.slice();
        template = combineArray(pbookRowData.values[i], template, git_values["v20"].template.pbookOffset);
        dataOut[i] = template;
      }
      const pbookOutPage = context.workbook.worksheets.getItem("Out Phonebook");
      const outPbookPageRange = pbookOutPage.getRange(`A2:N${dataOut.length+1}`);
      outPbookPageRange.values = dataOut;
    });
  } catch (error) {
    console.log(error);
  }

}

export async function exportPhonebook(){
  try {
    await Excel.run(async (context) => {
      const pbookOutPage = context.workbook.worksheets.getItem("Out Phonebook");
      const pbookOutRange = await advSelect(context, { columns: { start: "a", end: "n" } });
      const pbookOutRowData = pbookOutPage.getRange(pbookOutRange);
      pbookOutRowData.load("values");
      await context.sync();

      const newArray = [git_values["v20"].pages[3].data].concat(pbookOutRowData.values);
      let dataOut = new Blob([newArray.join("\n")], {type: "text/plain"});
      saveAs(dataOut, "Phonebook.csv")
    });
  } catch(error){
    console.log(error);
  }
}


export async function genExt() {
  try {
    await Excel.run(async (context) => {
      const extPage = context.workbook.worksheets.getItem("Extentions");
      const extRange = await advSelect(context, { columns: { start: "a", end: "k" } });
      const extRowData = extPage.getRange(extRange);
      extRowData.load("values");
      await context.sync();
      const dataOut = new Array(extRowData.values.length);
      for (let i = 0; i < extRowData.values.length; i++) {
        let template = git_values["v20"].template.ext.slice();
        template = combineArray(extRowData.values[i], template, git_values["v20"].template.extOffset);

        template[11] = (extRowData.values[i][0]+extRowData.values[i][1]).replace(/\s/g, '');;
        template[20] = importConfigData.globalACPRMSET;
        template[22] = generateRandomPin();

        if(extRowData.values[i][9] != ""){
          const phoneConfig = retrivePhoneInfo(extRowData.values[i][7]);
          let router = "";
          if(phoneConfig != null){
            template[13] = phoneConfig.xml;
            template[17] = phoneConfig.ring[0]
            template[18] = phoneConfig.ring[1]
          }
          if(extRowData.values[i][10] == "Router"){
            router = extRowData.values[i][9];
          }
          if(extRowData.values[i][10] != "Router" && extRowData.values[i][10] != "Local"){
            if(retriveSBC(extRowData.values[i][10])){
              router = retriveSBC(extRowData.values[i][10]);
            }
          }
          template[15] = router;
        }
        dataOut[i] = template;
      }
      const extOutPage = context.workbook.worksheets.getItem("Out Extentions");
      const outExtPageRange = extOutPage.getRange(`A2:AX${dataOut.length+1}`);
      console.log({extRowData: extRowData.values, dataOut});
      outExtPageRange.values = dataOut;
    });
  } catch (error) {
    console.log(error);
  }

}

export async function exportExts(){
  try {
    await Excel.run(async (context) => {
      const extOutPage = context.workbook.worksheets.getItem("Out Extentions");
      const extOutRange = await advSelect(context, { columns: { start: "a", end: "ax" } });
      const extOutRowData = extOutPage.getRange(extOutRange);
      extOutRowData.load("values");
      await context.sync();

      const newArray = [git_values["v20"].pages[2].data].concat(extOutRowData.values);
      let dataOut = new Blob([newArray.join("\n")], {type: "text/plain"});
      saveAs(dataOut, "Extentions.csv")
    })
  } catch(error){
    console.log(error);
  }
}

function exportPhonebookBackup(){
  const selElm = document.getElementById("t-phonebook-list").value;
  var zip = new JSZip();
  console.log("test")
  if(selElm != "null"){
    for (let i = 0; i < backupPhoneBooks.data.length; i++) {
      const dataOut = [];
      dataOut.push(git_values["v20"].pages[3].data.join(","));
      for (let j = 0; j < backupPhoneBooks.data[i].data.length; j++) {
        const template = git_values["v20"].template.pbook.slice();
        let tmpArrary = combineArray(backupPhoneBooks.data[i].data[j], template, git_values["v20"].template.backupImportOffset);
        dataOut.push(tmpArrary.join(","));
        console.log(dataOut);
      }
      zip.file(`${backupPhoneBooks.data[i].name}.csv`, dataOut.join("\n"));
    }
    zip.generateAsync({type:"blob"}).then(function (blob) {
      saveAs(blob, "PhonebookBackup.zip");
    });
  }
}

function retrivePhoneInfo(model){
  for (let i = 0; i < git_values.phones.length; i++) {
    console.log(git_values.phones[i]);
    if(git_values.phones[i].models.includes(model)){
      return {xml: git_values.phones[i].xml, ring: git_values.phones[i].ring}
    }
  }
  return null;
} 

function retriveSBC(sbc) {
  for (let i = 0; i < importConfigData.sbc.length; i++) {
    if(importConfigData.sbc[i].DisplayName == sbc){
      if(importConfigData.sbc[i].PhoneMAC == ""){
        return importConfigData.sbc[i].Name
      }
      return importConfigData.sbc[i].PhoneMAC;
    }
  }
  return null;
}

function generateRandomPin() {
  const digits = '0123456789';
  let pin = '';

  for (let i = 0; i < 6; i++) {
    const randomIndex = Math.floor(Math.random() * digits.length);
    pin += digits[randomIndex];
  }

  return pin;
}

// ADV SELECT
async function advSelect(context, settings) {
  //LOAD SELECTED RANGE
  let tmpRange = context.workbook.getSelectedRange();
  tmpRange.load("address");
  await context.sync();
  let range = ((tmpRange.address).split("!"))[1];

  // AVOID 1ST ROW SELECTION
  let rows = range.split(":");
  for (let i = 0; i < rows.length; i++) {
    rows[i] = rows[i].replace(regexRows, "");
    if (rows[i] == 1) {
      rows[i] = 2;
    }
  }

  //COMBINE SELECTED ROWS WT PASSED COLUMN
  if (rows.length == 1 && settings.columns.start == settings.columns.end) {
    return `${settings.columns.start}${rows[0]}`;
  }
  if (rows.length == 1 && settings.columns.start != settings.columns.end) {
    return `${settings.columns.start}${rows[0]}:${settings.columns.end}${rows[0]}`;;
  }
  return `${settings.columns.start}${rows[0]}:${settings.columns.end}${rows[1]}`;;
}

function genPagesBtn() {
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
        const columns = numberToColumn(page.data.length - 1);
        const selPage = context.workbook.worksheets.getItem(page.name);
        const range = selPage.getRange(`A1:${columns}1`);
        range.load("values");
        range.values = [page.data];
        const range2 = selPage.getRange(`A:${columns}`);
        range2.numberFormat = "@";
      });
    });

  } catch (error) {
    console.log(error);
  }
}


/* LIBS */
function numberToColumn(number) {
  let a = "";
  let result = Math.trunc(number / 26);
  if (result > 0) {
    a += String.fromCharCode(65 + result - 1)
    a += String.fromCharCode(65 + number % 26);
  } else {
    a += String.fromCharCode(65 + number % 26);
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

function combineArray(sourceArray, destArray, offsets, length) {

  console.log({sourceArray, destArray});
  if (length) {
    for (let i = 0; i < length; i++) {
      destArray[offsets[i]] = sourceArray[i];
    }
  } else {
    for (let i = 0; i < sourceArray.length; i++) {
      destArray[offsets[i]] = sourceArray[i];
    }
  }
  return destArray;
}

function filterArray(sourceArray, filter) {
  let destArray = new Array(filter.length).fill("");
  for (let i = 0; i < filter.length; i++) {
    destArray[i] = sourceArray[filter[i]];
  }
  return destArray;
}

gitVal();