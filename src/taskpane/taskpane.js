/*  SETUP VAR AND EVENT LISTENER  */
var pbxVersion = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("home-btn1").addEventListener("click", showHome);
    document.getElementById("home-btn2").addEventListener("click", showHome);
    document.getElementById("version-sel-btn20").addEventListener("click", showFunctions);
    document.getElementById("version-sel-btn18").addEventListener("click", showFunctions);
    document.getElementById("version-sel-btntools").addEventListener("click", showTools);
    document.getElementById("chrome").addEventListener("click", showHome);
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


/* LIBS */ 