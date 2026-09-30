/*************************************************
 * DashboardConfig.gs
 * EDEN BI V2
 *************************************************/

const DASHBOARD_CONFIG_SHEET = "Dashboard_Config";

function getDashboardConfigSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let sheet = ss.getSheetByName(DASHBOARD_CONFIG_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(DASHBOARD_CONFIG_SHEET);

    sheet.getRange(1,1,1,3).setValues([
      ["Section","Property","Value"]
    ]);

    sheet.getRange(2,1,14,3).setValues([
      ["Dashboard","Title","Transactions Dashboard"],
      ["Dashboard","Description","Overview for projects, sales, branches, meters and deals."],

      ["Sidebar","Brand","EDEN DEVELOPMENT"],
      ["Sidebar","Subtitle","Live Dashboard"],

      ["Layout","Cards Per Row","5"],

      ["Visibility","Charts","TRUE"],
      ["Visibility","Cards","TRUE"],
      ["Visibility","Filters","TRUE"],
      ["Visibility","Table","TRUE"],

      ["Theme","Primary Color","#ffd21f"],
      ["Theme","Background","#08111f"],
      ["Theme","Panel","#111c2d"],

      ["Version","Current","2.0"],
      ["System","Last Update",new Date()]
    ]);

    sheet.setFrozenRows(1);
  }

  return sheet;
}

/*************************************************
 * Read Config
 *************************************************/

function getDashboardConfig(authToken){

  validateAuthToken_(authToken);

  const sheet=getDashboardConfigSheet_();

  const values=sheet.getDataRange().getValues();

  const cfg={};

  values.slice(1).forEach(r=>{

    const section=String(r[0]||"").trim();

    const property=String(r[1]||"").trim();

    const value=r[2];

    if(!cfg[section]) cfg[section]={};

    cfg[section][property]=value;

  });

  return cfg;

}

/*************************************************
 * Save Config
 *************************************************/

function saveDashboardConfig(authToken,data){

  assertUserAdmin_(authToken);

  const sheet=getDashboardConfigSheet_();

  const values=sheet.getDataRange().getValues();

  values.slice(1).forEach((row,i)=>{

      const section=row[0];

      const property=row[1];

      if(
          data[section] &&
          data[section][property]!==undefined
      ){

          sheet.getRange(i+2,3)
               .setValue(data[section][property]);

      }

  });

  sheet.getRange("C15").setValue(new Date());

  return {
      success:true
  };

}

/*************************************************
 * Reset Config
 *************************************************/

function resetDashboardConfig(authToken){

  assertUserAdmin_(authToken);

  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);

  const old=ss.getSheetByName(DASHBOARD_CONFIG_SHEET);

  if(old){

      ss.deleteSheet(old);

  }

  getDashboardConfigSheet_();

  return {
      success:true
  };

}