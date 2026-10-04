const TARGETS_SHEET_NAME = 'Targets';

function targetsSpreadsheet_(){
  try{
    if(typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID){
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
  }catch(e){}
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!ss) throw new Error('Unable to resolve the EDEN BI spreadsheet for Targets.');
  return ss;
}

function ensureTargetsSheet_(){
  const ss=targetsSpreadsheet_();
  let sh=ss.getSheetByName(TARGETS_SHEET_NAME);
  if(!sh){
    sh=ss.insertSheet(TARGETS_SHEET_NAME);
    sh.getRange(1,1,1,10).setValues([[
      'ID','Scope','Name','Period','Year','Month','Quarter','Target EGP','Target USD','Updated At'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,10).setFontWeight('bold');
    sh.autoResizeColumns(1,10);
  }
  return sh;
}

// Signed-in users can read targets; only admins can change them.
function targetsRequireAdmin_(token){
  const user=validateAuthToken_(token);
  const role=String(user&&user.role||'').trim().toLowerCase();
  if(role!=='admin'&&role!=='super admin')throw new Error('Only admins can change targets.');
  return user;
}

function getTargetsData(authToken){
  const user=validateAuthToken_(authToken);
  const out=getTargetsRows_();
  if(isSalesScopedUser_(user))out.rows=scopeTargetsForSales_(out.rows,user);
  return out;
}

function getTargetsRows_(){
  const sh=ensureTargetsSheet_();
  const last=sh.getLastRow();
  if(last<2)return {rows:[]};

  const values=sh.getRange(2,1,last-1,10).getValues();
  const rows=values.filter(r=>r.some(v=>v!=='' && v!==null)).map(r=>({
    id:String(r[0]||''),
    scope:String(r[1]||''),
    name:String(r[2]||''),
    period:String(r[3]||'Yearly'),
    year:Number(r[4])||0,
    month:String(r[5]||''),
    quarter:String(r[6]||''),
    targetEGP:Number(r[7])||0,
    targetUSD:Number(r[8])||0,
    updatedAt:r[9] instanceof Date?r[9].toISOString():String(r[9]||'')
  }));

  return {rows};
}

function saveTargetRecord(authToken,record){
  targetsRequireAdmin_(authToken);
  const sh=ensureTargetsSheet_();
  record=record||{};

  const scope=String(record.scope||'Company').trim();
  const name=String(record.name||'').trim();
  const period=String(record.period||'Yearly').trim();
  const year=Number(record.year)||new Date().getFullYear();
  const month=period==='Monthly'?String(record.month||'').trim():'';
  const quarter=period==='Quarterly'?String(record.quarter||'').trim():'';
  const targetEGP=Math.max(0,Number(record.targetEGP)||0);
  const targetUSD=Math.max(0,Number(record.targetUSD)||0);

  if(!name)throw new Error('Target name is required.');
  if(!['Company','Manager','Sales','Branch'].includes(scope))throw new Error('Invalid target scope.');
  if(!['Monthly','Quarterly','Yearly'].includes(period))throw new Error('Invalid target period.');
  if(period==='Monthly'&&!month)throw new Error('Select a month.');
  if(period==='Quarterly'&&!quarter)throw new Error('Select a quarter.');

  let id=String(record.id||'').trim();
  if(!id)id=Utilities.getUuid();

  const last=sh.getLastRow();
  let rowIndex=0;
  if(last>=2){
    const ids=sh.getRange(2,1,last-1,1).getDisplayValues().flat();
    const idx=ids.findIndex(x=>String(x)===id);
    if(idx>=0)rowIndex=idx+2;
  }

  const payload=[[id,scope,name,period,year,month,quarter,targetEGP,targetUSD,new Date()]];
  if(rowIndex){
    sh.getRange(rowIndex,1,1,10).setValues(payload);
  }else{
    sh.getRange(sh.getLastRow()+1,1,1,10).setValues(payload);
  }

  return {success:true,id};
}

function deleteTargetRecord(authToken,id){
  targetsRequireAdmin_(authToken);
  const sh=ensureTargetsSheet_();
  id=String(id||'').trim();
  if(!id)return {success:false};

  const last=sh.getLastRow();
  if(last<2)return {success:false};

  const ids=sh.getRange(2,1,last-1,1).getDisplayValues().flat();
  const idx=ids.findIndex(x=>String(x)===id);
  if(idx<0)return {success:false};

  sh.deleteRow(idx+2);
  return {success:true};
}
