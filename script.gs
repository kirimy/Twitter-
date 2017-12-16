// OAuth1認証用インスタンス
var twitter = TwitterWebService.getInstance(
  '***CONSUMER_KEY***',
  '***CONSUMER_SECRET***'
);

//OAuth1ライブラリを導入したうえで、getServiceを上書き
twitter.getService = function() {
  return OAuth1.createService('Twitter2')
    .setAccessTokenUrl('https://api.twitter.com/oauth/access_token')
    .setRequestTokenUrl('https://api.twitter.com/oauth/request_token')
    .setAuthorizationUrl('https://api.twitter.com/oauth/authorize')
    .setConsumerKey(twitter.consumer_key)
    .setConsumerSecret(twitter.consumer_secret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
}

// 認証を行う（必須）
function authorize() {
  twitter.authorize();
}

// 認証をリセット
function reset() {
  twitter.reset();
}

// 認証後のコールバック（必須）
function authCallback(request) {
  return twitter.authCallback(request);
}

// ツイートを投稿
function postUpdateStatus() {
  var service  = twitter.getService();
  var response = service.fetch('https://api.twitter.com/1.1/statuses/update.json', {
    method: 'post',
    payload: { status: '***MESSAGE***' }
  });
  Logger.log(JSON.parse(response));
}


/***   シートのやつ   ****/
/***   twitter    ****/
/***   予約投稿するよ   ****/


  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('post');
  var shme = ss.getSheetByName("memory");
  var sh_trigger = ss.getSheetByName('trigger_');

function setTrigger() {
  var shme = ss.getSheetByName("memory");
  var date = new Date();
  var triggerDay = sh.getRange("J2").getValue();
  
  Logger.log(triggerDay);
  
  var tweet = sh.getRange("A2:F2");  
  var targetRange = shme.getRange(shme.getLastRow()+1, 1, 1, 6);
  
  tweet.copyTo(targetRange);
  
  if(date.getDate() == triggerDay.getDate()){
    var trigger = ScriptApp.newTrigger("main").timeBased().at(triggerDay).create();
    var idRange = shme.getRange(shme.getLastRow(), 7);
    Logger.log(trigger.getTriggerSource());
    idRange.setValue(trigger.getUniqueId());
  }
  
  var shx = ss.getSheetByName("backup");
  
  var backupRange = shx.getRange(shx.getLastRow()+1, 1, 1, 6);
  tweet.copyTo(backupRange);
  var backupRangeID = shx.getRange(shx.getLastRow(), 7);
  backupRangeID.setValue(trigger.getUniqueId());
}

/***  日付が変わった時に、その日分のトリガーを仕込む。 ***/
function changeTriggers(){
  var trigger =[];
  var time = [];
  var tweet = [];
  var idRow = ID_lastRow(7)+1;
  for(var i=0; i < sh_trigger.getLastRow(); i++) {
    tweet[i] = sh_trigger.getRange(i+1, 3).getValue();
    time[i] = sh_trigger.getRange(i+1, 9).getValue();
    if(typeof time[i]=="object"){
      trigger[i] = ScriptApp.newTrigger("main").timeBased().at(time[i]).create().getUniqueId();
      shme.getRange(idRow +i, 7).setValue(trigger[i]);
    }
  }
  Logger.log(tweet);
  Logger.log(time);
  Logger.log(trigger);
}

// UNIQUE IDが一致したトリガーを削除
function deleteTrigger(row) {
  var id = shme.getRange(row, 7).getValue();
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < triggers.length; i++) {
    var triggerId = triggers[i].getUniqueId();
    if (triggerId == id) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }  
}

// 実行したいスクリプト本体
function main() {
  var shme = ss.getSheetByName("memory");
  
  var compRow = getLastRowNumber_Column() +1;
  shme.getRange(compRow, 8).setValue("実行済み");
  Logger.log(compRow);
  deleteTrigger(compRow);
  
  var msg = shme.getRange(compRow, 3).getValue();
  Logger.log(msg);
  var service  = twitter.getService();
  var response = service.fetch('https://api.twitter.com/1.1/statuses/update.json', {
    method: 'post',
    payload: { status: msg }
  //  muteHttpExceptions:true,
  });
  Logger.log(JSON.parse(response));
    
}

/*H列の最終行取得*/
function getLastRowNumber_Column(num){
  var bk = SpreadsheetApp.getActiveSpreadsheet();
  var sh = bk.getSheetByName("memory");
  var last_row = sh.getLastRow();

  for(var i = last_row; i >= 1; i--){
    if(sh.getRange(i, 8).getValue() != ''){
      num = i;
      break;
    }
  }
  return num;
}

function deleteTriggerByID(){
  var deleteId = sh.getRange("G6").getValue();
  var targetRow = findRow(deleteId);
  var status = shme.getRange(targetRow, 3).getValue();
  if(shme.getRange(targetRow, 8).getValue() =='実行済み'){
    Browser.msgBox('このツイートは投稿済みです');
    return;
  }
  
  if(Browser.msgBox('ID：'+deleteId+'「'+status+'」 のツイートを削除してもよろしいですか？', Browser.Buttons.OK_CANCEL)){
    var triggers = ScriptApp.getProjectTriggers();
    for(var i=0; i < triggers.length; i++) {
      var triggerId = triggers[i].getUniqueId();
      if (triggerId == deleteId) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    shme.deleteRow(targetRow);
  }else{
    return;
  }   
}

function findRow(val){ //シート内検索
  var lastRow=shme.getDataRange().getLastRow(); //対象となるシートの最終行を取得 
  for(var i=1;i<=lastRow;i++){
    if(shme.getRange(i,7).getValue() === val){
      return i;
    }
  }
  return 0;
}

function ID_lastRow(val){ //memoryのval列からデータの入力されている最終行番号を取得する
　var last_row = shme.getLastRow();

　for(var i = last_row; i >= 1; i--){
　　if(shme.getRange(i, val).getValue() != ''){
　　　return i;
　　}
　}
}

