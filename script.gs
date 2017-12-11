// OAuth1認証用インスタンス
var twitter = TwitterWebService.getInstance(
  'feEwUJqXZn22lVaivGP7V30MC',
  'FjfnCSfGPBqeoO38byxwT9cqqQRlqzrh10dXp84BEsxArfk0uJ'
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

// タイムラインを取得
function getUserTimeline() {
  var service  = twitter.getService();
  var response = service.fetch('https://api.twitter.com/1.1/statuses/user_timeline.json');
  Logger.log(JSON.parse(response));
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


function setTrigger() {
 
  var triggerDay = sh.getRange("J2").getValue();
  
  Logger.log(triggerDay);
  
  var trigger = ScriptApp.newTrigger("main").timeBased().at(triggerDay).create();
  var tweet = sh.getRange("A2:F2");
  
  var targetRange = shme.getRange(shme.getLastRow()+1, 1, 1, 6);
  
  tweet.copyTo(targetRange);
  
  var idRange = shme.getRange(shme.getLastRow(), 7)
  idRange.setValue(trigger.getUniqueId());
  
  var shx = ss.getSheetByName("backup");
  
  var buckupRange = shx.getRange(shx.getLastRow()+1, 1, 1, 6);
  tweet.copyTo(buckupRange);
  
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
