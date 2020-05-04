var CHANNEL_ACCESS_TOKEN = ''; // Channel_access_tokenを登録
var DEFAULT_INFORMATION_SPREADSHEET_ID ="";
var DEFAULT_USERID ='';  //LINE送信先ユーザID
var DEFAULT_SENDLIST = "";  //送信先ID記録用SpreadsheetID
var SETTING_SPREADSHEETID ="";  //参照先SpreadSheetID
var DEFAULT_FOLDER_ID = "";  //監視するフォルダID(デフォルト値)
var DEFAULT_UPDATECHECK_SSID = ""; //フォルダアップデートを記録するSpreadSheetID
var EMAIL = "";
var ScheduleURL="";  //予定登録用URL(GAS)

function ShowId(ID){
  var spreadsheet;
  var sheet;
  try{
    spreadsheet = SpreadsheetApp.openById(SETTING_SPREADSHEETID);
    sheet = spreadsheet.getSheetByName('シート1');
  }
  catch(e){
    Logger.log("【重大なエラー】設定ファイルを読み込めませんでした。");
  }
  
  var id;
  switch (ID){
    case "CHANNEL_ACCESS_TOKEN":
      id=sheet.getRange(2,2).getValue();
      if(id=="" || id===undefined){
         id= CHANNEL_ACCESS_TOKEN ; 
      }
      break;
    case "INFO_SPREADSHEET":
      id= sheet.getRange(5,2).getValue();
      if(id=="" || id===undefined){
         id=DEFAULT_INFORMATION_SPREADSHEET_ID;
      }
      break;
    case "USER_ID":
      id=sheet.getRange(3,2).getValue();
      if(id=="" || id===undefined){
         id=DEFAULT_USERID;
      }
      break;
    case "SEND_LIST":
      id=sheet.getRange(7,2).getValue();
      if(id=="" || id===undefined){
         id=DEFAULT_SENDLIST;
      }
      break;
    case "FOLDER_ID":
      id=sheet.getRange(6,2).getValue();
      if(id=="" || id===undefined){
         id=DEFAULT_FOLDER_ID;
      }
      break;
    case "UPDATE_SHEET_ID":
      id=sheet.getRange(9,2).getValue();
      if(id=="" || id===undefined){
         id=DEFAULT_UPDATECHECK_SSID;
      }
      break;
    default:
      break;
  }
  
  return id;
  
}

//ユーザID参照
function userid(){
  
  //UserID設定の構成
  //シート1
  //--タイムスタンプ --USERID --GROUPID
  //GROUPIDが優先されて送信される。
  //固定長にする場合は全てにグローバル変数USER_IDを設定する･
  
  var spreadsheet = SpreadsheetApp.openById(ShowId("SEND_LIST"));  //UserIDを登録しているSpreadSheetID
  var sheet = spreadsheet.getSheetByName('シート1');
  var lastRow = sheet.getLastRow();
  var userID=sheet.getRange(2,2).getValue();
  var groupID=sheet.getRange(2,3).getValue();
  
  if(groupID!=null){
    return groupID;
  }else if(userID!=null){
    return userID;
  }else{
    return USER_ID; 
  }
  
}

//ファイル更新用のMessageAPI送信
function push(text) {
//メッセージを送信(push)する時に必要なurlでこれは、皆同じなので、修正する必要ありません。
//この関数は全て基本コピペで大丈夫です。
  var url = "https://api.line.me/v2/bot/message/push";
  var headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    'Authorization': 'Bearer ' + ShowId("CHANNEL_ACCESS_TOKEN"),
  };

  //toのところにメッセージを送信したいユーザーのIDを指定します。(toは最初の方で自分のIDを指定したので、linebotから自分に送信されることになります。)
  //textの部分は、送信されるメッセージが入ります。createMessageという関数で定義したメッセージがここに入ります。
  var postData = {
    "to" : userid(),
    "messages" : [
      {
        'type':'text',
        'text':text,
      }
    ]
  };

  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };

  return UrlFetchApp.fetch(url, options);
}

var TARGET_FOLDER_ID = ShowId("FOLDER_ID");  //監視するGoogleDriveフォルダID
var UPDATE_SHEET_ID = ShowId("UPDATE_SHEET_ID"); //更新を記録する作業用SpreadsheetID
var UPDATE_SHEET_NAME = "シート1";  //上記のシート名

//１時間に１回のメインプログラム
function updateCheck() {

  
  //
  //以下フォルダ更新確認//
  //
  //
  var targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  var folders = targetFolder.getFolders();
  var files = targetFolder.getFiles();

  //フォルダ内を再帰的に探索してすべてのファイルIDを配列にして返す
  function getAllFilesId(targetFolder){
    var filesIdList = [];
    
    var files = targetFolder.getFiles();
    while(files.hasNext()){
      filesIdList.push(files.next().getId());
    }
    
    var child_folders = targetFolder.getFolders();
    while(child_folders.hasNext()){
      var child_folder = child_folders.next();
      //Logger.log( 'child_folder :' + child_folder );

      //Logger.log('getAllFilesId(child_folder):'+ getAllFilesId(child_folder));
      filesIdList = filesIdList.concat( getAllFilesId(child_folder) );
    }
    return filesIdList;
  }
  //Logger.log('getAllFilesId(targetFolder):' + getAllFilesId(targetFolder));
  var allFilesId = getAllFilesId(targetFolder);
  var lastUpdateMap = {};
  //Logger.log(folders)
  allFilesId.forEach(
    function( value, i ){
      var file =DriveApp.getFileById( value );
      lastUpdateMap[file.getName()] = {lastUpdate : file.getLastUpdated(), fileId: file.getId()};
    }
  );          
 
  // スプレッドシートに記載されているフォルダ名と更新日時を取得。
  var spreadsheet = SpreadsheetApp.openById(UPDATE_SHEET_ID);
  var sheet = spreadsheet.getSheetByName(UPDATE_SHEET_NAME);
  //Logger.log(sheet)
  var data = sheet.getDataRange().getValues();
  //Logger.log('data: ' + data)
  // 取得したデータをMapに変換。
  var sheetData = {};
  for (var i = 0; i < data.length; i++) {
    sheetData[data[i][0]] = {name : data[i][0], lastUpdate : data[i][1], rowNo : i + 1};
  }

  // 実際のフォルダとスプレッドシート情報を比較。
  var updateFolderMap = [];
  for (key in lastUpdateMap) {
    if( UPDATE_SHEET_ID == lastUpdateMap[key].fileId ){
      continue;
    }
    if(key in sheetData) {
      // フォルダ名がシートに存在する場合。
      if(lastUpdateMap[key].lastUpdate > sheetData[key].lastUpdate) {
        // フォルダが更新されている場合。
        sheet.getRange(sheetData[key].rowNo, 2).setValue(lastUpdateMap[key].lastUpdate);
        sheet.getRange(sheetData[key].rowNo, 3).setValue(lastUpdateMap[key].fileId);
        updateFolderMap.push({filename:key, lastUpdate:lastUpdateMap[key].lastUpdate, fileId:lastUpdateMap[key].fileId});
      }
    } else {
      // フォルダ名がシートに存在しない場合。
      var newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, 1).setValue(key);
      sheet.getRange(newRow, 2).setValue(lastUpdateMap[key].lastUpdate);
      sheet.getRange(newRow, 3).setValue(lastUpdateMap[key].fileId);
      updateFolderMap.push({filename:key, lastUpdate:lastUpdateMap[key].lastUpdate, fileId:lastUpdateMap[key].fileId});
    }
  }
  //Logger.log('updateFolderMap:' + updateFolderMap)
  // 新規及び更新された情報をメール送信。
  var updateText=' ';
  var updateURL=[];
  for( key in updateFolderMap ){
    item = updateFolderMap[key];
    updateText+= item.filename + '\n更新日時：' + Utilities.formatDate(item.lastUpdate, "JST", "yyyy/MM/dd HH:mm") + "\n"+ DriveApp.getFileById(item.fileId).getUrl()+"\n\n";
  }
  
  if (updateFolderMap.length != 0) {
    
 
    push(
                        "【" + targetFolder.getName() + "】が更新されました。\n\n"+
                        updateText+"\n\n");
 

  }
}


///
///1日1回実行///
///
//

function check_notification(){
  inspection_calendar();
  infoKanshi();  //お知らせ確認
}

///
///
///
///

///////////////////////////////////////////////////////////////////////////////////////////////////
//予定監視//////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////////////
//お知らせ監視///////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////

//spreadSheetの構成
//0.タイムスタンプ	1.メールアドレス	2.予定名	3.場所	4.予定の説明	5.イベントの長さ	6.開始日程	7.終了日程	8.開始日程	9.終了日程	10.登録者名	11.MARK

function infoKanshi(){
  
  var spreadsheet = SpreadsheetApp.openById(ShowId("INFO_SPREADSHEET"));  //お知らせをキロクするSpreadSheetID
  var sheet = spreadsheet.getSheetByName('フォームの回答 1');
  var lastrow= sheet.getLastRow();
  var today=new Date();
  today=today.getTime();
  
  
  for(var i=0;i<lastrow;i++){
    var ans=sheet.getRange(i+2,1,1,11).getValues();
    var regday=sheet.getRange(i+2,5).getValue();
    var sendDay=new Date(ans[0][9]);
    regday=new Date([0][4]);
    regday=regday.getTime();
    sendDay=sendDay.getTime();
    var strDate=new Date(ans[0][4]);
    var endDate=new Date(ans[0][5]);
    strDate=strDate.getTime();
    endDate=endDate.getTime();
   
    if(today>regday){
      if(sheet.getRange(i+2,11).getValue()!='o'){
        var e=sheet.getRange(i+2,1,1,11).getValues();
         pushMSG(e);
         sheet.getRange(i+2,11).setValue('o');
      }
      
    }
    if(today>=sendDay && ans[0][8]=="指定日" && strDate>=today && endDate <=today ){
      if(sheet.getRange(i+2,11),getValue()!='o'){
         var e=sheet.getRange(i+2,1,1,11).getValues();
         pushMSG(e);
         sheet.getRange(i+2,11).setValue('o');
      }
    }
  }
}
  




///////////////////////////////////////////////////////////////////////////////////////////////////
//お知らせ用FLEXメッセージ////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////

//お知らせ用
//0.タイムスタンプ	1.属性	2.発信概要	3.発信事項	4.掲載開始日	5.掲載終了日	6.登録者	7.ADDRESS	8.MARK
//引数に送信メッセージ

function pushMSG(e) {
  
  var userID = userid();
  var postData;
  postData = {
    "to": userID,
    "messages": [{
      "type": "flex",
      "altText": '【' + e[0][1] + '】' + e[0][2],
      "contents": {
        "type": "bubble",
        "hero": {
          "type": "box",
          "layout": "vertical",
          "contents": [{
            "type": "text",
            "text": "■"+e[0][1],
            "weight": "bold",
            "position": "relative",
            "align": "start",
            "gravity": "top",
            "wrap": true,
            "size": "lg",
            "margin": "md",
            "decoration": "none",
            "offsetStart": "10px",
            "offsetTop": "5px",
            "color": "#f1933b"
          }, {
            "type": "text",
            "text": "お知らせが追加されました",
            "size": "xs",
            "offsetStart": "10px",
            "offsetTop": "2px",
            "color": "#0b82ff"
          }]
        },
        "body": {
          "type": "box",
          "layout": "vertical",
          "contents": [{
            "type": "text",
            "text": e[0][2],
            "weight": "bold",
            "size": "xl"
          }, {
            "type": "box",
            "layout": "vertical",
            "margin": "lg",
            "spacing": "sm",
            "contents": [{
              "type": "box",
              "layout": "baseline",
              "spacing": "sm",
              "contents": [{
                "type": "text",
                "text": "詳細",
                "color": "#aaaaaa",
                "size": "sm",
                "flex": 1
              }, {
                "type": "text",
                "text": e[0][3],
                "wrap": true,
                "color": "#666666",
                "size": "sm",
                "flex": 5
              }]
            }]
          }]
        },
        "footer": {
          "type": "box",
          "layout": "vertical",
          "spacing": "sm",
          "contents": [{
            "type": "button",
            "style": "link",
            "height": "sm",
            "action": {
              "type": "uri",
              "label": "他のお知らせを確認",
              "uri": ScheduleURL
            }
          }, {
            "type": "spacer"
          }],
          "flex": 0
        }
      }
    }]
  };

  push_flex(postData);
  
}

function inspection_calendar()
{
  var myCals = CalendarApp.getCalendarById(EMAIL); //今日の日付配列
  Logger.log(myCals);
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate()+1);
  var myEvents=myCals.getEventsForDay(tomorrow);  
  var isShow;
  
  for(var i=0;i<myEvents.length;i++){
    isShow = myEvents[i].getTitle();
    if(isShow.indexOf('中止') !== -1 || isShow.indexOf('延期') !== -1 || isShow.indexOf('非通知') !== -1 ){
      continue;
    }
    push_calendar(myEvents[i]);
  }
 
  
}



function push_calendar(e) {
 var userID=userid();
 var postData;
 var startDate;
 var endDate=new Date();
 var DateA;
 endDate.setMinutes(endDate.getMinutes()-1);
 
 if(e.isAllDayEvent() == true){
   startDate = Utilities.formatDate(e.getAllDayStartDate(), "JST", "MM'/'dd");
   endDate = Utilities.formatDate(endDate, "JST", "MM'/'dd");
   if(startDate == endDate){
     DateA = startDate;
   }else{
     DateA = startDate + " ~ " + endDate;
   }
   
   Logger.log(endDate);
 
 }else{
   var tmp_start_date = e.getStartTime();
   var tmp_end_date = e.getEndTime();
   
   if((Utilities.formatDate(tmp_start_date, "JST", "MM'/'dd")) == (Utilities.formatDate(tmp_end_date, "JST", "MM'/'dd"))){
     DateA = Utilities.formatDate(tmp_start_date, "JST", "MM'/'dd' 'hh:mm' ~ '") + Utilities.formatDate(tmp_end_date, "JST","hh:mm");
   }else{
     DateA = Utilities.formatDate(tmp_start_date, "JST", "MM'/'dd' 'hh:mm' ~ '") + Utilities.formatDate(tmp_end_date, "JST","MM'/'dd' 'hh:mm");
   }
 }
  
  
  postData = {
    "to": userID,
    "messages": [{
      "type": "flex",
      "altText": '【リマインド】',  //トークプレビューのタイトル
      "contents": {
        "type": "bubble",
        "hero": {
          "type": "box",
          "layout": "vertical",
          "contents": [{
            "type": "text",
            "text": "☆リマインド☆",  //タイトル
            "weight": "bold",
            "position": "relative",
            "align": "start",
            "gravity": "top",
            "wrap": true,
            "size": "lg",
            "margin": "md",
            "decoration": "none",
            "offsetStart": "10px",
            "offsetTop": "5px"
          }, {
            "type": "text",
            "text": "明日です!",  //body
            "size": "xs",
            "offsetStart": "10px",
            "offsetTop": "2px"
          }]
        },
        "body": {
          "type": "box",
          "layout": "vertical",
          "contents": [{
            "type": "text",
            "text": e.getTitle(), //予定名
            "weight": "bold",
            "size": "xl"
          }, {
            "type": "box",
            "layout": "vertical",
            "margin": "lg",
            "spacing": "sm",
            "contents": [{
              "type": "box",
              "layout": "baseline",
              "spacing": "sm",
              "contents": [{
                "type": "text",
                "text": "場所",
                "color": "#aaaaaa",
                "size": "sm",
                "flex": 1
              }, {
                "type": "text",
                "text": e.getLocation()+" ", //場所
                "wrap": true,
                "color": "#666666",
                "size": "sm",
                "flex": 5
              }]
            }, {
              "type": "box",
              "layout": "baseline",
              "spacing": "sm",
              "contents": [{
                "type": "text",
                "text": "日時",
                "color": "#aaaaaa",
                "size": "sm",
                "flex": 1
              }, {
                "type": "text",
                "text": DateA, //日時
                "wrap": true,
                "color": "#666666",
                "size": "sm",
                "flex": 5
              }]
            }]
          }, {
            "type": "box",
            "layout": "baseline",
            "spacing": "sm",
            "contents": [{
              "type": "text",
              "text": "詳細",
              "color": "#aaaaaa",
              "size": "sm",
              "flex": 1
            }, {
              "type": "text",
              "text": e.getDescription()+" ",  //description
              "wrap": true,
              "color": "#666666",
              "size": "sm",
              "flex": 5
            }]
          }]
        },
        "footer": {
          "type": "box",
          "layout": "vertical",
          "spacing": "sm",
          "contents": [{
            "type": "button",
            "style": "link",
            "height": "sm",
            "action": {
              "type": "uri",
              "label": "他の日程を確認",
              "uri": ScheduleURL //他の日程
            }
          }, {
            "type": "spacer"
          }],
          "flex": 0
        }
      }
    }]
  };

  push_flex(postData);
  

}

function push_flex(post_data){

  var url = "https://api.line.me/v2/bot/message/push";
  var headers = {
    "Content-Type": "application/json",
    'Authorization': 'Bearer ' + ShowId("CHANNEL_ACCESS_TOKEN"),
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(post_data)
  };
  
  var response = UrlFetchApp.fetch(url, options);
  
}
