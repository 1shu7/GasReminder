//Himawari(必須)
var CHANNEL_ACCESS_TOKEN = ''; // Channel_access_tokenを登録

//ShowID関数専用
var DEFAULT_INFORMATION_SPREADSHEET_ID ='';
var DEFAULT_USERID ='';
var DEFAULT_SENDLIST = '';  
var SETTING_SPREADSHEETID ='';  //設定の記録先 showidの管理用(コードに直接書く場合はshowid関数を使用せず,直接入力)
var EMAIL_ADDRESS="";
//ここまで

function ShowId(ID){
  Logger.log("呼ばれたよ");
  Logger.log(ID);
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
         id=CHANNEL_ACCESS_TOKEN; 
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
    default:
      break;
  }
  
  return id;
  
}

function record(event){
  var spreadsheet = SpreadsheetApp.openById(ShowId("SEND_LIST"));
  var sheet = spreadsheet.getSheetByName('シート1');
  var sheet2 = spreadsheet.getSheetByName('シート2');
  var lastRow = sheet.getLastRow();
  
  if((sheet2.getRange(2,2).getValue())==(event.source.userId)){
    sheet.getRange(2,2).setValue(event.source.userId);
    sheet.getRange(2,3).setValue(event.source.groupId);
    Logger.log("1");
    var today=Utilities.formatDate(new Date(),"JST","yyyy/MM/dd HH:mm:ss");
    sheet.getRange(2,1).setValue(today);
  }
}

function changeAdmin(event){
  var spreadsheet = SpreadsheetApp.openById(ShowId("SEND_LIST"));
  var sheet = spreadsheet.getSheetByName('シート2');
  sheet.getRange(2,2).setValue(event.source.userId);
  sheet.getRange(2,3).setValue(event.source.groupId);
  Logger.log("1");
  var today=Utilities.formatDate(new Date(),"JST","yyyy/MM/dd HH:mm:ss");
  sheet.getRange(2,1).setValue(today);
  
  
}

function doPost(e) {
  var event = JSON.parse(e.postData.contents).events[0];
  var replyToken= event.replyToken;

  if (typeof replyToken === 'undefined') {
    return; // エラー処理
  }
  var userId = event.source.userId;
  var nickname = getUserProfile(userId);
  
  

  if(event.type == 'follow') { 
    // ユーザーにbotがフォローされた場合に起きる処理
    pushHelp(event,"友達登録ありがとうございます！\nHimawariはグループに日程のリマインドをします。また、メニューからお知らせを確認できます");
    
  }

  if(event.type == 'message') {
    switch(event.message.text){
      case '$SYS?kanri=change':
        record(event);
        pushHelp(event,"200 OK");
        break;
      case '$SYS?admin=me':
        changeAdmin(event);
        break;
      case '$help':
        pushHelp(event,"$SYS?kanri=change\n送信先の変更\n\n$SYS?admin=me\n管理者の変更");
        break;
    
        
      case '今後の予定は？':
         Logger.log("呼ばれたよ");
         var myCals = CalendarApp.getCalendarById(EMAIL_ADDRESS); //今日の日付配列
         var oneMonth = new Date();
         oneMonth.setDate(oneMonth.getDate()+30);
         var myEvents=myCals.getEvents(new Date(),oneMonth);;  
         var isShow;
         var title="今後の30日間の日程です";
         for(var i=0;i<myEvents.length;i++){
           isShow = myEvents[i].getTitle();
           
           if(isShow.indexOf('中止') !== -1 || isShow.indexOf('延期') !== -1 || isShow.indexOf('非通知') !== -1 ){
              continue;
           }
           
           title += "\n";
           if( myEvents[i].isAllDayEvent() == true){
             title += Utilities.formatDate(myEvents[i].getAllDayStartDate(), "JST", "MM'/'dd")+" ";
             title += myEvents[i].getTitle();
             
           }else{
             title += Utilities.formatDate(myEvents[i].getStartTime(), "JST", "MM'/'dd' 'HH:mm")+" ";
             title += myEvents[i].getTitle();
           }
           
         }
         
         if(title == "今後の30日間の日程です"){
           title = "今日から30日間は何も登録されていません";
         }
         
         pushHelp(event,title);
         break;
 
      case '中止された予定は？':
         var myCals = CalendarApp.getCalendarById( EMAIL_ADDRESS); //今日の日付配列
         var oneMonth = new Date();
         oneMonth.setDate(oneMonth.getDate()+30);
         var myEvents=myCals.getEvents(new Date(),oneMonth);;  
         var isShow;
         var title="今後の30日間の日程です";
         for(var i=0;i<myEvents.length;i++){
           isShow = myEvents[i].getTitle();
           
           if(isShow.indexOf('中止') !== -1 || isShow.indexOf('延期') !== -1  ){      
             title += "\n";
             title += Utilities.formatDate(myEvents[i].getAllDayStartDate(), "JST", "MM'/'dd")+" ";
             title += myEvents[i].getTitle();  
           }
           
         }
         
         if(title == "今後の30日間の日程です"){
           title = "今日から30日間は中止されたものはありません";
         }
         
         pushHelp(event,title);
         break;
        
      default:
         break;
    }

    var userMessage = event.message.text;
    return ContentService.createTextOutput(
      JSON.stringify({'content': 'post ok'})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function pushHelp(event,text){
  var url = "https://api.line.me/v2/bot/message/push";
  var headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
  };
  
  

  //toのところにメッセージを送信したいユーザーのIDを指定します。(toは最初の方で自分のIDを指定したので、linebotから自分に送信されることになります。)
  //textの部分は、送信されるメッセージが入ります。createMessageという関数で定義したメッセージがここに入ります。
  var postData = {
    "to" : event.source.userId,
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
  
// profileを取得してくる関数
function getUserProfile(userId){ 
  var url = 'https://api.line.me/v2/bot/profile/' + userId;
  var userProfile = UrlFetchApp.fetch(url,{
    'headers': {
      'Authorization' :  'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
  })
  return JSON.parse(userProfile).displayName;
}
