FormApp.getActiveForm();


var CHANNEL_ACCESS_TOKEN = ''; // Channel_access_tokenを登録
var DEFAULT_INFORMATION_SPREADSHEET_ID ="";
var DEFAULT_USERID ='';
var DEFAULT_SENDLIST = "";
var SETTING_SPREADSHEETID ="";
var DEFAULT_FOLDER_ID = "";  //監視するフォルダID(デフォルト値)
var DEFAULT_UPDATECHECK_SSID = ""; //フォルダアップデートを記録するSpreadSheetID
var E_MAILADDRESS = ""; //予定が追加されたら送るメールアドレス
var ScheduleURL ="" ; //予定表示用URL


//0.予定名	1.場所	2.予定の説明	3.イベントの長さ	4.開始日程(指定)	5.終了日程(指定)	6.登録者名 7.今すぐ？ 
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

function submitForm(e) {
　
  var itemResponses = e.response.getItemResponses();
  var tempAns = [];
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    var answer = itemResponse.getResponse();
    tempAns.push(answer);
  }

  tempAns[4]=new Date(tempAns[4]);
  tempAns[5]=new Date(tempAns[5]);
 
  const ans=tempAns.concat();
  Logger.log("concat="+ans);
  
  var today=new Date();
  
  var calendar = CalendarApp.getDefaultCalendar();
  var rtoday=Utilities.formatDate(today, "JST", "yyyy'年'MM'月'dd'日'")
  var option={
    description:ans[2],
    location:ans[1]
  };
  
  if(ans[7]=="はい")
  {
    pushflexMessage(ans);
    
   var date = new Date(ans[8]);
    if((date.getTime())<(today.getTime())){
      
      pushflexMessage(ans);
      
    }
  }
  
  if(ans[3]=="終日"){
    ans[5].setMinutes(ans[5].getMinutes+1); 
    calendar.createAllDayEvent(ans[0],ans[4],ans[5],option);
  }else{
    calendar.createEvent(ans[0],ans[4],ans[5], option);
  }
  

  MailApp.sendEmail(E_MAILADDRESS, ans[0]+'がGoogleFormで追加されました',"登録内容は以下の通りです。\n"+ans, {noReply: true});

}

///////////////////////////////////////////////////////////////////////////////////////////////////
//↓LINEID////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////



function userid(){
  var spreadsheet = SpreadsheetApp.openById(ShowId("SEND_LIST"));
  var sheet = spreadsheet.getSheetByName('シート1');
  var lastRow = sheet.getLastRow();
  var userID=sheet.getRange(lastRow,2).getValue();
  var groupID=sheet.getRange(lastRow,3).getValue();
  
  if(groupID!=null){
    return groupID;
  }else if(userID!=null){
    return userID;
  }else{
    return ShowId("USER_ID"); 
  }
  
}

///////////////////////////////////////////////////////////////////////////////////////////////////
//↓LINEflex送信部////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////

//0.予定名	1.場所	2.予定の説明	3.イベントの長さ	4.開始日程	5.終了日程	6.開始日程	7.終了日程	8.登録者名
//  
function pushflexMessage(e) {
  
  var postData;
  const StartDate=Utilities.formatDate(e[4], "JST", "yyyyMMdd");
  const EndDate=Utilities.formatDate(e[5], "JST", "yyyyMMdd");
  var DateA;
  var userID = userid();
  Logger.log("MessageAPI\n"+e);
  
  
  if (e[3] == "終日") {
    if (StartDate != EndDate) {
      DateA = Utilities.formatDate(e[4], "JST", "MM'/'dd' '") + '~' + Utilities.formatDate(e[5], "JST", "MM'/'dd' '") + '\n';
    } else {
      DateA = Utilities.formatDate(e[4], "JST", "MM'/'dd' '") + '\n';
    }
  } else {
    if (StartDate != EndDate) {
      DateA = Utilities.formatDate(e[4], "JST", "MM'/'dd' 'HH:mm") + '~' + Utilities.formatDate(e[5], "JST", "MM'/'dd' 'HH:mm") + '\n';
    } else {
      DateA = Utilities.formatDate(e[4], "JST", "MM'/'dd' 'HH:mm") + '~' + Utilities.formatDate(e[5], "JST", "HH:mm") + '\n';
    }
  }
  //0.予定名	1.場所	2.予定の説明	3.イベントの長さ	4.開始日程	5.終了日程 8.登録者名
  postData = {
    "to": userID,
    "messages": [{
      "type": "flex",
      "altText": '【' + e[0] + '】',
      "contents": {
        "type": "bubble",
        "hero": {
          "type": "box",
          "layout": "vertical",
          "contents": [{
            "type": "text",
            "text": "イベント追加",
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
            "text": "以下の日程が登録されました",
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
            "text": e[0], //予定名
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
                "text": e[1], //場所
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
              "text": e[2],
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
            "type": "button",
            "action": {
              "type": "uri",
              "label": "Googleカレンダーに登録",
              "uri": RegistURL(e) //google_URL
            }
          }, {
            "type": "spacer"
          }],
          "flex": 0
        }
      }
    }]
  };

  var url = "https://api.line.me/v2/bot/message/push";
  var headers = {
    "Content-Type": "application/json",
    'Authorization': 'Bearer ' + ShowId("CHANNEL_ACCESS_TOKEN"),
  };
  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(postData),
    muteHttpExceptions: true,
  };
  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response);
}

///////////////////////////////////////////////////////////////////////////////////////////////////
//↑LINEflex送信部ここまで////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////////////
//カレンダー作成////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////

//0.予定名	1.場所	2.予定の説明	3.イベントの長さ	4.開始日程	5.終了日程	6.開始日程	7.終了日程	8.登録者名
function RegistURL(ans) {
  var EndDate=new Date();
  
  if (ans[3] == "終日") {
    EndDate.setMinutes(ans[5].getMinutes()-1);
    var DateS = Utilities.formatDate(ans[4], "JST", "yyyyMMdd");
    var DateE = Utilities.formatDate(EndDate, "JST", "yyyyMMdd");
  } else {
    var DateS = Utilities.formatDate(ans[4], "JST", "yyyyMMdd'T'HHmmss");
    var DateE = Utilities.formatDate(ans[5], "JST", "yyyyMMdd'T'HHmmss");
  }
  
  var text = ans[0];
  var location = ans[1];
  var details = ans[2];
  if (location == "") {
    location = " ";
  }
  if (details == "") {
    details = "詳細情報はありません";
  }
  
  return 'https://www.google.com/calendar/event?action=TEMPLATE'+
    '&text='    + encodeURIComponent(text) +
    '&dates='   + DateS + '/' + DateE+
    '&details=' + encodeURIComponent(details) +
    '&location='+ encodeURIComponent(location);
}
