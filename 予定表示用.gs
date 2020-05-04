var CHANNEL_ACCESS_TOKEN = ''; // Channel_access_tokenを登録
var DEFAULT_INFORMATION_SPREADSHEET_ID ="";
var DEFAULT_USERID ='';
var DEFAULT_SENDLIST = "";
var SETTING_SPREADSHEETID ="";
var DEFAULT_FOLDER_ID = "";  //監視するフォルダID(デフォルト値)
var DEFAULT_UPDATECHECK_SSID = ""; //フォルダアップデートを記録するSpreadSheetID
var E_MAILADDRESS = "";

function escapeHTML(str) {
 str = str.replace(/&/g, '&amp;');
 str = str.replace(/</g, '&lt;');
 str = str.replace(/>/g, '&gt;');
 str = str.replace(/"/g, '&quot;');
 str = str.replace(/'/g, '&#39;');
 return str;
}

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

function doGet(){
  return HtmlService.createHtmlOutput(URL_calendar());
}


function URL_calendar(){
  
  //css設定
  var html="<style>h3 {position: relative;color: white;padding:0.5em 0.5em 0.5em 2em;background-color: #446689;border-radius:1.5em;}h3::after {position: absolute;top: 50%;left:1em;transform:translateY(-50%);content: '';width: 13px;height:13px;background-color: white; border-radius:100%;}</style>\n";
  
  html+=''; //google カレンダー埋め込み
  //何も登録されてない場合
  
  //先頭メッセージ
  html+='<div><h2><font color="orangered">テスト</font></h2><a href="google.co.jp" target="_blank">大学のサイトを確認</a></div>';
  //
  var myCals=CalendarApp.getCalendarById(E_MAILADDRESS); //特定のIDのカレンダーを取得
  var today = new Date();
  var to_endday = new Date();
  to_endday.setDate(today.getDate() + 200);
  var Events = myCals.getEvents(today,to_endday);
  var tmp_event;
   
  
  for(var i=0;i<Events.length;i++){
    var start_date;
    var end_date;
    var DateA;
    tmp_event = Events[i];
    var isShow = tmp_event.getTitle();
    
    if(isShow.indexOf('中止') !== -1 || isShow.indexOf('延期') !== -1 || isShow.indexOf('非通知') !== -1 ){
      continue;
    }
    
    html+='<div class="content">' ;
    html+="<h3>"+escapeHTML(tmp_event.getTitle())+"</h3>";
    html+="<h4>日時:";
    
    
    if(tmp_event.isAllDayEvent() == true){
      start_date = Utilities.formatDate(tmp_event.getAllDayStartDate(),"JST","MM'/'dd");
      var tmp = tmp_event.getAllDayEndDate();
      tmp.setMinutes(tmp.getMinutes()-1);
      end_date = Utilities.formatDate(tmp,"JST","MM'/'dd");
      
      if(start_date == end_date){
        DateA=start_date;
      }else{
        DateA = start_date + " ~ " + end_date;
      }
      
    }else{
      var tmp_start_date = tmp_event.getStartTime();
      var tmp_end_date = tmp_event.getEndTime();
      start_date = Utilities.formatDate(tmp_start_date, "JST", "MM'/'dd' 'HH:mm' ~ '");
      end_date = Utilities.formatDate(tmp_end_date, "JST", "MM'/'dd' 'hh:mm");
      if((Utilities.formatDate(tmp_start_date,"JST", "MM'/'dd")) == (Utilities.formatDate(tmp_end_date, "JST", "MM'/'dd"))){
        DateA = start_date + Utilities.formatDate(tmp_end_date, "JST", "HH:mm");
      }else{
        DateA = start_date + end_date;
      }
    }
    
    html+= DateA + "</h4>";

    html+="</h4>";
    html+="<h4>場所:"+escapeHTML(tmp_event.getLocation()+" ")+"</h4>";
    html+=escapeHTML(tmp_event.getDescription());
    var isEmpty = tmp_event.isAllDayEvent();
    html+='<a href=\"'+RegistURL(tmp_event,isEmpty)+'\"'+ 'target="_blank"><br>カレンダーに登録</a>';
    
    html+="</div>";
    
  }
    //お問い合わせ欄
    html+='<br><br><a href="mailto:example@example.example?body=%E4%BB%A5%E4%B8%8B%E3%81%AE%E6%83%85%E5%A0%B1%E3%82%92%E5%85%A5%E5%8A%9B%E3%81%97%E3%81%A6%E3%80%81%E9%80%81%E4%BF%A1%E3%81%8A%E9%A1%98%E3%81%84%E3%81%97%E3%81%BE%E3%81%99%E3%80%82%0D%0A%0D%0A%E3%81%8A%E5%90%8D%E5%89%8D%3A%0D%0A%E3%81%B5%E3%82%8A%E3%81%8C%E3%81%AA%3A%0D%0A%E6%89%80%E5%B1%9E%28%E9%83%A8%E7%BD%B2%29%3A%0D%0A%E3%81%94%E6%84%8F%E8%A6%8B%3A">技術的なお問い合わせはこちらまで</a>';
  return html;
  
  

}



///////////////////////////////////////////////////////////////////////////////////////////////////
//カレンダー作成////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////

//タイムスタンプ	メールアドレス	予定名	場所	予定の説明	イベントの長さ	開始日程	終了日程	開始日程(終日)	終了日程	登録者名
// 0              1             2    3     4             5         6         7          8                9       10
function RegistURL(e,isEmpty){
  Logger.log(isEmpty);
  
  var DateS= new Date();
  var DateE= new Date();
  
    if(isEmpty == true){
      DateS= e.getAllDayStartDate();
      DateE= e.getAllDayStartDate();
      
      DateE.setDate(DateE.getDate() + 1);  //bag対応(不具合あったら消す)
      DateS=Utilities.formatDate(DateS,"JST","yyyyMMdd");
      DateE=Utilities.formatDate(DateE,"JST","yyyyMMdd"); 
      
    }else{
  
      DateS= e.getStartTime();
      DateE= e.getEndTime();
    
      DateS=Utilities.formatDate(DateS,"JST","yyyyMMdd'T'HHmmss");
      DateE=Utilities.formatDate(DateE,"JST","yyyyMMdd'T'HHmmss"); 
  }
    
    var text=e.getTitle();
    var location= e.getLocation();
    var details= e.getDescription();
  
  if(location==""){
    location=" ";
  }
   if(details==""){
    details="詳細情報はありません";
  }
    
  return 'https://www.google.com/calendar/event?action=TEMPLATE'+
    '&text='    + encodeURIComponent(text) +
    '&dates='   + DateS + '/' + DateE+
    '&details=' + encodeURIComponent(details) +
    '&location='+ encodeURIComponent(location);
}
