// 【コンタクトのリストから企業名一覧を取得する】

// 各値のセット
var CLIENT_ID = '自分のクライアントID';
var CLIENT_SECRET = '自分のクライアントシークレット';
var SCOPE = 'contacts';
var AUTH_URL = 'https://app.hubspot.com/oauth/authorize';
var TOKEN_URL = 'https://api.hubapi.com/oauth/v1/token';

// ライブラリから。各種値をゲット・セット（実は中身あまり分かってない）
function getService() {
 return OAuth2.createService('hubspot')
 .setTokenUrl(TOKEN_URL)
 .setAuthorizationBaseUrl(AUTH_URL)
 .setClientId(CLIENT_ID)
 .setClientSecret(CLIENT_SECRET)
 .setCallbackFunction('authCallback')
 .setPropertyStore(PropertiesService.getUserProperties())
 .setScope(SCOPE);
}

//　以下2つの関数で初回認証を行う。参考：https://medium.com/how-to-lean-startup/create-a-hubspot-custom-dashboard-with-google-spreadsheet-and-data-studio-27f9c08ade8d
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
     return HtmlService.createHtmlOutput('Success!');
  } else {
     return HtmlService.createHtmlOutput('Denied.');
  }
}
function authenticate() {
  var service = getService();
  if (service.hasAccess()) {
  } else {
     var authorizationUrl = service.getAuthorizationUrl();
     Logger.log('Open the following URL and re-run the script: %s',authorizationUrl);
  }
}

// スプレからこの関数を実行する。他関数の実行を行う。
function exeFunc(){
 // スプレ上のヘッダーになる名前と、取得したいリストID　（HuBspotのリストURLの最後の数字列）
 var list_dict = {"名前（例えば問い合わせリストとか）":'リストID',"名前":'リストID',"名前":'リストID',"名前":'リストID'}
 var sheet_name = "参照シート(HubSpot各ステージ)";
 var colnum = 0;
 deleteResults(sheet_name); //更新時にシートの内容を消しておく。
 for (key in list_dict){
   each_list_data_result = getContactList(list_dict[key], key);
   colnum ++;
   writeResults(sheet_name, [each_list_data_result],colnum);  //setValues関数使う用に[]を重ねる。
 };
};

// 諸々の処理を行う関数。
function getContactList(list_num, lavel_name) {
 var service = getService();
 var options = {headers: {'Authorization': 'Bearer ' +   service.getAccessToken()}};
 var numResults = 0;
 var each_list_data = new Array();
 each_list_data.push(lavel_name);

 // 1リクエストでのデータ数に制限があるのでページがなくなるまで行う。
 var go = true;
 var hasMore = false;
 var offset = 0;
 while (go)
 {
   var url_query = 'https://api.hubapi.com/contacts/v1/lists/'+ list_num +'/contacts/all';
   if (hasMore)
   {
     //各パラメータセット（他のオプションはドキュメント参照）
     url_query += "?count=100"+　"&vidOffset=" + offset;
   }
   var response = UrlFetchApp.fetch(url_query, options).getContentText();
   response = JSON.parse(response);
   hasMore = response['has-more'];
   offset = response['vid-offset'];

   if (!hasMore)
   {
     go = false;
   }
   // hasOwnPropertyで探してなければ”NA”を返す。後で重複を消すので表記ゆれを綺麗にしておく。
   response.contacts.forEach(function(item) {
     var company = (item.properties.hasOwnProperty('company')) ? item.properties.company.value.replace(/株式会社|（株）| | |Inc.|NPO法人|有限会社|（株）|(有)|㈱|()|,/g, ""): "NA";
     each_list_data.push(company);
     //データ数を追加
     numResults++;
   });
 }
 var dd_each_list_data = each_list_data.filter(function (x, i, self) {
   return self.indexOf(x) === i;
 });
 return dd_each_list_data ;

}
// スプレへの書き込み関数を定義
function writeResults(sheetName, results, colnum)
{
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
 var setRange = sheet.getRange(1, colnum, results[0].length, results.length);
 var _ = Underscore.load();
 var transData = _.zip.apply(_, results);
 setRange.setValues(transData);
}

// シートのデータリセット関数を定義
function deleteResults(sheetName)
{
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
 sheet.clear();
}
