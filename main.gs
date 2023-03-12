/*
{
    "action": "insert", //シートの操作、追加:"insert",削除:"delete",変更:"replace",トークン発行:"issue"
    "sheetName": "issue-token", //使用するシートの名前
    "rows": [
        {messageID: "1070596170981843004", //トークン付与されたDiscordメッセージのID
        userID: "1068645423327215697", //トークン付与された人のDicordID
        issuerID:"977472703856513054",//トークン付与者のID、ファウンダーや管理者など？
        issue: 3} //トークン付与数
    ]
}
*/
function doPost(e){
  let postContent = e.postData.getDataAsString();
  main(postContent);
  let output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  output.setContent(JSON.stringify({ message: "success" }));
  return output;
}

function test(){
  let spsheet = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = spsheet.getSheetByName("issue-token");
  let headerRowNum = getHeaderRowNum(sheet);
  let datas = sheet.getDataRange().getValues();
  let headerRow = datas[headerRowNum - 1];
  console.log(headerRow);
}

function main(message){
  let json = JSON.parse(message);
  let spsheet = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = spsheet.getSheetByName(json.sheetName);
  let headerRowNum = getHeaderRowNum(sheet);
  let sheetDatas = sheet.getDataRange().getValues();
  let headerRow = sheetDatas[headerRowNum - 1];
  let idColomnNum = getIdColomnNum(headerRow);

  if (json.action == "replace"){
    replaceRows(sheet, headerRowNum, idColomnNum, json);
  }else if(json.action == "delete"){
    deleteRows(sheet, headerRowNum, idColomnNum, json)
  }else if(json.action == "insert"){
    insertRows(sheet, headerRow, json)
  }
}
// 固定されている行数を取得する。なかったら1を返却（初めの行）
function getHeaderRowNum(sheet) {
  let frozen = sheet.getFrozenRows()
  if ( frozen == 0 ){
    return 1
  } else {
    return frozen
  }
}
// データの左端カラムの位置を取得。
function getIdColomnNum(headerRow){
  let i = 0
  for (const header of headerRow){
    if ( header === ""){
      i++;
      continue;
    } else {
      break;
    }
  }
  return i + 1;
}
//カラム名からカラムidの表を生成する (col_hash)
function createColIndex(sheet, headerRowNum, idColomnNum){
  let currentColomn = idColomnNum
  let result = {}
  let currentRange = sheet.getRange(headerRowNum,currentColomn)
  while(currentRange.getValue() != ""){
    result[currentRange.getValue()] = currentColomn
    currentColomn+=1;
    currentRange = sheet.getRange(headerRowNum, currentColomn)
  }
  return result;
}
function currentRowFromId(sheet, headerRowNum, idColomnNum, id){
  let last_row = sheet.getLastRow();
  let id_datas = sheet.getRange(headerRowNum, idColomnNum, last_row - headerRowNum + 1).getDisplayValues();
  for(let row_num = 1; row_num < id_datas.length; row_num++){
    if(id_datas[row_num][0] === id){
      return headerRowNum + row_num;
    }
  }
  return last_row + 1;
}

//replaceメソッド
function modifyRow(sheet, row, colIndex, currentRow){
  for (const colName of Object.keys(row)){
    sheet.getRange(currentRow, colIndex[colName]).setValue(row[colName])
  }
}
function replaceRows(sheet, headerRowNum, idColomnNum, jsonMessage){
  let colIndex = createColIndex(sheet, headerRowNum, idColomnNum);
  for (const row of jsonMessage.rows){
    let id_hash = Object.entries(row)[0]
    let currentRow = currentRowFromId(sheet, headerRowNum, idColomnNum, id_hash[1]);
    modifyRow(sheet, row, colIndex, currentRow);
  }
}

// insertメソッド
function insertRows(sheet, headerRow, jsonMessage){
  for (let row of jsonMessage.rows){
    let rowArray = []
    for (let colName of headerRow){
      rowArray.push(row[colName])
    }
    //作成した空配列とデータの配列を結合し末尾に挿入
    sheet.appendRow(rowArray);
  }
}

// deleteメソッド
function deleteRows(sheet, headerRowNum, idColomnNum, jsonMessage){
  for (const row of jsonMessage.rows){
    let id_hash = Object.entries(row)[0]
    let currentRow = currentRowFromId(sheet, headerRowNum, idColomnNum, id_hash[1]);
    sheet.deleteRows(currentRow);
  }
}


/////////////////////////////////
/////////////////////////////////

//5分おきにこの関数を実行するように時限トリガーを設定
const retainGlitch = () => {
  //あらかじめ設定しておいたプロパティを呼び出す
  const p = PropertiesService.getScriptProperties().getProperties();
  const glitchURL = p.glitchURL;

  const data = {}
  const headers = { 'Content-Type': 'application/json; charset=UTF-8' }
  const params = {
    method: 'post',
    payload: JSON.stringify(data),
    headers: headers,
    muteHttpExceptions: true
  }
  //特に中身のないデータをGlitchへPOST
  const response = UrlFetchApp.fetch(glitchURL, params);
  console.log(response);
}


///////////////////////////
///初期設定
function initialSettings(p){
  const functionName = 'retainGlitch';
  const triggers = ScriptApp.getProjectTriggers();
  console.log(triggers);

  // Check if trigger already exists using map function
  var triggerExists = triggers.map(function(trigger) {
    return (trigger.getHandlerFunction() === functionName);
  }).indexOf(true) !== -1;

  // If trigger does not exist, create new trigger
  if (!triggerExists) {
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .everyMinutes(5)
      .create();
      Browser.msgBox("Glitch常時起動のためのトリガーを設定しました");
  }

  if (!p.folderId) {
    p.folderId = makeFolder();
  }
  if (!p.glitchURL) {
    p.glitchURL = saveScriptProperty("常時起動するGlitchのURLを入力してください","glitchURL");
  }
  if (!p.email) {
    p.email = saveScriptProperty("CSVデータを送信するメールアドレスを入力してください","email");
  }
}

//CSVの保存フォルダの作成
function makeFolder(){
  //スプレッドシートのIDからフォルダのIDを取得、保存用のフォルダを作成
  const currentFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
  const newFolder = currentFolder.createFolder("CSV-LOG");
  const newFolderId = newFolder.getId();
  //スクリプトプロパティに保存フォルダIDを保存
  PropertiesService.getScriptProperties().setProperty('folderId', newFolderId);
  return newFolderId;
}

//プロパティの登録・変更
function saveScriptProperty(prompt,property) {
  const newProperty = Browser.inputBox(prompt);
  if(newProperty === "cancel"){
    return;
  }
  PropertiesService.getScriptProperties().setProperty(property, newProperty);
  return newProperty;
}

//ThirdWEBへアップロードするCSVファイルの作成
function issueTokenCSV(){
  //スクリプトプロパティに保存されているデータをまとめて取得
  //p = {folderid,glitchURL,email}
  const p = PropertiesService.getScriptProperties().getProperties();
  initialSettings(p); //初期設定
  
  let spsheet = SpreadsheetApp.getActiveSpreadsheet()
  const issueTokenSheet = spsheet.getSheetByName("issue-token");
  const issueLogSheet = spsheet.getSheetByName("issue-log");
  const issuerTableDatas = spsheet.getSheetByName("issuer-table").getDataRange().getValues();
  const addlessDatas = spsheet.getSheetByName("address-table").getDataRange().getValues();
  const issueLogDatas = issueLogSheet.getDataRange().getValues();
  let sheetDatas = issueTokenSheet.getDataRange().getValues();
  
  //検索しやすいように発行元シートのデータをオブジェクトの配列に変換
  const issueDatas = array2Object(sheetDatas);
  //ウォレットアドレスのテーブルと発行者テーブルをオブジェクトの配列に変換
  const objAddlessDatas = array2Object(addlessDatas);
  const objIssuerDatas = array2Object(issuerTableDatas);

  const userWallets = {};
  objAddlessDatas.forEach(user => userWallets[user.userID] = user.walletAddress);

  //発行者IDが登録されているもののみユーザーごとにトークンを加算
  const tokens = {};  
  issueDatas
    .filter(message => objIssuerDatas.some(issuer => issuer.issuerID === message.issuerID))
    .forEach(message => {
      const userID = message.userID;
      if (tokens[userID] === undefined) {
        tokens[userID] = 0;
      }
      tokens[userID] += message.issue;
    });

  const tokenList = Object.keys(tokens).map(userID => ({
    walletAddress: userWallets[userID],
    token: tokens[userID]
  }));
  
  console.log(tokenList);
  //出力するCSVに見出しを設定、最後は改行
  let csv = 'walletAddless,token\r\n';
  tokenList.forEach(token => {
    csv += `${token.walletAddress},${token.token}\r\n`;
  });
  //CSVの長さで判定
  if(csv.length>22){//見出しのみの長さが21
    sendCsvToMail(csv,p);
    Browser.msgBox("申請されていたトークンを認証しました");
  }else{
    //申請されているデータがないなら終了
    Browser.msgBox("トークンが申請されていません");
    return;
  }
  //issue-tokenシートからissue-logシートへ移動
  issueLogSheet.getRange(issueLogDatas.length+1,1,sheetDatas.length,sheetDatas[0].length).setValues(sheetDatas);
  //issue-tokenシートをクリア
  issueTokenSheet.getRange(2,1,sheetDatas.length+1,sheetDatas[0].length).clearContent();
}

//指定のGoogleドライブフォルダへ保存
function saveToDrive(csv,fileName,p) {
  var file = DriveApp.createFile(fileName, csv, MimeType.CSV);
  let folder;
  try{
    folder = DriveApp.getFolderById(p.folderId);
    }catch(e){
      console.log("フォルダ作成")
      folder = DriveApp.getFolderById(makeFolder());
    }
  folder.addFile(file);
  var fileId = file.getId();
  return fileId;
}

//CSVファイルをメール送信
function sendCsvToMail(csv,p) {
  var outputdate = new Date();　　//CSVファイルの作成日を今日の日付で取得
  outputdate = Utilities.formatDate(outputdate,"JST","yyyy/MM/dd");
  const fileName = `${outputdate}.csv` 

  // CSVファイルをGoogle Driveに保存
  saveToDrive(csv,fileName,p);

  //メール添付用にblob作成
  var blob = Utilities.newBlob("", 'text/comma-separated-values', fileName);
  blob.setDataFromString(csv, "utf-8");
  var options = {attachments:[blob]};
  
  // メールを送信
  let to = p.email;
  let subject = `${outputdate}作成のトークン配布リスト`;
  let body = "CSVファイルを添付して送信します。";
  MailApp.sendEmail(to, subject, body, options);;
}


/////////////////////////////
//二次元配列を連想配列に変換する
//
function array2Object(valuesArray) {
  //Key　と　値を　二次元配列上で分離します。
  let keys = valuesArray[0];
  valuesArray.shift();
  //二次元配列を連想配列に変換
  let array = valuesArray.map(function(values) {
    let hash = {};
    values.map(function(value, index) {
      hash[keys[index]] = value
    })
    return hash;
  })
  return array;
}

////////////////////////
//連想配列を二次元配列に変換する
//
function object2Array(objectdata){
  //連想配列を見出しを含めた二次元配列に変換する
  let keys = Object.keys(objectdata[0]);
  let values = objectdata.map(data => keys.map(key=>data[key]))
  let arrayList = [keys];
  arrayList = arrayList.concat(values);
  return arrayList;
}