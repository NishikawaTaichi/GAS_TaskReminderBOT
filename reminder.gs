function main() {
  // 全タスクを取得
  const taskList = getTaskList();

  // 「タスク期限が1日前」で「ステータスが未完了」のタスクのみをフィルタリング
  const today = new Date();
  const imcompleteTaskList = taskList.filter((task) => {
    let sheetDay = new Date(task.deadline);
    return task.status === "未完了" && isDeadline(today, sheetDay);
  });

  // slackに送信する全メッセージをリスト形式で作成
  const messageList = createMessageList(imcompleteTaskList);

  // slackにmessageを送信
  for(let i = 0; i < messageList.length; i++){
    sendSlack(messageList[i]);
  }
}


/*
 * タスクが1日前か判定するメソッド 
 */
function isDeadline (today, sheetDay) {
  const isMatchMonth = today.getMonth() == sheetDay.getMonth();　// 同じ月かどうか 
  const isOneDayBefore = (sheetDay.getDate()- today.getDate()) == 1;  // 1日前かどうか
  return isOneDayBefore && isMatchMonth;
}

/*
 * シート内の全タスクを取得するメソッド
 */
function getTaskList () {
  const activeSheet = SpreadsheetApp.getActiveSheet(); // アクティブシート
  if(activeSheet.getName() != "タスク管理"){
    return;
  }
  let taskList = [];
  const lastRow = activeSheet.getLastRow();
  for(let i = 2; i <= lastRow; i++) {
    if(activeSheet.getRange(i, 1).getValue() == null) continue;
    // タスク
    let task = {
      taskNo: activeSheet.getRange(i, 1).getValue(),
      client: activeSheet.getRange(i, 2).getValue(),
      content: activeSheet.getRange(i, 3).getValue(),
      host: activeSheet.getRange(i, 4).getValue(),
      deadline: activeSheet.getRange(i, 5).getValue(),
      status: activeSheet.getRange(i, 6).getValue()
    }
    taskList.push(task);
  }
  return taskList;
}

/*
 * 通知メッセージをリスト形式で作成
 */
function createMessageList (imcompleteTaskList) {
  // 環境変数SHEET_URLの参照
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetURL = scriptProperties.getProperty('SHEET_URL');

  let messageList = [];
  imcompleteTaskList.forEach((task) => {
    let message = "「タスク管理」No." + task.taskNo + " のタスクが未完了です\n" +
    "担当者の"+ task.host +"さんに連絡してください\n" + sheetURL + (task.taskNo + 1);
    messageList.push(message);
  })
  return messageList;
} 

/*
 * slackに通知するメソッド
 */
function sendSlack(text){
  // 環境変数WEBHOOK_URLの参照
  const scriptProperties = PropertiesService.getScriptProperties();
  const webhookURL = scriptProperties.getProperty('WEBHOOK_URL');

  const jsonData =
      {
        "channel": "#プログラミング",   // 通知したいチャンネル
        "username": "タスクリマインドBOT", // Botの表示名
        "icon_emoji": "ロボット", // Botのアイコン,
        "unfurl_links": true, // 送信したリンクを展開する
        "text": text,
      };

  const payload = JSON.stringify(jsonData);
  const options =
      {
        "method" : "post",
        "contentType" : "application/json",
        "payload" : payload,
      };
  
  UrlFetchApp.fetch(webhookURL, options);　// 送信
}
