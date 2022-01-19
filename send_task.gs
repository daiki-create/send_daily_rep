function send_task() {
  var response = Browser.msgBox("最重要タスクのみ送信しますか？",Browser.Buttons.YES_NO);
  if(response==='yes'){
    var sheet=get_sheet();

    //確認モーダル（送信、日報のみ送信、最重要タスクのみ送信、キャンセル）を表示

    //各ブロックの開始行を取得
    var row_report=get_title_row('勤務状況報告',10);

    var last_row=sheet.getLastRow();

    // 最重要タスクを二次元配列に格納
    var report=sheet.getRange(row_report+2,11,3,1).getValues().
                concat(
                  sheet.getRange(row_report+6,11).getValue()
                );
    Logger.log(report);

    // 最重要タスクのテキストを作成
    var task_text='■本日の最重要タスク■　'+report[0]+'~'+report[1]+'\n'+
      report[2]+'\n'+
      '■次回の最重要タスク■　'+'\n'+
      report[3];
      

    //送信処理

    //テスト用
    user_url={
      '真・小野寺 卍':'https://chat.googleapis.com/v1/spaces/AAAAofQIqxI/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=vZGYUNMAL7hoKvvlU01ZlT5fNa5Ey3PF4sC-l06W2dU%3D',
      '山崎2':'https://chat.googleapis.com/v1/spaces/AAAAofQIqxI/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=JaHlR2zdBSKP-XfIZVBMj9p6Ks8TtyWBA6i4sti4SA0%3D'
    }

    var sheet_name=sheet.getSheetName();
    var url=user_url[sheet_name]
    var thread='spaces/AAAAofQIqxI/threads/ACCDA8XzWUs';

    //本番用
    // var url='https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=k1e998r7jleOXxQXjW7zOFvFXDdRlnVhQ4cZxSf1LHw%3D';
    // var thread = "spaces/AAAAzBbo-48/threads/eu_6pNpCK28"

    var payload = {
      "text" : task_text,
    "thread": {
      "name": thread
    }
      }
    var json = JSON.stringify(payload);
    var options = {
      "method" : "POST",
      "contentType" : 'application/json; charset=utf-8',
      "payload" : json
    }

    var response = UrlFetchApp.fetch(url, options);
    Logger.log(JSON.parse(response.getContentText()))
  }

}
