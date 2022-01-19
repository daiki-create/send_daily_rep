function send() {
  var response = Browser.msgBox("日報と最重要タスクを送信しますか？",Browser.Buttons.YES_NO);
  if(response==='yes'){
    var sheet=get_sheet();

    //確認モーダル（送信、日報のみ送信、最重要タスクのみ送信、キャンセル）を表示

    //各ブロックの開始行を取得
    var row_work_today=get_title_row('本日の作業',2);
    var row_ingenuity=get_title_row('工夫したところ',6);
    var row_work_next=get_title_row('次回の作業',2);
    var row_comment=get_title_row('コメント',6);
    var row_trouble=get_title_row('困っていること',6);
    var row_report=get_title_row('勤務状況報告',10);

    var last_row=sheet.getLastRow();

    //シートの各ブロックを配列に格納  
    //宛先を変数に格納
    var to=sheet.getRange(3,3,2,3).getValue();
    var to2=sheet.getRange(3,6,2,3).getValue();
    Logger.log(to);

    //サブジェクトを配列に格納
    var subject=sheet.getRange(7,3).getValue()+sheet.getRange(7,4).getValue()+' '+sheet.getRange(7,5,1,2).getValue();
    Logger.log(subject)

    //時間割を二次元配列に格納
    var timetable=sheet.getRange(11,2,row_work_today-13,2).getValues();
    Logger.log(timetable);

    //本日の作業を二次元配列に格納
    var work_today=sheet.getRange(row_work_today+1,2,row_work_next-row_work_today-2,2).getValues();
    Logger.log(work_today);

    //工夫したところを変数に格納
    var ingenuity=sheet.getRange(row_ingenuity+1,6).getValue();
    Logger.log(ingenuity);

    //次回の作業二次元配列に格納
    var work_next=sheet.getRange(row_work_next+1,2,last_row-row_work_next,2).getValues();
    Logger.log(work_next);

    //コメントを変数に格納
    var comment=sheet.getRange(row_comment+1,6).getValue();
    Logger.log(comment);

    //困っていることを変数に格納
    var trouble=sheet.getRange(row_trouble+1,6).getValue();
    Logger.log(trouble);

    // 最重要タスクを二次元配列に格納
    var report=sheet.getRange(row_report+2,11,3,1).getValues().
                concat(
                  sheet.getRange(row_report+6,11).getValue()
                );
    Logger.log(report);

    //サブジェクトを文字列にする
    var subject_text=subject;
    Logger.log(subject_text);

    //メールのテキストを作成
    var mail_text='';

    //時間割を追記
    for(i=0;i<timetable.length;i++){
      mail_text+=timetable[i][0]+'　'+timetable[i][1]+'\n';
    }

    //本日の作業を追記
    mail_text+='\n';
    mail_text+='【本日の作業】'+'\n';
    for(i=0;i<work_today.length;i++){
      if(work_today[i][0]!=''){
        mail_text+=work_today[i][0]+'\n　'+work_today[i][1]+'\n';
      }
      else{
        mail_text+='　'+work_today[i][1]+'\n';
      }
    }

    //工夫したところを追記
    mail_text+='\n';
    mail_text+='【工夫したところ】'+'\n';
    mail_text+=ingenuity+'\n';

    //次回の作業を追記
    mail_text+='\n';
    mail_text+='【次回の作業】'+'\n';
    for(i=0;i<work_next.length;i++){
      if(work_next[i][0]!=''){
        mail_text+=work_next[i][0]+'\n　'+work_next[i][1]+'\n';
      }
      else{
        mail_text+='　'+work_next[i][1]+'\n';
      }
    }
    //コメントを追記
    mail_text+='\n';
    mail_text+='【コメント】'+'\n';
    mail_text+=comment+'\n';

    //困っていることを追記
    mail_text+='\n';
    mail_text+='【困っていること】'+'\n';
    mail_text+=trouble+'\n';

    Logger.log(mail_text);

    // 最重要タスクのテキストを作成
    var task_text='■本日の最重要タスク■　'+report[0]+'~'+report[1]+'\n'+
      report[2]+'\n'+
      '■次回の最重要タスク■　'+'\n'+
      report[3];
      

    //送信処理
    var sheet_name=sheet.getSheetName();

    //テスト用
    user_url={
      '真・小野寺 卍':'https://chat.googleapis.com/v1/spaces/AAAAofQIqxI/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=vZGYUNMAL7hoKvvlU01ZlT5fNa5Ey3PF4sC-l06W2dU%3D',
      '山崎2':'https://chat.googleapis.com/v1/spaces/AAAAofQIqxI/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=JaHlR2zdBSKP-XfIZVBMj9p6Ks8TtyWBA6i4sti4SA0%3D'
    }
    var url=user_url[sheet_name]
    var thread='spaces/AAAAofQIqxI/threads/ACCDA8XzWUs';

    //本番用
    // user_url={
    //   '徳長':'https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=K8ODxoPnWisAvlTrtyQYL6BK-mElXgq7h2C5sYHEd24%3D',
    //   '永井':'https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=3PsrXA_qg2KaRrdrDXnBiO0hg1SjQcCftkXqyDZxcVc%3D',
    //   '小野寺':'https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=gXoCANcy0_8lnvqLAQ-nQYgR6yEWccR6pBkZC4KWQr0%3D',
    //   '本田':'https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=GbDVAX3cET7qOxO8DkiBKKkcyRXrlQmC7VNC8btSnMY%3D',
    //   '山崎':'https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=z3OWdltpHSJJ9_2OxA1Ax5MbzyfAdjjUYdrhj4pspHE%3D',
    //   '今村':'https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=FPRiv2c-lEhRIP-syXCvt6qP19WbfhmgzrdyEceQIsc%3D',
    //   '大塚':'https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=7bW0ua08gsvKlD936063GCrtPi5M-GfghxO4qmY5Nzg%3D',
    //   '鈴木':'https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=GQDvRI4EHQVF8AeZPPU8ARldO8uaO9dK3q11Vm7wcgs%3D',
    //   '福原':'https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=xnF2lcb9--Yg65-csJfHR7P5HlLLrsnRB6cvRES41uE%3D',
    //   '加藤':'https://chat.googleapis.com/v1/spaces/AAAAzBbo-48/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=kfiEIwDexFYSjFbLBpcRsKoMQpesrA37PA5Jzds3LfI%3D',
    // }
    // var url=user_url[sheet_name]
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

    if(to==''){
      MailApp.sendEmail({to:to2, subject:subject_text, name:sheet.getRange(7,5,1,2).getValue(), body:mail_text});
      var response = UrlFetchApp.fetch(url, options);
      Logger.log(JSON.parse(response.getContentText()))
    }
    else{
      MailApp.sendEmail({to:to+','+to2, subject:subject_text, name:sheet.getRange(7,5,1,2).getValue(),body:mail_text});
      var response = UrlFetchApp.fetch(url, options);
      Logger.log(JSON.parse(response.getContentText()))
    }
  }

}
