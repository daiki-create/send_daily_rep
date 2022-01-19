function empty_3() {

//16:30退勤用スクリプト

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row_ingenuity=get_title_row("工夫したところ",6);
  const row_comment=get_title_row("コメント",6);
  const row_trouble=get_title_row("困っていること",6);
  const row_report=get_title_row("勤務状況報告",10);


//【タイムテーブル】ごっそり消す
  const range_timetable_delete = sheet.getRange(11,2,get_title_row('本日の作業',2)-3-11+1,3);
  range_timetable_delete.deleteCells(SpreadsheetApp.Dimension.ROWS);

//新しいタイムテーブルの中身
  const array_new_timetable = 
  [
    ['9:30','朝礼',''],
    ['10:00','',''],
    ['12:00','昼休み',''],
    ['13:00','',''],
    ['16:30','終業',''],

  ];

//範囲指定して
  const range_timetable_insert = sheet.getRange("B11:D15");
//セル挿入して中身いれる
  range_timetable_insert.insertCells(SpreadsheetApp.Dimension.ROWS);
  range_timetable_insert.setValues(array_new_timetable);

  //プルダウン追加
  const values_timetable = ['追加','削除'];
  const rule_timetable = SpreadsheetApp.newDataValidation().requireValueInList(values_timetable).build();
  const range_input_pulldown_timetable = sheet.getRange("D11:D15");
  range_input_pulldown_timetable.setDataValidation(rule_timetable);

  //枠線更新
  range_timetable_insert.setBorder(true, true, true, true, false, false);

//【タイムテーブル】終わり


//【本日の作業】はじまり

//【本日の作業】で消すとこの範囲指定と消す操作
  const range_today_work = sheet.getRange(get_title_row('本日の作業',2)+1,2,get_title_row('次回の作業',2)-get_title_row('本日の作業',2)-2,3);
  range_today_work.deleteCells(SpreadsheetApp.Dimension.ROWS);


//【次回の作業】はじまり

//例によって範囲指定と消す操作
  const range_next_work_delete = sheet.getRange(get_title_row('次回の作業',2)+1,2,10,3);
  range_next_work_delete.deleteCells(SpreadsheetApp.Dimension.ROWS);

  //【次回の作業】更新後の中身用意
  const range_next_work_input = sheet.getRange(get_title_row('次回の作業',2)+1,2,1,3);
  const value_new_next_work = [ ['','',''] ];
  //ぶちこむ
  range_next_work_input.setValues(value_new_next_work);

  //プルダウンの用意と挿入
  const range_input_pulldown_nextwork = sheet.getRange(get_title_row('次回の作業',2)+1,4);
  const values_nextwork = ['追加','削除'];
  const rule_nextwork = SpreadsheetApp.newDataValidation().requireValueInList(values_nextwork).build();
  range_input_pulldown_nextwork.setDataValidation(rule_nextwork);

  //枠線更新
  const range_border_next_work = sheet.getRange(get_title_row('次回の作業',2),2,2,3);
  range_border_next_work.setBorder(true, true, true, true, false, false);


  //「工夫したところ」をクリア
  sheet.getRange(row_ingenuity+1,6,1,3).setValue("");
  //「コメント」をクリア
  sheet.getRange(row_comment+1,6,1,3).setValue("");
  //「困っていること」をクリア
  sheet.getRange(row_trouble+1,6,1,3).setValue("");
  //勤務状況報告をクリア
  sheet.getRange(row_report+2,11,3,1).setValue("");
  sheet.getRange(row_report+6,11).setValue("");

}

