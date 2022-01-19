function empty2(){
  var sheet=get_sheet();

  var lastRow=sheet.getLastRow();
  var row_ingenuity=get_title_row("工夫したところ",6);
  var row_comment=get_title_row("コメント",6);
  var row_trouble=get_title_row("困っていること",6);
  var row_report=get_title_row("勤務状況報告",10);

  //「本日の作業」のある行と「次回の作業」のある行の間の行のC列を削除
  var row_work_today=get_title_row("本日の作業",2);
  var row_work_next=get_title_row("次回の作業",2);
  //「本日の作業」最終行のエラー解消
  if(row_work_next-2==row_work_today){
    row_work_next=row_work_today+3
  }
  sheet.getRange(row_work_today+1,3,row_work_next-row_work_today-2,1).setValue("");

  //「本日の作業」のある行と「次回の作業」のある行の間の行のB列が空白の行を削除
  for(i=row_work_next-2;i>=row_work_today+1;i--){
    if(sheet.getRange(i,2).getValue()==="" && sheet.getRange(i,2).getRow()!=row_work_today+1){
      sheet.getRange(i,2,1,3).deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
  }
  
  var row_work_next=get_title_row("次回の作業",2);
  var row_last_work_next=sheet.getRange(row_work_next,2,20,3).getLastRow();
  //「次回の作業」最終行のエラー解消
  if(row_last_work_next==row_work_next){
    row_last_work_next=row_work_next+1
  }
  //「次回の作業」のある行と「次回の作業」の最終行の間の行のC列を削除
  sheet.getRange(row_work_next+1,3,row_last_work_next-row_work_next,1).setValue("");

  //「次回の作業」のある行と最終行の間の行のB列が空白の行を削除
  for(i=row_last_work_next;i>=row_work_next+1;i--){
    if(sheet.getRange(i,2).getValue()==="" && sheet.getRange(i,2).getRow()!=row_work_next+1){
      sheet.getRange(i,2,1,3).deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
  }

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


