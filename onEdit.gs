//セルの編集時にイベント発生
function onEdit(e){

  //2021/6/18 追加は全部完了　削除実装中
  //2021/6/21 【タイムテーブル】削除実装完了
  //2021/6/21　枠線更新の変数名がおかしい？
  //2021/6/21　枠線更新の変数名修正完了
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();

  const row_selecting = cell.getRow();
  //どこ選択しているかの確認
  //Logger.log('row_selecting:'+row_selecting);

  //【タイムテーブル】関連行
  let timetable_start = 11;
  let timetable_end = get_title_row('本日の作業',2)-3;
  //Logger.log('timetable_start:'+ timetable_start);
  //Logger.log('timetable_end:'+ timetable_end);

  //【本日の作業】関連行
  let today_work_start = get_title_row('本日の作業',2);
  let today_work_end = get_title_row('次回の作業',2)-2;
  //Logger.log('today_work_start:'+today_work_start);
  //Logger.log('today_work_end:'+today_work_end);

  //【次回の作業】関連行
  let next_work_start = get_title_row('次回の作業',2);
  let next_work_end = next_work_start + 3;
  //Logger.log('next_work_start:'+next_work_start);
  //Logger.log('next_work_end:'+next_work_end);



  //-----【タイムテーブル】D列での操作-----//
  if(timetable_start <= row_selecting && row_selecting <= timetable_end){   

    if(cell.getValue()==='追加' && cell.getColumn()===4){
      //動作チェック
      Logger.log('【タイムテーブル】D列操作ver.内に入った');

      //行の追加
      const range_add_timetable = sheet.getRange(cell.getRow()+1,2,1,3);
      range_add_timetable.insertCells(SpreadsheetApp.Dimension.ROWS);
      
      //Logger.log('【タイムテーブル】で行追加した');

      //タイムテーブル最終行での操作時には枠線の更新を行う
      if(cell.getRow() == timetable_end){
        
        //【タイムテーブル】最終行の更新と枠線の更新

        //変数再宣言
        timetable_end = get_title_row('本日の作業',2)-3;

        //枠線の更新
        const border_range_timetable = sheet.getRange(timetable_start,2,timetable_end-timetable_start+1,3);
        border_range_timetable.setBorder(true, true, true, true, false, false);

        //プルダウン挿入

        const values = ['追加','削除'];
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
        const inputCell = sheet.getRange(cell.getRow()+1,4);
        inputCell.setDataValidation(rule);
 
        
      }

      cell.clearContent();

    }

    if(cell.getValue()==='削除' && cell.getColumn()===4){

      Logger.log('【タイムテーブル】D列削除入った');

      //行削除前にgetValue
      const range_delete_timetable = sheet.getRange(row_selecting,3,1,1)
      const cell_value = range_delete_timetable.getValue();

      //値取得の確認
      Logger.log('cell_value:'+cell_value);

      //【タイムテーブル】の範囲を選択
      const range_timetable = sheet.getRange(11,3,get_title_row(('本日の作業'),2)-2-11,1);
     //【タイムテーブル】の内容を配列に格納
      const timetable_container = range_timetable.getValues();

      Logger.log('timetable_container:'+timetable_container);


      //delete_count:消すのが上から何番目なのか把握するためのもの（1番上を0番目とする）
      let delete_count = -1;

      for(i=0; i <= row_selecting-11; i++){
        if(timetable_container[i] == cell_value){

        
          delete_count += 1;

        }
      }
      // count:【タイムテーブル】にcell_valueがいくつあるか。1つなら出力０
      let count = -1;

      for(i=0; i<timetable_container.length; i++){
        //Logger.log(timetable_container[i]);

        if(timetable_container[i] == cell_value){

          count += 1;
        }
      } 

      //ここで選択した行消す（タイムテーブル）
      const range_delete_today_work = sheet.getRange(row_selecting,2,1,3);
      range_delete_today_work.deleteCells(SpreadsheetApp.Dimension.ROWS);

      //【タイムテーブル】の枠線更新
      const border_range_timetable = sheet.getRange(11,2,get_title_row('本日の作業',2)-2-11,3);
      border_range_timetable.setBorder(true, true, true, true, false, false);


      //----------------//
      //【タイムテーブル】先頭のものだった場合
      if(delete_count == 0  ){

        //作業名が１つだった場合は【本日の作業】該当行をただ消すだけ

        if(count == 0){
          const range_delete_today_work = sheet.getRange(get_title_row(cell_value,2)+delete_count,2,1,3);
          //range_delete_today_work.setBackground('red')
          range_delete_today_work.deleteCells(SpreadsheetApp.Dimension.ROWS);
        
          //【本日の作業】罫線更新

          //変数再宣言
          today_work_start = get_title_row('本日の作業',2);
          next_work_start = get_title_row('次回の作業',2);

          //【本日の作業】の枠線更新
          const border_range_today_work = sheet.getRange(today_work_start,2,next_work_start-2-today_work_start+1,3);
          border_range_today_work.setBorder(true, true, true, true, false, false);


        }
        //【タイムテーブル】に作業名が2つ以上かつ該当行が先頭だった時
        else
        {
          //ひとつ下の行に作業名追加してから行を消す
          const ValueSetCell = sheet.getRange(get_title_row(cell_value,2)+1,2);
          ValueSetCell.setValue(cell_value);
          //【本日の作業】で消す場所を指定
          const range_delete_today_work = sheet.getRange(get_title_row(cell_value,2)+delete_count,2,1,3);
          //range_delete_today_work.setBackground('red')
          range_delete_today_work.deleteCells(SpreadsheetApp.Dimension.ROWS);

        }

      }
      //作業名が【タイムテーブル】で2つ以上だった場合で選択行が先頭じゃないとき
      else
      {
        //【本日の作業】で消す場所を指定　delete_count分下を消す
        const range_delete_today_work = sheet.getRange(get_title_row(cell_value,2)+delete_count,2,1,3);
        //range_delete_today_work.setBackground('red')
        range_delete_today_work.deleteCells(SpreadsheetApp.Dimension.ROWS);

        //変数再宣言
        today_work_start = get_title_row('本日の作業',2);
        next_work_start = get_title_row('次回の作業',2);

        //【本日の作業】の枠線更新
        const border_range_today_work = sheet.getRange(today_work_start,2,next_work_start-2-today_work_start+1,3);
        border_range_today_work.setBorder(true, true, true, true, false, false);
      }

      //--------//
    
    }

  }

  //-----【タイムテーブル】C列での操作-----//
  if(timetable_start <= row_selecting && row_selecting <= timetable_end && cell.getColumn()===3){
    
    
    //動作チェック
    Logger.log('【タイムテーブル】C列操作ver.内に入った');

    const cell_value = cell.getValue();
    Logger.log('cell_value: '+cell_value);
      
    //入力値が空白でない場合
    if(cell_value !== '' && cell_value !== '夕礼' && cell_value !== '朝礼' && cell_value !== '昼休み'){
      //【タイムテーブル】の範囲を選択
      const range_timetable = sheet.getRange(timetable_start,3,timetable_end - timetable_start+1,1);
      //【タイムテーブル】の内容を配列に格納
      const array_timetable = range_timetable.getValues();
  
      Logger.log(array_timetable);

      //cell_valueが【タイムテーブル】にいくつあるかチェック
      //便宜上、実際の個数-1している

      let count_timetable = -1;
      for(i=0; i<array_timetable.length; i++){
        //Logger.log(array_timetable[i]);

        if(array_timetable[i] == cell_value){

          count_timetable += 1;
        }
      } 

      //count_timetableのチェック
      Logger.log('count_timetable :'+count_timetable); 

      //【タイムテーブル】に初めて追加された場合（count_timetable=0）、【本日の業務】最終行B列に入力値をセットした状態のセル群を追加する。

      if(count_timetable == 0 ){

        //動作チェック
        Logger.log('(count_timetable==0)内に入った');

        
        const range_add_today_work = sheet.getRange(get_title_row('次回の作業',2)-2+1,2,1,3);
        range_add_today_work.insertCells(SpreadsheetApp.Dimension.ROWS);
        //const cellAddWord = ;

        //変数再宣言
        today_work_start = get_title_row('本日の作業',2);
        today_work_end = get_title_row('次回の作業',2)-2;
        next_work_start = get_title_row('次回の作業',2);

        //罫線更新
        //2021/6/18　getRangeの引数不具合
        //上記修正完了

        Logger.log('today_work_start:'+ today_work_start);
        Logger.log('next_work_start: '+ next_work_start);
        

        const border_range_today_work_from_timetable = sheet.getRange(today_work_start,2,next_work_start - today_work_start-1,3);
        border_range_today_work_from_timetable.setBorder(true, true, true, true, false, false);

        //以下でB列に入力値をセット
        const cell_add_cell_value  = sheet.getRange(get_title_row('次回の作業',2)-2,2);
        //wordSetCell.setBackground("blue");
        cell_add_cell_value.setValue(cell_value); 

        //プルダウン挿入
        const values = ['追加','削除'];
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
        const inputCell = sheet.getRange(get_title_row(cell_value,2),4);
        inputCell.setDataValidation(rule);

      }  

      //count≠0の場合：該当作業最終行に空のセル群を追加する
      else if(count_timetable >= 1)
      {
        
        Logger.log('count_timetable >= 1に入った');
        //【本日の作業】内の該当作業名がある行数+count行名に空のセルを追加
        const range_add_today_work = sheet.getRange(get_title_row(cell_value,2)+count_timetable,2,1,3);
        //range_add_today_work.setBackground("red")
        range_add_today_work.insertCells(SpreadsheetApp.Dimension.ROWS);
        Logger.log('count_timetable >= 1で行追加');

        //罫線更新
        const border_range_today_work = sheet.getRange(today_work_start,2,next_work_start-today_work_start,3);
        border_range_today_work.setBorder(true, true, true, true, false, false);
        Logger.log('count_timetable >= 1で行更新');

        //プルダウン挿入
        const values = ['追加','削除'];
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
        const inputCell = sheet.getRange(get_title_row(cell_value,2)+count_timetable,4);
        inputCell.setDataValidation(rule);
        Logger.log('count_timetable >= 1でプルダウン挿入');

      }
    }
    
  }

  //-----【本日の作業】D列での操作 -----//
  if(today_work_start <= row_selecting && row_selecting <= today_work_end){

    if(cell.getValue()==='追加' && cell.getColumn()===4){

      //動作チェック
      Logger.log('【本日の作業】追加内に入った');

      //行の追加
      const range_add_today_work = sheet.getRange(cell.getRow()+1,2,1,3);
      range_add_today_work.insertCells(SpreadsheetApp.Dimension.ROWS);

      Logger.log('【本日の作業】で行追加した');

      
      //【本日の作業】最終行での操作時には枠線の更新を行う
      if(cell.getRow() == today_work_end){

        //【本日の作業】関連行の再宣言
        today_work_end = get_title_row('次回の作業',2)-2;
        today_work_start = get_title_row('本日の作業',2);
        
        //【本日の作業】最終行の更新と枠線の更新
        
        const border_range_today_work = sheet.getRange(today_work_start,2,today_work_end-today_work_start+1,3);
        border_range_today_work.setBorder(true, true, true, true, false, false);

        //プルダウン挿入
        const values = ['追加','削除'];
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
        const inputCell = sheet.getRange(cell.getRow()+1,4);
        inputCell.setDataValidation(rule);

        
      }

      cell.clearContent();

    }

    if(cell.getValue()==='削除' && cell.getColumn()===4){

      //動作チェック
      Logger.log('【本日の作業】削除内に入った');


      //【本日の業務】最終行なら行削除後に枠線の更新を行う
      if(row_selecting == today_work_end){

         Logger.log('削除【本日の作業】最終行ver.内に入った');

        //削除する場所指定と削除操作
        const rande_delete_today_work = sheet.getRange(row_selecting,2,1,3);
        rande_delete_today_work.deleteCells(SpreadsheetApp.Dimension.ROWS);

        today_work_start = get_title_row('本日の作業',2);
        next_work_start = get_title_row('次回の作業',2);
        next_work_end = next_work_start + 3;


        //枠線の更新
        const border_range_today_work = sheet.getRange(today_work_start,2,next_work_start-today_work_start-1,3);
        border_range_today_work.setBorder(true, true, true, true, false, false);
        //border_range_today_work.setBackground('red');
      }

      else
      {
        const range_delete_today_work = sheet.getRange(row_selecting,2,1,3);
        range_delete_today_work.deleteCells(SpreadsheetApp.Dimension.ROWS);
      }

      

    }
    
  }

  //-----【次回の作業】D列での操作-----//
  if(next_work_start <= row_selecting ){

     
    if(cell.getValue()==='追加' && cell.getColumn()===4){

      //動作チェック
      Logger.log('【次回の作業】追加内に入った'); 

      //行の追加
      const range_add_next_work = sheet.getRange(cell.getRow()+1,2,1,3);
      //range_add_next_work.insertCells(SpreadsheetApp.Dimension.ROWS);

      Logger.log('【次回の作業】で行追加した');

      //【次回の作業】関連行の再宣言
      next_work_end = cell.getRow()+1;
      next_work_start = get_title_row('次回の作業',2);
        
      //【次回の作業】最終行の更新と枠線の更新
      const border_range_timetable_next_work = sheet.getRange(next_work_start,2, next_work_end-next_work_start+1,3);
      border_range_timetable_next_work.setBorder(true, true, true, true, false, false);

      //プルダウン挿入
      const values = ['追加','削除'];
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
      const inputCell = sheet.getRange(cell.getRow()+1,4);
      inputCell.setDataValidation(rule);
      
      cell.clearContent();

    }

    if(cell.getValue()==='削除' && cell.getColumn()===4){

      //削除操作
      const range_delete_next_work = sheet.getRange(row_selecting,2,1,3);
      range_delete_next_work.deleteCells(SpreadsheetApp.Dimension.ROWS);

      next_work_start = get_title_row('次回の作業',2);
      next_work_end = next_work_start + 3;

      //枠線の更新
      const border_range_next_work = sheet.getRange(next_work_start,2,row_selecting-next_work_start,3);
      border_range_next_work.setBorder(true, true, true, true, false, false);

    }
    
  }

}