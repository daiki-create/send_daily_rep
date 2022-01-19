function update() {
  var response = Browser.msgBox("更新しますか？",Browser.Buttons.YES_NO);
  if(response==='yes'){
    var sheet=get_sheet();

    const date1=new Date();
  　month=date1.getMonth()+1;
    day=date1.getDate();
    date=month.toString()+'/'+day.toString();
    sheet.getRange(7,4).setValue(date);

    empty_3();
  }
}
