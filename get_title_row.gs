function get_title_row(title,col) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();

  for(var i=1;i<=lastRow;i++){
    if(sheet.getRange(i,col).getValue()===title){
        var title_row=i;
        return title_row;
    }
  }
}
function a(){
  const titleRow = get_title_row("本日の作業",2);
Logger.log(titleRow);
}
