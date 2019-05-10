function onEdit(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName("Form Responses 1");
  var range = e.range;
  var column = range.getColumn();
  var row = range.getRow();
  
  // Column W (Approved) = 23
  // Row 3
  // Colors from: https://flatuicolors.com/palette/us
  
  if(e.value == "Clear Format")
  {
    var rangeToMove = scheduleSheet.getRange(row, 3, 1,21);
    rangeToMove.setValue("");
    rangeToMove.clearFormat();
  }
  else if(column == 23 && e.value == 'Yes' || e.value == 'YES' || e.value == 'yes')
  {
    var rangeToMove = scheduleSheet.getRange(row, 3, 1,21);
    rangeToMove.setBackgroundRGB(0, 234, 148); // Light Green
    
  }

  else if(column == 23 && e.value == 'No' || e.value == 'NO' || e.value == 'no')
  {
    var rangeToMove = scheduleSheet.getRange(row, 3, 1,21);
    rangeToMove.setBackgroundRGB(255, 118, 117); // Light Red
  }
  else if(column == 23 && e.value == '' || (e.value) != "Yes" || (e.value) != "No" || (e.value) != "YES" || (e.value) != "NO") 
  {
    var rangeToMove = scheduleSheet.getRange(row, 3, 1,21);
    rangeToMove.setBackgroundRGB(255, 234, 167); // Light Yellow
  } 
}

function refreshButton() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var scheduleSheet = ss.getSheetByName("Form Responses 1");
  var rangeData = sheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow();
  var startRow = 5;
  var searchRange = sheet.getRange(5, 23, lastRow -1, lastColumn-1);
  var rangeValues = searchRange.getValues();
  
  // Only iterate through rows
  for(var rows = 5; rows <= lastRow; rows++)
  {
    // W (Approved) column = 23
    var rangeToCheck = sheet.getRange(rows, 23, 1, 1); // 1 columns starting with column 23, so W
    
    if(rangeToCheck.getValue() === "Yes")
    {
      var rangeToMove = scheduleSheet.getRange(rows, 1, 1,23);
      rangeToMove.setBackgroundRGB(0, 234, 148); // Light Green
    }
    else if(rangeToCheck.getValue() === "No")
    {
      var rangeToMove = scheduleSheet.getRange(rows, 1, 1,23);
      rangeToMove.setBackgroundRGB(255, 118, 117); // Light Green
    }
    else
    {
      var rangeToMove = scheduleSheet.getRange(rows, 1, 1,23);
      rangeToMove.setBackgroundRGB(255, 234, 167); // Light Green
    }
    
  }
}