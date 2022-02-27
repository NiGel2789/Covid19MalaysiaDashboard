function importDataFromMOH() {
  var ss = SpreadsheetApp.openById("1YoK969RWaTBLCTN5ViMm9dXV8zJTEiiWC3E8SjHiNts");
  SpreadsheetApp.setActiveSpreadsheet(ss);
  
  var date = new Date();
  date.setDate(date.getDate()-1);
  var todaysDate = Utilities.formatDate(date, "GMT+8", "MM/dd/yyyy");

  var newDayData = ss.insertSheet();
  newDayData.setName(todaysDate);

  //var firstCell = newDayData.getRange("A1"); // Please provide the row and column of your cell here 
  var mohLink = 'https://raw.githubusercontent.com/MoH-Malaysia/covid19-public/main/epidemic/clusters.csv';
  var contents = Utilities.parseCsv(UrlFetchApp.fetch(mohLink));
  newDayData.getRange(1, 1, contents.length, contents[0].length).setValues(contents);
  //firstCell.setFormula("=ImportData(" + mohLink + ")");

  console.log("Created sheet for " + todaysDate);
  
  var sheets = ss.getSheets();
  var noOfDaysRecorded = sheets.length;
  var names = new Array();
  for (var i = 1 ; i < sheets.length ; i++) names.push([sheets[i].getName()]);

    if (noOfDaysRecorded > 8)
  {
      var secondDate = new Date();
      secondDate.setDate(date.getDate()-8);
      secondDate = Utilities.formatDate(secondDate, "GMT+8", "MM/dd/yyyy");
      var oldSheet = ss.getSheetByName(secondDate);
      ss.deleteSheet(oldSheet);
      console.log("Deleted data for " + secondDate);
  }

  console.log("Waiting for delete to complete processing...");
  SpreadsheetApp.flush();

  var procSheet = ss.getSheetByName("Processed Data");
  ss.setActiveSheet(procSheet);
  firstCell = ss.getRange('A1');
  var select = '"select Col1, Col3, Col6, Col7, Col8, Col9, Col10, Col17 where Col7 = ' + "'active'";
  var query = "=QUERY({";
  for (var j = 0; j < names.length-1; j++)
  {
    query = query + "'" + names[j] + "'!A:Q";
    if (j < names.length-2)
    {
      query = query + ";";
    }
  }
  query = query + "},(" + select + '"))';
  // Example: =QUERY({'02/20/2022'!A:Q;'02/21/2022'!A:Q},("select Col1, Col3, Col6, Col7, Col8, Col9, Col10, Col17 where Col7 = 'active'"))

  firstCell.setFormula(query);

  console.log("Waiting for processing...");
  SpreadsheetApp.flush();

  var lastRow = newDayData.getLastRow();
  var range = newDayData.getRange('Q1');
  newDayData.setActiveRange(range);
  range.setValue("Date");

  range = newDayData.getRange('Q2:Q' + lastRow);
  newDayData.setActiveRange(range);
  range.setValue(todaysDate);
}

function main() {
  importDataFromMOH();
}
