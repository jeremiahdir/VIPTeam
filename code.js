function startup(list){    // Use once to create the company spreadsheet, then filter will trigger each form submit
    //Creates new sheet
  if (list.length == 0){  // if there's no file by queried name
    var newSheet = SpreadsheetApp.create("newSheet");  // makes new sheet with name 'newsheet'
    var newSheetURL = newSheet.getUrl();  // fetches ^^ url
  }  
  
  else {
    var rawFile = DriveApp.getFilesByName('newSheet');  // if sheet already exists, fetch file iterable
    var actualFile = []
    while (rawFile.hasNext() ){   
      var oldFile = rawFile.next(); 
      actualFile.push(oldFile); }  // takes file out of iterable and puts into empty list 
    
    var oldFile = actualFile[0];  
    var newSheetURL = oldFile.getUrl(); // fetches url from file within list
     
  }
  var rawFile2 = DriveApp.getFilesByName('Responses')
  var actualFile2 = []
  while (rawFile2.hasNext() ){
    var oldFile2 = rawFile2.next();
    actualFile2.push(oldFile2);}
  var oldFile2 = actualFile2[0];
  var dataSheetURL = oldFile2.getUrl();
  
  var rawFile3 = DriveApp.getFilesByName('config')
  var actualFile3 = []
  while (rawFile3.hasNext() ){
    var oldFile3 = rawFile3.next();
    actualFile3.push(oldFile3);}
  var oldFile3 = actualFile3[0];
  var configURL = oldFile3.getUrl();
  return [configURL,dataSheetURL,newSheetURL,]// In case you need to pull the URL
  }


function searchForFile(){ //Checks that the sheet doesn't already exist
  var file = DriveApp.getFilesByName('newSheet');// fetches sheet 
  var files = []
  while (file.hasNext()){ // while there's another entry
    var actualFile = file.next();     
    files.push(actualFile);  // add file name to file list
    
 }length = files.length  
  return files ; // returns either empty list or list with file's name 
}

function clearRange1(sheetURL) { //deletes old entries
  var sheet = SpreadsheetApp.openByUrl(sheetURL); // fetches new sheet
  var range = sheet.getDataRange().clearContent(); // deletes all data in data range
  
}




function filterExcel(sheetURL,datasheetUrl) {
//Filters the original data based on company code and places those rows in new sheet  
  
  //Fetch ss and its length
  var sheet = SpreadsheetApp.openByUrl(datasheetUrl);
  var range = sheet.getDataRange();  
  var lastRow = sheet.getLastRow();
  var response = sheet.get
  var responseRow = 'B' + lastRow.toString()
  var code =  sheet.getRange(responseRow).getValue(); 

  
  //Fetches new sheet
  var newSheet = SpreadsheetApp.openByUrl(sheetURL);
  
  // Fetch values for each row in the Range.
  var data = range.getValues();
  newSheet.appendRow(data[0])
  var totalResponses = 0;

  for (var i in data ) {  // iterates through
    var row = data[i];
    if ( row[1] == code )     // if the code input is the same as the company code, it adds the new row to the existing spreadsheet
    {
      newSheet.appendRow(row);
      totalResponses = totalResponses + 1
      i = i + 1; 
    }
    else
    { 
      i = i + 1; 
    } 

  } 
   return [totalResponses,code] 

}

function filterExcelMonthly(sheetURL,datasheetUrl,configUrl) {
//Filters the original data based on company code and places those rows in new sheet  
  var sheet2 = SpreadsheetApp.openByUrl(sheetURL)
  var range = sheet2.getDataRange().clearContent(); // deletes all data in data range
  //Fetch ss and its length
  var sheet = SpreadsheetApp.openByUrl(datasheetUrl);
  var range = sheet.getDataRange();  
  var lastRow = sheet.getLastRow();
  var doc = DocumentApp.openByUrl(configUrl)
  var body = doc.getBody();
  var rawText = body.getText();
  var text = rawText.split("\n");
  var orgList = text[5];
  var newList = orgList.split(" ");
  
  for (num in newList){  
    var sheet2 = SpreadsheetApp.openByUrl(sheetURL)
    var range2 = sheet2.getDataRange().clearContent();
    var code = newList[num]
    
    var newSheet = SpreadsheetApp.openByUrl(sheetURL);
    var data = range.getValues();
    //logger.log(data + ' Data ' )
    newSheet.appendRow(data[0])
    var totalResponses = 0;

    for (var i in data ) {  // iterates through
      var row = data[i];

      if ( row[1] == code )     // if the code input is the same as the company code, it adds the new row to the existing spreadsheet
      {
        newSheet.appendRow(row);
        totalResponses = totalResponses + 1
        i = i + 1; 
      }
      else
      { 
        i = i + 1; 
      } 

  }
   var tuple = [totalResponses,code] 
   var totalResponse = tuple[0]
   var code2 = tuple[1]
   var lastRowIndex = basicAnalytics(sheetURL,totalResponse); // applies basic analytics to entries and adds
   takeinConfig(lastRowIndex,sheetURL,configUrl)}

}



function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


function takeinConfig(lrIndex,sheetURL,configUrl){
  var doc = DocumentApp.openByUrl(configUrl)
  var body = doc.getBody();   // opens up config file
  
  var rawText = body.getText();   // extracts the text of the doc as a string
  var text = rawText.split("\n");
  
  var rawCCL = text[0];
  var rawAE = text[1];
  var rawSI = text[2];  // gets the raw ques num lists
  var rawAPP = text[3];
  var rawEE = text[4];
  var sheet = SpreadsheetApp.openByUrl(sheetURL)
  var lastLastIndex = sheet.getLastRow();
  var raws = [rawCCL,rawAE,rawSI,rawAPP,rawEE] 
  for (num in raws){   // iterates through each set of questions 
    var list1 = raws[num];   
    var list1 = list1.split(" ");
    var name = list1[0];  // extracts name and then removes it, leaving list of just ques nums
    list1.splice(0,1);
    var cellVals = [];
    var start = lrIndex + 2;  // establishes start value for cell placement
    
    for (num in list1){   // gets the appropriate row/cell place and appends to a list of them
      var value = list1[num]
      var number = Number(value);
      var number = number + start;  
      var string1 = number.toString();
      var excel = 'F' + string1
      cellVals.push(excel); 
    }
    
    var rangeStr = '=AVERAGE(';
    for (num in cellVals){
      val = cellVals[num];     // creates actual excel form that is appended
      var rangeStr = rangeStr + val + ',';}
    var rangeStr = rangeStr.slice(0,-1) + ')'
    
    var sheet = SpreadsheetApp.openByUrl(sheetURL) // appends to form
    sheet.appendRow([name,rangeStr])  
    }
var idkCells = [];
var idkAvgStr = '';
var oppCells = [];  
  for ( var i = lrIndex + 2; i <= lastLastIndex; i++){
    var cellVal = 'F' + i.toString();
    var cellVal1 = 'G' + i.toString();
      oppCells.push(cellVal);
      idkCells.push(cellVal1);
}
 var avgStr = '=AVERAGE(';
 var oppStr = '=AVERAGE(';
  for (num in oppCells){
    cell = oppCells[num];
    cell2 = idkCells[num];
    var avgStr = avgStr + cell2 + ',';
    var oppStr = oppStr + cell + ',';}
  var oppStr = oppStr.slice(0,-1) + ')';
  var avgStr = avgStr.slice(0,-1) + ')';
  
  var lastRow = sheet.getLastRow();
  var oppThreshholdStr = '=(B' + (lastRow+1) + '*1.5)';
  var idkThreshholdStr = '=(B' + (lastRow+2) + '*1.5)';
  sheet.appendRow(['oppAVG' , oppStr]);
  sheet.appendRow(['idkAVG' , avgStr]);
  sheet.appendRow(['OppThreshhold',oppThreshholdStr]);
  sheet.appendRow(['IDKThreshhold',idkThreshholdStr]);
  
}
  


function basicAnalytics(newSheetURL,totalResponses){  // Copies sheets formulas to new sheet where they auto-apply
  var sheet = SpreadsheetApp.openByUrl(newSheetURL); // fetch sheet and set active
  SpreadsheetApp.setActiveSpreadsheet(sheet);
  
  var lastRowIndex = sheet.getLastRow();  // gets last entry with data
  var lastdataindex = lastRowIndex.toString()
  var lastcol = sheet.getLastColumn();
  var strings = "A2:"+columnToLetter(lastcol)+lastdataindex
  var range = sheet.getRange(strings); // extracts data from sheet
  
  var numCols = range.getNumColumns();
  var numRows = range.getNumRows();
  
  var responses = ['']

  sheet.appendRow(responses);
  var range2 = sheet.getDataRange();
      
  //For initial questions
  var rowcount = lastRowIndex;
          for (var i = 2; i <= 10; i++) {
            var withString = 'Question '+(i-1).toString();
            var dataArray = [withString];
            var countArray = ["Response Count"];
            var perArray = ["Response Percent"];
            var totArray = ["Answered Question"];
            rowcount++;
            
            var colname = columnToLetter(i);
            var totresp = '=COUNTA('+ colname+ '1:'+ colname + lastdataindex+')'
            
              for (var j = 1; j <= numRows; j++) {
                var currentValue = range.getCell(j,i).getValue();
                var dup = false
                for( l in dataArray){
                  if(dataArray[l] == currentValue && currentValue !=''){
                    dup=true;
                    countArray[l]++;
                      //rowcount++;
                  }
                
                 }
                if(dup == false && currentValue !='' ){
                    dataArray.push(currentValue);
                    countArray.push(1);
                    rowcount++;
                    
                }
                perArray.push('=B' + rowcount +'/C' + rowcount + '*100');
                

                totArray.push(totresp);
              }
            
            for( h in dataArray)
            {
              sheet.appendRow([dataArray[h], countArray[h], totArray[h], perArray[h]]);

            }

            
          }    // END OF FOR LOOP 
  
  
 //For Agree/Disagree questions 
  var beforeQues = sheet.getLastRow();
  var firstQues = beforeQues + 2;
  sheet.appendRow(['Question',"Agree","Neither Agree or Disagree","Disagree",'Don'+"'"+'t Know','Opp','IDK']); //Headings for counts
  var rowStr = rowcount+2
  var markstart = rowStr
      
          for (var k = 11; k <= numCols; k++) {
            
            var colname = columnToLetter(k);  //get letter for column number to use in formulas below
            //construct formulas of the format: =COUNTIF(D1:D9,"Agree")
            var Agreestr = '=COUNTIF('+ colname+ '1:'+ colname + lastdataindex +';"Agree")'
            var NAgDstr = '=COUNTIF('+ colname+ '1:'+ colname + lastdataindex +';"Neither Agree or Disagree")'
            var Disgreestr = '=COUNTIF('+ colname+ '1:'+ colname + lastdataindex +';"Disagree")'
            var DKstr = '=COUNTIF('+ colname+ '1:'+ colname + lastdataindex +';"Don'+"'"+'t Know")'
            var oppPerc = '=IFERROR(SUM(C' + rowStr + ':D' + rowStr + ')/(SUM(B' + rowStr + ':D' + rowStr + '))' + '*100,0)'
            var idkPerc = '=E' + rowStr +'/IF(SUM(B' + rowStr + ':E' + rowStr + ')=0,1,SUM(B' + rowStr + ':E' + rowStr + ')' + ')*100'
            sheet.appendRow(['Question '+(k-1),Agreestr,NAgDstr,Disgreestr,DKstr,oppPerc,idkPerc]);
            rowStr++;
          } 
  var markend = rowStr;
  var totalQues = sheet.getLastRow() - lastRowIndex - 3
  var totalPossResp = (lastRowIndex - 1) * (totalQues)

  //Sorting
  var sheetFinal = SpreadsheetApp.openByUrl(newSheetURL);
  SpreadsheetApp.setActiveSpreadsheet(sheetFinal);
  var newSheetFinal = SpreadsheetApp.getActiveSpreadsheet();
  
  if (sheetFinal.getSheetByName("Sorted Opp") == null) {
    sheetFinal.insertSheet("Sorted Opp");
  }
  else{
    sheetFinal.getSheetByName("Sorted Opp").clear()
  }
  if (sheetFinal.getSheetByName("Key Metrics") == null) {
    sheetFinal.insertSheet("Key Metrics");
  }
  else{
    sheetFinal.getSheetByName("Key Metrics").clear()
  }
  if (sheetFinal.getSheetByName("Final Report") == null) {
    sheetFinal.insertSheet("Final Report");
  }
  else{
    sheetFinal.getSheetByName("Final Report").clear()
  }
                   

  
  //Copies over question numbers and opportunity score values into tab1 on newSheet
  SpreadsheetApp.setActiveSpreadsheet(sheetFinal);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheets()[0];
  var destination = ss.getSheets()[1];
  
  destination.appendRow(["Question Number","Opportunity Score"]);

  var range = source.getRange("A"+markstart+":A"+markend)
  range.copyValuesToRange(destination, 1, 1, 2, markend-markstart+1);
  
  SpreadsheetApp.setActiveSpreadsheet(sheetFinal);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheets()[0];
  var destination = ss.getSheets()[1];

  var range = source.getRange("F"+markstart+":F"+markend);
  range.copyValuesToRange(destination, 2, 2, 2, markend-markstart+1);
  
     //Sorts by the value of the opportunity score, keeps the opportunity score and corresponding question together
     //Sorts from highest to lowest 
  var ss1 = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss1.getSheets()[1];
  var string = "A2:B"+(markend-markstart+1)
  var range1 = sheet1.getRange(string);
  range1.sort({column: 2, ascending: false});
  
  //Top 20 and Bottom 20 Opportunity Scores
  var sheetFinal = SpreadsheetApp.openByUrl(newSheetURL);
  SpreadsheetApp.setActiveSpreadsheet(sheetFinal);
  var newSheetFinal = SpreadsheetApp.getActiveSpreadsheet();
  
     //Copies the top and bottom 20 opportunity scores and question numbers into tab2
  SpreadsheetApp.setActiveSpreadsheet(sheetFinal);
  var ss3 = SpreadsheetApp.getActiveSpreadsheet();
  var source2 = ss3.getSheets()[1];
  var destination1 = ss3.getSheets()[2];

  destination1.appendRow(["TOP OPPORTUNITIES (highest opportunity score)", "Opportunity Score", "TOP STRENGTHS (lowest opportunity score)", "Opportunity Score"]) 
  
  var range1 = source2.getRange("A2:A21");
  range1.copyValuesToRange(destination1, 1, 1, 2, 21);
  
  
  SpreadsheetApp.setActiveSpreadsheet(sheetFinal);
  var ss3 = SpreadsheetApp.getActiveSpreadsheet();
  var source2 = ss3.getSheets()[1];
  var destination1 = ss3.getSheets()[2];

  var range1 = source2.getRange("B2:B21");
  range1.copyValuesToRange(destination1, 2, 2, 2, 21);
  
  SpreadsheetApp.setActiveSpreadsheet(sheetFinal);
  var ss3 = SpreadsheetApp.getActiveSpreadsheet();
  var source2 = ss3.getSheets()[1];
  var destination1 = ss3.getSheets()[2];
  var tempstr = "A"+(markend-markstart+1-19)+":A"+(markend-markstart+1)
  var range1 = source2.getRange(tempstr);
  range1.copyValuesToRange(destination1, 3, 3, 2, 21);
  
  SpreadsheetApp.setActiveSpreadsheet(sheetFinal);
  var ss3 = SpreadsheetApp.getActiveSpreadsheet();
  var source2 = ss3.getSheets()[1];
  var destination1 = ss3.getSheets()[2];
  var tempstr2 = "B"+(markend-markstart+1-19)+":B"+(markend-markstart+1)
  var range1 = source2.getRange(tempstr2);
  range1.copyValuesToRange(destination1, 4, 4, 2, 21);
  
  return beforeQues
}




function emailNotification(code,newSheetUrl){   // Emails user a link to the new spreadsheet 
    var email = Session.getActiveUser().getEmail(); 
    var message = "A new response was logged. Here's the link for the new spreadsheet!" + newSheetUrl;
    var subject = "Company" + code + "Analytics Spreadsheet" ; 
    GmailApp.sendEmail(email, subject, message); // sends email to user
        
} 

function newChart() {
  SpreadsheetApp.setActiveSpreadsheet(sheetFinal);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheets()[0];
  var destination = ss.getSheets()[3];
  var chart = source.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(source.getRange('B8:D10'))
    .setPosition(1, 1, 0, 0)
    .build();
  destination.insertChart(chart);}

function mainMonthly(){
  var yesOrNo = searchForFile(); // returns whether or not to make new sheet
  var URLs = startup(yesOrNo) ; //makes and returns the results sheet's url
  var newSheetURL = URLs[2];
  clearRange1(newSheetURL); // deletes current entries
  var tuple =  filterExcelMonthly(newSheetURL,URLs[1],URLs[0]); //re generates entries
  newChart();

}
function main2() {    // Ties all functions together with email notification, triggers every month 
  var yesOrNo = searchForFile(); // returns whether or not to make new sheet
  var URLs = startup(yesOrNo) ; //makes and returns the results sheet's url
  var newSheetURL = URLs[2];
  clearRange1(newSheetURL); // deletes current entries
  var tuple =  filterExcel(newSheetURL,URLs[1]); //re generates entries
  var totalResponse = tuple[0]
  var code2 = tuple[1]
  var lastRowIndex = basicAnalytics(newSheetURL,totalResponse); // applies basic analytics to entries and adds
  takeinConfig(lastRowIndex,newSheetURL,URLs[0])
 // emailNotification(code2,newSheetURL); // sends email notification to user
  
} 