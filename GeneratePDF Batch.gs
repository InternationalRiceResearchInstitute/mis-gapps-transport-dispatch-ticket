function GetNewBatchProcess(){ //Function to Get New Process of rows from CombinedTrips sheet to produce PDF Trip Tickets
try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));
  
   var scriptProperties = PropertiesService.getScriptProperties();
   SortSheet("CombinedTrips",19);
   // Clear Trigger GeneratePDFBatch
     var triggers = ScriptApp.getScriptTriggers();
     if (triggers.length>0) {
        for (var i=0; i<triggers.length; i++) {
            var handlerFunction = triggers[i].getHandlerFunction();
            if (handlerFunction=='GeneratePDFBatch') {
            ScriptApp.deleteTrigger(triggers[i]);
            }   
        }
     }  
  
   // Reset Properties
   scriptProperties.setProperty("runCount", 0); 
   scriptProperties.setProperty("rowList", 0); 
   scriptProperties.setProperty("rowCount", 0); 
   scriptProperties.setProperty("runCount", 0);
   scriptProperties.setProperty('Interval', 60); // lowest Value
   scriptProperties.setProperty('LastDocumentID', '');
   scriptProperties.setProperty('PDFDocument','');
   scriptProperties.setProperty('BatchName','');
   scriptProperties.setProperty('MaxTimeout', 240000);
   scriptProperties.setProperty('BatchPDFID','');
   scriptProperties.setProperty('BatchPDFURL','');
   scriptProperties.setProperty('BatchPDFlink','');  

   // Get Lines to Process
   GetProcessBatch();
   var rowList = scriptProperties.getProperty('rowList');
   // Run the main Looping Function
  if (rowList != '') {
    var ListArray = rowList.split(",");
    var Length = ListArray.length; 
  } else var Length = 0;
    
    
  if (Length > 0) {
    var Event = Length + ' Lines to Process (' + rowList + ')'; 
    //CreateLog(Event,Date());
    var SheetTimerLog =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TestLog');
    SheetTimerLog.clear();
    var value = ['x','Total Time', 'Condition','Elapsed Time','runCount','MaxTimeout','Time Log','FileID']
    SheetTimerLog.appendRow(value);
    //GeneratePDFBatch();  //Main Loop of Creating Folders 
    // Create Trigger to Run Script to Avoid Timeout from the first Function Run Generate PDF. 
          var Interval = Number(scriptProperties.getProperty('Interval'));
          var date = new Date();
          var newDate = new Date(date);
          newDate.setSeconds(date.getSeconds() + Interval);
          ScriptApp.newTrigger('GeneratePDFBatch').timeBased().at(newDate).create();
    //
  } else { var Event = 'Nothing to Process'; 
  //CreateLog(Event,Date()); 
  }
//Catch Error
}
catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
} 
//End Catch Error 
}
  

// Gets the Rows to Process for Autocrat Merging and PDF Converstion 
function GetProcessBatch(){
try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));  
  
var scriptProperties = PropertiesService.getScriptProperties();
var sheetCombinedTrips = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');
var LastRowCombinedTrips = sheetCombinedTrips.getLastRow();
    var arrayList = [];   var y = 0
    for (var x = 0; x < LastRowCombinedTrips-1; x++){
      var BatchStatus = sheetCombinedTrips.getRange(x+2,21).getValue();
          if (BatchStatus == ''){
              arrayList[y] = x+2; 
              y++
          }
    }
  
    Logger.log(arrayList.toString());
    Logger.log(arrayList.length);
    scriptProperties.setProperty("rowList",  arrayList.toString());   // Store value to PropertiesService 
    scriptProperties.setProperty("rowCount", arrayList.length);
//Catch Error
}
catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
} 
//End Catch Error  
}

// Main Function to avoid the Time Out Loop of 5 minutes 

function GeneratePDFBatch(){
try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));  
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var rowList = scriptProperties.getProperty('rowList');
  var rowCount = Number(scriptProperties.getProperty('rowCount'));
  var runCount = Number(scriptProperties.getProperty('runCount'));
  var MaxTimeout = Number(scriptProperties.getProperty('MaxTimeout'));
  
  var SheetTimerLog =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TestLog');
  //SheetTimerLog.clear();
  var start = new Date();
  

  for (var x = runCount; x < rowCount; x++){
  //var SampleTime = now.getTime();
  //Logger.log(SampleTime);
  //now.getTime() - start.getTime() > 2400000; // 5 minutes
      
    // Tsarot begin counter
    var startA = new Date();
    // Process here
    var ListArray = rowList.split(",")
    var rowProcess = ListArray[x];
    Logger.log(rowProcess);
    CreateTicket(rowProcess);
    // End Process Here
    var now = new Date();
    var TimeProcess = now.getTime() - startA.getTime();
    
      
      if (now.getTime() - start.getTime() < MaxTimeout){
          var elapsedTime = now.getTime() - start.getTime();
          var LastDocumentID = scriptProperties.getProperty('LastDocumentID');
          var value = [x,elapsedTime, (now.getTime() - start.getTime() < MaxTimeout),TimeProcess,runCount,MaxTimeout,now,LastDocumentID]
          scriptProperties.setProperty("runCount", x+1);
          SheetTimerLog.appendRow(value);
           
      } else {
          // Clear Trigger first 
          var triggers = ScriptApp.getProjectTriggers();
          for (var i=0; i<triggers.length; i++) {
             var handlerFunction = triggers[i].getHandlerFunction();
             if (handlerFunction=='GeneratePDFBatch') {
             ScriptApp.deleteTrigger(triggers[i]);
             }   
          }
          // call timer trigger here
          var Interval = Number(scriptProperties.getProperty('Interval'));
          var date = new Date();
          var newDate = new Date(date);
          newDate.setSeconds(date.getSeconds() + Interval);
          ScriptApp.newTrigger('GeneratePDFBatch').timeBased().at(newDate).create();
          var message = "Pausing PDF Tickets. Will resume after 1 minute";
          var messagevalue = [message];
          SheetTimerLog.appendRow(messagevalue);
          break; 
          // terminate function GeneratePDFBatch if reached timeout Value
      }
    
     var z = x+1; 
     if (z == rowCount){
          // call pdf merge and log the final tally 
     var startB = new Date();  
     var value = ['Converting to PDF'];    
     SheetTimerLog.appendRow(value);  
     ConvertDocToPDF();
     var now = new Date();
     var TimeProcess = now.getTime() - startA.getTime();  
     var elapsedTime = now.getTime() - start.getTime();
     var BatchPDFURL = scriptProperties.getProperty('BatchPDFURL');  
     var value = [x,elapsedTime, (now.getTime() - start.getTime() < MaxTimeout),TimeProcess,runCount,MaxTimeout,now,BatchPDFURL]
     SheetTimerLog.appendRow(value);
       
     var startC = new Date(); 
     UpdateCombinedTripsToTripTicket();  
     UpdatePDFLinksToAssign();  
     UpdatePDFLinksToTripRequest();  
     // UpdatePDFLinksToTripRequest();  
     var now = new Date();
     var TimeProcess = now.getTime() - startC.getTime();  
     var elapsedTime = now.getTime() - start.getTime();
     var BatchPDFlink = scriptProperties.getProperty('BatchPDFlink');  
     var value = [x,elapsedTime, (now.getTime() - start.getTime() < MaxTimeout),TimeProcess,runCount,MaxTimeout,now,'Logged']
     SheetTimerLog.appendRow(value);  
     var BatchName = Number(scriptProperties.getProperty('BatchName'));
     Browser.msgBox(rowCount + " Trip Ticket(s) PDF Generated. \\n\\nYou may now View/Print the Tickets in PDF Viewer.\\n\\nPDF File - " + BatchName); 
       
       
     }
  }
  
//Catch Error
}
catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
} 
//End Catch Error   
}

function UpdateBatchPDFtoTables(){
try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));    
  
     var SheetTimerLog =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TestLog');
     var scriptProperties = PropertiesService.getScriptProperties();
     var MaxTimeout = Number(scriptProperties.getProperty('MaxTimeout'));
     var runCount = Number(scriptProperties.getProperty('runCount'));
     var rowCount = Number(scriptProperties.getProperty('rowCount'));
     var start = new Date();
     var startC = new Date(); 
     UpdateCombinedTripsToTripTicket();  
     UpdatePDFLinksToAssign();  
     UpdatePDFLinksToTripRequest();  
     // UpdatePDFLinksToTripRequest();  
     var now = new Date();
     var TimeProcess = now.getTime() - startC.getTime();  
     var elapsedTime = now.getTime() - start.getTime();
     var BatchPDFlink = scriptProperties.getProperty('BatchPDFlink');  
     var value = [rowCount,elapsedTime, (now.getTime() - start.getTime() < MaxTimeout),TimeProcess,runCount,MaxTimeout,now,'Logged']
     SheetTimerLog.appendRow(value);  
     var BatchName = Number(scriptProperties.getProperty('BatchName'));
     Browser.msgBox(rowCount + " Trip Ticket(s) PDF Generated. \\n\\nYou may now View/Print the Tickets in PDF Viewer.\\n\\nPDF File - " + BatchName); 

//Catch Error
}
catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
} 
//End Catch Error  
  
}

function CreateTicket(j){
try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));   
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var numNew = 0; 
  var docIDs = []; var docIndex = 0; var charlimit = []; var trim = ""; var charlength = []; 
  var data = [];
     var i = 1;
     // Execute Data Trimmings
     for (var x = 0; x < 20; x++) {
           //var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
           var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');
           data[x] = sheetTicket.getRange(j, x+1).getValue();
           var sheetCharLimit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CharLimit');
           charlimit[x] = sheetCharLimit.getRange(2, x+1).getValue();   
           if (data[x] instanceof Date) { Logger.log("Its a date:" + x); } else {charlength[x] = data[x].length; Logger.log(x + " " + data[x] + " " + charlength[x] + " " + charlimit[x]); }
           if (charlength[x] >= charlimit[x]) { Logger.log("Long"); trim = data[x].substring(0, charlimit[x]) + "..."; data[x] = trim; Logger.log(data[x]);} else {Logger.log("Short");}
           if (data[16] == '') {data[16] = '__________' } //Replaces ___________ if Vehicle ID is not present
           if (data[17] == '') {data[17] = '__________________' } //Replaces ___________ if Driver is not present
           if (data[19] == '') {data[19] = '__________' } //Replaces ___________ if Plate Number is not present
            
     }
  var sheetFileID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FileIDList');
  var DocumentTemplate = sheetFileID.getRange(7,2).getValue(); 
  var fileId = DocumentTemplate;
  var DocumentFolderID = sheetFileID.getRange(10,2).getValue(); 
  var fileTayp = DriveApp.getFileById(fileId);       
  var NewDocumentFolder = DriveApp.getFolderById(DocumentFolderID); //Destination Folder ID
  var mimetype = fileTayp.getMimeType();
          //  var fileTayp1 = DriveApp.getFilesByType();
          //Logger.log(fileTayp + " ++ " + mimetype);
       if ((mimetype=="application/vnd.google-apps.document")||(mimetype=="application/vnd.google-apps.document")) { 
          var template = DocumentApp.openById(fileId);
          var title = template.getName();
          var NewDocID = DriveApp.getFileById(fileId).makeCopy(data[0],NewDocumentFolder).getId();
          //var NewDocID = DriveApp.getFileById(fileId).makeCopy(TRIPTICKET, NewDocumentFolder).getId();
          var NewDocName = DocumentApp.openById(NewDocID).getName();        
          var copyDoc = DocumentApp.openById(NewDocID);
          // Get the document’s body section and replace the details
          var copyBody = copyDoc.getActiveSection();  

          var ConvertedDate = Utilities.formatDate(data[2], "GMT+8","M/d/YYYY");
          copyBody.replaceText('<<TripDate>>', ConvertedDate);
          copyBody.replaceText('<<TripTicket>>', data[0]);
          copyBody.replaceText('<<PR Number>>', data[3]);
          copyBody.replaceText('<<VehicleID>>', data[16]);
          copyBody.replaceText('<<Driver>>', data[17]);  
          var ConvertTime = Utilities.formatDate(data[18], "GMT+8","h:mm a");
          copyBody.replaceText('<<PlateNo>>', data[19]);
          copyBody.replaceText('<<Departure>>', ConvertTime);
          copyBody.replaceText('<<BUS>>', data[4]); 
          copyBody.replaceText('<<Requestor>>', data[5]);
          copyBody.replaceText('<<OU>>', data[6]);
          copyBody.replaceText('<<Passenger>>', data[7]);
          if (data[8] instanceof Date){
          data[8]  = Utilities.formatDate(data[8], "GMT+8","h:mm a");
          } else { data[8] = data[8]; }  
            
          copyBody.replaceText('<<Pickup Time>>', data[8]);
          copyBody.replaceText('<<Pickup Place>>', data[9]);
          copyBody.replaceText('<<Destination>>', data[11]);
          copyBody.replaceText('<<Nature>>', data[12]);
          copyBody.replaceText('<<Other Instructions>>', data[10]);
          copyBody.replaceText('<<Trips>>', data[13]);          
          copyDoc.saveAndClose();
          
          //convert to PDF
          var source = DriveApp.getFileById(NewDocID);
          var DocURL = source.getUrl();
          var DocTitle = source.getName();
          docIDs[x] = NewDocID;
          //docIDs[docIndex] = NewDocID;
          var Doclink = [['=hyperlink("' + DocURL + '", "' + DocTitle + '")']];
          //docIndex++;
          //docIndex++;
        }
  
        // Get the Value of the LastDocument from Properties
        var LastDocumentID = scriptProperties.getProperty('LastDocumentID');
        var BatchDocID = LastDocumentID;
  
        // Check if Batch Document Exists? 
        if (LastDocumentID.length > 0) {
        // Staple the New Doc to the BatchDocID if BatchDocID Exists  
        var baseDoc = DocumentApp.openById(BatchDocID);
        var body = baseDoc.getActiveSection();
        //for (var x = 1; x < docIDs.length; ++x ) {
        var otherBody = DocumentApp.openById(NewDocID).getActiveSection();
        var totalElements = otherBody.getNumChildren();
        for( var z = 0; z < totalElements; ++z ) {
          var element = otherBody.getChild(z).copy();
          var type = element.getType();
          if( type == DocumentApp.ElementType.PARAGRAPH )
            body.appendParagraph(element);
          else if( type == DocumentApp.ElementType.TABLE )
            body.appendTable(element);
          else if( type == DocumentApp.ElementType.LIST_ITEM )
            body.appendListItem(element);
          else
            throw new Error("Unknown element type: "+type);
            } 
        //}
        baseDoc.saveAndClose();
        DriveApp.getFileById(NewDocID).setTrashed(true);
        //Delete the Old Doc 
        Drive.Files.remove(NewDocID);
        } else {
        // If Batch Document does not exist then Create a Batch Document     
        var NewDocumentFolder = DriveApp.getFolderById(DocumentFolderID);  //PDF Folder ID
        var sheetAutoNumber = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');  
        var BatchNum = sheetAutoNumber.getRange(5, 2).getValue();
        var Batchname = 'Batch' + (Number(10000) + Number(BatchNum));
        //Copy the New Document to a New File Batchname
        var BatchDocID = DriveApp.getFileById(NewDocID).makeCopy(Batchname,NewDocumentFolder).getId();
        //Delete the Old Doc 
        DriveApp.getFileById(NewDocID).setTrashed(true);
        Drive.Files.remove(NewDocID);
        // Save the BatchDoc ID in Properties 
        scriptProperties.setProperty('LastDocumentID', BatchDocID);
        }

//Catch Error
}
catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
} 
//End Catch Error 
  
}

function ConvertDocToPDF(){
try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));   
  
        var scriptProperties = PropertiesService.getScriptProperties();
        var sheetFileID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FileIDList');
        var DocumentFolderID = sheetFileID.getRange(10,2).getValue(); 
        var LastDocumentID = scriptProperties.getProperty('LastDocumentID');
        var logemail = Session.getActiveUser().getEmail();
        var DateTime = Utilities.formatDate(new Date, "GMT+8","YYYYMMDD h:mm a");
        var source = DriveApp.getFileById(LastDocumentID);
        var PDFName = source.getName(); 
        var blob = source.getAs('application/pdf');
        var PDFfile = DriveApp.getFolderById(DocumentFolderID).createFile(blob);
        var BatchPDFID = PDFfile.getId();
        var BatchPDFURL = PDFfile.getUrl();
        var BatchPDFlink = [['=hyperlink("' + BatchPDFURL + '", "' + PDFName + '")']];

        scriptProperties.setProperty('BatchPDFID',BatchPDFID);
        scriptProperties.setProperty('BatchPDFURL',BatchPDFURL);
        scriptProperties.setProperty('BatchPDFlink',BatchPDFlink);  
  
        DriveApp.getFileById(LastDocumentID).setTrashed(true);
        Drive.Files.remove(LastDocumentID);
        scriptProperties.setProperty('PDFDocument',BatchPDFURL);
        scriptProperties.setProperty('BatchName',PDFName);
        //Update Batch Number
        var sheetAutoNumber = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');
        var BatchNum = sheetAutoNumber.getRange(5, 2).getValue();
        BatchNum++;
        sheetAutoNumber.getRange(5, 2).setValue(BatchNum);
//Catch Error
}
catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
} 
//End Catch Error         
}

function UpdateCombinedTripsToTripTicket(){ //updated code 9/26/2017 2:05 PM
try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('(')); 
//start function code   
      var dataCombinedtrips = []; 
      var scriptProperties = PropertiesService.getScriptProperties();
      var BatchPDFID = scriptProperties.getProperty('BatchPDFID');
      var BatchName = scriptProperties.getProperty('BatchName');   
      var BatchPDFURL = scriptProperties.getProperty('BatchPDFURL');
      var logemail = Session.getActiveUser().getEmail(); // gets the email of the user who made the PDF Trip Ticket. 
      var CurrentDateTime = Utilities.formatDate(new Date(), "GMT+8","M/d/YYYY+h:mm a");
      var DateTimeLog = 'Document successfully merged by: ' +logemail+ " " + CurrentDateTime; 
      var BatchPDFlink = [['=hyperlink("' + BatchPDFURL + '", "' + BatchName + '")']];
      Logger.log('BatchPDFID'+'--'+BatchPDFID);
      Logger.log('BatchName'+'--'+BatchName);
      Logger.log('BatchPDFURL'+'--'+BatchPDFURL);
      var UpdateRow = [BatchPDFID,BatchPDFURL,BatchPDFlink,DateTimeLog]; 
      
      Logger.log(UpdateRow);
      var sheetCombineTrips = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');   
      var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket'); 
      var DatasheetCombinedTrips = sheetCombineTrips.getDataRange().getValues();
      var DatasheetTripTickets = sheetTripTicket.getDataRange().getValues(); 
      var sheetTEST = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TEST'); 
      
      //sheetTripTicket.getRange(k, 24).setValue('Document successfully merged by: ' +logemail+" " + CurrentDateTime); 
      Logger.log(UpdateRow.length);
      //sheetTEST.getRange(6, 4, 1, UpdateRow.length).setValues([UpdateRow]);   
      for (var x = 1; x<DatasheetCombinedTrips.length; x++){
          Logger.log(DatasheetCombinedTrips[x][0]); 
          for (var y = 0; y<DatasheetTripTickets.length; y++){
               if(DatasheetCombinedTrips[x][0] == DatasheetTripTickets[y][0]){
               var rowx = x+1; // row of CombinedTrip reference
               var rowy = y+1; // row of Trip Ticket reference
               //Logger.log(DatasheetCombinedTrips[x][0] +'-- Match --'+ DatasheetTripTickets[y][0]);
               //Logger.log('rowx - ' + rowx +';rowy -'+rowy); 
               sheetTripTicket.getRange(rowy, 21, 1, UpdateRow.length).setValues([UpdateRow]);   // logs the PDF ID, BatchName, PDFhyperlink, datetime user 
               var BegTimeLog = Utilities.formatDate(DatasheetCombinedTrips[x][18], "GMT+8","h:mm:ss a"); //logs time departure of the driver
               sheetTripTicket.getRange(rowy, 29).setValue(BegTimeLog);
               //Logger.log('BegTimeLog: '+ DatasheetCombinedTrips[x][18]); 
               }
          }
      }
//end function code 
//Catch Error
}
catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
} 
//End Catch Error       
}

function UpdatePDFLinksToAssign(){ //updated code 9/26/2017 3:05
try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));      
//start function code  
var sheetCombinedTrips = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');    
var DataCombinedTrips = sheetCombinedTrips.getDataRange().getValues(); 
      var scriptProperties = PropertiesService.getScriptProperties();
      var BatchPDFID = scriptProperties.getProperty('BatchPDFID');
      var BatchPDFURL = scriptProperties.getProperty('BatchPDFURL');
      var BatchName = scriptProperties.getProperty('BatchName'); 
      var BatchPDFlink = [['=hyperlink("' + BatchPDFURL + '", "' + BatchName + '")']];
      for (var k = 1; k < DataCombinedTrips.length; k++) {
          var rowz = k+1; 
          sheetCombinedTrips.getRange(rowz, 21).setValues(BatchPDFlink);
    }
//ednd function code
//Catch Error
}
catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
} 
//End Catch Error    
}

function UpdatePDFLinksToTripRequest(){ //Enhancement for Updating PDFLink from CombinedTrips to TripRequest 9/27/2017
try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));      
  
var sheetTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');    
var DataTripRequest = sheetTripRequest.getDataRange().getValues();  
var sheetCombinedTrips = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');  
var DataCombinedTrips = sheetCombinedTrips.getDataRange().getValues()
var scriptProperties = PropertiesService.getScriptProperties();
var BatchPDFURL = scriptProperties.getProperty('BatchPDFURL');
var PDFHyperLink ='<b><a href="'+BatchPDFURL+'" style="text-decoration:none;background-color:transparent" target="_blank">ⓘ</a></b>';
var countupdate = 0;   
//Logger.log(BatchPDFURL);
    for(var v = 1; v<DataCombinedTrips.length; v++){
         var rowv = v+1; 
         //Logger.log(DataCombinedTrips[v][0] + ' - ' + rowv);
         for (var u = 1; u<DataTripRequest.length; u++){
             var rowu = u+1; //row of TripRequest
             if(DataTripRequest[u][21] == DataCombinedTrips[v][0]){
             Logger.log(DataCombinedTrips[v][0]); 
             //Logger.log(DataCombinedTrips[v][0] +' -- '+ 'rowv: '+rowv+ ' rowu: ' +rowu);  
             //countupdate++;
             sheetTripRequest.getRange(rowu, 22).setValue(PDFHyperLink);
             }
         }
        
    }

//Catch Error
}
catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
} 
//End Catch Error  
}





  


 
