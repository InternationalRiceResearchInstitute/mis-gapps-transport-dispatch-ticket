//var formURL = 'https://docs.google.com/a/irri.org/forms/d/1cITLFbnswWenBs1iSB9vjbIYhYo5SS2pgpBcfPvAKVU/viewform';
var formURL = 'https://docs.google.com/a/irri.org/forms/d/1700ClnHUBTTz33rqSe74huHz1tP8C2xYijqsZ_GMWG0/viewform';
var SaveLogSheet = '1gq-3kfz3j_Wax2yMDirHV-wkONUYCnmPlg9oAD9VUBs'
var sheetName = 'TripRequest';
var FormLink = 16;
var status = '';

function getResponseLink(){
  var sheetrange = SpreadsheetApp.getActiveRange(), data = sheetrange.getValues();
  var output = [];
  for(var i = 0, iLen = data.length; i < iLen; i++) {
  var timestamp = data[i][0];
 
  var formSubmitted = FormApp.openByUrl(formURL).getResponses(timestamp);
  // var formSubmitted = form.getResponses(timestamp);
  var editResponseUrl = formSubmitted[0].getEditResponseUrl();
  Logger.log(editResponseUrl);
       var link = [[editResponseUrl]];
       output.push([editResponseUrl]);
       Logger.log(output);
       sheetrange.offset(0,15).setValues(output);
  }
}

function getURLform(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
  var lastRow = sheet.getLastRow();
  var data = sheet.getDataRange().getValues();
  var sheetFileID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FileIDList');
  var FormLinkID = sheetFileID.getRange(4,2).getValue();
  var formURL = 'https://docs.google.com/a/irri.org/forms/d/'+FormLinkID+'/viewform'; 
  var form = FormApp.openByUrl(formURL); //What's wrong? 
  var timestamp =  new Date(sheetFileID.getRange(7458,1).getValue());
  var formSubmitted = form.getResponses(timestamp);
  Logger.log(timestamp);
  Logger.log(formSubmitted);
  if(formSubmitted.length < 1) {
  var editResponseUrl = formSubmitted[0].getEditResponseUrl();
  Logger.log(edirResponseUrl); 
                               }
}

function getEditResponseUrls(){
try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));  
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
  var lastRow = sheet.getLastRow();
  var data = sheet.getDataRange().getValues();
  var YearToday = Utilities.formatDate(new Date(), "GMT+8","YYYY");
//  var RequestNo =  YearToday + (Number(10000) + Number(lastRow-1)); // Autonumber ticketing
//  sheet.getRange(lastRow, 13).setValue(RequestNo);
  var DateToday = Utilities.formatDate(new Date(), "GMT+8","YYYYMMdd");
  sheet.getRange(lastRow, 18).setValue(DateToday);
  sheet.getRange(lastRow, 19).setValue("0");                          // Set default value of Combine Trip to Zero
  var sheetFileID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FileIDList');
  var FormLinkID = sheetFileID.getRange(4,2).getValue();
  var formURL = 'https://docs.google.com/a/irri.org/forms/d/'+FormLinkID+'/viewform';  
  var form = FormApp.openByUrl(formURL); //What's wrong? 
  for(var i = 2; i < data.length; i++) {
    if(data[i][0] != '' && data[i][FormLink-1] == '') {
      var timestamp = data[i][0];
      var formSubmitted = form.getResponses(timestamp);
      if(formSubmitted.length < 1) continue;
      var editResponseUrl = formSubmitted[0].getEditResponseUrl();
      sheet.getRange(i+1, 16).setValue(editResponseUrl);
      var sheetAutoNum = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');
      var YearToday = Utilities.formatDate(new Date(), "GMT+8","YYYY");
      var newTripNum = Number(sheetAutoNum.getRange(1, 2).getValue());
      var RequestNo =  YearToday + (Number(100000) + Number(newTripNum));
      var Tags = sheet.getRange(i+1, 19).getValue();
      newTripNum++;
      var rowi = i+1; 
      sheetAutoNum.getRange(1, 2).setValue(newTripNum);
      sheet.getRange(i+1, 17).setValue(RequestNo);
      sheet.getRange(i+1, 18).setValue(DateToday);
      var dateentry = sheet.getRange(i+1, 8).getValues();
      var DateSource = new Date(dateentry);
      //=if(H5609<>"",TEXT(H5609,"M/d/yyyy"),"")
      var textvalue = '=if(H'+rowi+'<>"",TEXT(H'+rowi+',"M/d/yyyy"),"")'; 
      sheet.getRange(i+1, 26).setValue(textvalue);
      if (Tags != 0){ sheet.getRange(i+1, 19).setValue(Tags);} else {var Tags = 0; sheet.getRange(i+1, 19).setValue(Tags);}   
      var DateToday = Utilities.formatDate(new Date(), "GMT+8","YYYYMMdd");
    //  sheet.getRange(i+1, 17).setValue(DateToday);                  //Update the Date of Entry / Edit
    }
  }
  
//Catch Error
}
catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
} 
//End Catch Error   
}


function getEditResponseUrls2(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
  var lastRow = sheet.getLastRow();
  var data = sheet.getDataRange().getValues();
  var YearToday = Utilities.formatDate(new Date(), "GMT+8","YYYY");
//  var RequestNo =  YearToday + (Number(10000) + Number(lastRow-1)); // Autonumber ticketing
//  sheet.getRange(lastRow, 13).setValue(RequestNo);
  var DateToday = Utilities.formatDate(new Date(), "GMT+8","YYYYMMdd");
                      // Set default value of Combine Trip to Zero
  //sheet.getRange(lastRow, 15).setValue("0");   
  var form = FormApp.openByUrl(formURL); //What's wrong? 
  for(var i = 2; i < data.length; i++) {
    if(data[i][0] != '' && data[i][FormLink-1] == '') {
      var timestamp = data[i][0];
      var formSubmitted = form.getResponses(timestamp);
      if(formSubmitted.length < 1) continue;
      var editResponseUrl = formSubmitted[0].getEditResponseUrl();
      sheet.getRange(i+1, 16).setValue(editResponseUrl);
      var sheetAutoNum = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');
      var YearToday = Utilities.formatDate(new Date(), "GMT+8","YYYY");
      var newTripNum = Number(sheetAutoNum.getRange(1, 2).getValue());
      var RequestNo =  YearToday + (Number(10000) + Number(newTripNum));

      newTripNum++;
      sheetAutoNum.getRange(1, 2).setValue(newTripNum);
      sheet.getRange(i+1, 17).setValue(RequestNo);
      var DateToday = Utilities.formatDate(new Date(), "GMT+8","YYYYMMdd");
      //  sheet.getRange(i+1, 17).setValue(DateToday);                  //Update the Date of Entry / Edit
    }
  }
}


function onOpen() {
  FirstMenu();
  var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
  var TripTicketCount = sheetTripTicket.getLastRow(); 
  var a1NotationTripticket = "A3:AC"+TripTicketCount;
  var dataTripTicket = sheetTripTicket.getRange(a1NotationTripticket).getValues();
  dataTripTicket.length
  Browser.msgBox("TripTickets Active: "+dataTripTicket.length);
  if (dataTripTicket.length > 250) {
  Browser.msgBox("Trip Tickets are now more than 250. Please consider Archiving now.");
  }
}


function completeMenu(){
 SpreadsheetApp.getUi()
    .createMenu("Transaction")
    .addItem("Select Trip Request Day","TripRequestSummary")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Generate Trip Ticket')
                .addItem("All List","CombineTripStandard")
                .addItem("Additional","CombineTripAdditional"))
    .addSeparator()
    .addItem("Generate PDF Trip Tickets","PDFTripTicket")
    .addItem("Select Trip Ticket Day","TripTicketSummary")
    .addToUi()  

}


function FirstMenu(){
 SpreadsheetApp.getUi()
    .createMenu("Transaction")  
    .addItem("Generate Regular Trip", "GenerateRegularTrips")
    .addItem("Select Trip Request Day","TripRequestSummary")
    .addItem("Save Combination Tag","SaveTripMarkings")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Generate Trip Ticket')
                .addItem("All List","CombineTripStandard")
                .addItem("Additional","CombineTripAdditional"))
    .addItem("Proceed to Trip Ticket Commands","TripTicketCommandList")
    .addToUi()  
}

function SecondMenu(){
   SpreadsheetApp.getUi()
    .createMenu("Transaction")
    .addItem("Generate PDF Trip Tickets","PDFTripTicket")
     //.addItem("Backup Run PDF Trip Tickets ","GetNewBatchProcess")
    .addItem("Select Trip Ticket Day","TripTicketSummary")
    .addItem("Save Trip Ticket Details","SaveTripTicketDetails")
    .addItem("Back to Trip Request Commands","TripTrequestCommands")
    .addItem("Archive Trip Tickets","ArchiveTripTickets") 
    .addToUi()  
}

function TripTicketCommandList(){
SecondMenu();
}

function TripTrequestCommands(){
FirstMenu();

}


function TripRequestSummary() {
 var PickupDate =  Browser.inputBox('Select Trip Request Day', 'Enter Trip Request Pick Up Date M/D/YYYY fomat', Browser.Buttons.OK);
  ScriptProperties.setProperty('TripRequestSummary', PickupDate);
  Logger.log (PickupDate);
  if (PickupDate != 'cancel')
  { ListTripRequest(PickupDate); }
}

function CombineTripStandard() {
autosortFormulaStandard();
//var Additional = 'False'; 
//return Additional; 
CombineTripCheck();
var sheetCombine = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests')
var PickUpDate = sheetCombine.getRange(2, 6).getValue();
var PickupDate = Utilities.formatDate(new Date(PickUpDate), "GMT+8","M/d/YYYY"); 

ListTripTicket(PickupDate);  
  SecondMenu();
}

function CombineTripAdditional() {
autosortFormulaAdditional();
//var Additional = 'True'; 
//return Additional; 
CombineTripCheck();
var sheetCombine = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests')
var PickUpDate = sheetCombine.getRange(2, 6).getValue();
var PickupDate = Utilities.formatDate(PickUpDate, "GMT+8","M/d/YYYY"); 

ListTripTicket(PickupDate);
  SecondMenu();
}


function TripTicketSummary() {
 var PickupDate =  Browser.inputBox('Trip Ticket Date to List', 'Enter Trip Ticket Pick Up Date MM/DD/YYYY fomat', Browser.Buttons.OK);
  ScriptProperties.setProperty('TripTicketSummary', PickupDate);
  if (PickupDate !='cancel'){ ListTripTicket(PickupDate); }
  
}

function SaveTripMarkings() {
UpdatePrimaryTable();
}


function CombineTripCheck() {

   var sheetcomb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoSort');
   var LastRowComb = sheetcomb.getLastRow();
   var TripDateMerge = sheetcomb.getRange(2, 6).getValue();
   //Logger.log("Rows: " + LastRowComb);
   Logger.log(TripDateMerge);
   
   if (TripDateMerge instanceof Date){  // if P
    
    //Logger.log("Valid Date");  
    UpdatePrimaryTable();
        if (LastRowComb > 1) {
          var ConvertedDate = Utilities.formatDate(TripDateMerge, "GMT+8","M/d/YYYY");
          //Logger.log(ConvertedDate);
          //Browser.msgBox("Processing "+ (LastRowComb-1) + " Rows. Date: " + ConvertedDate); 
            CombineMergeTrip(ConvertedDate);
            
        }
        if (LastRowComb == 1) {
            Browser.msgBox("No Rows Found To Process"); 
        }
        
    //Browser.msgBox("MergeTrip Processing for Date: " + ConvertedDate + "/n" + "Number of Rows to Process: " + (LastRowComb-1));
    //Run Trip Ticket ID Generation
  }
  else if (TripDateMerge ==  "") {
    //Logger.log("No Date Found");
    Browser.msgBox("No Date Found")
  }
  else {
   //Logger.log("Not a Valid Date"); //Output if Date Detected is not a Date format
   Browser.msgBox("No Date Found");
  }
  
}

function SaveTripTicket() {
 UpdateTripTicket();
}



function PDFTripTicket() {
  var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');
  var lastRowAssign = sheetAssign.getLastRow(); 
  var row = []; var invalidrow = 0;
  // check if there are blank departure Time. 
  for (var x = 2; x < lastRowAssign+1; x++){
       var DepartTime = sheetAssign.getRange(x,19).getValue();
       if (DepartTime == "") {
       }      
       if (DepartTime instanceof Date){} else {
       row[x-2] = x; invalidrow++;
       }
  } var invalidlists = row.join(", ");
  
  
  if (invalidrow > 0) { Browser.msgBox("Row(s) " + invalidlists + " has blank or invalid Depart Time Format\\n\\nPlease use H:MM AM/PM - sample 8:00 AM \\n\\nPDF Tickets not Generated.");
  } else {// Logger.log("Process PDF Save and Print"); 
  //Browser.msgBox("Departure Date Complete");
    UpdateTripTicket(); //activate this
  //MergeDocument(); 
  }
}



function autosortFormulaStandard() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoSort');
  sheet.getRange(1, 1).setFormula('=query(SelectedTripRequests!$A:$R,"SELECT A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R WHERE M is not null ORDER BY F,N,O")');
}
function autosortFormulaAdditional() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoSort');
  sheet.getRange(1, 1).setFormula('=query(SelectedTripRequests!$A:$R,"SELECT A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R WHERE P=1 ORDER BY F,N,O")');
}

function testpickupdate(){
  var date = '12/31/2017';
  ListTripRequest(date);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
  var DataTripRequest = sheet.getDataRange().getValues(); 
  Logger.log(DataTripRequest.length);
}


function ListTripRequest(PickupDate){
  try{
      var funcName = arguments.callee.toString();
      funcName = funcName.substr('function '.length);
      funcName = funcName.substr(0, funcName.indexOf('('));
//function start
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest') 
   var lastRow = sheet.getLastRow()  
   var DataTripRequestList = sheet.getRange(1, 3, lastRow, 21).getValues(); 
   //Logger.log(DataTripRequestList); 
   var sheetpool = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests');
   var headers =["PR Number","BUS Code","PR Requestor","OU","Passenger(s)","Pickup Date","Pick Up Time","Return Date","Pick Up Location","Point of Destination","Other Instructions","Nature of Trip","RequestNo","Combine Num","Sequence","Additional","Trip Ticket","Depart Time"]
   sheetpool.clearContents();
   sheetpool.appendRow(headers);
   sheetpool.setFrozenRows(1);  
   var numNew = 0;
   var TripRequestNo = []; 
    Logger.log(DataTripRequestList.length); 
      for (var x = 2; x<DataTripRequestList.length;x++){
           
          var CompareDate = Utilities.formatDate(new Date(DataTripRequestList[x][5]), "GMT+8","M/d/YYYY");
          var PickUpTime = Utilities.formatDate(new Date(DataTripRequestList[x][6]), "GMT+8","hh:mm a") 
          var ReturnDate = Utilities.formatDate(new Date(DataTripRequestList[x][7]), "GMT+8","M/d/YYYY");
          
          if (DataTripRequestList[x][12] == 'Ongoing' && CompareDate == PickupDate){
          //Logger.log(DataTripRequestList[x])
          //var values = [Passenger,PickUpDate,deptime,PickUpTime,ReturnDate,PickupLocation,Destination]
          //Logger.log(String(DataTripRequestList[x][14]));
          //var RequestNo = String(DataTripRequestList[x][14]); 
          TripRequestNo[numNew] = String(DataTripRequestList[x][14]);
          
          var ReferenceID = TripRequestNo.join(" / ");
          //var values = [DataTripRequestList[x][0],CompareDate,"",PickUpTime,ReturnDate,DataTripRequestList[x][8],DataTripRequestList[x][9]]
          //var values = [PRNumber,BUSCode,Requestor,OU,Passenger,PickUpDate,PickUpTime,ReturnDate,PickupLocation,Destination,Instructions,Nature,RequestNo,CombineNum,Sequence,Additional,TripTicket,DepartTime]
          var values = [DataTripRequestList[x][0],
                        DataTripRequestList[x][1],
                        DataTripRequestList[x][2],
                        DataTripRequestList[x][3],
                        DataTripRequestList[x][4],
                        CompareDate,
                        PickUpTime,
                        ReturnDate,
                        DataTripRequestList[x][8],
                        DataTripRequestList[x][9],
                        DataTripRequestList[x][10],
                        DataTripRequestList[x][11],
                        DataTripRequestList[x][14],
                        DataTripRequestList[x][16],
                        DataTripRequestList[x][17],
                        DataTripRequestList[x][18],
                        DataTripRequestList[x][19],
                        DataTripRequestList[x][20]]
          sheetpool.appendRow(values);
          numNew++;
          }
      }
    //Logger.log(ReferenceID);
      FomatPlainText("SelectedTripRequests",1);
  if (numNew != 0) {} else {var ReferenceID  = 'None';}
  var sheetFileID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FileIDList');
  var LogSheetID = sheetFileID.getRange(5,2).getValue(); 
  var LogSheetTable = SpreadsheetApp.openById(LogSheetID).getSheetByName('Event_Logs');  
  var logemail = Session.getActiveUser().getEmail();
  var DateTime = Utilities.formatDate(new Date, "GMT+8","YYYYMMdd h:mm:ss a");
  var Logs = [logemail,DateTime,'Select Trip Request Day',numNew +' Selected Trip Request for date: ' + PickupDate,'Trip Request',ReferenceID];
  LogSheetTable.appendRow(Logs);

  if (numNew == 0) {
  Browser.msgBox("No Pickup Date Found");
  }
         sheetpool.setActiveSelection("N2");
         sheetpool.setTabColor("ff0000");    
    
//function end
}
catch(e){
    MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Trip Request List', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})

}     
 
}

function UpdatePrimaryTable(){
    var sheetupdate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests');
    var sheetUpdateLastRow = sheetupdate.getLastRow(); 
    var DataUpdateTripRequest = sheetupdate.getRange(1, 13, sheetUpdateLastRow, 6).getValues();                       //Array Source
    
    var sheetdestination = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
    var sheetdestinationLastRow = sheetdestination.getLastRow();      
    var DataTripRequestNo = sheetdestination.getRange(1, 17, sheetdestinationLastRow, 1).getValues();                 //Array Destination
    var sheetTestLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PrimaryUpdateTimer');
    //sheetTestLog.getRange(1, 10, DataTripRequestNo.length, DataTripRequestNo[0].length).setValues(DataTripRequestNo)  
    
    var sheetTimer = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PrimaryUpdateTimer');
    sheetTimer.clear();
    var value = ['x','Total Time', 'Condition','Elapsed Time','runCount','MaxTimeout','Time Log','FileID'];
    sheetTimer.appendRow(value);
    var MaxTimeout = 240000;
    var start = new Date();
        for(var i = 1;i<DataUpdateTripRequest.length;i++){
           var startA = new Date();
           Logger.log(DataUpdateTripRequest[i])
           var rowsource = i+1; 
           for(var j = 2; j<DataTripRequestNo.length; j++){
                 if (DataUpdateTripRequest[i][0] == DataTripRequestNo[j]){
                        var rowupdate = j+1;       
                        sheetdestination.getRange(rowupdate, 19).setValue(DataUpdateTripRequest[i][1]);
                        sheetdestination.getRange(rowupdate, 20).setValue(DataUpdateTripRequest[i][2]);
                        sheetdestination.getRange(rowupdate, 21).setValue(DataUpdateTripRequest[i][3]);
                        sheetdestination.getRange(rowupdate, 22).setValue(DataUpdateTripRequest[i][4]);
                        sheetdestination.getRange(rowupdate, 23).setValue(DataUpdateTripRequest[i][5]);
                 }
           }
            //Put Timer Log here
    var now = new Date();
    var TimeProcess = now.getTime() - startA.getTime();
         // if (now.getTime() - start.getTime() < MaxTimeout){
          var elapsedTime = now.getTime() - start.getTime();
          var value = [i,elapsedTime, (now.getTime() - start.getTime() < MaxTimeout),TimeProcess,"none",MaxTimeout,now,DataUpdateTripRequest[i][0]]
          sheetTimer.appendRow(value);
         // }
    //End Timer Log Here   
        }
    //Logger.log(DataRequestNumberRow);
}





function UpdateTripTicket() { // Saves/Update the Trip Ticket from AssignTable Sheet to TripTicket Sheet details of VehicleID, Driver, Departure Time, Plate Number
     var sheetCombinedTrips = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');
     var DataCombinedTrips = sheetCombinedTrips.getRange(1, 1, sheetCombinedTrips.getLastRow(), 20).getValues();

     var sheetSampleLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SampleLog');
     sheetSampleLog.clearContents();
     var countblank = 0;
     var TripTicket = []; var VehicleID = []; var Driver = []; var DepartTime = []; var PlateNo = [];  var PickupTime = []; var PickupPlace = []; var OtherInstructions = []; var PickupTime2 = [];
                 for (var v = 1; v < DataCombinedTrips.length; v++) {
                          var TripTicket = DataCombinedTrips[v][0];        
                          var PickupTime = DataCombinedTrips[v][8];         
                          var PickupPlace = DataCombinedTrips[v][9];         
                          var OtherInstructions = DataCombinedTrips[v][10];  
                          var PickupTime2 = DataCombinedTrips[v][14];   
                          var VehicleID = DataCombinedTrips[v][16]; 
                          var Driver = DataCombinedTrips[v][17]; 
                          var DepartTime = DataCombinedTrips[v][18];
                          var PlateNo = DataCombinedTrips[v][19];
                          if (DepartTime instanceof Date){
                            DepartTime = Utilities.formatDate(DepartTime, "GMT+8","h:mm:ss a");
                               //Logger.log(v + " " + DepartTime[v]);
                          } else { }
                          if (VehicleID == "" || Driver == "" || DepartTime == "" || PlateNo == "") { countblank ++; } //Checks for Blank details in Yellow Section
                          //var logsample = [TripTicket[v],PickupTime[v],PickupPlace[v],OtherInstructions[v],PickupTime2[v],VehicleID[v],Driver[v],DepartTime[v],PlateNo[v]];
                          //sheetSampleLog.appendRow(logsample)
                 }
                 //Browser.msgBox("Countblank: " + countblank);
                 //Logger.log("Countblank: " + countblank); //test data
                 var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
                 var DataTripTicket = sheetTripTicket.getRange(1, 1, sheetTripTicket.getLastRow(), 29).getValues();
                 var sheetTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
                 var DataTripRequest = sheetTripRequest.getRange(1, 1, sheetTripRequest.getLastRow(), 25).getValues();
                 
                 if (countblank == 0) { 

                      for (var v = 1; v < DataCombinedTrips.length; v++) {
                                   var TripTicketCombined = DataCombinedTrips[v][0];
                                   Logger.log(DataCombinedTrips.length);
                                   //Logger.log(DataTripTicket); 
                                   for (var w = 2; w < DataTripTicket.length; w++) {
                                      var TripTicketID = DataTripTicket[w][0];

                                      if (TripTicketID == TripTicketCombined) {
                                      var rowW = w+1;
                                      if (DataCombinedTrips[v][8] instanceof Date){ 
                                      var PickupTime =  Utilities.formatDate(DataCombinedTrips[v][8], "GMT+8","hh:mm:ss a")  
                                      } else { var PickupTime = DataCombinedTrips[v][8]}
                                      sheetTripTicket.getRange(rowW, 9).setValue(PickupTime);                                        //PickupTime
                                      sheetTripTicket.getRange(rowW, 10).setValue(DataCombinedTrips[v][9]);                          //PickupPlace 
                                      sheetTripTicket.getRange(rowW, 11).setValue(DataCombinedTrips[v][10]);                         //OtherInstructions
                                      if (DataCombinedTrips[v][8] instanceof Date){ 
                                      var PickupTime2 =  Utilities.formatDate(DataCombinedTrips[v][14], "GMT+8","hh:mm:ss a")  
                                      } else { var PickupTime2 = DataCombinedTrips[v][14]}
                                      sheetTripTicket.getRange(rowW, 15).setValue(PickupTime2);                                      //PickupTime2
                                      sheetTripTicket.getRange(rowW, 17).setValue(DataCombinedTrips[v][16]);                         //VehicleID
                                      sheetTripTicket.getRange(rowW, 18).setValue(DataCombinedTrips[v][17]);                         //Driver
                                       if (DataCombinedTrips[v][8] instanceof Date){ 
                                      var DepartTime =  Utilities.formatDate(DataCombinedTrips[v][18], "GMT+8","hh:mm:ss a")  
                                      } else { var DepartTime = DataCombinedTrips[v][18]}
                                      sheetTripTicket.getRange(rowW, 19).setValue(DepartTime);                                       //DepartTime                            
                                      sheetTripTicket.getRange(rowW, 28).setValue(DepartTime);
                                      sheetTripTicket.getRange(rowW, 20).setValue(DataCombinedTrips[v][19]);                         //PlateNumber
                                      }
                                      
                                   }
                                   for (var z = 2; z < DataTripRequest.length; z++) {
                                   var TripTicketIDTripRequest = DataTripRequest[z][21]; //sheetTripRequest.getRange(z, 22).getValue();  
                                   //Logger.log("TripTicket Request: " + TripTicketIDTripRequest);    
                                         if (TripTicketIDTripRequest == TripTicketCombined) {
                                              //sheetTripRequest.getRange(w, 18).setValue(DepartTime[v]);
                                              var DepartTime =  Utilities.formatDate(DataCombinedTrips[v][18], "GMT+8","hh:mm a")                                             
                                              //sheetTripRequest.getRange(z+1, 23).setValue(DepartTime);
                                         }
                                   }
                      }  var status = 'print'; //You may now run Generate PDF Tickets Browser.msgBox("Saving Trip Ticket","Trip Ticket Updated the following details: \\nVehicleID, Driver\\n\\nGenerating PDF Tickets now.",Browser.Buttons.OK);
                 } else { if (Browser.msgBox("Saving Trip Ticket","Incomplete Vehicle, Driver details. \\nDo you want to save the Trip Ticket with blank details? ", Browser.Buttons.YES_NO) == 'yes') {
                           //----------------------------------------------------------------------------------------------------------------------   
                                for (var v = 1; v < DataCombinedTrips.length; v++) {
                                   var TripTicketCombined = DataCombinedTrips[v][0];
                                   Logger.log(DataCombinedTrips.length);
                                   //Logger.log(DataTripTicket); 
                                   for (var w = 2; w < DataTripTicket.length; w++) {
                                      var TripTicketID = DataTripTicket[w][0];

                                      if (TripTicketID == TripTicketCombined) {
                                      var rowW = w+1;
                                      if (DataCombinedTrips[v][8] instanceof Date){ 
                                      var PickupTime =  Utilities.formatDate(DataCombinedTrips[v][8], "GMT+8","hh:mm:ss a")  
                                      } else { var PickupTime = DataCombinedTrips[v][8]}
                                      sheetTripTicket.getRange(rowW, 9).setValue(PickupTime);                                        //PickupTime
                                      sheetTripTicket.getRange(rowW, 10).setValue(DataCombinedTrips[v][9]);                          //PickupPlace 
                                      sheetTripTicket.getRange(rowW, 11).setValue(DataCombinedTrips[v][10]);                         //OtherInstructions
                                      if (DataCombinedTrips[v][8] instanceof Date){ 
                                      var PickupTime2 =  Utilities.formatDate(DataCombinedTrips[v][14], "GMT+8","hh:mm:ss a")  
                                      } else { var PickupTime2 = DataCombinedTrips[v][14]}
                                      sheetTripTicket.getRange(rowW, 15).setValue(PickupTime2);                                      //PickupTime2
                                      sheetTripTicket.getRange(rowW, 17).setValue(DataCombinedTrips[v][16]);                         //VehicleID
                                      sheetTripTicket.getRange(rowW, 18).setValue(DataCombinedTrips[v][17]);                         //Driver
                                       if (DataCombinedTrips[v][8] instanceof Date){ 
                                      var DepartTime =  Utilities.formatDate(DataCombinedTrips[v][18], "GMT+8","hh:mm:ss a")  
                                      } else { var DepartTime = DataCombinedTrips[v][18]}
                                      sheetTripTicket.getRange(rowW, 19).setValue(DepartTime);                                       //DepartTime                            
                                      sheetTripTicket.getRange(rowW, 28).setValue(DepartTime);
                                      sheetTripTicket.getRange(rowW, 20).setValue(DataCombinedTrips[v][19]);                         //PlateNumber
                                      }
                                      
                                   }
                                   for (var z = 2; z < DataTripRequest.length; z++) {
                                   var TripTicketIDTripRequest = DataTripRequest[z][21]; //sheetTripRequest.getRange(z, 22).getValue();  
                                   //Logger.log("TripTicket Request: " + TripTicketIDTripRequest);    
                                         if (TripTicketIDTripRequest == TripTicketCombined) {
                                              //sheetTripRequest.getRange(w, 18).setValue(DepartTime[v]);
                                              var DepartTime =  Utilities.formatDate(DataCombinedTrips[v][18], "GMT+8","hh:mm a")                                             
                                              //sheetTripRequest.getRange(z+1, 23).setValue(DepartTime);
                                         }
                                   }
                                } Browser.msgBox("Saving Trip Ticket","Trip Ticket Details updated the with Incomplete Details\\n\\nGenerating PDF Tickets now.",Browser.Buttons.OK); var status = 'print'; //\\nYou may now run Generate PDF Tickets.
                           //----------------------------------------------------------------------------------------------------------------------
                           } else {
                           Browser.msgBox("Saving Trip Ticket","Trip Ticket Not Saved. \\n\\nGenerating PDF Ticket(s) Cancelled",Browser.Buttons.OK)
                           var status = 'cancelled'; 
                           
                           }
                        }
                   
                     BlankVehicle(); //Clear the Blank Vehicle
                     //Logger.log(status); 
                     if (status == 'print') { 
                       SortSheet("CombinedTrips",19)
                       //MergeDocument();
                       GetNewBatchProcess();
                     }
    
}


function SaveTripTicketDetails() { // Saves/Update the Trip Ticket from AssignTable Sheet to TripTicket Sheet details of VehicleID, Driver, Departure Time, Plate Number
     var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');
     var lastRowAssign = sheetAssign.getLastRow(); 
     var countblank = 0;
     var TripTicket = []; var VehicleID = []; var Driver = []; var DepartTime = []; var PlateNo = [];  var PickupTime = []; var PickupPlace = []; var OtherInstructions = []; var PickupTime2 = [];
                 for (var v = 0; v < lastRowAssign-1; v++) {
                          
                          TripTicket[v] = sheetAssign.getRange(v+2, 1).getValue(); 
                          
                          PickupTime[v] = sheetAssign.getRange(v+2, 9).getValue(); 
                          PickupPlace[v] = sheetAssign.getRange(v+2, 10).getValue(); 
                          OtherInstructions[v] = sheetAssign.getRange(v+2, 11).getValue(); 
                          PickupTime2[v] = sheetAssign.getRange(v+2, 15).getValue(); 
                   
                          VehicleID[v] = sheetAssign.getRange(v+2, 17).getValue(); 
                          Driver[v] = sheetAssign.getRange(v+2, 18).getValue(); 
                          DepartTime[v] = sheetAssign.getRange(v+2, 19).getValue(); 
                          PlateNo[v] = sheetAssign.getRange(v+2, 20).getValue(); 
                          
                          if (DepartTime[v] instanceof Date){
                               DepartTime[v] = Utilities.formatDate(new Date(DepartTime[v]), "GMT+8","h:mm a");
                               //Logger.log(v + " " + DepartTime[v]);
                          } else { }
                                 
                               if (VehicleID[v] == "" || Driver[v] == "" || DepartTime[v] == "" || PlateNo[v] == "") { countblank ++; } //Checks for Blank details in Yellow Section
              
                 }
                 //Logger.log("Countblank: " + countblank); //test data
                 if (countblank == 0) { 
                      var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
                      var lastRowTicket = sheetTicket.getLastRow();
                      var sheetTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
                      var LastRowTripRequest = sheetTripRequest.getLastRow();
                      
                      for (var v = 0; v < lastRowAssign+1; v++) {
                          for (var w = 3; w < lastRowTicket+1; w++) {
                              var TripTicketID = sheetTicket.getRange(w, 1).getValue(); 
                              if (TripTicketID == TripTicket[v]) {
                                
                                   sheetTicket.getRange(w, 9).setValue(PickupTime[v]); 
                                   sheetTicket.getRange(w, 10).setValue(PickupPlace[v]); 
                                   sheetTicket.getRange(w, 11).setValue(OtherInstructions[v]); 
                                   sheetTicket.getRange(w, 15).setValue(PickupTime2[v]); 
                                
                                             sheetTicket.getRange(w, 17).setValue(VehicleID[v]); 
                                             sheetTicket.getRange(w, 18).setValue(Driver[v]); 
                                             sheetTicket.getRange(w, 19).setValue(DepartTime[v]); 
                                             sheetTicket.getRange(w, 20).setValue(PlateNo[v]);
                              }
                          }
                                   for (var z = 3; z < LastRowTripRequest+1; z++) {
                                   var TripTicketIDTripRequest = sheetTripRequest.getRange(z, 22).getValue();  
                                   //Logger.log("TripTicket Request: " + TripTicketIDTripRequest);    
                                         if (TripTicketIDTripRequest == TripTicket[v]) {
                                              //sheetTripRequest.getRange(w, 18).setValue(DepartTime[v]);
                                              var rangeTripReuest = sheetTripRequest.getRange(z, 22);
                                              rangeTripReuest.offset(0, 1).setValue(DepartTime[v])
                                         }
                                   }
                      }  var status = 'print'; //You may now run Generate PDF Tickets Browser.msgBox("Saving Trip Ticket","Trip Ticket Updated the following details: \\nVehicleID, Driver\\n\\nGenerating PDF Tickets now.",Browser.Buttons.OK);
                 } else { if (Browser.msgBox("Saving Trip Ticket","Incomplete Vehicle, Driver details. \\nDo you want to save the Trip Ticket with blank details? ", Browser.Buttons.YES_NO) == 'yes') {
                           //----------------------------------------------------------------------------------------------------------------------   
                                var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
                                var sheetTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
                                var LastRowTripRequest = sheetTripRequest.getLastRow();
                                var lastRowTicket = sheetTicket.getLastRow();
                                for (var v = 0; v < lastRowAssign+1; v++) {
                                    for (var w = 3; w < lastRowTicket+1; w++) {
                                        var TripTicketID = sheetTicket.getRange(w, 1).getValue(); 
                                        if (TripTicketID == TripTicket[v]) {
                                             sheetTicket.getRange(w, 9).setValue(PickupTime[v]); 
                                             sheetTicket.getRange(w, 10).setValue(PickupPlace[v]); 
                                             sheetTicket.getRange(w, 11).setValue(OtherInstructions[v]); 
                                             sheetTicket.getRange(w, 15).setValue(PickupTime2[v]); 
                                             
                                          
                                             sheetTicket.getRange(w, 17).setValue(VehicleID[v]); 
                                             sheetTicket.getRange(w, 18).setValue(Driver[v]); 
                                             sheetTicket.getRange(w, 19).setValue(DepartTime[v]); 
                                             sheetTicket.getRange(w, 20).setValue(PlateNo[v]);
                                        }
                                    }
                                             for (var z = 3; z < LastRowTripRequest+1; z++) {
                                             var TripTicketIDTripRequest = sheetTripRequest.getRange(z, 22).getValue();  
                                             //Logger.log("TripTicket Request: " + TripTicketIDTripRequest);    
                                                   if (TripTicketIDTripRequest == TripTicket[v]) {
                                                        //sheetTripRequest.getRange(w, 18).setValue(DepartTime[v]);
                                                        var rangeTripReuest = sheetTripRequest.getRange(z, 22);
                                                        rangeTripReuest.offset(0, 1).setValue(DepartTime[v])
                                                   }
                                             }
                                } Browser.msgBox("Saving Trip Ticket","Trip Ticket Details updated the with Incomplete Details.",Browser.Buttons.OK); //var status = 'print'; //\\nYou may now run Generate PDF Tickets.
                           //----------------------------------------------------------------------------------------------------------------------
                           } else {
                           Browser.msgBox("Saving Trip Ticket","Trip Ticket Not Saved.",Browser.Buttons.OK)
                           //var status = 'cancelled'; 
                           }
                        }
                     BlankVehicle(); //Clear the Blank Vehicle
}

function ListTripTicket(PickUpDate) {
      
      var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket'); 
      var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');
      sheetAssign.clearContents()
      var AssignHeader = ["TripTicket","RequestNo","DateTrip","PR Number","BUS","Requestor","OU","Passenger","Pickup Time","Pickup Place","Other Instructions","Destination","Nature","Trips","Primary PickupTime","User Log Email","VehicleID","Driver","Departure Time","PlateNo","PDF-ID"];
      sheetAssign.appendRow(AssignHeader)
      var lastRowTicket = sheetTicket.getLastRow();
      var countMatchDates = 0; var numNew = 0; 
      var getdata = []; var TripTicketNo = [];
      //var PickUpDate = Date(PickUpDate);
         //var collectdata = [][];
          for (var x = 3; x < lastRowTicket+1; x++) {
                   var PickUpDateTT = sheetTicket.getRange(x, 3).getValue();     
                   var PickUpDateTT = Utilities.formatDate(new Date(PickUpDateTT), "GMT+8","M/d/YYYY");
                   //Logger.log(PickUpDateTT);
                   if (PickUpDate == PickUpDateTT) {
                        countMatchDates++;
                        for (var y = 1; y < 22; y++) {
                        getdata[y-1] = sheetTicket.getRange(x, y).getValue();
                          
                        }  
                         
                        var PDFID = getdata[20];
                        if (PDFID != '') {
                        var PDFName = sheetTicket.getRange(x, 23).getValue();
                        var PDFURL = DriveApp.getFileById(PDFID).getUrl(); 
                        getdata[20] ='=hyperlink("' + PDFURL + '", "' + PDFName + '")';
                        }
                     
                        //Logger.log(PDFID + " " + PDFName + " " + PDFlink);
                        //getdata.push(PDFlink);
                        //Logger.log(getdata[18]);
                        //Logger.log(getdata);
                        sheetAssign.appendRow(getdata)
                        //sheetAssign.appendRow(TripLine);
                        numNew++;
                        TripTicketNo[numNew-1] = getdata[0];
                        var ReferenceID = TripTicketNo.join(" / ");
                        
                   } 
          }
          if (numNew != 0) {} else {var ReferenceID  = 'None';}
          var sheetFileID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FileIDList');
          var LogSheetID = sheetFileID.getRange(5,2).getValue(); 
          var LogSheetTable = SpreadsheetApp.openById(LogSheetID).getSheetByName('Event_Logs');
          var logemail = Session.getActiveUser().getEmail();
          var DateTime = Utilities.formatDate(new Date, "GMT+8","YYYYMMdd h:mm:ss a");
          var Logs = [logemail,DateTime,'Select Trip Ticket Day',numNew +' Selected Trip Tricket for date: ' + PickUpDate,'Trip Request',ReferenceID];
          LogSheetTable.appendRow(Logs);
    
          if (countMatchDates == 0) { Browser.msgBox("No Trip Ticket Date Found");} 
          else { 
          FomatPlainText("CombinedTrips",4);
          FormatTime("CombinedTrips",9)
          FormatTime("CombinedTrips",15) 
          FormatTime("CombinedTrips",19)  
          Browser.msgBox("Trip Ticket List", countMatchDates + " Trip Ticket Dates Found", Browser.Buttons.OK); 
          //var range = sheetAssign.getRange("02")
          sheetAssign.setActiveSelection("O2");     
          }
}



function runtemp(){
var ConvertedDate = '5/16/2017';  
  
CombineMergeTrip(ConvertedDate)

}


function CombineMergeTrip(ConvertedDate) {
    
    //var ConvertedDate = '6/18/2016';
    //var ConvertedDate = Utilities.formatDate(TripDateMerge, "GMT+8","M/d/YYYY");
    //var PickupDateSort = ConvertedDate; 
    //Logger.log('Additional Status: ',Additional)
    var sheetSorted = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoSort');  // Source of Sheet for merging is AUTOSORT Sheet
    var lastRowSort = sheetSorted.getLastRow();
    var YearToday = Utilities.formatDate(new Date(), "GMT+8","YYYY");
    var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  // Source of Sheet for merging is AUTOSORT Sheet
    var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');
    sheetAssign.clearContents();
    var AssignHeader = ["TripTicket","RequestNo","DateTrip","PR Number","BUS","Requestor","OU","Passenger","Pickup Time","Pickup Place","Other Instructions","Destination","Nature","Trips","VehicleID","Driver","Departure Time","PlateNo"];
    sheetAssign.appendRow(AssignHeader)
    var lastRowMerge = sheetTicket.getLastRow();
    var numNew = 0;  
    var ReqNo1 = []; var PR1 = []; var BUS1 = [];  var REQ1 = []; var OU1 = []; var PASS1 = []; 
    var PickUpTime1 = []; var Instructions1 = []; var PickUpPlace1 = []; var Dest1 = []; var Nature1 = [];
    var TripTicketNo = [];
    //var log = [][];
    //Logger.log("TripDate: " + ConvertedDate);
    //Logger.log("AutoSort Rows: " + (lastRowSort-1));
    var countprint = 0;
    var tripnum = 0;
    for (var i = 2; i < lastRowSort+1; i++) {
    var CombineNum = sheetSorted.getRange(i, 14).getValue();
    //Logger.log("i=" + i);  
    //Logger.log(CombineNum);
              if (CombineNum == 0) { // Zero for Solo Trips 
                  var sheetSorted = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoSort');   // Source of Sheet for merging is AUTOSORT Sheet
                  var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  // Destination Sheet for Merging
                  var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');  
                  var lastRowMerge = sheetTicket.getLastRow();
                  var sheetAutoNum = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');
                  var newTripTicket = Number(sheetAutoNum.getRange(2, 2).getValue());                   // get the New Trip Ticket Number from the Bucket Holder Sheet AutoNumber
                  var TripTicket = "TRIP" + YearToday + (Number(500000) + Number(newTripTicket));        // autonumber ticketing
                  newTripTicket++;
                  sheetAutoNum.getRange(2, 2).setValue(newTripTicket);                                  //Update the Next Trip Ticket Number
                  var RequestNo = String(sheetSorted.getRange(i, 13).getValue());                       // RequestNo Log
                  var PR = String(sheetSorted.getRange(i, 1).getValue());                               //PR Number
                  var BUS = String(sheetSorted.getRange(i, 2).getValue());                              //BUS
                  var Requestor = String(sheetSorted.getRange(i, 3).getValue());                        //Requestor
                  var OU = String(sheetSorted.getRange(i, 4).getValue());                               //OU
                  var Passenger = String(sheetSorted.getRange(i, 5).getValue());                        //Passenger
                  var PickUpDate = String(sheetSorted.getRange(i, 6).getValue());                       //PickupDate
                  var PickUpTime = sheetSorted.getRange(i, 7).getValue();                               //Pickuptime
                  var PickUpTime = Utilities.formatDate(PickUpTime, "GMT+8","h:mm a");
                  var ReturnDate = String(sheetSorted.getRange(i, 8).getValue());                       //Return Date
                  var PickUpLocation = String(sheetSorted.getRange(i, 9).getValue());                   //Pick up Place 
                  var Destination = String(sheetSorted.getRange(i, 10).getValue());                     //Destination     
                  var Instruction = String(sheetSorted.getRange(i, 11).getValue());                     //Instruction
                  var Nature = String(sheetSorted.getRange(i, 12).getValue());                          //Nature of Trip
                  var UserLogEmail = Session.getActiveUser().getEmail();
                  //          generated         13       data      1    2    3      4        5         6              7         11          10
                  var values = [TripTicket,RequestNo,ConvertedDate,PR,BUS,Requestor,OU,Passenger,PickUpTime,PickUpLocation,Instruction,Destination,Nature,"1",PickUpTime,UserLogEmail,"","","","","","","","","","","Ongoing",PickUpTime];
                  var requestNoUpdate = RequestNo;
                  //Logger.log("RequestNumber: " + "Solo" + " " + RequestNo + " TripTicket: "+ TripTicket);
                  //Logger.log(values);
                  updateRequestTrip(requestNoUpdate,TripTicket);
                  updateCombineTrip(requestNoUpdate,TripTicket);
                  //var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
                  //sheet1.appendRow(values);
                  sheetAssign.appendRow(values); 
                  sheetTicket.appendRow(values); 
                  countprint++;  // count print log
                  numNew++;
                  TripTicketNo[numNew-1] = TripTicket;
                  var ReferenceID = TripTicketNo.join(" / ");
              //  Logger.log(TripTicket + " " + RequestNo + " " + PickupDateSort + " " + PR);
              } 
              if (CombineNum > 0) {
                   tripnum++;
                   var sheetSorted = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoSort');
                   var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
                   var CombineNumCurrent = sheetSorted.getRange(i, 14).getValue();     //Get Combine Number from AutoSort
                   var h = i +1; 
                   var CombineNumNext = sheetSorted.getRange(h, 14).getValue();   
                     
                   var ReqNogetmerge = String(sheetSorted.getRange(i, 13).getValue()); //Requestor Mercge
                   var ReqNometmerge = ReqNo1.push(ReqNogetmerge); 
                   var mergeReqNo = ReqNo1.join(" / ");
                
                   var PRgetmerge = String(sheetSorted.getRange(i, 1).getValue()); //getColIndexByName("PR Number StringFilter") for the PR Number
                   
                   var PRmetmerge = PR1.push(PRgetmerge);
                   var TripCount = PRmetmerge;
                   var mergePR = PR1.join(" / ");
          
                   var BUSgetmerge = String(sheetSorted.getRange(i, 2).getValue()); //BUS
                   var BUSmetmerge = BUS1.push(BUSgetmerge);
                   var mergeBUS = BUS1.join(" / ");
          
                   var REQgetmerge = String(sheetSorted.getRange(i, 3).getValue()); //Requestor
                   var REQmetmerge = REQ1.push(REQgetmerge);
                   var mergeREQ = REQ1.join(" / ");
          
                   var OUgetmerge = String(sheetSorted.getRange(i, 4).getValue()); //OU
                   var OUmetmerge = OU1.push(OUgetmerge);
                   var mergeOU = OU1.join(" / ");
          
                   var PASSgetmerge = String(sheetSorted.getRange(i, 5).getValue()); //Passenger
                   var PASSmetmerge = PASS1.push(PASSgetmerge);
                   var mergePASS = PASS1.join(" / ");
                   //Logger.log(String(sheetSorted.getRange(i, 5).getValue()))
                 
                   var PickUpTimegetmerge = sheetSorted.getRange(i, 7).getValue(); //Pickuptime
                   var ConvertTime = Utilities.formatDate(PickUpTimegetmerge, "GMT+8","h:mm a");         // convert the time properly        
                   var PickUpTimemetmerge = PickUpTime1.push(ConvertTime);
                   var mergePickUpTime = PickUpTime1.join(" / ");
              
                   var PickUpPlacegetmerge = String(sheetSorted.getRange(i, 9).getValue()); //Pick up Place
                   var PickUpPlacemetmerge = PickUpPlace1.push(PickUpPlacegetmerge);
                   var mergePickUpPlace = PickUpPlace1.join(" / ");
                   
                   var Destgetmerge = String(sheetSorted.getRange(i, 10).getValue()); //Destination
                   var Destmetmerge = Dest1.push(Destgetmerge);
                   var mergeDest = Dest1.join(" / ");
                   
                   var Instructionsgetmerge = String(sheetSorted.getRange(i, 11).getValue()); //Pick up Place
                   var Instructionsmetmerge = Instructions1.push(Instructionsgetmerge);
                   var mergeInstructions = Instructions1.join(" / ");
                
                   var Naturegetmerge = String(sheetSorted.getRange(i, 12).getValue()); //Pick up Place
                   var Naturemetmerge = Nature1.push(Naturegetmerge);
                   var mergeNature = Nature1.join(" / ");
                   //Logger.log("Current: " + CombineNumCurrent);
                   //Logger.log("Next: " + CombineNumNext);
                   if (CombineNumCurrent == CombineNumNext) { } 
                   else {
                             //Logger.log("Print Now the Merge Values");
                             var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  // Destination Sheet for Merging
                             var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');  // Destination Sheet for Merging
                             var lastRowMerge = sheetTicket.getLastRow();
                             var sheetAutoNum = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');
                             var newTripTicket = Number(sheetAutoNum.getRange(2, 2).getValue());                   // get the New Trip Ticket Number from the Bucket Holder Sheet AutoNumber
                             var TripTicket = "TRIP" + YearToday + (Number(500000) + Number(newTripTicket));        // autonumber ticketing
                             newTripTicket++;
                             sheetAutoNum.getRange(2, 2).setValue(newTripTicket); 
                             //var TripTicket = "TRIP" + YearToday + (Number(50000) + Number(newTripTicket));
                             //var values = [mergeReqNo,mergeBUS];
                             //          generated         13       data      1    2    3      4        5         6              7         11          10
                             //var values = [TripTicket,mergeReqNo,ConvertedDate,mergePR,mergeBUS,mergeREQ,mergeOU,mergePASS,mergePickUpTime,mergePickUpPlace,mergeInstructions,mergeDest,mergeNature,TripCount];
                             var UserLogEmail = Session.getActiveUser().getEmail();
                             var values = [TripTicket,mergeReqNo,ConvertedDate,mergePR,mergeBUS,mergeREQ,mergeOU,mergePASS,mergePickUpTime,mergePickUpPlace,mergeInstructions,mergeDest,mergeNature,TripCount,PickUpTime1[0],UserLogEmail,"","","","","","","","","","","Ongoing",PickUpTime];
                             sheetAssign.appendRow(values); 
                             sheetTicket.appendRow(values); 
                                  for (var k = 0; k < ReqNometmerge; k++) {
                                    var requestNoUpdate = ReqNo1[k];
                                    //Logger.log("RequestNumber " + k + "Merged " + ReqNo1[k] + " TripTicket: "+ TripTicket);
                                    updateRequestTrip(requestNoUpdate,TripTicket);
                                    updateCombineTrip(requestNoUpdate,TripTicket);
                                  }
                             numNew++;
                             TripTicketNo[numNew-1] = TripTicket;
                             var ReferenceID = TripTicketNo.join(" / ");
                             PRmetmerge = 0;
                             var ReqNo1 = []; 
                             var PR1 = []; var BUS1 = [];
                             var REQ1 = [];  
                             var OU1 = []; 
                             var PASS1 = [];  
                             var PickUpTime1 = []; 
                             var Instructions1 = [];
                             var PickUpPlace1 = [];    
                             var Dest1 = []; var Nature1 = []; 
                             countprint++;  //count print log
                     
                            
                    }
              }
          //count the printed rows        
    }
    //end of for loop
  
    //var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
    //var lastRowprint = sheetTicket.getLastRow();
    //var FocusRow = (Number(lastRowprint)-Number(countprint-1));
    //var countprint = 0;   
    //var rangerow = 'O' + FocusRow;
    //sheetTicket.setActiveSelection(rangerow);

    if (numNew != 0) {} else {var ReferenceID  = 'None';}
    var sheetFileID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FileIDList');
    var LogSheetID = sheetFileID.getRange(5,2).getValue(); 
    var LogSheetTable = SpreadsheetApp.openById(LogSheetID).getSheetByName('Event_Logs');
    var logemail = Session.getActiveUser().getEmail();
    var DateTime = Utilities.formatDate(new Date, "GMT+8","YYYYMMdd h:mm:ss a");
    var Logs = [logemail,DateTime,'Generate Trip Ticket',numNew +' Generated Trip Ticket(s) for date: ' + ConvertedDate,'Trip Request',ReferenceID];
    LogSheetTable.appendRow(Logs);
    var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');
    var lastRowprint = sheetAssign.getLastRow();
    var FocusRow = (Number(lastRowprint)-Number(countprint-1));
    var countprint = 0;   
    var rangerow = 'O' + FocusRow;
    sheetAssign.setActiveSelection(rangerow);  // Focus the Assign Vehicle
    //Browser.msgBox("Assign the Vehicle and Driver");
    Browser.msgBox("Trip Tickets created from Selected Trip Requests. \\n\\nEnter values for vehicle, driver and departure time.");
}

function updateRequestTrip(requestNoUpdate,TripTicket) {
            var sheetTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
            var DatasheetTripRequest = sheetTripRequest.getRange(1,17,sheetTripRequest.getLastRow(),8).getValues();
    
            var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  // Searches for the existint TripTicket from TripRequest and Deletes the Row
            var DatasheetTripTicket = sheetTripRequest.getRange(1,1,sheetTripTicket.getLastRow(),1).getValues();
   
            for (var j = 2; j < DatasheetTripRequest.length; j++) {
            var RequestNoDest = String(DatasheetTripRequest[j][0]);  
                 if (RequestNoDest == requestNoUpdate) {
                 var status = "";  
                 var rowUpdate = j+1; 
                 sheetTripRequest.getRange(rowUpdate, 21).setValue(status); 
                 Logger.log(requestNoUpdate+" -- "+ rowUpdate); 
                 var TripTicketDest = String(DatasheetTripRequest[j][5]);
                     if (TripTicketDest == ""){
                     sheetTripRequest.getRange(rowUpdate, 22).setValue(TripTicket);     
                     } else { 
                       Logger.log("Delete the Row: " + TripTicketDest);
                       for (var k = 2; k < DatasheetTripTicket.length;k++){
                       var TripTicketDelete = DatasheetTripTicket[k][0];
                       var TripTicketRowDelete = k+1; 
                            if (TripTicketDest == TripTicketDelete) {
                               sheetTripTicket.deleteRow(TripTicketRowDelete);
                               sheetTripRequest.getRange(rowUpdate, 22).setValue(TripTicket);
                            }  
                       }
                     }
                 }
            }
}

function updateCombineTrip(requestNoUpdate,TripTicket) {
            var sheetcomb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests');
            var DataSelectedTripRequest = sheetcomb.getRange(1,13,sheetcomb.getLastRow(),1).getValues();
            
            for (var l = 1; l < DataSelectedTripRequest.length; l++) {
            var row = l+1;   
            var RequestNoDest = String(DataSelectedTripRequest[l][0]);
                  if (RequestNoDest == requestNoUpdate) {
                       var status = "";
                       //sheetcomb.getRange(l, 16).setValue(status);
                       sheetcomb.getRange(row, 17).setValue(TripTicket);
                  }
            }
}
function updateCombineTripAdditional(requestNoUpdate,TripTicket) {
            var sheetcomb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests');
            var lastRowE = sheetcomb.getLastRow(); 
            for (var l = 2; l < lastRowE+1; l++) {
            var RequestNoDest = sheetcomb.getRange(l, 13).getValue();
                  if (RequestNoDest == requestNoUpdate) {
                       var status = "";
                       //sheetcomb.getRange(l, 16).setValue(status);
                       sheetcomb.getRange(l, 18).setValue(TripTicket);
                  }
            }
}



function checkvalue() {
   var sheetAutoNum = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');
   var value1 = sheetAutoNum.getRange(1,2).getValue();
   Logger.log(value1);
}



function VehicleUpdate() {
  var range = SpreadsheetApp.getActiveRange(), data = range.getValues();
  var output = []; 
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet().getName();
 //Logger.log(data[0][0]);
            Logger.log(sheet);
            var sheetvehicle = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VehicleList');
            var lastRowV = sheetvehicle.getLastRow(); 
            //Logger.log(lastRowV);
            for (var o = 2; o < lastRowV+1; o++) {
            var vehicleID = sheetvehicle.getRange(o, 1).getValue();
                  if (vehicleID == data[0][0]) {
                    //Logger.log('Vehicle ID Found');
                    var PlateNum =  sheetvehicle.getRange(o, 2).getValue(); 
                    output.push([PlateNum]);
                    range.offset(0,3).setValues(output);
                  }
               
            }
}

function BlankVehicle() {  //Clears PlateNumbers with Blank Vehicle ID in AssignTicket sheet
        var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');  
        var lastRowAssign = sheetAssign.getLastRow();
        for (var f = 2; f < lastRowAssign+1; f++) {
        var VehicleIDAssign = sheetAssign.getRange(f, 15).getValue(); 
               if (VehicleIDAssign == '') {
               sheetAssign.getRange(f,18).setValue('');
               }
        }
}

function BlankVehicle2() {
        var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  
        var lastRowAssign = sheetTripTicket.getLastRow();
        for (var f = 2; f < lastRowAssign+1; f++) {
        var VehicleIDAssign = sheetTripTicket.getRange(f, 15).getValue(); 
               if (VehicleIDAssign == '') {
               sheetTripTicket.getRange(f,18).setValue('');
               }
        }
}

function FomatPlainText(SheetName,Column){
     // Begin Try catch error
     try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));
     // Function Begins here
      var sheetReference = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
      var MaxRows = sheetReference.getMaxRows();
      sheetReference.getRange(1, Column, MaxRows).setNumberFormat('@STRING@');
     //Catch Error
     }
     catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
     } 
     //End Catch Error        
}

function SortSheet(SheetName,Column){
     // Begin Try catch error
     try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));
     // Function Begins here   
      var sheetReference = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
      var MaxRows = sheetReference.getLastRow()
      var MaxColumn = sheetReference.getLastColumn(); 
      sheetReference.insertRowsAfter(sheetReference.getMaxRows(), 1);
      var range = sheetReference.getRange(2, 1, MaxRows, MaxColumn)
      range.sort({column: Column, ascending: true})
     //Catch Error
     }
     catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {
         htmlBody: "Function Name: "+funcName+
                   "<br>Filename:  "+e.fileName+
                   "<br>Message:   "+e.message+
                   "<br>Line no:   "+e.lineNumber+
                   "<br>SheetName: "+e.Sheetname})
     } 
     //End Catch Error 
}

function FormatTime(SheetName,Column){
     // Begin Try catch error
     try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));
     // Function Begins here   
  
     var sheetReference = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
     var MaxRows = sheetReference.getMaxRows();
     sheetReference.getRange(1, Column, MaxRows).setNumberFormat("h:mm AM/PM");

     //Catch Error
     }
     catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
     } 
     //End Catch Error 
}

function FomatNumber(SheetName,Column){
     // Begin Try catch error
     try{
           var funcName = arguments.callee.toString();
           funcName = funcName.substr('function '.length);
           funcName = funcName.substr(0, funcName.indexOf('('));
     // Function Begins here   
      var sheetReference = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
      var MaxRows = sheetReference.getMaxRows();
      sheetReference.getRange(1, Column, MaxRows).setNumberFormat("0");
     //Catch Error
     }
     catch(e){
         MailApp.sendEmail('m.delrosario@irri.org', 'TS Dispatch Error', '', {htmlBody: "Function Name: "+funcName+"<br>Filename: "+e.fileName+"<br> Message: "+e.message+"<br> Line no: "+e.lineNumber})
     } 
     //End Catch Error 
}







