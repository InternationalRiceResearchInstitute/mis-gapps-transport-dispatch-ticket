function UpdateSelectedTripRequestToTripRequest() {
    var sheetSelectedTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests');
    var sheetTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');  
    var lastRowSelectedTripRequest = sheetSelectedTripRequest.getLastRow();
    var lastRowTripRequest = sheetTripRequest.getLastRow();
 
    var data1 = sheetSelectedTripRequest.getRange(1,13,lastRowSelectedTripRequest,6).getValues(); 
    var data2 = sheetTripRequest.getRange(3, 17, lastRowTripRequest).getValues();   
    // transfer data1 to data2
    //get matching rows in data2
    //Logger.log(data2.length)
    for(var x = 1; x < lastRowSelectedTripRequest;x++){
    //var index = binarySearch(data, data1[x][0],0)
    //Logger.log(data1[x][0]);
    
    var index = binarySearch(data2, data1[x][0],0);
    var row = index +3;
    Logger.log(data1[x][0]+ " can be found in row "+row);
    sheetTripRequest.getRange(row, 19).setValue(data1[x][1]); //updates combine
    sheetTripRequest.getRange(row, 20).setValue(data1[x][2]); //updates combine
    sheetTripRequest.getRange(row, 22).setValue(data1[x][4]); //updates combine
    sheetTripRequest.getRange(row, 23).setValue("");
    sheetTripRequest.getRange(row, 24).setValue("");
    //sheetTripRequest.getRange(row, 26).setValue(data1[x][1]); //updates combine
    //sheetTripRequest.getRange(row, 27).setValue(data1[x][2]); //updates combine
    //sheetTripRequest.getRange(row, 28).setValue(data1[x][4]); //updates combine  
    
    //var start = new Date();
    }
}

function UpdateTripTicketToTripRequest() {
    var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
    var sheetTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');  
    var lastRowTripTicket = sheetTripTicket.getLastRow()
    var lastRowTripRequest = sheetTripRequest.getLastRow();
 
    var data1 = sheetTripTicket.getRange("A7:V62").getValues(); 
    //Logger.log(data1[0]);
    var data2 = sheetTripRequest.getRange("V3:V871").getValues();  
    //var data2 = sheetTripRequest.getRange(3, 17, lastRowTripRequest).getValues();   
    // transfer data1 to data2
    //get matching rows in data2
    for(var x = 7; x<63;x++){
    var Tripdata = sheetTripTicket.getRange(x, 1, 1, 22).getValues();
    Logger.log(Tripdata[0][0]+Tripdata[0][18])   
    var ConvertTime = Utilities.formatDate(Tripdata[0][18], "GMT+8","h:mm a");
    //Logger.log(Tripdata[0][21]);
    
    var PDFHyperLink ='<b><a href="'+Tripdata[0][21]+'" style="text-decoration:none;background-color:transparent" target="_blank">ⓘ</a></b>';
    //Logger.log(PDFHyperLink);
    Logger.log(ConvertTime);
    Logger.log(PDFHyperLink);
       for(var y = 0; y<data2.length;y++){
    //var row = index + 3; 
           if(data2[y]==Tripdata[0][0]){
           var row = y+3;
           Logger.log(Tripdata[0][0]+ " can be found in row "+row);
           //sheetTripTicket.getRange(row,23).setValue(ConvertTime);
           //sheetTripTicket.getRange(row,24).setValue(PDFHyperLink);
           sheetTripRequest.getRange(row,23).setValue(ConvertTime);
           sheetTripRequest.getRange(row,24).setValue(PDFHyperLink);
           }
       }  
    }
    
}


function binarySearch(list, item,column) {
    var min = 0;
    var max = list.length - 1;
    var guess;
    var column = column || 0
    while (min <= max) {
        guess = Math.floor((min + max) / 2);
        if (list[guess][column] === item) {
            return guess;
        }
        else {
            if (list[guess][column] < item) {
                min = guess + 1;
            }
            else {
                max = guess - 1;
            }
        }
    }
    return -1;
}


function splittoarray() {
var data = '2017100687 / 2017100647';
var ss = [];
ss= data.split(" / ");
Logger.log(ss)
}

function ListTripRequestOld(PickupDate){
   
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest') 
   var lastRow = sheet.getLastRow()
   var sheetpool = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests');
   var headers =["PR Number","BUS Code","PR Requestor","OU","Passenger(s)","Pickup Date","Pick Up Time","Return Date","Pick Up Location","Point of Destination","Other Instructions","Nature of Trip","RequestNo","Combine Num","Sequence","Additional","Trip Ticket","Depart Time"]
   sheetpool.clearContents();
   
   sheetpool.appendRow(headers);
   sheetpool.setFrozenRows(1);    
   var numNew = 0; 
   //Browser.msgBox(PickupDate);
   var TripRequestNo = []; 
   for (var i = 3; i < lastRow+1; i++) {
      var CheckPickUpDate = sheet.getRange(i, 8).getValue(); //Pickup Date  
      var Status = sheet.getRange(i, 15).getValue();
      var CompareDate = Utilities.formatDate(CheckPickUpDate, "GMT+8","M/d/YYYY");
      //Logger.log(Status);
      //Browser.msgBox(CompareDate);
      //var display = i + " " + CheckPickUpDate;
     // var testdata = Utilities.formatDate(PickupDate, "GMT+8","M/d/YYYY");
      if (CompareDate == PickupDate && (Status =='Ongoing')) {  
         numNew++; 
        
         var PRNumber = sheet.getRange(i, 3).getValue();          //  PR Number	
         var BUSCode = sheet.getRange(i, 4).getValue();           //  PR Number
         var Requestor = sheet.getRange(i, 5).getValue();         //  PR Requestor	
         var OU = sheet.getRange(i, 6).getValue();                //  PR Requestor	
         var Passenger = sheet.getRange(i, 7).getValue();         //  Passenger(s)	
         var PickUpDate = sheet.getRange(i, 8).getValue();        //  Pickup Date	
         var PickUpTime = sheet.getRange(i, 9).getValue();        //  Pick Up Time	
         var ReturnDate = sheet.getRange(i, 10).getValue();        //  Pick Up Date	
         var PickupLocation = sheet.getRange(i, 11).getValue();   //  Pick Up Location 	
         var Destination = sheet.getRange(i, 12).getValue();      //  Point of Destination
         var Instructions = sheet.getRange(i, 13).getValue();     //  Other Instructions	
         var Nature = sheet.getRange(i, 14).getValue();           //  Nature
         var Username = sheet.getRange(i, 15).getValue();         //  Nature
         var FormLink = sheet.getRange(i, 16).getValue();         //  FormLink
         var RequestNo = String(sheet.getRange(i, 17).getValue());        //  RequestNo     
         TripRequestNo[numNew-1] = RequestNo;
         var ReferenceID = TripRequestNo.join(" / ");
         var DateEntered = sheet.getRange(i, 18).getValue();      //  DateEntered
         var CombineNum = sheet.getRange(i, 19).getValue();       //  CombineNum
         var Sequence = sheet.getRange(i, 20).getValue();         //  Sequence
         var Additional = sheet.getRange(i, 21).getValue();       //  Additional
         var TripTicket = sheet.getRange(i, 22).getValue();       //  TripTicket  
         var DepartTime = sheet.getRange(i, 23).getValue();       //  Depart Time
         
         var values = [PRNumber,BUSCode,Requestor,OU,Passenger,PickUpDate,PickUpTime,ReturnDate,PickupLocation,Destination,Instructions,Nature,RequestNo,CombineNum,Sequence,Additional,TripTicket,DepartTime]

         sheetpool.appendRow(values); 
     }
  }  
  
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
}

// Implemented 5/17/2017 as replacement for ListTripRequest

function ListTripRequestFIX(PickupDate){
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
      for (var x = 2; x<DataTripRequestList.length;x++){
           
          var CompareDate = Utilities.formatDate(DataTripRequestList[x][5] , "GMT+8","M/d/YYYY");
          var PickUpTime = Utilities.formatDate(DataTripRequestList[x][6], "GMT+8","hh:mm a") 
          var ReturnDate = Utilities.formatDate(DataTripRequestList[x][7], "GMT+8","M/d/YYYY");
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
    Logger.log(ReferenceID);
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

//Fix Implemented 5/23/2017
// formerly function updateRequestTrip(requestNoUpdate,TripTicket) {
function updateTripRequestOldandSlow(requestNoUpdate,TripTicket) {
            var sheetdestination = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
            var lastRowD = sheetdestination.getLastRow(); 
            for (var j = 2; j < lastRowD+1; j++) {
            var RequestNoDest = sheetdestination.getRange(j, 17).getValue();
                  if (RequestNoDest == requestNoUpdate) {
                       var status = "";
                       sheetdestination.getRange(j, 21).setValue(status); 
                       //sheetdestination.getRange(j, 22).setValue(TripTicket);
                       //------------------------------------------------------------------------------------------------------
                        var TripTicketDest = sheetdestination.getRange(j, 22).getValue(); 
                        Logger.log("TripTicketDest: " + TripTicketDest);
                        if (TripTicketDest == "") {                                                                 // Checks if Trip Request already has TripTicket
                        sheetdestination.getRange(j, 22).setValue(TripTicket);                    
                        } else {  Logger.log("Delete the Row: " + TripTicketDest);
                        var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  // Searches for the existint TripTicket from TripRequest and Deletes the Row
                        var LastRowTripTicket = sheetTripTicket.getLastRow();
                              for (var g = 3; g < LastRowTripTicket+1; g++) {
                              var TripTicketDelete = sheetTripTicket.getRange(g, 1).getValue();
                                    Logger.log("TripTicketDelete: " + TripTicketDelete);
                                   if (TripTicketDest == TripTicketDelete) {
                                         sheetTripTicket.deleteRow(g);
                                         sheetdestination.getRange(j, 22).setValue(TripTicket);
                                   }
                              }
                        }
                      //------------------------------------------------------------------------------------------------------
                  }
            }
} 

function testUpdateTripRequest(){
var requestNoUpdate ='2017102032';
var TripTicket ='TRIP201753884';
  Logger.log(requestNoUpdate); 
updateTripRequestTripFIX(requestNoUpdate,TripTicket)

}

function updateTripRequestTripFIX(requestNoUpdate,TripTicket){
            var sheetTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
            var DatasheetTripRequest = sheetTripRequest.getRange(1,17,sheetTripRequest.getLastRow(),8).getValues();
    
            var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  // Searches for the existint TripTicket from TripRequest and Deletes the Row
            
   
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
                       var DatasheetTripTicket = []; 
                       var DatasheetTripTicket = sheetTripTicket.getRange(1,1,sheetTripTicket.getLastRow(),1).getValues();
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

//formerly function updateCombineTrip(requestNoUpdate,TripTicket) {
function updateCombineTripOldAndSlow(requestNoUpdate,TripTicket) {  
            var sheetcomb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests');
            var lastRowE = sheetcomb.getLastRow(); 
            for (var l = 2; l < lastRowE+1; l++) {
            var RequestNoDest = sheetcomb.getRange(l, 13).getValue();
                  if (RequestNoDest == requestNoUpdate) {
                       var status = "";
                       //sheetcomb.getRange(l, 16).setValue(status);
                       sheetcomb.getRange(l, 17).setValue(TripTicket);
                  }
            }
}

function updateCombineTripFix(requestNoUpdate,TripTicket) {  
            var sheetcomb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests');
            var DataSelectedTripRequest = sheetcomb.getRange(1,13,sheetcomb.getLastRow(),1).getValues();
            
            for (var l = 2; l < DataSelectedTripRequest.length; l++) {
            var row = l+1;   
            var RequestNoDest = DataSelectedTripRequest[l][0];
                  if (RequestNoDest == requestNoUpdate) {
                       var status = "";
                       //sheetcomb.getRange(l, 16).setValue(status);
                       sheetcomb.getRange(row, 17).setValue(TripTicket);
                  }
            }
}
//Fix Implemented 6/1/2017
// formerly function UpdateTripTicket(requestNoUpdate,TripTicket) {
function UpdateTripTicketFix() { // Saves/Update the Trip Ticket from AssignTable Sheet to TripTicket Sheet details of VehicleID, Driver, Departure Time, Plate Number
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

function updateTripTicketTemp(){
                 var sheetCombinedTrips = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');
                 var DataCombinedTrips = sheetCombinedTrips.getRange(1, 1, sheetCombinedTrips.getLastRow(), 20).getValues();
                 var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
                 var DataTripTicket = sheetTripTicket.getRange(1, 1, sheetTripTicket.getLastRow(), 29).getValues();
                 var sheetTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
                 var DataTripRequest = sheetTripRequest.getRange(1, 1, sheetTripRequest.getLastRow(), 25).getValues();
                 //Logger.log(DataTripTicket); 
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
                                }
}