
/*
CODE A.1 replacement    - UpdatePrimaryTable
Replaced date 5/16/2017 
Code Start: 5/15/2017 
Code by:                 Melchor del Rosario
Replaced Function Name:  UpdatePrimaryTable 
Archive Code Name     :  UpdatePrimaryTableArchive   

Reason: Getting slow due to slow update of sheet because of 2 for loops and getvalue per column due  large rows Trip request. 
see reference analysis -- https://docs.google.com/spreadsheets/d/1op9MyWbBz5C_glNaFIj8hbhmrHGqn8RBE98c4NzObeY/edit#gid=0

*/

function UpdatePrimaryTableArchive() { //update Trip Request Number with its updated combination tag
    var sheetupdate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SelectedTripRequests');
    var sheetdestination = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');
    var lastRow = sheetupdate.getLastRow();
    //Logger.log(lastRow);
    for (var i = 2; i < lastRow+1; i++) {
    var RequestNo = sheetupdate.getRange(i, 13).getValue(); 
    var CombineNum = sheetupdate.getRange(i, 14).getValue(); 
    var Sequence   = sheetupdate.getRange(i, 15).getValue();   
    var Additional = sheetupdate.getRange(i, 16).getValue(); 
    var TripTicket = sheetupdate.getRange(i, 17).getValue();    
    var DepartTime = sheetupdate.getRange(i, 18).getValue();  
    //  Logger.log(RequestNo + " " + i + " from CombineTrip");
            var lastRowD = sheetdestination.getLastRow(); 
            for (var j = 2; j < lastRowD+1; j++) {
            var RequestNoDest = sheetdestination.getRange(j, 17).getValue();
                  if (RequestNoDest == RequestNo) {
                        var TripTicketDest = sheetdestination.getRange(j, 22).getValue(); 
                        Logger.log("TripTicketDest: " + TripTicketDest);
                        if (TripTicketDest == "") {                                                                 // Checks if Trip Request already has TripTicket
                        sheetdestination.getRange(j, 19).setValue(CombineNum);
                        sheetdestination.getRange(j, 20).setValue(Sequence);
                        sheetdestination.getRange(j, 21).setValue(Additional);
                        sheetdestination.getRange(j, 22).setValue(TripTicket);
                        sheetdestination.getRange(j, 23).setValue(DepartTime);
                        } else {  Logger.log("Delete the Row: " + TripTicketDest);
                        var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  // Searches for the existing TripTicket from TripRequest and Deletes the Row
                        var LastRowTripTicket = sheetTripTicket.getLastRow();
                              for (var g = 2; g < LastRowTripTicket+1; g++) {
                              var TripTicketDelete = sheetTripTicket.getRange(g, 1);
                                   if (TripTicketDest == TripTicketDelete) {
                                         //sheetTripTicket.deleteRow(g);
                                         sheetdestination.getRange(j, 19).setValue(CombineNum);
                                         sheetdestination.getRange(j, 20).setValue(Sequence);
                                         sheetdestination.getRange(j, 21).setValue(Additional);
                                         sheetdestination.getRange(j, 22).setValue(TripTicket);
                                         sheetdestination.getRange(j, 23).setValue(DepartTime);
                                   }
 
                              }
                        }
                   }
            }
    } 
}

function UpdatePrimaryFIX(){
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
           //Logger.log(DataUpdateTripRequest[i])
           var rowsource = i+1; 
           for(var j = 2; j<DataTripRequestNo.length; j++){
                 if (DataUpdateTripRequest[i][0] == DataTripRequestNo[j]){
                        var rowupdate = j+1;       
                        sheetdestination.getRange(j, 19).setValue(DataUpdateTripRequest[i][1]);
                        sheetdestination.getRange(j, 20).setValue(DataUpdateTripRequest[i][2]);
                        sheetdestination.getRange(j, 21).setValue(DataUpdateTripRequest[i][3]);
                        sheetdestination.getRange(j, 22).setValue(DataUpdateTripRequest[i][4]);
                        sheetdestination.getRange(j, 23).setValue(DataUpdateTripRequest[i][5]);
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
