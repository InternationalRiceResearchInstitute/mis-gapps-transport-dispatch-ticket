
var MainListTable = '1iXor3AQlOzMvHs_CeXpxoOvbi30p96AN2NrkJJFPNa8';
var sheetFileIDSource = '1iXor3AQlOzMvHs_CeXpxoOvbi30p96AN2NrkJJFPNa8';
var EventLogsold = '1gq-3kfz3j_Wax2yMDirHV-wkONUYCnmPlg9oAD9VUBs';
var EventLogs = '1vl_atyiNaSXnQlo5zGmlJYOsI7Tg54QHvu_ZETMxC1c';

function GenerateRegularTrips(){
  var data = []; var ticketrange = [];
  var SheetRegularTrip = SpreadsheetApp.openById(MainListTable).getSheetByName('QuerySort');
  var CountTripTicket = 0;
  var LastRowRegularTrip = SheetRegularTrip.getLastRow();
  var LastColumnRegularTrip = SheetRegularTrip.getLastColumn();
  for (var z = 2; z<LastRowRegularTrip+1;z++){
        for (var y = 0; y < 20; y++){
        data[y] = SheetRegularTrip.getRange(z,y+1).getValue();  
        } 
        
        var YearToday = Utilities.formatDate(new Date(), "GMT+8","YYYY");
        var newTripTicket = 0;
        var sheetAutoNum = SpreadsheetApp.openById(MainListTable).getSheetByName('AutoNumber');
        var newTripTicket = Number(sheetAutoNum.getRange(2, 2).getValue());    
        var TripTicket = "TRIP" + YearToday + (Number(500000) + Number(newTripTicket));
       
        var sheetAutoNum = SpreadsheetApp.openById(MainListTable).getSheetByName('AutoNumber');
        
        var newTripNum = Number(sheetAutoNum.getRange(6, 2).getValue());
        var RequestNo =  YearToday + (Number(1000000) + Number(newTripNum))+"R";     
        //var RequestNo =   YearToday + (Number(100000) + Number(newTripNum));  
        ticketrange[z-2] = TripTicket;
        
        var daysOffset = 1; 
        var date = new Date();
        date.setDate(date.getDate() + daysOffset);
        var DateTrip = Utilities.formatDate(date, "GMT+8","M/d/YYYY");
        var logemail = Session.getActiveUser().getEmail();
        
        var PickupTime = Utilities.formatDate(data[3], "GMT+8","h:mm a");
        Logger.log(PickupTime)
        //var value = [TripTicket,RequestNo,DateTrip,'',data[9],'','',data[7],data[3],'',data[8],data[7],'Regular','R',data[3],logemail,data[2],data[1],data[3],data[19],'','','','','','','Ongoing',data[3],data[4]];
        var value = [TripTicket,RequestNo,DateTrip,'',data[9],'','',data[7],data[3],'',data[8],data[7],'Regular','1',data[3],logemail,data[2],data[1],data[3],data[19],'','','','','','','Ongoing',data[3],data[4]];
        CountTripTicket++;
        newTripNum++;           
        newTripTicket++;   
        sheetAutoNum.getRange(6, 2).setValue(newTripNum);             
        sheetAutoNum.getRange(2, 2).setValue(newTripTicket);
        var sheetTripTicket = SpreadsheetApp.openById(MainListTable).getSheetByName('TripTicket');
        sheetTripTicket.appendRow(value);
        Logger.log(value);
  }
  Browser.msgBox(CountTripTicket + " Regular Trips Generated. \\n\\nYou may now View in the Trip Ticket Commands."); 
  var sheetFileID = SpreadsheetApp.openById(sheetFileIDSource).getSheetByName('FileIDList');
  var LogSheetID = sheetFileID.getRange(5,2).getValue(); 
  var LogSheetTable = SpreadsheetApp.openById(LogSheetID).getSheetByName('Event_Logs');
  var LogSheetTable = SpreadsheetApp.openById(EventLogs).getSheetByName('Event_Logs');
  var logemail = Session.getActiveUser().getEmail();
  var DateTime = Utilities.formatDate(new Date, "GMT+8","YYYYMMdd h:mm:ss a");
  var Logs = [logemail,DateTime,'Generate Regular Trip Tickets',CountTripTicket + ' Generated Trip Ticket(s) for date: '+DateTrip,'Regular',ticketrange[0] + ' to ' + ticketrange[LastRowRegularTrip-2]];
  Logger.log(Logs); 
  LogSheetTable.appendRow(Logs);
}