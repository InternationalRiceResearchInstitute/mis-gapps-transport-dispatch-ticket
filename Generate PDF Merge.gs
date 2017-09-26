function getData() {
  var sheetFileID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FileIDList');
  var fileId = sheetFileID.getRange(7,2).getValue(); 
  var DocumentFolder = sheetFileID.getRange(10,2).getValue(); 
  var LogSheetID = sheetFileID.getRange(5,2).getValue(); 
  Logger.log(LogSheetID); 
  
}

function MergeDocument() {
  
  //UpdateTripTicket(); //Saves Trip TicketDetails from the Clone Data List - sheet AssignTicket
  BlankVehicle2();
  var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
  var lastRowTicket = sheetTicket.getLastRow(); 
  var numNew = 0; 
  var docIDs = []; var docIndex = 0; var charlimit = []; var trim = ""; var charlength = []; 
 
var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');    
var lastRowAssign = sheetAssign.getLastRow();  
var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  
var lastRowTripTicket = sheetTripTicket.getLastRow();
var RowPrint = []; var i = 0;
     for (var g = 2; g < lastRowAssign+1; g++) {
          var TripTicketforPrinting = sheetAssign.getRange(g,1).getValue();
          for(var h = 3; h < lastRowTripTicket+1; h++) {
               var TripTicketforUpdate = sheetTripTicket.getRange(h,1).getValue();
               if( TripTicketforPrinting == TripTicketforUpdate) {
               RowPrint[i] = h;
               i++;
               }
          }
     }
     var length = RowPrint.length;
     for (var k = 0; k < RowPrint.length; k++ ) {
     var j = RowPrint[k];
     //Logger.log(j);
     
  
  //---------------------------------------------------------- AAA
   //if (startrow > 0) {
   //for ( var j = startrow; j < lastRowTicket+1; j++){
    var data = [];
    var i = 1;
    for (var x = 0; x < 20; x++) {
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
  //var fileId = '1tAR41bRRrATI75SS4zZAJcHRiJ15xcrmwJzlCrXq-LA'; //Template Source
  var fileTayp = DriveApp.getFileById(fileId);       
  var NewDocumentFolder = DriveApp.getFolderById(DocumentFolderID); //Destination Folder ID
  
  //  var  = '';
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
          //if (data[8] == '') { data[8] = '__________'; }
          //if (data[8] instanceof Date){ 
          //Logger.log("Valid Date");
          //var PickupTime = Utilities.formatDate(data[8], "GMT+8","h:mm a");
          //} else { PickupTime = data[8]; }  
          //if (data[8] instanceof Date){ var PickupTime = Utilities.formatDate(data[8], "GMT+8","h:mm a");}
          //else { PickupTime = data[8] }
          if (data[8] instanceof Date){
          data[8]  = Utilities.formatDate(data[8], "GMT+8","h:mm a");
          } else { data[8] = ""; }  
            
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
          docIDs[docIndex] = NewDocID;
          var Doclink = [['=hyperlink("' + DocURL + '", "' + DocTitle + '")']];
          
          //sheetTicket.getRange(j, 19).setValue(NewDocID); 
          //sheetTicket.getRange(j, 20).setValue(DocURL); 
          //sheetTicket.getRange(j, 21).setValue(Doclink); 
          //sheetTicket.getRange(j, 22).setValue('Document successfully merged'); 
          //var blob = source.getAs('application/pdf');
          //var PDFfolder = '0B6yi2TYFy9gaMzFBd3VQWk95N3M'; 
//          var PDFfile = DriveApp.getFolderById(PDFfolder).createFile(blob);
//          var PDFLink = PDFfile.getId();
//          var PDFURL = "https://docs.google.com/open?id=" + PDFfile.getId();
//          var PDFId = PDFfile.getId();
//          Logger.log("PDF ID: " + PDFId);
//          Logger.log("PDF URL: " + PDFURL);
          
          //autoCrat_trashDoc(NewDocID);
          docIndex++;
        }
   }   // J ends here
     
        // Document Stapling  Starts here-----------------------------------------------------------------------
        var NewDocumentFolder = DriveApp.getFolderById(DocumentFolderID);  //PDF Folder ID
        var sheetAutoNumber = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');
        var BatchNum = sheetAutoNumber.getRange(5, 2).getValue();
        var Batchname = 'Batch' + (Number(10000) + Number(BatchNum));
        var BatchDocID = DriveApp.getFileById(docIDs[0]).makeCopy(Batchname,NewDocumentFolder).getId();
        DriveApp.getFileById(docIDs[0]).setTrashed(true);
        Drive.Files.remove(docIDs[0]);
        docIDs[0] = BatchDocID;
        var baseDoc = DocumentApp.openById(docIDs[0]);
        var body = baseDoc.getActiveSection();
        for (var x = 1; x < docIDs.length; ++x ) {
        var otherBody = DocumentApp.openById(docIDs[x]).getActiveSection();
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
        }
        baseDoc.saveAndClose();
        // Document Stapling  ends here-----------------------------------------------------------------------
     
        //Logger.log("Doczero: " + docIDs[0]);
        //Logger.log("BatchDocID: " + BatchDocID);
        // Convert to PDF Script
       // var PDFfolder = DriveApp.getFolderById(''); // Destination File
        //var pdfreference = DriveApp.getFileById(BatchDocID);
        var logemail = Session.getActiveUser().getEmail();
        var DateTime = Utilities.formatDate(new Date, "GMT+8","YYYYMMDD h:mm a");
        var source = DriveApp.getFileById(BatchDocID);
        var blob = source.getAs('application/pdf');
        var PDFfile = DriveApp.getFolderById('0B6yi2TYFy9gaY2llSGZvVWJocVE').createFile(blob);
        var PDFfile = DriveApp.getFolderById('0B6yi2TYFy9gaY2JXc0t0T1Q3V1E').createFile(blob);
        var BatchPDFID = PDFfile.getId();
        var BatchPDFURL = PDFfile.getUrl();
        var BatchPDFlink = [['=hyperlink("' + BatchPDFURL + '", "' + Batchname + '")']];
        //Record the Logs in the row Script
        for (var m = 0; m < RowPrint.length; m++ ) {
        var k = RowPrint[m];
        //for (var k = startrow; k < lastRowTicket+1; k++){
        sheetTicket.getRange(k, 21).setValue(BatchPDFID); 
        sheetTicket.getRange(k, 22).setValue(BatchPDFURL); 
        sheetTicket.getRange(k, 23).setValue(BatchPDFlink); 
         var CurrentDateTime = Utilities.formatDate(new Date(), "GMT+8","M/d/YYYY+h:mm a");
          sheetTicket.getRange(k, 24).setValue('Document successfully merged by: ' +logemail+" " + CurrentDateTime); 
        }
        
        BatchNum++;
        sheetAutoNumber.getRange(5, 2).setValue(BatchNum); //Update the Batch Code
        var length = docIDs.length;
        //Logger.log("docsIDs Length: " + length)
        //Logger.log("BatchDocID " + BatchDocID);
        //Delete the Docs Generated after PDF Convertion
         DriveApp.getFileById(BatchDocID).setTrashed(true)
        //autoCrat_trashDoc(BatchDocID);
         //File Deleting Starts here
         for (var x = 1; x < docIDs.length+1; ++x ) {
         var y= x-1; 
         //Logger.log(y + " " + docIDs[y]);
         DriveApp.getFileById(docIDs[y]).setTrashed(true)
         Drive.Files.remove(docIDs[y]);
         } 
         
   
  // } else { Browser.msgBox("No Rows found to Generate PDF"); }
  //if (countdoc > 0) {
  
    if (numNew != 0) {} else {var ReferenceID  = 'None';}
    var sheetFileID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FileIDList');
    var LogSheetID = sheetFileID.getRange(5,2).getValue(); 
    var LogSheetTable = SpreadsheetApp.openById(LogSheetID).getSheetByName('Event_Logs');
    var logemail = Session.getActiveUser().getEmail();
    var DateTime = Utilities.formatDate(new Date, "GMT+8","YYYYMMdd h:mm:ss a");
    var Logs = [logemail,DateTime,'Generate PDF Trip Ticket',RowPrint.length +' PDF Generated Trip Ticket(s) for date: ' + ConvertedDate,'Trip Request',Batchname];
    LogSheetTable.appendRow(Logs);
  
  UpdatePDFLinksToAssign();
  UpdatePDFLinksToTripRequest()
  Browser.msgBox(RowPrint.length + " Trip Ticket(s) PDF Generated. \\n\\nYou may now View/Print the Tickets in PDF Viewer.\\n\\nPDF File - " + Batchname);
  //}
  FirstMenu();
}

function UpdatePDFLinksToAssign123(){
var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');    
var lastRowAssign = sheetAssign.getLastRow();  
var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  
var lastRowTripTicket = sheetTripTicket.getLastRow();

    for (var x = 2; x < lastRowAssign+1; x++) {
       var TripTicketAssign = sheetAssign.getRange(x,1).getValue();
       for (var y = 3; y < lastRowTripTicket+1; y++){
       var TripTicketPDF = sheetTripTicket.getRange(y,1).getValue();
                 if (TripTicketPDF == TripTicketAssign) {
                        var PDFID = sheetTripTicket.getRange(y, 21).getValue();
                        if (PDFID != '') {
                        var PDFName = sheetTripTicket.getRange(y, 23).getValue();
                        var PDFURL = DriveApp.getFileById(PDFID).getUrl(); 
                        var PDFHyperLink ='=hyperlink("' + PDFURL + '", "' + PDFName + '")';
                        sheetAssign.getRange(x,21).setValue(PDFHyperLink);
                        }
                  }
             }
     }
}

function UpdatePDFLinksToTripRequest(){
var sheetTripRequest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripRequest');    
var lastRowTripRequest = sheetTripRequest.getLastRow();  
var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  
var lastRowTripTicket = sheetTripTicket.getLastRow();

    for (var x = 3; x < lastRowTripRequest+1; x++) {
       var TripTicketRequest = sheetTripRequest.getRange(x,22).getValue();
       for (var y = 3; y < lastRowTripTicket+1; y++){
       var TripTicketPDF = sheetTripTicket.getRange(y,1).getValue();
                 if (TripTicketPDF == TripTicketRequest) {
                        var PDFID = sheetTripTicket.getRange(y, 21).getValue();
                        if (PDFID != '') {
                        var PDFName = sheetTripTicket.getRange(y, 23).getValue();
                        var PDFURL = DriveApp.getFileById(PDFID).getUrl(); 
                        var PDFHyperLink ='<b><a href="'+PDFURL+'" style="text-decoration:none;background-color:transparent" target="_blank">ⓘ</a></b>';
                        sheetTripRequest.getRange(x,24).setValue(PDFHyperLink);
                        }
                  }
             }
     }
}


function checkforprinting(){
var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');    
var lastRowAssign = sheetAssign.getLastRow();  
var sheetTripTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');  
var lastRowTripTicket = sheetTripTicket.getLastRow();
var RowPrint = []; var i = 0;
     for (var g = 2; g < lastRowAssign+1; g++) {
          var TripTicketforPrinting = sheetAssign.getRange(g,1).getValue();
          for(var h = 3; h < lastRowTripTicket+1; h++) {
               var TripTicketforUpdate = sheetTripTicket.getRange(h,1).getValue();
               if( TripTicketforPrinting == TripTicketforUpdate) {
               RowPrint[i] = h;
               i++;
               }
          }
     }
     var length = RowPrint.length;
     for (var k = 0; k < RowPrint.length; k++ ) {
     var j = RowPrint[k];
     Logger.log(j);
     }
}
  
function trial(){
    var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CombinedTrips');
    var data = sheetAssign.getRange(2,17).getValue();
    if (data == '') { data = '__________'; }
    if (data instanceof Date){ 
        Logger.log("Valid Date");
        var PickupTime = Utilities.formatDate(data, "GMT+8","h:mm a");
        Logger.log(PickupTime);
        } else { 
        Logger.log("Invalid Date");
        PickupTime = data 
        Logger.log(PickupTime)
        }
}


