function MergeDocumentA() {
  BlankVehicle2();
  var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
  var lastRowTicket = sheetTicket.getLastRow(); 
  var docIDs = []; var docIndex = 0; var charlimit = []; var trim = ""; var charlength = []; 
  //------------------------------------------------------------
  var countdoc = 0;
    //Logger.log(lastRowTicket);
         for (var z = 3; z < lastRowTicket; z++) {
             var docID = sheetTicket.getRange(z, 19).getValue();
             var VehicleID = sheetTicket.getRange(z, 15).getValue();
               if (docID == "") {
                    countdoc++;
               }
               if (VehicleID == "") {
               var VehicleID = sheetTicket.getRange(z, 18).setValue("");  
               }
         }
   var startrow = Number(lastRowTicket) - Number(countdoc);
   var VehicleID = sheetTicket.getRange(z, 19).getValue();
  //----------------------------------------------------------
   if (startrow > 0) {
   for ( var j = startrow; j < lastRowTicket+1; j++){
    var data = [];
    var i = 1;
    for (var x = 0; x < 18; x++) {
             data[x] = sheetTicket.getRange(j, x+1).getValue();
             var sheetCharLimit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CharLimit');
             charlimit[x] = sheetCharLimit.getRange(2, x+1).getValue();
           
             if (data[x] instanceof Date) { Logger.log("Its a date:" + x); } else {charlength[x] = data[x].length; Logger.log(x + " " + data[x] + " " + charlength[x] + " " + charlimit[x]); }
             if (charlength[x] >= charlimit[x]) { Logger.log("Long"); trim = data[x].substring(0, charlimit[x]) + "..."; data[x] = trim; Logger.log(data[x]);} else {Logger.log("Short");}
             if (data[14] == '') {data[14] = '__________' } 
             if (data[17] == '') {data[17] = '__________' } 
  }
  var fileId = '1kVN1OtCqEfzTrhFd3OP5LxqNJ0AF97PITVmzIrfnaCQ'; //Template Source
  var fileTayp = DriveApp.getFileById(fileId);
  var NewDocumentFolder = DriveApp.getFolderById('0B6yi2TYFy9gaMzFBd3VQWk95N3M'); //Destination File
  
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
          
       
          copyBody.replaceText('<<VehicleID>>', data[14]);
          copyBody.replaceText('<<PlateNo>>', data[17]);
          var ConvertTime = Utilities.formatDate(data[16], "GMT+8","h:mm a");
          copyBody.replaceText('<<Departure>>', ConvertTime);
          
          copyBody.replaceText('<<Driver>>', data[15]);
          copyBody.replaceText('<<BUS>>', data[4]); 
          copyBody.replaceText('<<Requestor>>', data[5]);
          copyBody.replaceText('<<OU>>', data[6]);
          copyBody.replaceText('<<Passenger>>', data[7]);
          if (data[8] instanceof Date){ var PickupTime = Utilities.formatDate(data[8], "GMT+8","h:mm a");}
          else { PickupTime = data[8] }
          copyBody.replaceText('<<Pickup Time>>', PickupTime);
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
   }
        var NewDocumentFolder = DriveApp.getFolderById('0B6yi2TYFy9gaMzFBd3VQWk95N3M');
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
        //Logger.log("Doczero: " + docIDs[0]);
        //Logger.log("BatchDocID: " + BatchDocID);
        // Convert to PDF Script
       // var PDFfolder = DriveApp.getFolderById(''); // Destination File
        //var pdfreference = DriveApp.getFileById(BatchDocID);
        var source = DriveApp.getFileById(BatchDocID);
        var blob = source.getAs('application/pdf');
        var PDFfile = DriveApp.getFolderById('0B6yi2TYFy9gaMzFBd3VQWk95N3M').createFile(blob);
        var BatchPDFID = PDFfile.getId();
        var BatchPDFURL = PDFfile.getUrl();
        var BatchPDFlink = [['=hyperlink("' + BatchPDFURL + '", "' + Batchname + '")']];
        //Record the Logs in the row Script
        for (var k = startrow; k < lastRowTicket+1; k++){
        sheetTicket.getRange(k, 19).setValue(BatchPDFID); 
        sheetTicket.getRange(k, 20).setValue(BatchPDFURL); 
        sheetTicket.getRange(k, 21).setValue(BatchPDFlink); 
         var CurrentDateTime = Utilities.formatDate(new Date(), "GMT+8","M/d/YYYY+h:mm a");
        sheetTicket.getRange(k, 22).setValue('Document successfully merged ' + CurrentDateTime); 
        }
        
        BatchNum++;
        sheetAutoNumber.getRange(5, 2).setValue(BatchNum); //Update the Batch Code
        var length = docIDs.length;
        //Logger.log("docsIDs Length: " + length)
        //Logger.log("BatchDocID " + BatchDocID);
        //Delete the Docs Generated after PDF Convertion
         DriveApp.getFileById(BatchDocID).setTrashed(true)
        //autoCrat_trashDoc(BatchDocID);

         for (var x = 1; x < docIDs.length+1; ++x ) {
         var y= x-1; 
         //Logger.log(y + " " + docIDs[y]);
         DriveApp.getFileById(docIDs[y]).setTrashed(true)
         Drive.Files.remove(docIDs[y]);
         }
   
   } else { Browser.msgBox("No Rows found to Generate PDF"); }
  if (countdoc > 0) {
    Browser.msgBox(countdoc + " Trip Ticket(s) PDF Generated \\n\\nYou may now view the Tickets in PDF Viewer.");
  }
}



function autoCrat_trashDoc(docId) {  // Forever Delete File
  //var docId = '1Yt6uGjA0zyX0NTgX51s1vU6iynoYXqN7Vj3wkNrpwzw';
  var file = DriveApp.getFileById(docId);
  DriveApp.getFileById(docId).setTrashed(true)
  Drive.Files.remove(docId);
}

function datetoday() {
  //var newPickupDate = Utilities.formatDate(PickUpDate, "GMT+8","M/d/YYYY");
  var CurrentDateTime = Utilities.formatDate(new Date(), "GMT+8","YYYYMMdd h:mm a");
  Logger.log(CurrentDateTime);
}