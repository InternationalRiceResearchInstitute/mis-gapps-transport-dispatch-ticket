function testingsdfsa(){

   var scriptProperties = PropertiesService.getScriptProperties();
 scriptProperties.setProperty("runCount", 0); 

}

function myFunction2() {

var array = ['First', 'Second','Third','Fourth','Fifth'];

Logger.log(array.length);

array[0] = 'This is the new first element.';

Logger.log(array);
  

var emptyArray = [];

emptyArray = ['not empty anymore'];

Logger.log(emptyArray);

var arrayString = array.join(' ');

Logger.log(arrayString);

}




function pickuptimecheck(){ 
  var sheetSorted = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoSort');
  
  var PickUpTime = sheetSorted.getRange(2, 7).getValue();    
  var ConvertTime = Utilities.formatDate(PickUpTime, "GMT+8","hh:mm a");
  Logger.log("Original Time: " + PickUpTime);
  Logger.log("Converted Time: " + ConvertTime);
}

function CharLimit() {
   var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CharLimit');
   var charlimit = []; 
   var data = []; var charlength = []; var trim
   for (var x = 0; x < 13; x++) {
      charlimit[x] = sheetTicket.getRange(2, x+1).getValue();
      data[x] = sheetTicket.getRange(3, x+1).getValue();
      if (data[x] instanceof Date) { Logger.log("Its a date:" + x); } else {charlength[x] = data[x].length; Logger.log(x + " " + data[x] + " " + charlength[x] + " " + charlimit[x]); }
      if (charlength[x] >= charlimit[x]) { Logger.log("Long"); trim = data[x].substring(0, charlimit[x]) + "..."; data[x] = trim; Logger.log(data[x]);} else {Logger.log("Short");}
      //Logger.log("Charlimit " + d + ": " + charlimit[d]);
   }
   //Logger.log(charlimit.length)
}

function CharLimit2() {
   var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CharLimit');
   var charlimit = []; 
   var charsample = []; var charlength = []; var trim
   for (var d = 0; d < 13; d++) {
      charlimit[d] = sheetTicket.getRange(2, d+1).getValue();
      charsample[d] = sheetTicket.getRange(3, d+1).getValue();
      if (charsample[d] instanceof Date) { Logger.log("Its a date:" + d); } else {charlength[d] = charsample[d].length; Logger.log(d + " " + charsample[d] + " " + charlength[d] + " " + charlimit[d]); }
      if (charlength[d] >= charlimit[d]) { Logger.log("Long"); trim = charsample[d].substring(0, charlimit[d]) + "..."; charsample[d] = trim; Logger.log(charsample[d]);} else {Logger.log("Short");}
      //Logger.log("Charlimit " + d + ": " + charlimit[d]);
   }
   //Logger.log(charlimit.length)
}

function charleng() {
var str = "Hello World!";
var n = str.length;
  Logger.log(n);
}

function trimchar() {
  var str = "The quick brown fox jumps over the lazy dog.";
  var n = str.lenth; 
     Logger.log(n);
  var res = str.substring(0, 20) + "...";
 // Logger.log(res);
}

function MergeDocument2() {
  //var TRIPTICKET = 'TRIP201650010';
  // var PRNumber = '120311111';
  //var Driver = 'Nestor Marcelo';
  
  var sheetTicket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripTicket');
  var lastRowTicket = sheetTicket.getLastRow(); 
  var docs = []; var docIndex = 0; 
  //------------------------------------------------------------
  var countdoc = 0;
    //Logger.log(lastRowTicket);
         for (var z = 2; z < lastRowTicket; z++) {
             var docID = sheetTicket.getRange(z, 19).getValue();
               if (docID == "") {
                    countdoc++;
               }
         }
   var startrow = Number(lastRowTicket) - Number(countdoc);
   var VehicleID = sheetTicket.getRange(z, 19).getValue();
   if (startrow > 0) {
   for ( var j = startrow; j < lastRowTicket+1; j++){
   //-----------------------------------------------------------
  
  var data = [];
  var i = 1;
  for (var x = 1; x < 19; x++) {
             data[x-1] = sheetTicket.getRange(j, x).getValue();
             //Logger.log(data[x]);
  }
  var fileId = '1kVN1OtCqEfzTrhFd3OP5LxqNJ0AF97PITVmzIrfnaCQ'; //Template Source
  var fileTayp = DriveApp.getFileById(fileId);
  var NewDocumentFolder = DriveApp.getFolderById('0B6yi2TYFy9gaMzFBd3VQWk95N3M'); //Destination File
  var PDFfolder = DriveApp.getFolderById('0B6yi2TYFy9gaMzFBd3VQWk95N3M'); // Destination File
  //  var  = '';
  var mimetype = fileTayp.getMimeType();
          if ((mimetype=="application/vnd.google-apps.document")||(mimetype=="application/vnd.google-apps.document")) { 
          var template = DocumentApp.openById(fileId);
          var title = template.getName();
          var NewDocID = DriveApp.getFileById(fileId).makeCopy(data[0],NewDocumentFolder).getId();
          //var NewDocID = DriveApp.getFileById(fileId).makeCopy(TRIPTICKET, NewDocumentFolder).getId();
          var NewDocName = DocumentApp.openById(NewDocID).getName();
          //Logger.log("Template: " + template);
          //Logger.log("Title: " + title);
          //Logger.log("NewDocID: " + NewDocID);
          //Logger.log("DocNAme: " + NewDocName);
          
          var copyDoc = DocumentApp.openById(NewDocID);
          // Get the document’s body section and replace the details
          var copyBody = copyDoc.getActiveSection();  
          var ConvertedDate = Utilities.formatDate(data[2], "GMT+8","M/d/YYYY");
          copyBody.replaceText('<<TripDate>>', ConvertedDate);
          copyBody.replaceText('<<TripTicket>>', data[0]);
          copyBody.replaceText('<<PR Number>>', data[3]);
          copyBody.replaceText('<<VehicleID>>', data[14]);
          var ConvertTime = Utilities.formatDate(data[16], "GMT+8","h:mm a");
          copyBody.replaceText('<<Departure>>', ConvertTime);
          copyBody.replaceText('<<PlateNo>>', data[17]);
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
          docs[docIndex] = NewDocID;
          var Doclink = [['=hyperlink("' + DocURL + '", "' + DocTitle + '")']];
          
          sheetTicket.getRange(j, 19).setValue(NewDocID); 
          sheetTicket.getRange(j, 20).setValue(DocURL); 
          sheetTicket.getRange(j, 21).setValue(Doclink); 
          sheetTicket.getRange(j, 22).setValue('Document successfully merged'); 
          var blob = source.getAs('application/pdf');
          var PDFfolder = '0B6yi2TYFy9gaMzFBd3VQWk95N3M'; 
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
     Logger.log(docs);
     var length = docs.length;
     Logger.log(length);
     
        var NewDocumentFolder = DriveApp.getFolderById('0B6yi2TYFy9gaMzFBd3VQWk95N3M');
        var sheetAutoNumber = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');
        var BatchNum = sheetAutoNumber.getRange(5, 2).getValue();
        var DateToday = Utilities.formatDate(new Date(), "GMT+8","YYYYMMdd")
        var Batchname = DateToday + 'Batch' + (Number(10000) + Number(BatchNum));
     
        var BatchDocID = DriveApp.getFileById(docIDs[0]).makeCopy(Batchname,NewDocumentFolder).getId();
         
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
        Logger.log("Doczero: " + docIDs[0]);
        Logger.log("BatchDocID: " + BatchDocID);
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
        sheetTicket.getRange(k, 22).setValue('Document successfully merged'); 
        }
        
        BatchNum++;
        sheetAutoNumber.getRange(5, 2).setValue(BatchNum); //Update the Batch Code
        
        
   } else { Browser.msgBox("No Rows found to Generate PDF"); }

}

function trimchar() {
  var input1 = "The quick brown fox jumps over the lazy dog.";
  var res = input1.substring(0, 20) + "...";
  Logger.log(res);
  
}

function Left(str, optLen) {
  return Mid( str, 1 , optLen);
}

function ListforPrinting() {
var sheetAssign = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AssignTicket');    
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

function BrowserMsg() {
  
 if (Browser.msgBox('Hello', 'Have you made a copy of the blank spreadsheet?', Browser.Buttons.YES_NO) == 'no') {
  Browser.msgBox('Please make a copy of the blank spreadsheet before entering data.  Thank you!')
}
  
}