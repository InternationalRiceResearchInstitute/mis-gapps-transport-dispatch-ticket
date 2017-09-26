function resetCounters() {
            var sheetAutoNum = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');
            var value1 = sheetAutoNum.getRange(1,2).getValue();
            var value2 = sheetAutoNum.getRange(2,2).getValue();
            var DateLog = Utilities.formatDate(new Date(), "GMT+8","yyyy-MM-dd' 'HH:mm:ss a");
  
            sheetAutoNum.getRange(3, 2).setValue(value1);    
            sheetAutoNum.getRange(4, 2).setValue(value2); 
            sheetAutoNum.getRange(3, 3).setValue(DateLog); 
            sheetAutoNum.getRange(4, 3).setValue(DateLog); 
            
            sheetAutoNum.getRange(1, 2).setValue('1');    
            sheetAutoNum.getRange(2, 2).setValue('1'); 
            sheetAutoNum.getRange(1, 3).setValue(DateLog);    
            sheetAutoNum.getRange(2, 3).setValue(DateLog); 
       //    sheetAutoNum.appendRow('Trip Request: '+ value1 , "Trip Ticket: " + value2,DateLog);
}

function revertCounters() {
            var sheetAutoNum = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoNumber');
            var value1 = sheetAutoNum.getRange(3,2).getValue();
            var value2 = sheetAutoNum.getRange(4,2).getValue();
            var DateLog = Utilities.formatDate(new Date(), "GMT+8","yyyy-MM-dd' 'HH:mm:ss a");
  
            sheetAutoNum.getRange(3, 2).setValue('1');    
            sheetAutoNum.getRange(4, 2).setValue('1'); 
            sheetAutoNum.getRange(3, 3).setValue(DateLog); 
            sheetAutoNum.getRange(4, 3).setValue(DateLog); 
            
            sheetAutoNum.getRange(1, 2).setValue(value1);    
            sheetAutoNum.getRange(2, 2).setValue(value2); 
            sheetAutoNum.getRange(1, 3).setValue(DateLog);    
            sheetAutoNum.getRange(2, 3).setValue(DateLog); 
       //    sheetAutoNum.appendRow('Trip Request: '+ value1 , "Trip Ticket: " + value2,DateLog);
            
    
}