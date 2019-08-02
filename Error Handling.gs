function getFunctionName(x) {

try{

    var functionName = x;
        functionName = functionName.substr('function '.length);
        functionName = functionName.substr(0, functionName.indexOf('('));
    var emailSubject = 'TS Dispatch Error';
    var email1 = 'm.delrosario@irri.org';
  
  
    return functionName;
    }
    catch(e) {
      MailApp.sendEmail(email1, emailSubject, '',
        {htmlBody: 'Function Name: '+ 'getFunctionName'+ '<br> Filename: '+ e.fileName + '<br> Message: '+ e.message+ '<br> Line No: '+e.lineNumber})
     
      throw e;
    }
}

