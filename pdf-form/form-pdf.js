function onFormSumbit(e){
    var formresponses = e.values;
    var email = formresponses[2];
    var name = '"'+ formresponses[1]+' '+formresponses[3]+'"';
    var address = formresponses[4]; 
    
    var TDate = new Date();
    var editedDate = Utilities.formatDate(TDate,"GMT","dd-MM-yyyy")
    var doc = DocumentApp.openById('DocumentApp');
  
    var body = doc.getBody();
  
    body.replaceText("<<email>>", email);
    body.replaceText("<<name>>", name);
    body.replaceText("<<address>>",address);
    body.replaceText("<<Date>>",editedDate);
  
    doc.saveAndClose()
  
    var pdf = doc.getAs('application/pdf');
  
    var empEmail = email;
    var subject = 'Your Interview';
    var body = "Hi " + name + ", \n\n Please read th attached file for your interview Info";
    
    MailApp.sendEmail(empEmail,subject,body,{attachments:[pdf]})
  
  
    lastVer(email,name,address,editedDate)
  }
  
  function lastVer (email,name,address,editedDate) {
    
    
    var doc = DocumentApp.openById('DocumentApp');
    var body = doc.getBody();
    
    body.replaceText(email, '<<email>>');
    body.replaceText(name, '<<name>>');
    body.replaceText(address, "<<address>>");
    body.replaceText(editedDate, "<<Date>>");
  
     doc.saveAndClose()
  }
  
  
  