var MAIN_SPREADSHEET_ID='';
var REPORT_FOLDER_ID = '';

function createTriger(formID) {

  var form = FormApp.openById(formID);
  ScriptApp.newTrigger('runCommand')
  .forForm(form)
  .onFormSubmit()
  .create();
}


var date_regex = /^\d{4}\-\d{1,2}\-\d{1,2}$/;



function runCommand(e) {
  var formResponses='';
  
  try
  {
    var items = e.response.getItemResponses();
   
    var form_id = e.response.getEditResponseUrl();
    Logger.log(form_id);
    formResponses += 'Form ID:'+form_id+"\n";
    var formid_Reg = /\/forms\/d\/e\/([^\/]*?)\//;  
    var match = formid_Reg.exec(form_id);
    form_id = match[1];
    formResponses += 'Form Link : '+e.response.getEditResponseUrl()+"\n";    
    var responder = e.response.getRespondentEmail();    
    formResponses += 'Respondent Email : '+responder+"\n";
    formResponses += '# of ItemResponses: '+items.length+"\n";
    var templateid = ''  
    var spreadsheetID = ''
    var spreadsheetIndex = ''
    var spreadsheetColumns = ''
    var sheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID).getSheetByName('Confic');
    formResponses += 'Going to access the sheet :'+MAIN_SPREADSHEET_ID+"\n";
    var formIDsWithTemplateIDs = sheet.getRange(2, 1, sheet.getLastRow(),5).getValues();
    var templateID='';
    
    Logger.log(formIDsWithTemplateIDs);
    
    Logger.log('Going to search:'+form_id)
    formResponses += 'Going to search:'+form_id+"\n";
    for(var id in formIDsWithTemplateIDs)
    {
      if(formIDsWithTemplateIDs[id][0]!='')
      {
        if(formIDsWithTemplateIDs[id][0]==form_id)
        {
          templateid = formIDsWithTemplateIDs[id][1];
          spreadsheetID = formIDsWithTemplateIDs[id][2];
          spreadsheetIndex = formIDsWithTemplateIDs[id][3];
          spreadsheetColumns = formIDsWithTemplateIDs[id][4];
        }   
      }
    }
    
    if(templateid!='')
    {    
      var docid = DriveApp.getFileById(templateid).makeCopy().getId();
      var doc = DocumentApp.openById(docid);
      var body = doc.getActiveSection();  
  
      var report_name='<<Certified Staff>>_<<Observation Date>>_Submitted_<<Timestamp>>';
      
      
      for(var i = 0; i< items.length; i++) 
      {
        var title = items[i].getItem().getTitle();
        var temp =  title;    
        temp += ' : ';
        temp += items[i].getResponse();            
        formResponses+=temp+"\n";                    
        var value = items[i].getResponse();
        
        var r = date_regex.exec(value);
        
        
        if((title.indexOf('Observation Date')>=0 || title.indexOf('Date of Observation')>=0 ) && value!='' && value!=null && value!='null' && r==null)
        {
          var dt = new Date(value);  
          value = Utilities.formatDate(dt, "GMT", "yyyy-MM-dd");
        }      
        
        body.replaceText("<<"+title+">>", value); 
        
        formResponses+='Going to replace [Form Value]:'+"<<"+title+">>"+' with '+value+"\n"            
        
        
        Logger.log(title+'>>>[Form Value]>>>>>>>'+value)            
        if(title=='Certified Staff')
        {
          report_name = report_name.replace('<<Certified Staff>>', items[i].getResponse());
        }      
        if(title=='Observation Date' || title=='Date of Observation')
        {        
          report_name = report_name.replace('<<Observation Date>>', items[i].getResponse());
        }      
      }  
      
      if(spreadsheetID!='')
      {      
        Logger.log('Spreadsheet id found:'+spreadsheetID)    
        Logger.log('Spreadsheet index found:'+spreadsheetIndex)   
        formResponses += 'Spreadsheet id found:'+spreadsheetID+"\n";
         
        formResponses += 'Spreadsheet index found:'+spreadsheetIndex+"\n";
          
        spreadsheetColumns=spreadsheetColumns.split(',');
        Logger.log('Spreadsheet Columns found:'+spreadsheetColumns) 
        formResponses += 'Spreadsheet Columns found:'+spreadsheetColumns+"\n";
        Utilities.sleep(10000)
        var responseSheet = SpreadsheetApp.openById(spreadsheetID).getSheets()[spreadsheetIndex];   
        var lastRow = SpreadsheetApp.openById(spreadsheetID).getSheets()[0].getLastRow();
        Logger.log('Last updated row is '+lastRow)
        formResponses += 'Last updated row is:'+lastRow+"\n";
        formResponses += 'Going to access the sheet :'+spreadsheetID+"\n";
        var headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
        Logger.log('Spreadsheet Headers'+headers);
        formResponses += 'Spreadsheet Headers:'+headers+"\n";
        var columnMap= {};
        
        for(var sc in spreadsheetColumns)
        {
          for(var h in headers)
          {           
            if(spreadsheetColumns[sc]==headers[h])
            {
              columnMap[spreadsheetColumns[sc]] = h;
            }
          }        
        }
        Logger.log('columnMap'+columnMap);
        formResponses += 'columnMap:'+columnMap+"\n";
        
        for(var cm in columnMap)
        {
          
          var value = responseSheet.getRange(lastRow,parseInt(columnMap[cm])+1).getValue();    
          formResponses+='Going to replace [Spreadsheet Value]:'+"<<"+cm+">>"+' with '+value+"\n";
          var r = date_regex.exec(value);
           
           
          if((cm.indexOf('Observation Date')>=0 || cm.indexOf('Date of Observation')>=0 ) && value!='' && value!=null && value!='null' && r==null)
          {
              var dt = new Date(value);              
              value = Utilities.formatDate(dt, "GMT", "yyyy-MM-dd");
          }  
          body.replaceText("<<"+cm+">>", value); 
          formResponses+='Going to replace [Spreadsheet Value]:'+"<<"+cm+">>"+' with '+value+"\n"                
          Logger.log(cm+'>>>[Spreadsheet Value]>>>>>>>'+value)                  
          if(cm=='Certified Staff')
          {
            report_name = report_name.replace('<<Certified Staff>>', value);
          }
          if(cm=='Observation Date' || cm=='Date of Observation')
          {          
            report_name = report_name.replace('<<Observation Date>>', value);
          }
        }      
      }        
      var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd-HH-mm-ss");   
      formResponses+='LEVEL 1 : Report name is '+report_name+"\n";      
      report_name = report_name.replace('<<Certified Staff>>', '');
      report_name = report_name.replace('<<Observation Date>>','');    
      report_name = report_name.replace('<<Timestamp>>', formattedDate); 
      formResponses+='LEVEL 2 : Report name is '+report_name+"\n";
      doc.setName(report_name);    
      doc.saveAndClose();    
      formResponses+='Please find the report template in the below link'+"\n";
      formResponses+=doc.getUrl();
      var file = DriveApp.getFileById(doc.getId());        
      var folder = DriveApp.getFolderById(REPORT_FOLDER_ID);
      //file.addToFolder(folder);
      //folder.addfile(file);
      file.makeCopy(folder)
    }
  } 
  catch(err)
  {
    formResponses+=err;
  }
  
  MailApp.sendEmail('hendersonr@rccsec.org', 'This is response form trigger',formResponses)
}
function deleteAllTriggers()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var formIDs = sheet.getRange(2, 1, sheet.getLastRow()).getValues();
  var myTriggers = ScriptApp.getProjectTriggers();
  Logger.log('Total triggers:'+myTriggers.length)
  for(var id in formIDs)
  {
    var triggerFound=0;
    if(formIDs[id]!='')
    {      
      for(var i=0;i<myTriggers.length;i++) 
      {     
        Logger.log(myTriggers[i].getTriggerSourceId());
        Logger.log(myTriggers[i].getUniqueId());
        Logger.log(formIDs[id])
        if(formIDs[id]==myTriggers[i].getTriggerSourceId())
        {       
           ScriptApp.deleteTrigger(myTriggers[i])
        }
      } 
    }
  } 
  
}




function initTrigges() {

  deleteAllTriggers();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var formIDs = sheet.getRange(2, 1, sheet.getLastRow()).getValues();
  var myTriggers = ScriptApp.getProjectTriggers();
  
  for(var id in formIDs)
  {
    var triggerFound=0;
    if(formIDs[id]!='')
    {
      for(var i=0;i<myTriggers.length;i++) 
      {                         
        if(formIDs[id]==myTriggers[i].getTriggerSourceId())
        {
          triggerFound=1;
        }
      }
      if(triggerFound==0)
      {
        Logger.log('Going to create trigger for Form with id:'+formIDs[id]);
        createTriger(formIDs[id])
      }
     }
  } 
}



 function onOpen() {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var menuEntries = [];   
   menuEntries.push({name: "Initiate Trigger", functionName: "initTrigges"});
   menuEntries.push(null); // line separator
   menuEntries.push({name: "Delete All Triggers", functionName: "deleteAllTriggers"});
   ss.addMenu("Manage Triggers", menuEntries);
 }
 
