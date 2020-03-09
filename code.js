function start(){
  
  // Define sheet names as variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName("master");
  var gantt = ss.getSheetByName("gantt");
  var startpage = ss.getSheetByName("Getting Started");
  var task_summary = ss.getSheetByName("task_summary");
  var empty_template = ss.getSheetByName("empty_template");
  
  
  // Make master sheet active
  master.showSheet();
  ss.setActiveSheet(master);
  // Delete prefilled data if existis
  master.getRange("C3:C11").setValue("");
  master.getRange("E16:F65").setValue("");
  // Set CountryName
  country = startpage.getRange("F66").getValue();
  master.getRange("C2").setValue(country);
  
  // add RQ and Finpro detials for Ghana in the master

  if(country == "Ghana"){
         master.getRange("E12").setValue("Ghana Research Quality Team");
         master.getRange("E13").setValue("Ghana Finance and Procurement Team");
         master.getRange("F12").setValue("mbidinlib@poverty-action.org");
         master.getRange("F13").setValue("finprocgh@poverty-action.org");
         master.getRange("E12:E13").setBackground("#9FC5E8")
         master.getRange("E12:F13").setBorder(true, true, true, true, true, true)
   }
 
  // Make other sheets visible
  
  // gantt and taks summary
  task_summary.showSheet();
  gantt.showSheet();
  //  project open and close
  empty_template.copyTo(ss).setName("project_open");
  empty_template.copyTo(ss).setName("project_close");
  ss.getSheetByName("project_close").showSheet();
  ss.getSheetByName("project_open").showSheet();
  
  var project_open = ss.getSheetByName("project_open");
  var project_close = ss.getSheetByName("project_close");
 
  // Reorder sheets
  ss.setActiveSheet(project_open)
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(2); 
  ss.setActiveSheet(project_close)
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(3);
  ss.setActiveSheet(gantt)
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(4);
  ss.setActiveSheet(task_summary)
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(5);
  ss.setActiveSheet(master)
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(1);
  
  // Set master as active sheet
  ss.setActiveSheet(master);  
  
  // Hide Getting Started sheet
  startpage.hideSheet() 
  
}




// Get back to StartPage

function getback(){

  
  // Message Box to confirm deletion of sheets and returning to getting started 
  
  var name = Browser.msgBox("Restart Alert!","Are you sure you want to restart? Continuing will delete all sheets and progress", Browser.Buttons.OK_CANCEL);
  
  if (name == "ok"){
   
    // Define sheet names
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var startpage = ss.getSheetByName("Getting Started")
    var master = ss.getSheetByName("master");
    var gantt = ss.getSheetByName("gantt");
    var project_open = ss.getSheetByName("project_open");
    var project_close = ss.getSheetByName("project_close");
    var task_sumary = ss.getSheetByName("task_summary");
    
    // Make Getting Started sheet active
    startpage.showSheet();
    ss.setActiveSheet(startpage);
    
    // Hide and delete sheets
    gantt.hideSheet();
    task_sumary.hideSheet();
    ss.deleteSheet(project_open)
    ss.deleteSheet(project_close)
    // Delete created sheets and hide other sheets
    
    // Reformat and Hide Master sheet
    var master = ss.getSheetByName("master");
    master.getRange("E12:E13").setBackground("");
    master.getRange("E12:F13").setValue("");
    master.getRange("E12:F13").setBorder(false, false, false, false, false,false);
    master.getRange("E12:F12").setBorder(true, false, false, false, false, false)
    master.hideSheet()
  
  }
  
  
}




// Function to update sheetes and rows

function onEdit(e) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gant = ss.getSheetByName("gantt");
  var master = ss.getSheetByName("master");
  var project_open = ss.getSheetByName("project_open");
  var project_close = ss.getSheetByName("project_close");
  var health_check = ss.getSheetByName("health_check");
  var task_summary = ss.getSheetByName("task_summary");
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetname = sheet.getSheetName();
  var name = sheetname.split("_")[1]
  
  
  //adding new line at the top
  if (name == "tasks" | name == "survey" | name == "impl" | name =="open" | name == "close") {
      var editcell = sheet.getActiveCell().getValue();
      var editcol = sheet.getActiveCell().getColumn();
      var editrow = sheet.getActiveCell().getRow();
      
      if (editcol == 1 && editcell == "New"){
          sheet.insertRowsBefore(editrow, 1);    // inserts a row before
          
          // remove new values and leave blank
          sheet.getRange(editrow+1, editcol).setValue("")
        
          //Add delete validation for new items added
          sheet.getRange(editrow, editcol).setDataValidation(null)
          var valrange = sheet.getRange(editrow, editcol);
          var rule = SpreadsheetApp.newDataValidation().requireValueInList(["Completed","New","Delete"]).build();
          valrange.setDataValidation(rule);   
      } 
     
      // Delete row that is marked as delete    
      if (editcol == 1 && editcell == "Delete"){
            sheet.deleteRow(editrow)
        } 
        
        // Change fontColor of completed tasks    
      if (editcol == 1 && editcell == "Completed"){
           var completed_range =  sheet.getRange(editrow,editcol+1)
            completed_range.setFontColor("Gray");
            var completed_task = completed_range.getValue();
            var country_name = master.getRange("C2").getValue();
        
             // Trigger an email to RQ Team (Ghana)
            if(country_name == "Ghana"){
                var project_name = master.getRange("C3").getValue();
                var rq_email = master.getRange("F12").getValue();
                var link = SpreadsheetApp.getActiveSpreadsheet().getUrl();
              
               MailApp.sendEmail(
                rq_email,                                                                
                "MyRA_Updates_"+ project_name ,                                         
                "Hello RQ Team,\n\nThere has been a recent MyRa update for " +            
                project_name + "Project. \nTask: " + completed_task + "\nAction: Completed" +
                "\nMyRA Link: " + link); 
            }     
      }
    
            
   }
   
  
}



 
function addchecklist(){

  
  function onOpen() {
    var ui = SpreadsheetApp.getUi();  
    ui.createMenu('Form')
      .addItem('add Item', 'addItem')
      .addToUi();
  }
  
  var html = HtmlService.createHtmlOutputFromFile('form');
  SpreadsheetApp.getUi() 
      .showModalDialog(html, 'Add New Checklist');
}


function addNewItem(form_data){
   // Define sheet names as variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName("master");
  var gantt = ss.getSheetByName("gantt");
  var startpage = ss.getSheetByName("Getting Started");
  var task_summary = ss.getSheetByName("task_summary");
  var empty_template = ss.getSheetByName("empty_template");
  var create_template = ss.getSheetByName("create_template")
  
  var type = form_data.checklist_type
  var sht = form_data.sheet_name
  
  Logger.log(type)
  Logger.log(sht)
  Logger.log(sht + "_" + type)
  
  // for survey
    if(type == "survey"){
      empty_template.copyTo(ss).setName(sht + "_survey");
      ss.getSheetByName(sht + "_survey").showSheet();
    }
    // for implementation
    if(type == "impl"){
      empty_template.copyTo(ss).setName(sht + "_impl");
      ss.getSheetByName(sht + "_impl").showSheet();
    }
  //form empty tasks
    if(type == "empty"){
      create_template.copyTo(ss).setName(sht + "_tasks");
      ss.getSheetByName(sht + "_tasks").showSheet();       
    }
    
  // activate newly created sheet  
  var new_sheet = ss.getSheetByName(sht + "_" + type)
  ss.setActiveSheet(new_sheet)
  ss.moveActiveSheet(3); 
  
  
  
  
}
  
 

