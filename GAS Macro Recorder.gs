/***********************************
____________________________________
***    | GAS MACRO RECORDER |    ***
------------------------------------
***********************************/

/*   
   Copyright 2012 Rahman Mohamud Faisal MOORABY 

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/


// function onOpen()
// Runs on opening of the spreadhseet
// Adds menus and submenus
function onOpen() {
  
  // Get current spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create submenus 'Record Macro' and 'Stop Recording Macro'
  var entries = [
    {name : "Record Macro", functionName : "showApp"},
    {name : "Stop Recording Macro", functionName : "recStop"},
    
  ];
  
  // Create a menu 'Recording Macro App', with the above submenues, to the GUI
  sheet.addMenu("Recording Macro App", entries);
    
}
    
    
    
// function showApp()
// The function to initiate record the macros
// Recording the macros is dependent on the onEdit trigger
// This function records the initial values in the active sheet, active range
// and then set the onEdit trigger to run the recordMacro(e) - which basically record the values of active cells (at this stage)
function showApp(){
 
  // Delete all programatical triggers - to eliminate dupication of triggers 
  deleteTriggers(); 
  
    // Get elements of the spreadsheet:
  var sheet = SpreadsheetApp.getActiveSheet();         // the active sheet
  var range = sheet.getActiveRange();                  // the active range
  var values = sheet.getActiveRange().getValues().toString();  // the values in the range
  
  // output the values in the range in the string out_str
  var out_str = "";
  out_str = "SpreadsheetApp.setActiveSheet(\""+sheet.getName()+"\");\n";
  out_str = out_str+"var sheet = SpreadsheetApp.getActiveSheet();\n";
  out_str = out_str+"sheet.setActiveRange("+ range.getRowIndex() +","+ range.getColumnIndex() +","+ range.getHeight() +","+ range.getWidth() +");\n";
  out_str = out_str+"var range = sheet.getActiveRange();\n";
  out_str = out_str+"range.setValues(\""+ values +"\");\n";
  out_str = out_str+"var values = range.getvalues();\n";
  
  // for global value passing (for other functions to use the values), output the string value to a property called 'CODE' in ScriptProperties
  ScriptProperties.setProperties({'CODE': out_str});
  
  // Create the onEdit Trigger (set to use recordMacro function)
  createonEditTrigger();
    
}

    
// function recordMacro(e) 
// This function records values from active cells and ranges when edited
// This function is triggered by the onEdit Event

// This function will evolve to look at further property changes within the cells and also looking at
// other objects addition, removal or modification (e.g. charts and sheets)
function recordMacro(e) {

    // Get spreadsheet elements
    var sheet1 = SpreadsheetApp.getActiveSheet(); // The active sheet
    var range1 = sheet1.getActiveRange(); // the active range - range being modified
    var values1 = range1.getValues(); // the values added
    
    var out_str = ScriptProperties.getProperty('CODE'); // Get the global value of out_str (contains previous outputs as off the start of the recording)

   // Start formating the current value to 'macro' format and append to the global parameter out_str
    out_str = out_str+"\n\n"; // add some extra blank lines
      out_str = out_str+"SpreadsheetApp.setActiveSheet(\""+sheet1.getName()+"\");\nsheet = SpreadsheetApp.getActiveSheet();\n"; //The spreadsheet part
      out_str = out_str+"sheet.setActiveRange("+ range1.getRowIndex() +","+ range1.getColumnIndex() +","+ range1.getHeight() +","+ range1.getWidth() +");\nvar range = sheet.getActiveRange();\n";// the range part
      out_str = out_str+ "range.setValues(\""+ values1 +"\");\nvalues = range.getvalues();\n";// the set value part
  
    
  // Add the output to the 'CODE' property Globally
  ScriptProperties.setProperties({'CODE': out_str});
      
    


}

    
// function recStop()
// This function stops the recording process by
// Outputing the 'recorded macro' to a separate sheet
// and clearing all global properties used in this macro
function recStop(){
 
  var spreadsheet = SpreadsheetApp.getActive(); // Get the active Sheet
    
  // Get the Macro string from the global property 'CODE' where the outputs have been saved
  var out_str = ScriptProperties.getProperty('CODE');
  
  // If there were none, output blank instead of 'null'
  if (out_str ==null) out_str="";
  
  // Get the sheet 'RECORDER_OUTPUT or create it
  var recsheet = spreadsheet.getSheetByName('RECORDER_OUTPUT');
  if (recsheet!=null){
    // If it exists, clear all its content
    recsheet.clear();
  }
  else{
    // else, create it
    recsheet = spreadsheet.insertSheet('RECORDER_OUTPUT');
  }
  
  spreadsheet.setActiveSheet(recsheet); // Set the 'Recorder_Output as the active sheet

  // The macro will be outputted in the first cell 'A1',
  recsheet.setActiveCell('A1');
  var range = recsheet.getActiveCell(); //Set the cell as active

  // Add some comments, the function paranthese and name around the macro output
  out_str = "\\**********************************************************************************\n**  COPY THIS OUTPUT TO A MACRO SHEET IN SCRIPT EDITOR  **\n**********************************************************************************\\\nfunction MacroRecord(){\n\n\t"+out_str+"\n\n}";
  range.setValue(out_str); // finally add the output to the cell
  range.setWrap(true);
  spreadsheet.setColumnWidth(1, 800); // Make the width large enough
  
  // Clear all global parameters
  ScriptProperties.deleteProperty('CODE');
  
  // Delete all programaticall triggers to prevent conflicts and duplications
  deleteTriggers();
  
 }


// function createonEditTrigger()
// This function assigns the onEdit trigger to the function recordMacro();
function createonEditTrigger(){
    // Get spreadsheet key
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = spreadsheet.getId();

  // Create onEdit trigger using the Spreadsheet to use the recordMacro() function
  var onEditTrigger = ScriptApp.newTrigger("recordMacro")
      .forSpreadsheet(spreadsheet)
      .onEdit()
      .create();

 
    }

// function deleteTriggers()
// Delete all triggers to elimate conflict and duplication
function deleteTriggers() {
 
    // Get All triggers
    var allTriggers = ScriptApp.getScriptTriggers(); // Get all triggers in the array allTriggers
    var numTriggers = allTriggers.length; // Get the number of triggers
    
    // For all triggers in the array
    for (var i=0; i<numTriggers; i++){

    ScriptApp.deleteTrigger(allTriggers[i]); // Delete the trigger 
    }
    
}
    
    
 


