var patching_date = '02/16/2021 8-12 p.m.';
var hosts_sheet_name = 'hosts';
var host_patching_schedule = 'patchingSchedule';
var form_title = '01 Patching Exceptions Form'; //'Mac Patching Exceptions 20210216'
var form_description = 'Please submit this form for systems to be rescheduled from standard patching.';
var host_selection_title = 'Systems';
var host_selection_help_text = 'Please select all systems the you would like to reschedule.';
var rescheduled_title = 'Date and Time';
var rescheduled_help_text = 'Please select the date and time on which you would like to reschedule the patch of these system(s).';
var comment_title = 'Comment';
var comment_help_text = 'Provide a comment about rescheduling the patching of system(s)';

function setUpPatchingExceptionsForm() {
  var hosts_sheet = SpreadsheetApp.getActive();
  var range = hosts_sheet.getDataRange();
  var host_names = range.getValues();
  setUpForm(hosts_sheet,host_names);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(hosts_sheet).onFormSubmit().create();
  ScriptApp.newTrigger('onEdit').forSpreadsheet(hosts_sheet).onEdit().create();
  hosts_sheet.getSheetByName(hosts_sheet_name).setName(host_patching_schedule);
  var lastRow = hosts_sheet.getLastRow();
  var lastColumn = hosts_sheet.getLastColumn();
  hosts_sheet.getRange("B1").setValue('Date and Time');
  hosts_sheet.getRange("B2:B"+lastRow).setValue(patching_date);
}

function setUpForm(hosts_sheet, host_names) {
  var form = FormApp.create(form_title);
  form.setCollectEmail(true)
  form.setDescription(form_description);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, hosts_sheet.getId());
  form.setCollectEmail(true)
  var hostSelection = form.addCheckboxItem().setRequired(true);
  hostsArray = []
  for(var i=1;i<host_names.length;i++) {
    hostsArray.push(host_names[i][0])
  }
  hostSelection.setTitle(host_selection_title).setChoiceValues(hostsArray);
  hostSelection.setHelpText(host_selection_help_text)
  var rescheduled = form.addDateTimeItem();
  rescheduled.setTitle(rescheduled_title);
  rescheduled.setHelpText(rescheduled_help_text);
  var comment = form.addParagraphTextItem();
  comment.setTitle(comment_title);
  comment.setHelpText(comment_help_text)
}

function onFormSubmit(e) {
  onEdit(e);
}

function onEdit(e) {
  var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  responseSheet.activate();
  var lastRowNumber = responseSheet.getLastRow();
  var lastColumnNumber = responseSheet.getLastColumn();
  var timestamp = responseSheet.getRange(lastRowNumber,1).getValue();
  var email = responseSheet.getRange(lastRowNumber,2).getValue();;
  var optionsArray = responseSheet.getRange(lastRowNumber,3).getValue().split(", ");
  var dateTime = responseSheet.getRange(lastRowNumber,4).getValue()
  var comment = responseSheet.getRange(lastRowNumber,5).getValue();
  editSchedule(optionsArray,dateTime);
}

function editSchedule(optionsArray,dateTime) {
  var patchSchedule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(host_patching_schedule);
  patchSchedule.activate();
  var data = patchSchedule.getDataRange();
  var rows = data.getNumRows();
  var columns = data.getNumColumns();
  for(i=2;i<rows;i++) {
    var host = data.getCell(i, 1).getValue();
    for(j=0;j<optionsArray.length;j++){
      if(host == optionsArray[j]) {
        //Logger.log(dateTime)
        data.getCell(i,2).setValue(dateTime);
      }
    }    
  }
}
