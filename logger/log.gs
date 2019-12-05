//Logging the data which is fed as input by the user


function myFunction_prompt(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Getting to know you', 'May I know your name?', ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.YES) {
    Logger.log('The user's name is %s.', response.getResponseText());
  } else if (response.getSelectedButton() == ui.Button.NO) {
    Logger.log('The user didn't want to provide a name.');
  } else {
    Logger.log('The user clicked the close button in the dialog's title bar.');
  }
}
