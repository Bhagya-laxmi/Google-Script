function myTutorial(){
  var name = Browser.inputBox('Enter your name');
  Browser.msgBox(name);
  // The code below sets the value of name to the name input by the user, or 'cancel'.
  var name1 = Browser.inputBox('Enter your name', Browser.Buttons.OK_CANCEL);
  
  // The code below sets the value of name to the name input by the user, or 'cancel'.
  var name2 = Browser.inputBox('ID Check', 'Enter your name', Browser.Buttons.OK_CANCEL);
  
  
  // The code below displays "hello world" in a dialog box with an OK button
  Browser.msgBox('hello world');
  
  // The code below displays "hello world" in a dialog box with OK and Cancel buttons.
  Browser.msgBox('hello world', Browser.Buttons.OK_CANCEL);
  
  // The code below displays "hello world" in a dialog box with a custom title and Yes and
  // No buttons
  Browser.msgBox('Greetings', 'hello world', Browser.Buttons.YES_NO);
}
