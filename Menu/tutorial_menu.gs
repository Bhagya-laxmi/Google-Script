function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('My Menu')
      .addItem('My Menu Item', 'myFunction')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('My Submenu')
          .addItem('One Submenu Item', 'mySecondFunction')
          .addItem('Another Submenu Item', 'myThirdFunction'))
      .addToUi();
}

function mySecondFunction(){
  Browser.msgBox('second function');

}

function myThirdFunction(){
  var name = Browser.inputBox('third function');
  Browser.msgBox(name);
}
