const onOpen = () => {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Accounting')
    .addItem('Add an item', 'displaySidebarForm')
    .addToUi();
}

const include = (filename) => {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

const displaySidebarForm = () => {
  const spreadsheetUi = SpreadsheetApp.getUi();

  let template = HtmlService.createTemplateFromFile('main_form');
  spreadsheetUi.showSidebar(template.evaluate());
}

const getUsers = (mainSheet) => {
  return mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 4)
    .getValues()
    .filter((userRow) => (userRow[0] && userRow[3]))
    .map((userRow) => ({
      lastName: userRow[0],
      firstName: userRow[1],
      nickName: userRow[2],
      number: userRow[3],
    })
  )
};

const getManips = (manipSheet) => {
  return manipSheet.getRange(2, 1, manipSheet.getLastRow() - 1, 3)
    .getDisplayValues()
    .map((manipRow) => {
      const dateData = manipRow[1].split('/').map((obj) => parseInt(obj));
      return {
        'name': manipRow[0],
        'date': new Date(dateData[2], dateData[1] - 1, dateData[0])
      }
    });
}

const onSidebarLoaded = () => {
  const mainSpreadsheet = SpreadsheetApp.getActive();
  return JSON.stringify({
    "userArray": getUsers(mainSpreadsheet.getSheetByName(MAIN_SHEET)),
    "manipArray": getManips(mainSpreadsheet.getSheetByName(MANIP_SHEET))
  })
}
