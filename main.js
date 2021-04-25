/**
 * onOpen function
 */
const onOpen = () => {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Accounting")
    .addItem("Add an item", "displaySidebarForm")
    .addToUi();
};

/**
 * include function
 * @param {string} filename
 * @return {Blob}
 */
const include = (filename) => {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
};

/**
 * displaySidebarForm function
 */
const displaySidebarForm = () => {
  const spreadsheetUi = SpreadsheetApp.getUi();
  let template = HtmlService.createTemplateFromFile("main_form");
  spreadsheetUi.showSidebar(template.evaluate());
};

/**
 * getUsers function
 * @param {Sheet} mainSheet
 * @return {object[]{}}
 */
const getUsers = (mainSheet) => {
  return mainSheet
    .getRange(2, 1, mainSheet.getLastRow() - 1, 4)
    .getValues()
    .filter((userRow) => userRow[0] && userRow[3])
    .map((userRow) => ({
      lastName: userRow[0],
      firstName: userRow[1],
      nickName: userRow[2],
    }));
};

/**
 * getManips function
 * @param {Sheet} manipSheet
 * @return {object[]{}}
 */
const getManips = (manipSheet) => {
  return manipSheet
    .getRange(2, 1, manipSheet.getLastRow() - 1, 3)
    .getDisplayValues()
    .map((manipRow) => {
      const dateData = manipRow[1].split("/").map((obj) => parseInt(obj));
      return {
        name: manipRow[0],
        date: new Date(dateData[2], dateData[1] - 1, dateData[0]),
      };
    });
};

/**
 * onSidebarLoaded function
 * @return {object}
 */
const onSidebarLoaded = () => {
  const mainSpreadsheet = SpreadsheetApp.getActive();
  return JSON.stringify({
    userArray: getUsers(mainSpreadsheet.getSheetByName(MAIN_SHEET)),
    manipArray: getManips(mainSpreadsheet.getSheetByName(MANIP_SHEET)),
  });
};

/**
 * onSubmit function
 * @return {object}
 */
const onSubmit = async (selectedUser, inputObject) => {
  console.log(selectedUser);
  console.log(inputObject);

  const mainSpreadsheet = SpreadsheetApp.getActive();
  const mainSheet = mainSpreadsheet.getSheetByName(MAIN_SHEET);
  const personNicknameList = mainSheet
    .getRange(2, 3, mainSheet.getLastRow() - 1, 1)
    .getValues()
    .map((nickName) => nickName[0])
    .filter((nickName) => nickName);
  let personIndex;
  for (personIndex in personNicknameList) {
    if (personNicknameList[personIndex] === selectedUser) break;
  }
  if (!personIndex) throw new Error("User not found");
  const personRow = personIndex + 3;

  let manipName;
  let input;
  let sheet;
  let promiseArray = [];
  for (manipName in inputObject) {
    promiseArray.push(
      new Promise((resolve, reject) => {
        input = inputObject[manipName];
        console.log(manipName);
        console.log(input);
        sheet = mainSpreadsheet.getSheetByName(manipName);
        if (!Object.keys(sheet).length) {
          const err = new Error(`No sheet found with the name : ${manipName}`);
          reject(err);
        }
        sheet.getRange(personRow, 7, 1, 1).setValue(input.amount);
        sheet
          .getRange(personRow, 9, 1, 2)
          .setValues([[input.option, input.date]]);
        resolve();
      })
    );
  }

  await Promise.all(promiseArray);
  return { status: "ok", message: "User input correctly handled." };
};
