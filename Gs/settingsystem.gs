const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
const idfolder = settingSheet.getRange('B1').getDisplayValue();
const sigFolder = settingSheet.getRange('B2').getDisplayValue();
const imageFolder = settingSheet.getRange('B3').getDisplayValue();
const sheetData = settingSheet.getRange('B4').getDisplayValue();
const sheetDataSet = settingSheet.getRange('B5').getDisplayValue();
const clientId = settingSheet.getRange('B6').getDisplayValue();
const clientSecret = settingSheet.getRange('B7').getDisplayValue();
const redirectUri = settingSheet.getRange('B8').getDisplayValue();
const logoUrl = settingSheet.getRange('B9').getDisplayValue();
const nameSystem = settingSheet.getRange('B10').getDisplayValue();

const getSet = () => {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
  var data = ss.getRange("B1:B").getDisplayValues();
  return data;
}

const getLineSet = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
  const data = sheet.getRange('A1:B' + sheet.getLastRow()).getValues();

  const settings = {};

  data.forEach(row => {
    const key = row[0];
    const value = row[1];
    settings[key] = value;
  });

  return settings;
};

const settingGS = (data) => {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Setting');
  var valuesToSet = [];
  for (let i = 1; i <= 11; i++) {
    valuesToSet.push([data[`set${i}`]]);
  }
  var range = sheet.getRange(1, 2, valuesToSet.length, 1);
  range.setValues(valuesToSet);
}

const selectDataFromSheet = (sheetName) => {
  var sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName(sheetName);
  var getLastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 2, getLastRow - 1, 1).getValues().flat();
  return data;
}

const selectDepartment = () => selectDataFromSheet("Department");
const selectGroup = () => selectDataFromSheet("Group");
const selectObjective = () => selectDataFromSheet("Objective");

const getTodos = (sheetName) => {
  var sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName(sheetName);
  var data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();
  return data.flat().filter(Boolean);
}

const saveTodos = (data) => {
  var sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName(data.sheetName);
  sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).clearContent();
  data.todos.forEach((todo, index) => {
    sheet.getRange(index + 2, 2).setValue(todo);
  });
}

const getsetMenuItems = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UsersMenu");
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  return data;
}

const updateMenuStatus = (menuItem, role, isChecked) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UsersMenu");
  const data = sheet.getDataRange().getValues();
  const index = data.findIndex(row => row[1] === menuItem);
  if (index !== -1) {
    const roleColumn = role === 'SuperAdmin' ? 3 : role === 'Admin' ? 4 : role === 'SuperUser' ? 5 : 6;
    const range = sheet.getRange(index + 1, roleColumn);
    range.setValue(isChecked ? "TRUE" : "FALSE");
  }
}

const getMenuItems = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UsersMenu");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const menuItems = {};

  for (let i = 1; i < data.length; i++) {
    const item = data[i][1];
    menuItems[item] = {};
    for (let j = 2; j < headers.length; j++) { 
      const cellValue = String(data[i][j]).toUpperCase() || "FALSE"; 
      menuItems[item][headers[j]] = cellValue === "TRUE";
    }
  }
  return menuItems;
};

function shortenURL(longURL) {
  var apiUrl = "http://tinyurl.com/api-create.php?url=" + encodeURI(longURL);
  var response = UrlFetchApp.fetch(apiUrl);
  return response.getContentText();
}


