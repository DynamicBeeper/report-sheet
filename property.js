class property {
  static get(key) {
    return PropertiesService.getScriptProperties().getProperty(key);
  }
  
  static set(key, value) {
    PropertiesService.getScriptProperties().setProperty(key, value);
  }
}


function applySettings() {
  const spread_sheet = SpreadsheetApp.getActive();
  const settings_sheet = spread_sheet.getSheetByName("設定");
  const records = settings_sheet.getRange(2, 1, settings_sheet.getLastRow()-1, 3).getValues();
  
  for (record of records) {
    const key = record[2];
    const value = record[1];
    if (key) {
      if (value === true) {
        property.set(key, "True");
      } else if (value === false) {
        property.set(key, "False");
      }
    }
  }
  showSuccessDialog();
}