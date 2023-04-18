function initialize() {
  const now = new Date();
  property.set("last_error_occurred", "True");
  property.set("last_error_message", "初期設定が正常に終了しませんでした。もう一度やり直してください。");
  property.set("last_error_time", now.toLocaleString("ja-JP"));
  while (property.get("report_api_key") == null) {
    showSetKeyDialog();
  }

  while (property.get("mail_address") == null) {
    showSetMailAddressDialog();
  }

  const triggers = ScriptApp.getProjectTriggers();
  for (trigger of triggers) {
    if (trigger.getHandlerFunction() == "routine_update") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  const hour = Math.floor(Math.random() * 7);
  ScriptApp.newTrigger("routine_update")
      .timeBased()
      .everyDays(1)
      .atHour(hour)
      .create();

  init_subjects();
  setGradients("#00FF00", "#FBBC04");
  applySettings();
  property.set("initialized", "True");
  property.set("last_error_occurred", "False");
  const spread_sheet = SpreadsheetApp.getActive();
  spread_sheet.getSheetByName("初期設定").hideSheet();
}


function init_subjects(){
  const spread_sheet = SpreadsheetApp.getActive();
  const subjects_data_sheet = spread_sheet.getSheetByName("subjects_data");
  const progress_sheet = spread_sheet.getSheetByName("進捗");
  const key = property.get("report_api_key");
  const report = new report_api(key);
  const subjects = report.getAllSubjects();

  const subjects_data_input = []
  for (subject of subjects) {
    subjects_data_input.push([subject["id"], subject["title"], subject["total_count"]])
  }
  subjects_data_sheet.clear();
  subjects_data_sheet.getRange(1, 1, subjects_data_input.length, subjects_data_input[0].length).setValues(subjects_data_input);

  const progress_input = [];
  for (subject of subjects) {
    progress_input.push(subject["title"]);
    [...Array(subject["total_count"])].map((_, i) => progress_input.push(i+1));
  }
  progress_input.push("前日との差分", "エラー内容")
  progress_sheet.getRange(1, 2, 1, progress_input.length).setValues([progress_input]);
  progress_sheet.autoResizeColumns(1, progress_input.length);
  progress_sheet.setColumnWidth(1, 72);
}


function setGradients(color1, color2) {
  const spread_sheet = SpreadsheetApp.getActive();
  const progress_sheet = spread_sheet.getSheetByName("進捗");
  const last_column = progress_sheet.getLastColumn();
  const rules = [];
  for (let column=2; column<last_column-1; column++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
          .setGradientMaxpoint(color1)
          .setGradientMinpoint("#FFFFFF")
          .setRanges([progress_sheet.getRange(3, column, 1000)])
          .build());
  }
  rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint(color2)
        .setGradientMinpoint("#FFFFFF")
        .setRanges([progress_sheet.getRange(3, last_column-1, 1000)])
        .build());
  progress_sheet.setConditionalFormatRules(rules);
}