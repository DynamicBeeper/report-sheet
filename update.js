function updateProgress() {
  const sleep_time = 15000
  const spread_sheet = SpreadsheetApp.getActive();
  const progress_sheet = spread_sheet.getSheetByName("進捗");
  const key = property.get("report_api_key");
  const report = new report_api(key);
  const subject_data = new subjects_data();

  const total_counts_input = [];
  const progress_input = [];
  const today = new Date();
  progress_input.push(today);
  const subject_ids = subject_data.getIds();

  try {
    for (const subject_id of subject_ids) {
      progress = report.getProgress(subject_id);
      total_counts_input.push(...progress[0]);
      progress_input.push(...progress[1]);
      Utilities.sleep(sleep_time);
    }
  } catch(e) {
    const last_record_range = progress_sheet.getRange(progress_sheet.getLastRow(), 2, 1, progress_sheet.getLastColumn()-3);
    const last_record = last_record_range.getValues().flat();
    const progress_error_input = [today];
    progress_error_input.push(...last_record);
    progress_error_input.push(0, e.message);
    progress_sheet.appendRow(progress_error_input);
    const now = new Date();
    property.set("last_error_occurred", "True");
    property.set("last_error_message", e.message);
    property.set("last_error_time", now.toLocaleString("ja-JP"));
    return e;
  }
  progress_sheet.getRange(3, 2, 1, total_counts_input.length).setValues([total_counts_input]);
  progress_sheet.appendRow(progress_input);

  if (progress_sheet.getLastRow() > 4) {
    const yesterday = new Date();
    yesterday.setDate(today.getDate()-1);
    const diffs = calcDetailDiff(today, yesterday);
    progress_sheet.getRange(progress_sheet.getLastRow(), progress_sheet.getLastColumn()-1).setValue(diffs[1]);
  }
  return 0;
}