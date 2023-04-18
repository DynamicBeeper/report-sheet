function calcDetailDiff(minuend_day, subtrahend_day) {
  const spread_sheet = SpreadsheetApp.getActive();
  const progress_sheet = spread_sheet.getSheetByName("進捗");
  const rows = findDatesRow([minuend_day, subtrahend_day]);
  const last_column = progress_sheet.getLastColumn();
  if (rows[0] == -1 || rows[1] == -1) {
    throw new Error("Date not found!");
  }
  const minuend_progress = progress_sheet.getRange(rows[0], 2, 1, last_column-3).getValues()[0];
  const subtrahend_progress = progress_sheet.getRange(rows[1], 2, 1, last_column-3).getValues()[0];
  const subject_data = new subjects_data();
  const subjects = subject_data.getSubjects();
  const diff = [0, 0, []];
  let index = 0;
  for (const subject of subjects) {
    const diff_big = minuend_progress[index]-subtrahend_progress[index];
    index += 1;
    let diff_little = 0;
    for (let i=0; i<subject.total_count; i++) {
      diff_little += minuend_progress[index]-subtrahend_progress[index];
      index += 1;
    }
    diff[0] += diff_big;
    diff[1] += diff_little;
    diff[2].push([subject.title, diff_big, diff_little]);
  }
  return diff;
}


function getDiffs(date_array) {
  const spread_sheet = SpreadsheetApp.getActive();
  const progress_sheet = spread_sheet.getSheetByName("進捗");
  const last_column = progress_sheet.getLastColumn();
  const last_row = progress_sheet.getLastRow();
  const diffs_column = progress_sheet.getRange(1, last_column-1, last_row).getValues().flat();
  const rows = findDatesRow(date_array);
  let diffs = [];
  for (let i=0; i<date_array.length; i++) {
    const row = rows[i]
    if (row == -1) {
      diffs.push(-1);
    } else {
      diffs.push(diffs_column[row-1]);
    }
  }
  return diffs;
}


function findDatesRow(date_array) {
  const spread_sheet = SpreadsheetApp.getActive();
  const progress_sheet = spread_sheet.getSheetByName("進捗");
  const date_column = progress_sheet.getRange(1, 1, progress_sheet.getLastRow()).getValues().flat();
  const dates_row = [];
  for (let i=0; i<date_column.length; i++) {
    if (isDateValid(date_column[i])){
      date_column[i] = Utilities.formatDate(date_column[i], "JST", "yyyy/MM/dd");
    }
  }
  for (date of date_array) {
    const target_date = Utilities.formatDate(date, "JST", "yyyy/MM/dd");
    const index = date_column.indexOf(target_date);
    if (index == -1) {
      dates_row.push(-1);
    } else {
      dates_row.push(index+1);
    }
  }
  return dates_row;
}


function isDateValid(obj) {
  if ( Object.prototype.toString.call(obj) !== "[object Date]" ){
    return false;
  }
  return !isNaN(obj.getTime());
}


function createDateArray(start_date, end_date) {
  const diff_day = Math.round((end_date.getTime() - start_date.getTime()) / (1000 * 60 * 60 * 24));
  let date_array = [];
  for (let i=0; i<=diff_day; i++) {
    const date = new Date();
    date.setDate(start_date.getDate()+i);
    date_array.push(date);
  }
  return date_array;
}