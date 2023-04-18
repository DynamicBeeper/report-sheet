function sendTestMail(mail_address) {
  const subject = "レポート進捗管理シートにメールアドレスが登録されました";
  const recipient = mail_address;
  const text = "このメールはテストです。HTMLメールが表示できませんでした。定期的な通知メールはHTML形式にのみ対応しています。";
  let html = HtmlService.createHtmlOutput("<p>このメールはテストです。</p>").getContent();
  const mail_options = {name: "レポート進捗管理シート", htmlBody: html};
  GmailApp.sendEmail(recipient, subject, text, mail_options);
}


function sendDailyMail() {
  const today = new Date();
  const yesterday = new Date();
  yesterday.setDate(today.getDate()-1);
  const subject = Utilities.formatDate(yesterday, "JST", "MM月dd日") + "のレポート進捗状況";
  const recipient = property.get("mail_address");
  const text = "HTMLメールが表示できませんでした";
  const html = generateDailyMailBody();
  const mail_options = {name: "レポート進捗管理シート", htmlBody: html};
  GmailApp.sendEmail(recipient, subject, text, mail_options);
}


function sendWeaklyMail(){
  const subject = "先週のレポート進捗状況";
  const recipient = property.get("mail_address");
  const text = "HTMLメールが表示できませんでした";
  const html = generateWeeklyMailBody();
  const week_chart = generateWeekChart().getBlob().setName("日別進捗").setContentType("image/jpeg");
  const mail_options = { name: "レポート進捗管理シート", htmlBody: html, inlineImages: {chart1: week_chart}};
  GmailApp.sendEmail(recipient, subject, text, mail_options);
}


function sendMonthlyMail(){
  const subject = "先月のレポート進捗状況";
  const recipient = property.get("mail_address");
  const text = "HTMLメールが表示できませんでした";
  const html = generateMonthlyMailBody();
  const month_chart = generateMonthChart().getBlob().setName("日別進捗").setContentType("image/jpeg");
  const mail_options = { name: "レポート進捗管理シート", htmlBody: html, inlineImages: {chart1: month_chart}};
  GmailApp.sendEmail(recipient, subject, text, mail_options);
}


function generateDailyMailBody() {
  let html = HtmlService.createHtmlOutputFromFile("Mail.html").getContent();
  const today = new Date();
  const yesterday = new Date();
  yesterday.setDate(today.getDate()-1);
  const diffs = calcDetailDiff(today, yesterday);
  let passed_count = "";
  if (diffs[0] > 0) {
    passed_count = diffs[1] + "コマ,　" + diffs[0] + "個";
  } else {
    passed_count = diffs[1] + "コマ";
  }
  html = html.replace("[title]", "日刊レポート進捗メール")
  html = html.replaceAll("[date]", Utilities.formatDate(yesterday, "JST", "MM月dd日"));
  html = html.replace("[passed_count]", passed_count);
  html = html.replace("[table_records]", generateTableRecords(diffs[2]));
  html = html.replace("[chart]", '');
  html = html.replaceAll("[spreadsheet_url]", SpreadsheetApp.getActive().getUrl());
  html = html.replaceAll("[form_url]", "");
  return html;
}


function generateWeeklyMailBody() {
  let html = HtmlService.createHtmlOutputFromFile("Mail.html").getContent();
  let passed_count = "";
  const today = new Date();
  const one_week_ago = new Date();
  one_week_ago.setDate(today.getDate()-7);
  const diffs = calcDetailDiff(today, one_week_ago);
  if (diffs[0] > 0) {
    passed_count = diffs[1] + "コマ,　" + diffs[0] + "個";
  } else {
    passed_count = diffs[1] + "コマ";
  }
  html = html.replace("[title]", "週刊レポート進捗メール")
  html = html.replaceAll("[date]", "先週");
  html = html.replace("[passed_count]", passed_count);
  html = html.replace("[table_records]", generateTableRecords(diffs[2]));
  html = html.replace("[chart]", '<div><img src="cid:chart1" width=”353″></div>');
  html = html.replaceAll("[spreadsheet_url]", SpreadsheetApp.getActive().getUrl());
  html = html.replaceAll("[form_url]", "");
  return html;
}


function generateMonthlyMailBody() {
  let html = HtmlService.createHtmlOutputFromFile("Mail.html").getContent();
  let passed_count = "";
  const today = new Date();
  const yesterday = new Date();
  yesterday.setDate(today.getDate()-1);
  const month_start_date = new Date(yesterday.getFullYear(), yesterday.getMonth(), 1);
  const diffs = calcDetailDiff(today, month_start_date);
  if (diffs[0] > 0) {
    passed_count = diffs[1] + "コマ,　" + diffs[0] + "個";
  } else {
    passed_count = diffs[1] + "コマ";
  }
  html = html.replace("[title]", "月刊レポート進捗メール")
  html = html.replaceAll("[date]", "先月");
  html = html.replace("[passed_count]", passed_count);
  html = html.replace("[table_records]", generateTableRecords(diffs[2]));
  html = html.replace("[chart]", '<div><img src="cid:chart1" width=”353″></div>');
  html = html.replaceAll("[spreadsheet_url]", SpreadsheetApp.getActive().getUrl());
  html = html.replaceAll("[form_url]", "");
  return html;
}


function generateTableRecords(records) {
  let table = "";
  for (const record of records) {
    if (record[2] > 0){
      table += '<tr><td style="color: #000; font-size: 14px; padding: 10px; border: 1px solid #d2d3d3;">';
      table += record[0];
      table += '</td><td style="color: #000; font-size: 14px; padding: 10px; border: 1px solid #d2d3d3; text-align: center;">';
      table += record[2];
      table += '</td><td style="color: #000; font-size: 14px; padding: 10px; border: 1px solid #d2d3d3; text-align: center;">';
      table += record[1];
      table += '</td></tr>';
    }
  }
  return table;
}


function generateWeekChart() {
  const today = new Date();
  const start_date = new Date();
  start_date.setDate(today.getDate()-6);
  const date_array = createDateArray(start_date, today);
  const diffs = getDiffs(date_array);
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, "Date")
    .addColumn(Charts.ColumnType.NUMBER, "Progress");
  for (let i=0; i<date_array.length; i++) {
    const date = date_array[i].getDate()-1;
    dataTable.addRow([String(date), diffs[i]]);
  }
  const chartBuilder = Charts.newColumnChart()
     .setTitle('日別進捗')
     .setDimensions(368, 250)
     .setDataTable(dataTable)
     .setLegendPosition(Charts.Position.NONE)
     .setOption("chartArea", {left:40,right:20,top:20,width:"80%",height:"85%"})
     .setOption("titleTextStyle" ,{fontSize: 30});
  const chart = chartBuilder.build();
  return chart;
}

function generateMonthChart() {
  const today = new Date();
  const yesterday = new Date();
  yesterday.setDate(today.getDate()-1);
  const start_date = new Date(yesterday.getFullYear(), yesterday.getMonth(), 2);
  const date_array = createDateArray(start_date, today);
  const diffs = getDiffs(date_array);
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.NUMBER, "Date")
    .addColumn(Charts.ColumnType.NUMBER, "Progress");
  for (let i=0; i<date_array.length; i++) {
    const date = date_array[i].getDate()-1;
    dataTable.addRow([date, diffs[i]]);
  }
  const chartBuilder = Charts.newColumnChart()
     .setTitle('日別進捗')
     .setDimensions(368, 250)
     .setDataTable(dataTable)
     .setLegendPosition(Charts.Position.NONE)
     .setOption("chartArea", {left:40,right:20,top:20,width:"80%",height:"85%"})
     .setOption("titleTextStyle" ,{fontSize: 30});
  const chart = chartBuilder.build();
  return chart;
}