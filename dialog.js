function showSetKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  while (true) {
    const answer = ui.prompt("キー設定", "N予備校のキーを入力してください", ui.ButtonSet.OK_CANCEL);
    if (answer.getSelectedButton() == ui.Button.CANCEL || answer.getSelectedButton() == ui.Button.CLOSE) {
      return -1;
    } else {
      const report = new report_api(answer.getResponseText());
      const exception = report.verifyKey();
      if (exception == 0) {
        report.getAllSubjects();
        property.set("report_api_key", answer.getResponseText());
        showSuccessDialog();
        return 0;
      } else {
        ui.alert("失敗しました", exception.message, ui.ButtonSet.OK);
      }
    }
  }
}

function showSetMailAddressDialog() {
  const ui = SpreadsheetApp.getUi();
  while (true) {
    const answer = ui.prompt("メールアドレス設定", "送信先のメールアドレスを入力してください", ui.ButtonSet.OK_CANCEL);
    if (answer.getSelectedButton() == ui.Button.CANCEL || answer.getSelectedButton() == ui.Button.CLOSE) {
      return -1;
    } else {
      try {
        sendTestMail(answer.getResponseText());
        const confirm = ui.alert("メールの確認", answer.getResponseText()+"にメールを送信しました。受信できましたか?", ui.ButtonSet.YES_NO);
        if (confirm == ui.Button.YES) {
          property.set("mail_address", answer.getResponseText())
          showSuccessDialog();
          return 0;
        }
      } catch(e) {
        ui.alert("失敗しました", e.message, ui.ButtonSet.OK);
      }
    }
  }
}

function showErrorDialog() {
  const ui = SpreadsheetApp.getUi();
  const error_time = property.get("last_error_time");
  const error_message = property.get("last_error_message");
  const dialog_message = "最終エラー発生日時:\n" + error_time + "\n最終エラー内容:\n" + error_message;
  ui.alert("エラーが発生しました", dialog_message, ui.ButtonSet.OK);
  property.set("last_error_occurred", "False");
}

function showSuccessDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("成功！", "", ui.ButtonSet.OK);
}