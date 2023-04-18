function onOpen() {
  const ui = SpreadsheetApp.getUi();
  if (property.get("initialized") != "True") {
    ui.createMenu("初期設定はこちら！！！！！！！！！！！！！！")
        .addItem("初期設定", "initialize")
        .addItem("シートにボタンがある場合はそれを使っても良いです", "initialize")
        .addToUi();
  } else {
    ui.createMenu("機密情報の設定")
        .addItem("メールアドレス設定", "showSetMailAddressDialog")
        .addItem("APIのキー", "showSetKeyDialog")
        .addToUi();
  }
  if (property.get("last_error_occurred") == "True") {
    showErrorDialog()
  }
}