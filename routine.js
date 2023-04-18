function routine_update() {
  const exception = updateProgress();
  ScriptApp.newTrigger("routine_mail").timeBased().after(10000).create();
  if (exception) {
    const message = "APIにアクセスできませんでした。" + exception.message;
    throw message;
  }
}

function routine_mail() {
  const sleep_time = 2000
  const today = new Date()
  if (property.get("daily_mail") == "True") {
    sendDailyMail();
  }
  Utilities.sleep(sleep_time);
  if (property.get("weekly_mail") == "True" && today.getDay() == 0){
    sendWeaklyMail();
  }
  Utilities.sleep(sleep_time);
  if (property.get("monthly_mail") == "True" && today.getDate() == 1){
    sendMonthlyMail();
  }
}