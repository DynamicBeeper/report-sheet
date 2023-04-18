class report_api {
  constructor(key) {
    this.domain_name = "example.com"
    this.key = key;
  }

  verifyKey() {
    try {
      this.getJson("https://api."+this.domain_name+"/v1/users");
    } catch(e) {
      return e;
    }
    return 0;
  }

  getAllSubjects(){
    const package_ids = [];
    const packages = this.getJson("https://api."+this.domain_name+"/v2/material/packages")["services"][0]["packages"];
    for (const package_ of packages) {
      package_ids.push(package_["id"]);
    }
    const subjects = [];
    for (const package_id of package_ids) {
      const subject_categories = this.getJson("https://api."+this.domain_name+"/v2/material/packages/"+package_id)["package"]["subject_categories"];
      for (const subject_category of subject_categories) {
        const courses = subject_category["courses"];
        for (const course of courses) {
          if (course["selected"] == true) {
            const subject = {
              "id": course["id"],
              "title": subject_category["title"],
              "total_count": course["progress"]["total_count"],
            };
            subjects.push(subject)
          }
        }
      }
    }
    return subjects;
  }

  getProgress(subject_id){
    const url = "https://api."+this.domain_name+"/v2/material/courses/" + subject_id;
    const json = this.getJson(url);
    const course = json["course"];
    const chapters = course["chapters"];
    const progress = [[], []];
    progress[0].push(course["progress"]["total_count"]);
    progress[1].push(course["progress"]["passed_count"]);
    for (const chapter of chapters) {
      progress[0].push(chapter["progress"]["total_count"]);
      progress[1].push(chapter["progress"]["passed_count"]);
    }
    return progress;
  }

  getJson(url){
    const cookie = {
      "_zane_session": this.key
    }
    let cookieStr = "";
    for (const i in cookie) {
      cookieStr += i + "=" + cookie[i] + "; ";
    }
    const headers = { Cookie: cookieStr };
    const get_options = {
      method: "get",
      headers: headers,
    };
    const response = UrlFetchApp.fetch(url, get_options);
    const text = response.getContentText('UTF-8');
    const json = JSON.parse(text);
    return json;
  }
}
