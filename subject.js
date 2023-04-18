class subjects_data {
  constructor() {
    const sheet = SpreadsheetApp.getActive().getSheetByName("subjects_data");
    const records = sheet.getDataRange().getValues();
    const subjects = [];
    for (const record of records) {
      const [id, title, total_count] = record;
      const subject = {};
      subject.id = id;
      subject.title = title;
      subject.total_count = total_count;
      subjects.push(subject);
    }
    this.subjects = subjects;
  }

  getIds() {
    const ids = [];
    for (const subject of this.subjects) {
      ids.push(subject.id);
    }
    return ids;
  }

  getTitles() {
    const titles = [];
    for (const subject of this.subjects) {
      titles.push(subject.title);
    }
    return titles;
  }

  getTotalCounts() {
    const total_counts = [];
    for (const subject of this.subjects) {
      total_counts.push(subject.total_count);
    }
    return total_counts;
  }

  getSubjects() {
    return this.subjects;
  }
}