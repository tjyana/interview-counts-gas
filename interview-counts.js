



function countInterviews() {
  // Get date ranges
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("counts");
  const startCell = sheet.getRange("B1").getValue();
  const endCell   = sheet.getRange("B2").getValue();
  const CAL_ID = "moneyforward.co.jp_10thsudeuufd1ej0a0i95pv9us@group.calendar.google.com";
  const START = new Date(startCell);   // inclusive
  const END_RAW  = new Date(endCell);   // inclusive

  const END = new Date(END_RAW);
  END.setDate(END.getDate() + 1);


  const keywords = [
  // Casual
  { label: "カジュアル面談", pattern: /Casual Interview|カジュアル|casual/i },

  // 1st
  { label: "1次面接", pattern: /First Interview|1st interview|1st|1次面接|１次面接|一次面接/i },

  // 2nd
  { label: "2次面接", pattern: /Second Interview|2nd interview|2nd|2次面接|２次面接|二次面接|2次|二次|2st/i },

  // HR
  { label: "人事面談", pattern: /HR Mee|人事面談|人事面接/i },

  // Final
  { label: "最終面接", pattern: /Final Interview|最終面接/i },

  // その他面接
  { label: "その他面接", pattern: /追加|技術面接|カルチャーマッチ|3rd Interview|3次|三次/i },

  // Offer
  { label: "オファー面談", pattern: /Offer Meeting|オファー面談|オフィスツアー|顔合わせ|座談会|期待値|Office Tour/i },

  // Irregulars
  { label: "その他面談・会食", pattern: /会食|派遣面談|入社前面談|採用ランチ|室長面接|フォロー面談|Meeting with|職場面接|再面談|Introductory|オファー前|Online Meeting|Recruitment Lunch/i },

  // 対応済み/辞退
  { label: "(対応ステータス/リマインド系)", pattern: /対応済|対応中|辞退|リマインド済|格納済|対応不要|ASHIATO|技術課題|TOEIC|通訳|総会|Backcheck|Work block|System Design Interview/i },

  // 祝日
  { label: "(祝日)", pattern: /の日|祝日/i },
 
 
  ];

  const cal = CalendarApp.getCalendarById(CAL_ID);
  const events = cal.getEvents(START, END);

  const counts = {};
  keywords.forEach(k => counts[k.label] = 0);
  counts["(Other)"] = 0;

const classified = [];

events.forEach(e => {
  const title = e.getTitle() || "";
  let matched = false;
  let matchedLabel = "(Other)";

  // Log every event title
  Logger.log("Event: " + title);

  // Match against keywords
  for (const k of keywords) {
    if (k.pattern.test(title)) {
      Logger.log(" → MATCHED: " + k.label);
      counts[k.label]++;
      matched = true;
      matchedLabel = k.label;
      break;
    }
  }

  if (!matched) {
    Logger.log(" → OTHER");
    counts["(Other)"]++;
  }

  classified.push([title, matchedLabel]);

});

  // output counts
  sheet.getRange("A7:B18").clearContent();
  // sheet.getRange("A4").setValue("Type");
  // sheet.getRange("B4").setValue("Count");

  let row = 7;
  Object.entries(counts).forEach(([type, c]) => {
    sheet.getRange(row, 1).setValue(type);
    sheet.getRange(row, 2).setValue(c);
    row++;
  });

  // output events
  sheet.getRange("A21:B999").clearContent();
  // sheet.getRange("A18").setValue("Event Title");
  // sheet.getRange("B18").setValue("Category");

  // sort by category, then title
  classified.sort((a, b) => {
    const catA = a[1], catB = b[1];
    if (catA < catB) return -1;
    if (catA > catB) return 1;

    const titleA = a[0], titleB = b[0];
    if (titleA < titleB) return -1;
    if (titleA > titleB) return 1;
    return 0;
  });


  if (classified.length > 0) {
    sheet.getRange(21, 1, classified.length, 2).setValues(classified);
  }


}






function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "counts") return;

  const a1 = e.range.getA1Notation();
  if (a1 === "B1" || a1 === "B2") {
    countInterviews();
  }
}
