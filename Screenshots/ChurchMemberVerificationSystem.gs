function setupChurchMemberSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename("Church Member Verification System");

  const requiredSheets = [
    "Members",
    "Form Responses",
    "Dashboard",
    "Follow Up",
    "Settings"
  ];

  requiredSheets.forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      ss.insertSheet(name);
    }
  });

  setupMembersSheet(ss);
  setupResponsesSheet(ss);
  setupDashboardSheet(ss);
  setupFollowUpSheet(ss);
  setupSettingsSheet(ss);
  createChurchForm(ss);
  generateFakeResponses(500);
  buildDashboard();
  buildFollowUpList();

  SpreadsheetApp.getUi().alert("Project created successfully with 500 records, charts, dashboard, ministry stats, attendance analysis, and follow up list.");
}

function setupMembersSheet(ss) {
  const sheet = ss.getSheetByName("Members");
  sheet.clear();

  const headers = [
    "Member ID",
    "Full Name",
    "Phone",
    "Email",
    "Age Group",
    "Gender",
    "Ministry",
    "Active Member",
    "Attendance Frequency",
    "Preferred Contact",
    "Contact Info Correct",
    "Wants To Volunteer",
    "Status",
    "Missing Info",
    "Follow Up Required",
    "Follow Up Reason",
    "Last Updated"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length);
}

function setupResponsesSheet(ss) {
  const sheet = ss.getSheetByName("Form Responses");
  sheet.clear();

  const headers = [
    "Timestamp",
    "Full Name",
    "Phone",
    "Email",
    "Age Group",
    "Gender",
    "Ministry",
    "Active Member",
    "Attendance Frequency",
    "Preferred Contact",
    "Contact Info Correct",
    "Wants To Volunteer",
    "Notes"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length);
}

function setupDashboardSheet(ss) {
  const sheet = ss.getSheetByName("Dashboard");
  sheet.clear();
}

function setupFollowUpSheet(ss) {
  const sheet = ss.getSheetByName("Follow Up");
  sheet.clear();

  const headers = [
    "Member ID",
    "Full Name",
    "Phone",
    "Email",
    "Ministry",
    "Active Member",
    "Attendance Frequency",
    "Preferred Contact",
    "Follow Up Reason"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length);
}

function setupSettingsSheet(ss) {
  const sheet = ss.getSheetByName("Settings");
  sheet.clear();

  const rows = [
    ["Setting", "Value"],
    ["Church Name", "Example Community Church"],
    ["Admin Email", Session.getActiveUser().getEmail()],
    ["Project Name", "Church Member Verification System"]
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  styleHeader(sheet, 2);
}

function createChurchForm(ss) {
  const form = FormApp.create("Church Member Verification Form");

  form.setDescription("Please verify your church membership details so church leadership can keep accurate records.");

  form.addTextItem().setTitle("Full Name").setRequired(true);
  form.addTextItem().setTitle("Phone").setRequired(true);
  form.addTextItem().setTitle("Email").setRequired(false);

  form.addMultipleChoiceItem()
    .setTitle("Age Group")
    .setChoiceValues(["Under 18", "18 to 25", "26 to 35", "36 to 45", "46 to 60", "60 plus"])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle("Gender")
    .setChoiceValues(["Male", "Female", "Prefer not to say"])
    .setRequired(true);

  form.addListItem()
    .setTitle("Ministry")
    .setChoiceValues(["Youth", "Worship", "Ushering", "Media", "Children", "Intercession", "Outreach", "Hospitality", "None"])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle("Active Member")
    .setChoiceValues(["Yes", "No"])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle("Attendance Frequency")
    .setChoiceValues(["Weekly", "Twice a month", "Monthly", "Rarely"])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle("Preferred Contact")
    .setChoiceValues(["WhatsApp", "SMS", "Email", "Phone Call"])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle("Contact Info Correct")
    .setChoiceValues(["Yes", "No"])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle("Wants To Volunteer")
    .setChoiceValues(["Yes", "No", "Maybe"])
    .setRequired(true);

  form.addParagraphTextItem().setTitle("Notes").setRequired(false);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  const settings = ss.getSheetByName("Settings");
  settings.getRange("A6").setValue("Form Link");
  settings.getRange("B6").setValue(form.getPublishedUrl());
}

function generateFakeResponses(totalRecords) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responses = ss.getSheetByName("Form Responses");
  const members = ss.getSheetByName("Members");

  responses.getRange(2, 1, responses.getMaxRows(), responses.getMaxColumns()).clearContent();
  members.getRange(2, 1, members.getMaxRows(), members.getMaxColumns()).clearContent();

  const firstNames = [
    "Thabo", "Lerato", "Sipho", "Nomsa", "Karabo", "Blessing", "Faith", "Mpho", "Tshepo", "Amanda",
    "Grace", "Bongani", "Ayanda", "Neo", "Kabelo", "Zanele", "Nokuthula", "Sibusiso", "Refilwe", "Itumeleng",
    "Lindiwe", "Themba", "Dineo", "Katlego", "Boitumelo", "Nhlanhla", "Precious", "Hope", "Tumelo", "Lebogang"
  ];

  const surnames = [
    "Mokoena", "Dlamini", "Mahlangu", "Ndlovu", "Mabena", "Nkosi", "Mhlongo", "Khumalo", "Molefe", "Radebe",
    "Maseko", "Mthembu", "Mabaso", "Motaung", "Zulu", "Mabunda", "Mathebula", "Munyai", "Baloyi", "Ramalema"
  ];

  const ageGroups = ["Under 18", "18 to 25", "26 to 35", "36 to 45", "46 to 60", "60 plus"];
  const genders = ["Male", "Female", "Prefer not to say"];
  const ministries = ["Youth", "Worship", "Ushering", "Media", "Children", "Intercession", "Outreach", "Hospitality", "None"];
  const attendance = ["Weekly", "Twice a month", "Monthly", "Rarely"];
  const contact = ["WhatsApp", "SMS", "Email", "Phone Call"];
  const volunteer = ["Yes", "No", "Maybe"];

  const responseRows = [];
  const memberRows = [];

  for (let i = 1; i <= totalRecords; i++) {
    const fullName = randomItem(firstNames) + " " + randomItem(surnames);

    let phone = "07" + Math.floor(10000000 + Math.random() * 89999999);
    let email = fullName.toLowerCase().replace(" ", ".") + i + "@example.com";

    if (Math.random() < 0.07) {
      phone = "";
    }

    if (Math.random() < 0.12) {
      email = "";
    }

    const ageGroup = weightedPick([
      ["Under 18", 10],
      ["18 to 25", 22],
      ["26 to 35", 24],
      ["36 to 45", 18],
      ["46 to 60", 17],
      ["60 plus", 9]
    ]);

    const gender = weightedPick([
      ["Male", 47],
      ["Female", 51],
      ["Prefer not to say", 2]
    ]);

    const ministry = weightedPick([
      ["Youth", 15],
      ["Worship", 10],
      ["Ushering", 12],
      ["Media", 8],
      ["Children", 11],
      ["Intercession", 10],
      ["Outreach", 9],
      ["Hospitality", 8],
      ["None", 17]
    ]);

    const activeMember = weightedPick([
      ["Yes", 78],
      ["No", 22]
    ]);

    const attendanceFrequency = weightedPick([
      ["Weekly", 48],
      ["Twice a month", 25],
      ["Monthly", 17],
      ["Rarely", 10]
    ]);

    const preferredContact = weightedPick([
      ["WhatsApp", 62],
      ["SMS", 16],
      ["Email", 14],
      ["Phone Call", 8]
    ]);

    const contactInfoCorrect = weightedPick([
      ["Yes", 76],
      ["No", 24]
    ]);

    const wantsVolunteer = weightedPick([
      ["Yes", 32],
      ["No", 42],
      ["Maybe", 26]
    ]);

    const notes = generateNote(activeMember, attendanceFrequency, wantsVolunteer);
    const timestamp = randomDateWithinDays(90);

    const missingInfo = phone === "" || email === "" ? "Yes" : "No";
    const followUpReason = getFollowUpReason(phone, email, activeMember, attendanceFrequency, contactInfoCorrect, wantsVolunteer);
    const followUpRequired = followUpReason === "" ? "No" : "Yes";
    const status = activeMember === "Yes" && missingInfo === "No" ? "Verified Active" : "Needs Review";

    responseRows.push([
      timestamp,
      fullName,
      phone,
      email,
      ageGroup,
      gender,
      ministry,
      activeMember,
      attendanceFrequency,
      preferredContact,
      contactInfoCorrect,
      wantsVolunteer,
      notes
    ]);

    memberRows.push([
      "MEM" + String(i).padStart(4, "0"),
      fullName,
      phone,
      email,
      ageGroup,
      gender,
      ministry,
      activeMember,
      attendanceFrequency,
      preferredContact,
      contactInfoCorrect,
      wantsVolunteer,
      status,
      missingInfo,
      followUpRequired,
      followUpReason,
      timestamp
    ]);
  }

  responses.getRange(2, 1, responseRows.length, responseRows[0].length).setValues(responseRows);
  members.getRange(2, 1, memberRows.length, memberRows[0].length).setValues(memberRows);

  responses.autoResizeColumns(1, 13);
  members.autoResizeColumns(1, 17);
}

function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard");
  sheet.clear();

  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });

  sheet.getRange("A1").setValue("Church Member Verification Dashboard");
  sheet.getRange("A1").setValue("Church Member Verification Dashboard");
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").setFontSize(18).setFontWeight("bold");

  const metrics = [
    ["Metric", "Value"],
    ["Total Members", '=COUNTA(Members!B2:B)'],
    ["Active Members", '=COUNTIF(Members!H2:H,"Yes")'],
    ["Inactive Members", '=COUNTIF(Members!H2:H,"No")'],
    ["Want To Volunteer", '=COUNTIF(Members!L2:L,"Yes")'],
    ["Maybe Volunteer", '=COUNTIF(Members!L2:L,"Maybe")'],
    ["Records With Missing Info", '=COUNTIF(Members!N2:N,"Yes")'],
    ["Members Requiring Follow Up", '=COUNTIF(Members!O2:O,"Yes")'],
    ["Incorrect Contact Info", '=COUNTIF(Members!K2:K,"No")']
  ];

  sheet.getRange(3, 1, metrics.length, 2).setValues(metrics);
  styleHeader(sheet, 2, 3);

  sheet.getRange("D3").setValue("Ministry Statistics");
  sheet.getRange("D4").setFormula('=QUERY(Members!G2:G,"select G, count(G) where G is not null group by G label G \'Ministry\', count(G) \'Members\'",0)');

  sheet.getRange("G3").setValue("Attendance Analysis");
  sheet.getRange("G4").setFormula('=QUERY(Members!I2:I,"select I, count(I) where I is not null group by I label I \'Attendance\', count(I) \'Members\'",0)');

  sheet.getRange("J3").setValue("Preferred Contact Analysis");
  sheet.getRange("J4").setFormula('=QUERY(Members!J2:J,"select J, count(J) where J is not null group by J label J \'Contact Method\', count(J) \'Members\'",0)');

  sheet.getRange("A14").setValue("Follow Up Summary");
  sheet.getRange("A15").setFormula('=QUERY(Members!P2:P,"select P, count(P) where P is not null group by P label P \'Reason\', count(P) \'Members\'",0)');

  SpreadsheetApp.flush();

  const ministryChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("D4:E13"))
    .setPosition(14, 4, 0, 0)
    .setOption("title", "Members by Ministry")
    .setOption("legend", { position: "none" })
    .build();

  const attendanceChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange("G4:H8"))
    .setPosition(14, 8, 0, 0)
    .setOption("title", "Attendance Frequency")
    .build();

  const contactChart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange("J4:K8"))
    .setPosition(32, 4, 0, 0)
    .setOption("title", "Preferred Contact Method")
    .setOption("legend", { position: "none" })
    .build();

  const followUpChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("A15:B25"))
    .setPosition(32, 8, 0, 0)
    .setOption("title", "Follow Up Reasons")
    .setOption("legend", { position: "none" })
    .build();

  sheet.insertChart(ministryChart);
  sheet.insertChart(attendanceChart);
  sheet.insertChart(contactChart);
  sheet.insertChart(followUpChart);

  sheet.autoResizeColumns(1, 12);
}

function buildFollowUpList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const members = ss.getSheetByName("Members");
  const followUp = ss.getSheetByName("Follow Up");

  followUp.getRange(2, 1, followUp.getMaxRows(), followUp.getMaxColumns()).clearContent();

  const data = members.getRange(2, 1, members.getLastRow() - 1, 17).getValues();

  const followUpRows = data
    .filter(function(row) {
      return row[14] === "Yes";
    })
    .map(function(row) {
      return [
        row[0],
        row[1],
        row[2],
        row[3],
        row[6],
        row[7],
        row[8],
        row[9],
        row[15]
      ];
    });

  if (followUpRows.length > 0) {
    followUp.getRange(2, 1, followUpRows.length, followUpRows[0].length).setValues(followUpRows);
  }

  followUp.autoResizeColumns(1, 9);
}

function getFollowUpReason(phone, email, activeMember, attendanceFrequency, contactInfoCorrect, wantsVolunteer) {
  const reasons = [];

  if (phone === "") {
    reasons.push("Missing phone number");
  }

  if (email === "") {
    reasons.push("Missing email address");
  }

  if (activeMember === "No") {
    reasons.push("Inactive member");
  }

  if (attendanceFrequency === "Rarely") {
    reasons.push("Rare attendance");
  }

  if (contactInfoCorrect === "No") {
    reasons.push("Contact details need updating");
  }

  if (wantsVolunteer === "Yes" || wantsVolunteer === "Maybe") {
    reasons.push("Volunteer interest");
  }

  return reasons.join(", ");
}

function generateNote(activeMember, attendanceFrequency, wantsVolunteer) {
  if (activeMember === "No") {
    return "Needs membership status confirmation.";
  }

  if (attendanceFrequency === "Rarely") {
    return "May require pastoral follow up.";
  }

  if (wantsVolunteer === "Yes") {
    return "Interested in serving.";
  }

  if (wantsVolunteer === "Maybe") {
    return "Open to volunteering in future.";
  }

  return "";
}

function weightedPick(options) {
  const total = options.reduce(function(sum, item) {
    return sum + item[1];
  }, 0);

  let random = Math.random() * total;

  for (let i = 0; i < options.length; i++) {
    random -= options[i][1];

    if (random <= 0) {
      return options[i][0];
    }
  }

  return options[0][0];
}

function randomItem(list) {
  return list[Math.floor(Math.random() * list.length)];
}

function randomDateWithinDays(days) {
  const today = new Date();
  const pastDate = new Date(today.getTime() - Math.floor(Math.random() * days) * 24 * 60 * 60 * 1000);
  return pastDate;
}

function styleHeader(sheet, columnCount, rowNumber) {
  const row = rowNumber || 1;
  sheet.getRange(row, 1, 1, columnCount)
    .setFontWeight("bold")
    .setBackground("#d9ead3");

  sheet.setFrozenRows(1);
}

function refreshProject() {
  generateFakeResponses(500);
  buildDashboard();
  buildFollowUpList();

  SpreadsheetApp.getUi().alert("Project refreshed successfully.");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Church System")
    .addItem("Setup Full Project", "setupChurchMemberSystem")
    .addItem("Refresh Fake Data", "refreshProject")
    .addItem("Build Dashboard", "buildDashboard")
    .addItem("Build Follow Up List", "buildFollowUpList")
    .addToUi();
}
