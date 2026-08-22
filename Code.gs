// ===================================================================
// FACTORY MACHINE AUTOMATION - FREE BACKEND (Google Apps Script)
// फैक्ट्री मशीन ऑटोमेशन - मोफत बैकएंड
// ===================================================================
// SETUP STEPS / सेटअप स्टेप्स:
// 1. Google Sheets मध्ये नवीन Sheet बनवा (नाव द्या: Factory Automation)
// 2. वरती Extensions > Apps Script वर क्लिक करा
// 3. इथला सगळा code कॉपी करून तिथे paste करा (जुना code delete करून)
// 4. Save करा (Ctrl+S)
// 5. वरती "Deploy" > "New deployment" वर क्लिक करा
// 6. Type निवडा: "Web app"
// 7. Execute as: "Me"
// 8. Who has access: "Anyone"
// 9. "Deploy" क्लिक करा -> तुम्हाला एक URL मिळेल (Web app URL)
// 10. तो URL कॉपी करून HTML file मध्ये SCRIPT_URL च्या जागी टाका
// ===================================================================

function doGet(e) {
  if (e.parameter.action === 'list') {
    return getEntries();
  }
  return ContentService.createTextOutput('Factory Automation API is running');
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  return saveEntry(data);
}

function saveEntry(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Entries');
  if (!sheet) {
    sheet = ss.insertSheet('Entries');
    sheet.appendRow(['Timestamp', 'Operator Name', 'Machine No', 'Start Time', 'End Time',
      'Remark', 'Morning Reading Photo', 'Evening Reading Photo', 'Challan Photo', 'Latitude', 'Longitude']);
  }

  var folder = getOrCreateFolder();
  var morningUrl = data.readingMorning ? saveImage(data.readingMorning, folder, 'morning_' + Date.now()) : '';
  var eveningUrl = data.readingEvening ? saveImage(data.readingEvening, folder, 'evening_' + Date.now()) : '';
  var challanUrl = data.challanPhoto ? saveImage(data.challanPhoto, folder, 'challan_' + Date.now()) : '';

  sheet.appendRow([
    new Date(), data.opName, data.machNo, data.startTime, data.endTime, data.remark,
    morningUrl, eveningUrl, challanUrl,
    data.location ? data.location.lat : '', data.location ? data.location.lng : ''
  ]);

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveImage(base64Data, folder, name) {
  var parts = base64Data.split(',');
  var bytes = Utilities.base64Decode(parts[1]);
  var blob = Utilities.newBlob(bytes, 'image/jpeg', name + '.jpg');
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateFolder() {
  var folders = DriveApp.getFoldersByName('FactoryAppPhotos');
  return folders.hasNext() ? folders.next() : DriveApp.createFolder('FactoryAppPhotos');
}

function getEntries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Entries');
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[1]) continue;
    out.push({
      timestamp: new Date(r[0]).getTime(),
      dateStr: new Date(r[0]).toLocaleString(),
      opName: r[1], machNo: r[2], startTime: r[3], endTime: r[4], remark: r[5],
      readingMorning: r[6], readingEvening: r[7], challanPhoto: r[8],
      location: (r[9] && r[10]) ? { lat: r[9], lng: r[10] } : null
    });
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
